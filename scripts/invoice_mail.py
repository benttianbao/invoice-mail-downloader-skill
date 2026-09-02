#!/usr/bin/env python3
"""从 163/QQ 邮箱只读获取并归档电子发票。"""

from __future__ import annotations

import argparse
import base64
import contextlib
import datetime as dt
from email import policy
from email.header import decode_header, make_header
from email.message import Message
from email.parser import BytesParser
from email.utils import parseaddr
import getpass
import hashlib
from html.parser import HTMLParser
import imaplib
import ipaddress
import json
import os
from pathlib import Path, PurePosixPath
import re
import shutil
import socket
import ssl
import sys
import tempfile
import time
from typing import Any, Iterable
import urllib.error
import urllib.parse
import urllib.request
import zipfile


APP_NAME = "invoice-mail-downloader"
KEYRING_SERVICE = APP_NAME
MAX_FILE_BYTES = 50 * 1024 * 1024
MAX_ZIP_BYTES = 200 * 1024 * 1024
MAX_ZIP_FILES = 100
MAX_COMPRESSION_RATIO = 200
KEYWORDS = ("发票", "电子票", "开票", "票据", "invoice", "e-invoice", "einvoice")
SUPPORTED_SUFFIXES = {".pdf", ".ofd", ".zip"}
PROVIDERS = {
    "163": {"domain": "163.com", "imap": "imap.163.com", "port": 993},
    "qq": {"domain": "qq.com", "imap": "imap.qq.com", "port": 993},
}
EXCLUDED_FLAGS = {b"\\Sent", b"\\Drafts", b"\\Trash", b"\\Junk"}
EXCLUDED_NAMES = ("sent", "draft", "trash", "junk", "spam", "已发送", "草稿", "垃圾")


class SkillError(RuntimeError):
    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code


def error_code(exc: BaseException, fallback: str) -> str:
    return exc.code if isinstance(exc, SkillError) else fallback


def app_data_dir() -> Path:
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / APP_NAME
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
        return base / APP_NAME
    raise RuntimeError("首版仅支持 macOS 和 Windows。")


def default_invoice_root() -> Path:
    return Path.home() / "Documents" / "Invoices"


def atomic_json_write(path: Path, value: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix=path.name + ".", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(value, handle, ensure_ascii=False, indent=2, sort_keys=True)
            handle.flush()
            os.fsync(handle.fileno())
        if os.name != "nt":
            os.chmod(temp_name, 0o600)
        os.replace(temp_name, path)
    finally:
        with contextlib.suppress(FileNotFoundError):
            os.unlink(temp_name)


def load_json(path: Path, default: dict[str, Any]) -> dict[str, Any]:
    if not path.exists():
        return default
    with path.open("r", encoding="utf-8") as handle:
        value = json.load(handle)
    if not isinstance(value, dict):
        raise ValueError(f"无效 JSON 对象：{path}")
    return value


def config_path() -> Path:
    return app_data_dir() / "config.json"


def state_path() -> Path:
    return app_data_dir() / "state.json"


def load_config() -> dict[str, Any]:
    defaults = {
        "version": 2,
        "invoice_root": str(default_invoice_root()),
        "initial_lookback_days": 30,
        "first_run_confirmed": False,
        "trusted_domains": [],
        "accounts": [],
    }
    config = load_json(config_path(), defaults.copy())
    for key, value in defaults.items():
        config.setdefault(key, value)
    config["version"] = 2
    return config


def load_state() -> dict[str, Any]:
    return load_json(state_path(), {"version": 2, "folders": {}, "files": {}})


@contextlib.contextmanager
def exclusive_run_lock() -> Iterable[None]:
    lock = app_data_dir() / "run.lock"
    lock.parent.mkdir(parents=True, exist_ok=True)
    try:
        fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
    except FileExistsError as exc:
        raise SkillError("RUN_LOCKED", f"已有扫描正在运行；确认前次异常退出后才能删除精确锁文件：{lock}") from exc
    try:
        os.write(fd, str(os.getpid()).encode("ascii"))
        os.close(fd)
        yield
    finally:
        with contextlib.suppress(FileNotFoundError):
            lock.unlink()


def credential_set(email: str, secret: str) -> None:
    import keyring

    keyring.set_password(KEYRING_SERVICE, email.lower(), secret)


def credential_get(email: str) -> str | None:
    import keyring

    return keyring.get_password(KEYRING_SERVICE, email.lower())


def credential_delete(email: str) -> None:
    import keyring
    from keyring.errors import PasswordDeleteError

    try:
        keyring.delete_password(KEYRING_SERVICE, email.lower())
    except PasswordDeleteError:
        pass


def decode_mime(value: str | None) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except (LookupError, UnicodeError):
        return value


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def has_invoice_keyword(value: str) -> bool:
    lowered = value.lower()
    return any(word in lowered for word in KEYWORDS)


class LinkParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.links: list[tuple[str, str]] = []
        self._href: str | None = None
        self._text: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() == "a":
            self._href = dict(attrs).get("href")
            self._text = []

    def handle_data(self, data: str) -> None:
        if self._href is not None:
            self._text.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() == "a" and self._href is not None:
            self.links.append((self._href, normalize_text("".join(self._text))))
            self._href = None
            self._text = []


def extract_candidate_links(plain: str, html: str, message_is_candidate: bool = True) -> list[str]:
    candidates: list[str] = []
    parser = LinkParser()
    with contextlib.suppress(Exception):
        parser.feed(html)
    for url, label in parser.links:
        suffix = Path(urllib.parse.urlsplit(url).path).suffix.lower()
        if has_invoice_keyword(label + " " + url) or (message_is_candidate and suffix in SUPPORTED_SUFFIXES):
            candidates.append(url)

    url_re = re.compile(r"https?://[^\s<>\"']+", re.IGNORECASE)
    for match in url_re.finditer(plain):
        url = match.group(0).rstrip(".,;:!?)]}，。；：！？）】")
        window = plain[max(0, match.start() - 100) : min(len(plain), match.end() + 100)]
        suffix = Path(urllib.parse.urlsplit(url).path).suffix.lower()
        if has_invoice_keyword(window) or (message_is_candidate and suffix in SUPPORTED_SUFFIXES):
            candidates.append(url)
    return list(dict.fromkeys(candidates))


def message_content(message: Message) -> tuple[str, str, list[tuple[str, bytes]]]:
    plain_parts: list[str] = []
    html_parts: list[str] = []
    attachments: list[tuple[str, bytes]] = []
    for part in message.walk():
        if part.is_multipart():
            continue
        content_type = part.get_content_type().lower()
        disposition = part.get_content_disposition()
        filename = decode_mime(part.get_filename())
        payload = part.get_payload(decode=True) or b""
        if disposition == "attachment" or filename:
            attachments.append((filename or "unnamed", payload))
            continue
        charset = part.get_content_charset() or "utf-8"
        text = payload.decode(charset, errors="replace")
        if content_type == "text/plain":
            plain_parts.append(text)
        elif content_type == "text/html":
            html_parts.append(text)
    return "\n".join(plain_parts), "\n".join(html_parts), attachments


def redact_url(url: str) -> str:
    try:
        parts = urllib.parse.urlsplit(url)
        return urllib.parse.urlunsplit((parts.scheme, parts.hostname or "", parts.path, "", ""))
    except ValueError:
        return "<无效链接>"


def domain_is_trusted(hostname: str, trusted_domains: Iterable[str]) -> bool:
    hostname = hostname.rstrip(".").lower()
    return any(hostname == domain for domain in trusted_domains)


def validate_public_https(url: str, trusted_domains: Iterable[str]) -> None:
    parts = urllib.parse.urlsplit(url)
    if parts.scheme.lower() != "https":
        raise SkillError("DOWNLOAD_INSECURE_URL", "仅允许公网 HTTPS 下载链接")
    if parts.username or parts.password or not parts.hostname:
        raise SkillError("DOWNLOAD_INVALID_URL", "链接包含凭据或缺少主机名")
    if not domain_is_trusted(parts.hostname, trusted_domains):
        raise SkillError("UNTRUSTED_DOMAIN", f"域名未加入可信列表：{parts.hostname.lower()}")
    try:
        addresses = {item[4][0] for item in socket.getaddrinfo(parts.hostname, parts.port or 443, type=socket.SOCK_STREAM)}
    except socket.gaierror as exc:
        raise SkillError("DOWNLOAD_DNS_FAILED", "下载域名无法解析") from exc
    for address in addresses:
        ip = ipaddress.ip_address(address)
        if not ip.is_global:
            raise SkillError("DOWNLOAD_PRIVATE_ADDRESS", "拒绝访问内网、回环或保留地址")


class SafeRedirectHandler(urllib.request.HTTPRedirectHandler):
    def __init__(self, trusted_domains: Iterable[str]) -> None:
        super().__init__()
        self.trusted_domains = tuple(trusted_domains)

    def redirect_request(self, req: Any, fp: Any, code: int, msg: str, headers: Any, newurl: str) -> Any:
        validate_public_https(newurl, self.trusted_domains)
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def download_http(url: str, destination: Path, trusted_domains: Iterable[str]) -> tuple[str, str]:
    trusted = tuple(trusted_domains)
    validate_public_https(url, trusted)
    opener = urllib.request.build_opener(
        urllib.request.HTTPSHandler(context=ssl.create_default_context()),
        SafeRedirectHandler(trusted),
    )
    request = urllib.request.Request(url, headers={"User-Agent": "invoice-mail-downloader/1.0"})
    last_error: Exception | None = None
    for attempt in range(3):
        try:
            with opener.open(request, timeout=30) as response:
                content_type = response.headers.get_content_type()
                filename = response.headers.get_filename() or Path(urllib.parse.urlsplit(response.url).path).name
                total = 0
                with destination.open("wb") as handle:
                    while chunk := response.read(64 * 1024):
                        total += len(chunk)
                        if total > MAX_FILE_BYTES:
                            raise ValueError("文件超过 50 MB 限制")
                        handle.write(chunk)
                return content_type, filename
        except urllib.error.HTTPError as exc:
            last_error = exc
            if not 500 <= exc.code <= 599:
                raise SkillError("DOWNLOAD_HTTP_ERROR", f"HTTP {exc.code}") from exc
        except (urllib.error.URLError, TimeoutError, ConnectionError) as exc:
            last_error = exc
        if attempt < 2:
            time.sleep(2**attempt)
    raise SkillError("DOWNLOAD_NETWORK_FAILED", f"网络下载重试 3 次后失败：{type(last_error).__name__}")


def detect_file_type(path: Path, hinted_name: str = "") -> str:
    with path.open("rb") as handle:
        head = handle.read(8)
    if head.startswith(b"%PDF-"):
        return ".pdf"
    if head.startswith(b"PK\x03\x04"):
        try:
            with zipfile.ZipFile(path) as archive:
                names = {name.replace("\\", "/").lstrip("/").lower() for name in archive.namelist()}
            if "ofd.xml" in names:
                return ".ofd"
        except zipfile.BadZipFile:
            return ""
        return ".zip"
    suffix = Path(hinted_name).suffix.lower()
    return suffix if suffix in SUPPORTED_SUFFIXES else ""


def safe_zip_members(path: Path) -> list[zipfile.ZipInfo]:
    selected: list[zipfile.ZipInfo] = []
    total = 0
    with zipfile.ZipFile(path) as archive:
        for info in archive.infolist():
            member = PurePosixPath(info.filename.replace("\\", "/"))
            if info.flag_bits & 0x1:
                raise ValueError("拒绝加密 ZIP")
            if member.is_absolute() or ".." in member.parts:
                raise ValueError("拒绝包含路径穿越的 ZIP")
            if info.is_dir():
                continue
            suffix = member.suffix.lower()
            if suffix == ".zip":
                continue
            if suffix not in {".pdf", ".ofd"}:
                continue
            ratio = info.file_size / max(info.compress_size, 1)
            if ratio > MAX_COMPRESSION_RATIO:
                raise ValueError("拒绝异常压缩比 ZIP")
            total += info.file_size
            selected.append(info)
            if len(selected) > MAX_ZIP_FILES or total > MAX_ZIP_BYTES:
                raise ValueError("ZIP 超过文件数量或解压大小限制")
    return selected


def unpack_zip(path: Path, target: Path) -> list[Path]:
    output: list[Path] = []
    with zipfile.ZipFile(path) as archive:
        for index, info in enumerate(safe_zip_members(path), start=1):
            suffix = PurePosixPath(info.filename).suffix.lower()
            destination = target / f"item-{index}{suffix}"
            with archive.open(info) as source, destination.open("wb") as sink:
                shutil.copyfileobj(source, sink, 64 * 1024)
            actual = detect_file_type(destination, info.filename)
            if actual in {".pdf", ".ofd"}:
                if actual != suffix:
                    corrected = destination.with_suffix(actual)
                    destination.rename(corrected)
                    destination = corrected
                output.append(destination)
            else:
                destination.unlink(missing_ok=True)
    return output


def pdf_text(path: Path) -> str:
    from pypdf import PdfReader

    reader = PdfReader(str(path), strict=False)
    if reader.is_encrypted:
        with contextlib.suppress(Exception):
            reader.decrypt("")
    return "\n".join(page.extract_text() or "" for page in reader.pages[:5])


def ofd_text(path: Path) -> str:
    texts: list[str] = []
    with zipfile.ZipFile(path) as archive:
        for info in archive.infolist():
            if not info.filename.lower().endswith(".xml") or info.file_size > 5 * 1024 * 1024:
                continue
            raw = archive.read(info)
            decoded = raw.decode("utf-8", errors="replace")
            texts.extend(re.findall(r"<[^>]*TextCode[^>]*>(.*?)</[^>]*TextCode>", decoded, re.I | re.S))
            texts.extend(re.findall(r"\bValue=[\"']([^\"']+)[\"']", decoded, re.I))
    return normalize_text(" ".join(re.sub(r"<[^>]+>", "", item) for item in texts))


def extract_invoice_fields(text: str) -> dict[str, str]:
    flat = normalize_text(text)
    date_value = ""
    seller = ""
    amount = ""

    date_patterns = (
        r"开票日期\s*[：:]?\s*(20\d{2})[年./-](\d{1,2})[月./-](\d{1,2})日?",
        r"\b(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})\b",
    )
    for pattern in date_patterns:
        if match := re.search(pattern, flat):
            try:
                date_value = dt.date(*map(int, match.groups())).isoformat()
                break
            except ValueError:
                continue

    seller_patterns = (
        r"销售方(?:信息)?\s*(?:名称)?\s*[：:]?\s*([^：:]{2,80}?)(?=\s*(?:统一社会信用代码|纳税人识别号|地址|开户行|购买方|发票))",
        r"销方名称\s*[：:]?\s*([^：:]{2,80}?)(?=\s*(?:统一社会信用代码|纳税人识别号|地址|开户行|$))",
    )
    for pattern in seller_patterns:
        if match := re.search(pattern, flat):
            seller = normalize_text(match.group(1))
            break

    amount_patterns = (
        r"价税合计.{0,40}?(?:小写)?\s*[（(]?(?:小写)?[)）]?\s*[：:]?\s*[¥￥]?\s*([0-9][0-9,]*\.\d{2})",
        r"[（(]小写[)）]\s*[¥￥]?\s*([0-9][0-9,]*\.\d{2})",
    )
    for pattern in amount_patterns:
        if match := re.search(pattern, flat, re.I):
            amount = match.group(1).replace(",", "")
            break
    return {"date": date_value, "seller": seller, "amount": amount}


def safe_component(value: str, fallback: str, limit: int = 60) -> str:
    value = normalize_text(value)
    value = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", "_", value).strip(" ._")
    return (value or fallback)[:limit]


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while chunk := handle.read(1024 * 1024):
            digest.update(chunk)
    return digest.hexdigest()


def archive_invoice(source: Path, suffix: str, root: Path, state: dict[str, Any]) -> tuple[str, str, list[str]]:
    try:
        text = pdf_text(source) if suffix == ".pdf" else ofd_text(source)
        fields = extract_invoice_fields(text)
    except Exception:
        fields = {"date": "", "seller": "", "amount": ""}
    missing = [key for key in ("date", "seller", "amount") if not fields[key]]
    date_label = fields["date"] or "未知日期"
    seller_label = safe_component(fields["seller"], "未知销售方")
    amount_label = fields["amount"] or "未知金额"
    filename = f"{date_label}_{seller_label}_¥{amount_label}{suffix}"

    digest = sha256_file(source)
    existing = state.setdefault("files", {}).get(digest)
    if existing and Path(existing).exists():
        return "skipped", existing, missing

    if missing:
        destination_dir = root / "待确认"
    else:
        parsed_date = dt.date.fromisoformat(fields["date"])
        destination_dir = root / f"{parsed_date.year:04d}" / f"{parsed_date.month:02d}-{parsed_date.day:02d}"
    destination_dir.mkdir(parents=True, exist_ok=True)
    destination = destination_dir / filename
    if destination.exists():
        if sha256_file(destination) == digest:
            state["files"][digest] = str(destination)
            return "skipped", str(destination), missing
        destination = destination.with_name(f"{destination.stem}_{digest[:8]}{suffix}")
    shutil.copy2(source, destination)
    state["files"][digest] = str(destination)
    return "success", str(destination), missing


def decode_modified_utf7(value: str) -> str:
    """解码 IMAP Modified UTF-7，同时保留无法解析的原始名称。"""
    output: list[str] = []
    index = 0
    while index < len(value):
        if value[index] != "&":
            output.append(value[index])
            index += 1
            continue
        end = value.find("-", index)
        if end < 0:
            return value
        encoded = value[index + 1 : end]
        if not encoded:
            output.append("&")
        else:
            padded = encoded.replace(",", "/") + "=" * ((4 - len(encoded) % 4) % 4)
            try:
                output.append(base64.b64decode(padded).decode("utf-16-be"))
            except (ValueError, UnicodeError):
                return value
        index = end + 1
    return "".join(output)


def folder_entries(client: imaplib.IMAP4_SSL) -> list[tuple[str, str]]:
    status, lines = client.list()
    if status != "OK":
        raise RuntimeError("无法读取邮箱文件夹列表")
    result: list[tuple[str, str]] = []
    pattern = re.compile(rb"^\((.*?)\)\s+(?:\"[^\"]*\"|NIL)\s+(.+)$")
    for line in lines or []:
        if not isinstance(line, bytes) or not (match := pattern.match(line)):
            continue
        flags = set(match.group(1).split())
        raw_name = match.group(2).strip()
        if raw_name.startswith(b'"') and raw_name.endswith(b'"'):
            raw_name = raw_name[1:-1].replace(b'\\"', b'"').replace(b"\\\\", b"\\")
        wire_name = raw_name.decode("ascii", errors="replace")
        display_name = decode_modified_utf7(wire_name)
        lowered = display_name.lower()
        if flags & EXCLUDED_FLAGS or any(token in lowered for token in EXCLUDED_NAMES):
            continue
        result.append((wire_name, display_name))
    if "INBOX" not in {wire.upper() for wire, _display in result}:
        result.insert(0, ("INBOX", "INBOX"))
    return list(dict.fromkeys(result))


def parse_uid_search(status: str, values: list[bytes | None]) -> set[int]:
    if status != "OK":
        raise RuntimeError("邮件搜索失败")
    return {int(value) for value in (values[0] or b"").split() if value.isdigit()}


def search_date_uids(client: imaplib.IMAP4_SSL, from_date: dt.date, to_date: dt.date | None) -> list[int]:
    criteria: list[str] = ["SINCE", from_date.strftime("%d-%b-%Y")]
    if to_date:
        criteria.extend(["BEFORE", (to_date + dt.timedelta(days=1)).strftime("%d-%b-%Y")])
    status, values = client.uid("search", None, *criteria)
    return sorted(parse_uid_search(status, values))


def search_new_uids(client: imaplib.IMAP4_SSL, last_uid: int) -> list[int]:
    start = max(1, last_uid + 1)
    status, values = client.uid("search", None, "UID", f"{start}:*")
    return sorted(value for value in parse_uid_search(status, values) if value > last_uid)


def response_number(client: imaplib.IMAP4_SSL, name: str) -> int | None:
    _status, values = client.response(name)
    for value in values or []:
        if isinstance(value, bytes) and value.isdigit():
            return int(value)
    return None


def current_last_uid(client: imaplib.IMAP4_SSL) -> int:
    uid_next = response_number(client, "UIDNEXT")
    if uid_next is not None:
        return max(0, uid_next - 1)
    status, values = client.uid("search", None, "ALL")
    found = parse_uid_search(status, values)
    return max(found, default=0)


def send_imap_identity(client: imaplib.IMAP4_SSL, provider_name: str) -> None:
    if provider_name != "163":
        return
    imaplib.Commands["ID"] = ("AUTH",)
    try:
        status, _values = client._simple_command(  # noqa: SLF001 - imaplib 3.10-3.14 没有公开 ID 方法
            "ID",
            '("name" "invoice-mail-downloader" "version" "1.0" "vendor" "local-agent-skill")',
        )
    except imaplib.IMAP4.error as exc:
        raise SkillError("IMAP_ID_FAILED", "163 邮箱 IMAP ID 客户端身份声明失败") from exc
    if status != "OK":
        raise SkillError("IMAP_ID_FAILED", "163 邮箱拒绝 IMAP ID 客户端身份声明")


def report_item(
    status: str,
    source: str,
    subject: str,
    sender: str,
    detail: str,
    code: str = "",
) -> dict[str, str]:
    return {"status": status, "code": code, "source": source, "subject": subject, "sender": sender, "detail": detail}


def process_file(
    path: Path,
    hinted_name: str,
    root: Path,
    state: dict[str, Any],
    temp_dir: Path,
) -> list[tuple[str, str, str]]:
    if path.stat().st_size > MAX_FILE_BYTES:
        return [("incomplete", hinted_name, "文件超过 50 MB 限制")]
    suffix = detect_file_type(path, hinted_name)
    if not suffix:
        return [("incomplete", hinted_name, "内容不是受支持的 PDF、OFD 或 ZIP")]
    if suffix == ".zip":
        try:
            extracted = unpack_zip(path, temp_dir)
        except (ValueError, zipfile.BadZipFile) as exc:
            return [("incomplete", hinted_name, str(exc))]
        if not extracted:
            return [("incomplete", hinted_name, "ZIP 中没有有效 PDF/OFD")]
        results: list[tuple[str, str, str]] = []
        for extracted_file in extracted:
            results.extend(process_file(extracted_file, extracted_file.name, root, state, temp_dir))
        return results
    status, destination, missing = archive_invoice(path, suffix, root, state)
    detail = destination
    if missing:
        detail += "；待确认字段：" + "、".join(missing)
    return [(status, hinted_name, detail)]


def process_message(
    raw: bytes,
    root: Path,
    state: dict[str, Any],
    temp_dir: Path,
    trusted_domains: Iterable[str] = (),
) -> tuple[list[dict[str, str]], bool]:
    message = BytesParser(policy=policy.default).parsebytes(raw)
    subject = decode_mime(message.get("Subject"))
    sender = parseaddr(decode_mime(message.get("From")))[1]
    plain, html, attachments = message_content(message)
    candidate = has_invoice_keyword(" ".join((subject, plain, re.sub(r"<[^>]+>", " ", html))))
    candidate = candidate or any(has_invoice_keyword(name) for name, _ in attachments)
    if not candidate:
        return [], False

    items: list[dict[str, str]] = []
    with tempfile.TemporaryDirectory(dir=temp_dir) as message_tmp_name:
        message_tmp = Path(message_tmp_name)
        for index, (filename, payload) in enumerate(attachments, start=1):
            suffix = Path(filename).suffix.lower()
            if suffix not in SUPPORTED_SUFFIXES:
                continue
            if len(payload) > MAX_FILE_BYTES:
                items.append(report_item("incomplete", f"附件:{filename}", subject, sender, "文件超过脚本限制", "ATTACHMENT_TOO_LARGE"))
                continue
            attachment_path = message_tmp / f"attachment-{index}{suffix}"
            attachment_path.write_bytes(payload)
            for status, source, detail in process_file(attachment_path, filename, root, state, message_tmp):
                items.append(report_item(status, f"附件:{source}", subject, sender, detail))

        links = extract_candidate_links(plain, html, candidate)
        for index, url in enumerate(links, start=1):
            target = message_tmp / f"download-{index}"
            redacted = redact_url(url)
            try:
                content_type, filename = download_http(url, target, trusted_domains)
                if content_type in {"text/html", "application/xhtml+xml"} or target.read_bytes()[:256].lstrip().lower().startswith(b"<"):
                    page_html = target.read_text(encoding="utf-8", errors="replace")
                    page_links = [
                        urllib.parse.urljoin(url, child)
                        for child in extract_candidate_links("", page_html, True)
                        if urllib.parse.urljoin(url, child) != url
                    ]
                    target.unlink(missing_ok=True)
                    if page_links:
                        for child_index, child_url in enumerate(dict.fromkeys(page_links), start=1):
                            child_target = message_tmp / f"download-{index}-{child_index}"
                            child_source = redact_url(child_url)
                            try:
                                child_type, child_name = download_http(child_url, child_target, trusted_domains)
                                if child_type in {"text/html", "application/xhtml+xml"}:
                                    raise SkillError("DOWNLOAD_DYNAMIC_PAGE", "下载链接仍返回网页；脚本不会执行 JavaScript 或点击按钮")
                                for status, _source, detail in process_file(child_target, child_name, root, state, message_tmp):
                                    items.append(report_item(status, child_source, subject, sender, detail))
                            except Exception as exc:
                                items.append(report_item("incomplete", child_source, subject, sender, str(exc), error_code(exc, "DOWNLOAD_FAILED")))
                        continue
                    raise SkillError("DOWNLOAD_DYNAMIC_PAGE", "页面没有静态发票文件链接；脚本不会执行 JavaScript 或点击按钮")
                for status, _source, detail in process_file(target, filename, root, state, message_tmp):
                    items.append(report_item(status, redacted, subject, sender, detail))
            except Exception as exc:
                items.append(report_item("incomplete", redacted, subject, sender, str(exc), error_code(exc, "DOWNLOAD_FAILED")))
    if not items:
        items.append(
            report_item(
                "incomplete",
                "邮件",
                subject,
                sender,
                "邮件疑似与发票有关，但未发现可处理的 PDF、OFD、ZIP 或下载链接",
                "NO_DOWNLOAD_CANDIDATE",
            )
        )
    incomplete = any(item["status"] == "incomplete" for item in items)
    return items, incomplete


def fetch_message(client: imaplib.IMAP4_SSL, uid: int) -> bytes:
    status, values = client.uid("fetch", str(uid), "(BODY.PEEK[])")
    if status != "OK":
        raise RuntimeError("邮件读取失败")
    for value in values or []:
        if isinstance(value, tuple) and isinstance(value[1], bytes):
            return value[1]
    raise RuntimeError("邮件内容为空")


def run_account_once(
    account: dict[str, Any],
    secret: str,
    root: Path,
    state: dict[str, Any],
    from_date: dt.date,
    to_date: dt.date | None,
    temp_dir: Path,
    explicit_range: bool,
    trusted_domains: Iterable[str],
) -> list[dict[str, str]]:
    email = account["email"]
    provider_name = account["provider"]
    provider = PROVIDERS[provider_name]
    output: list[dict[str, str]] = []
    context = ssl.create_default_context()
    with imaplib.IMAP4_SSL(provider["imap"], provider["port"], ssl_context=context, timeout=30) as client:
        client.login(email, secret)
        send_imap_identity(client, provider_name)
        for wire_name, display_name in folder_entries(client):
            status, select_response = client.select(wire_name, readonly=True)
            if status != "OK":
                response_text = " ".join(
                    value.decode("utf-8", errors="replace") for value in select_response or [] if isinstance(value, bytes)
                )
                code = "IMAP_ID_FAILED" if provider_name == "163" and "unsafe login" in response_text.lower() else "FOLDER_SELECT_FAILED"
                output.append(
                    report_item(
                        "error",
                        f"文件夹:{display_name}",
                        "",
                        email,
                        "无法只读打开文件夹" + (f"：{response_text}" if response_text else ""),
                        code,
                    )
                )
                continue

            key = f"{email.lower()}::{wire_name}"
            folder_state = state.setdefault("folders", {}).setdefault(
                key,
                {"uidvalidity": None, "initialized": False, "last_uid": 0, "pending_uids": []},
            )
            uidvalidity = response_number(client, "UIDVALIDITY")
            if uidvalidity is None:
                output.append(
                    report_item("error", f"文件夹:{display_name}", "", email, "服务器未返回 UIDVALIDITY", "UIDVALIDITY_MISSING")
                )
                continue
            if folder_state.get("uidvalidity") != uidvalidity:
                folder_state.clear()
                folder_state.update({"uidvalidity": uidvalidity, "initialized": False, "last_uid": 0, "pending_uids": []})

            previous_pending = {int(value) for value in folder_state.get("pending_uids", [])}
            last_uid = int(folder_state.get("last_uid", 0))
            try:
                if explicit_range:
                    uids = search_date_uids(client, from_date, to_date)
                    pending = set(previous_pending)
                    checkpoint = last_uid
                elif not folder_state.get("initialized", False):
                    checkpoint = current_last_uid(client)
                    uids = [uid for uid in search_date_uids(client, from_date, None) if uid <= checkpoint]
                    pending = set()
                else:
                    uids = sorted(set(search_new_uids(client, last_uid)) | previous_pending)
                    pending = set()
                    checkpoint = max([last_uid, *uids])
            except Exception as exc:
                output.append(
                    report_item(
                        "error",
                        f"文件夹:{display_name}",
                        "",
                        email,
                        str(exc),
                        error_code(exc, "FOLDER_SCAN_FAILED"),
                    )
                )
                continue

            for uid in uids:
                try:
                    raw = fetch_message(client, uid)
                    items, incomplete = process_message(raw, root, state, temp_dir, trusted_domains)
                    output.extend(items)
                    if incomplete:
                        pending.add(uid)
                    else:
                        pending.discard(uid)
                except Exception as exc:
                    pending.add(uid)
                    output.append(
                        report_item(
                            "error",
                            f"邮件UID:{uid}",
                            "",
                            email,
                            str(exc),
                            error_code(exc, "MESSAGE_PROCESS_FAILED"),
                        )
                    )

            folder_state["pending_uids"] = sorted(pending)
            if not explicit_range:
                folder_state["last_uid"] = checkpoint
                folder_state["initialized"] = True
    return output


def run_account(
    account: dict[str, Any],
    root: Path,
    state: dict[str, Any],
    from_date: dt.date,
    to_date: dt.date | None,
    temp_dir: Path,
    explicit_range: bool,
    trusted_domains: Iterable[str],
) -> list[dict[str, str]]:
    email = account["email"]
    try:
        secret = credential_get(email)
    except Exception as exc:
        return [report_item("error", "邮箱", "", email, f"系统凭据库读取失败：{type(exc).__name__}", "KEYRING_ERROR")]
    if not secret:
        return [report_item("error", "邮箱", "", email, "系统凭据库中没有授权码", "CREDENTIAL_MISSING")]

    last_error: BaseException | None = None
    for attempt in range(3):
        try:
            return run_account_once(
                account,
                secret,
                root,
                state,
                from_date,
                to_date,
                temp_dir,
                explicit_range,
                trusted_domains,
            )
        except SkillError as exc:
            return [report_item("error", "邮箱", "", email, str(exc), exc.code)]
        except imaplib.IMAP4.abort as exc:
            last_error = exc
        except imaplib.IMAP4.error:
            return [report_item("error", "邮箱", "", email, "IMAP 登录失败，请检查服务开关和客户端授权码", "IMAP_AUTH_FAILED")]
        except (OSError, ssl.SSLError) as exc:
            last_error = exc
        if attempt < 2:
            time.sleep(2**attempt)
    return [
        report_item(
            "error",
            "邮箱",
            "",
            email,
            f"IMAP 连接重试 3 次后失败：{type(last_error).__name__}",
            "IMAP_NETWORK_FAILED",
        )
    ]


def validate_account(provider: str, email: str) -> str:
    email = email.strip().lower()
    if provider not in PROVIDERS:
        raise ValueError("未知邮箱服务商")
    if not email.endswith("@" + PROVIDERS[provider]["domain"]):
        raise ValueError(f"{provider} 服务商要求 @{PROVIDERS[provider]['domain']} 邮箱地址")
    return email


def cmd_configure(args: argparse.Namespace) -> int:
    email = validate_account(args.provider, args.email)
    if not sys.stdin.isatty():
        raise SkillError(
            "INTERACTIVE_TERMINAL_REQUIRED",
            "configure 必须由用户本人在本机交互式终端执行；禁止通过 Agent 工具或管道传入授权码",
        )
    secret = getpass.getpass(f"请输入 {email} 的客户端授权码（输入不会显示）：").strip()
    if not secret:
        raise SystemExit("授权码不能为空。")
    try:
        credential_set(email, secret)
    except Exception as exc:
        raise SkillError("KEYRING_ERROR", f"系统凭据库写入失败：{type(exc).__name__}") from exc
    config = load_config()
    accounts = config.setdefault("accounts", [])
    existing = next((item for item in accounts if item.get("email", "").lower() == email), None)
    if existing:
        existing.update({"provider": args.provider, "enabled": True})
    else:
        accounts.append({"email": email, "provider": args.provider, "enabled": True})
    atomic_json_write(config_path(), config)
    print(f"已安全配置：{email}")
    return 0


def cmd_accounts(_args: argparse.Namespace) -> int:
    config = load_config()
    print(
        json.dumps(
            {
                "invoice_root": config["invoice_root"],
                "first_run_confirmed": config["first_run_confirmed"],
                "trusted_domains": config["trusted_domains"],
                "accounts": config["accounts"],
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


def find_account(config: dict[str, Any], email: str) -> dict[str, Any]:
    email = email.lower()
    account = next((item for item in config.get("accounts", []) if item.get("email", "").lower() == email), None)
    if not account:
        raise ValueError(f"未配置邮箱：{email}")
    return account


def cmd_toggle(args: argparse.Namespace, enabled: bool) -> int:
    config = load_config()
    find_account(config, args.email)["enabled"] = enabled
    atomic_json_write(config_path(), config)
    print(f"已{'启用' if enabled else '停用'}：{args.email.lower()}")
    return 0


def cmd_remove(args: argparse.Namespace) -> int:
    config = load_config()
    email = args.email.lower()
    before = len(config.get("accounts", []))
    config["accounts"] = [item for item in config.get("accounts", []) if item.get("email", "").lower() != email]
    if len(config["accounts"]) == before:
        raise ValueError(f"未配置邮箱：{email}")
    credential_delete(email)
    atomic_json_write(config_path(), config)
    print(f"已删除账号配置和凭据：{email}；已下载发票未删除。")
    return 0


def cmd_set_root(args: argparse.Namespace) -> int:
    if not args.confirm:
        raise SkillError("CONFIRMATION_REQUIRED", "修改根目录需要用户确认后添加 --confirm")
    root = Path(args.path).expanduser().resolve()
    root.mkdir(parents=True, exist_ok=True)
    config = load_config()
    config["invoice_root"] = str(root)
    atomic_json_write(config_path(), config)
    print(f"默认发票目录已设置为：{root}")
    return 0


def normalize_domain(value: str) -> str:
    value = value.strip().rstrip(".").lower()
    if not value or "://" in value or any(char in value for char in "/\\:*?@"):
        raise ValueError("请输入不带协议、端口、路径或通配符的域名")
    try:
        ascii_value = value.encode("idna").decode("ascii")
    except UnicodeError as exc:
        raise ValueError("域名格式无效") from exc
    labels = ascii_value.split(".")
    if len(labels) < 2 or any(not re.fullmatch(r"[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?", label) for label in labels):
        raise ValueError("域名格式无效")
    return ascii_value


def cmd_trusted_domains(_args: argparse.Namespace) -> int:
    print(json.dumps({"trusted_domains": load_config()["trusted_domains"]}, ensure_ascii=False, indent=2))
    return 0


def cmd_trust_domain(args: argparse.Namespace, trusted: bool) -> int:
    if not args.confirm:
        raise SkillError("CONFIRMATION_REQUIRED", "修改可信域名需要用户确认后添加 --confirm")
    domain = normalize_domain(args.domain)
    config = load_config()
    domains = {normalize_domain(value) for value in config.get("trusted_domains", [])}
    if trusted:
        domains.add(domain)
    else:
        domains.discard(domain)
    config["trusted_domains"] = sorted(domains)
    atomic_json_write(config_path(), config)
    print(f"已{'加入' if trusted else '移出'}可信域名：{domain}")
    return 0


def parse_iso_date(value: str | None, name: str) -> dt.date | None:
    if not value:
        return None
    try:
        return dt.date.fromisoformat(value)
    except ValueError as exc:
        raise ValueError(f"{name} 必须是 YYYY-MM-DD") from exc


def scan_metadata(
    explicit_range: bool,
    from_date: dt.date,
    to_date: dt.date | None,
    today: dt.date,
    initial_lookback_days: int,
) -> dict[str, Any]:
    if explicit_range:
        return {
            "mode": "date_range_rescan",
            "from": str(from_date),
            "to": str(to_date or today),
            "last_uid_advanced": False,
            "pending_uids_updated": True,
        }
    return {
        "mode": "uid_incremental",
        "new_folder_initial_lookback_days": initial_lookback_days,
        "description": "新文件夹首次回溯指定天数；已初始化文件夹扫描 last_uid 之后及 pending_uids 中的邮件",
    }


def cmd_run(args: argparse.Namespace) -> int:
    config = load_config()
    accounts = [item for item in config.get("accounts", []) if item.get("enabled", True)]
    if args.account:
        accounts = [item for item in accounts if item.get("email", "").lower() == args.account.lower()]
    if not accounts:
        raise ValueError("没有可运行的已启用邮箱账号")
    first_run_needs_persisting = not config.get("first_run_confirmed", False)
    if first_run_needs_persisting and not args.confirm_first_run:
        emails = "、".join(item["email"] for item in accounts)
        raise SkillError(
            "FIRST_RUN_CONFIRMATION_REQUIRED",
            f"首次扫描需要用户确认；邮箱：{emails}；写入目录：{config['invoice_root']}。确认后添加 --confirm-first-run",
        )
    today = dt.date.today()
    explicit_from = parse_iso_date(args.from_date, "--from-date")
    to_date = parse_iso_date(args.to_date, "--to-date")
    explicit_range = bool(args.from_date or args.to_date)
    initial_lookback_days = int(config.get("initial_lookback_days", 30))
    from_date = explicit_from or today - dt.timedelta(days=initial_lookback_days)
    if to_date and from_date > to_date:
        raise ValueError("开始日期不能晚于结束日期")
    root = Path(config["invoice_root"]).expanduser().resolve()
    root.mkdir(parents=True, exist_ok=True)
    if first_run_needs_persisting:
        config["first_run_confirmed"] = True
        atomic_json_write(config_path(), config)
    state = load_state()
    results: list[dict[str, str]] = []
    with exclusive_run_lock(), tempfile.TemporaryDirectory(prefix="invoice-mail-") as temp_name:
        temp_dir = Path(temp_name)
        for account in accounts:
            results.extend(
                run_account(
                    account,
                    root,
                    state,
                    from_date,
                    to_date,
                    temp_dir,
                    explicit_range,
                    config.get("trusted_domains", []),
                )
            )
            atomic_json_write(state_path(), state)
    counts = {status: sum(item["status"] == status for item in results) for status in ("success", "skipped", "incomplete", "error")}
    print(
        json.dumps(
            {
                "invoice_root": str(root),
                "scan": scan_metadata(explicit_range, from_date, to_date, today, initial_lookback_days),
                "counts": counts,
                "items": results,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 2 if counts["error"] else (1 if counts["incomplete"] else 0)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="从 163/QQ 邮箱下载并整理电子发票")
    subparsers = parser.add_subparsers(dest="command", required=True)

    configure = subparsers.add_parser("configure", help="添加或更新邮箱账号")
    configure.add_argument("--provider", choices=sorted(PROVIDERS), required=True)
    configure.add_argument("--email", required=True)
    configure.set_defaults(handler=cmd_configure)

    accounts = subparsers.add_parser("accounts", help="查看非敏感配置")
    accounts.set_defaults(handler=cmd_accounts)

    set_root = subparsers.add_parser("set-root", help="修改默认发票目录")
    set_root.add_argument("path")
    set_root.add_argument("--confirm", action="store_true", help="确认修改目录")
    set_root.set_defaults(handler=cmd_set_root)

    trusted_domains = subparsers.add_parser("trusted-domains", help="查看可信下载域名")
    trusted_domains.set_defaults(handler=cmd_trusted_domains)

    for command, trusted in (("trust-domain", True), ("untrust-domain", False)):
        domain_parser = subparsers.add_parser(command, help=f"{'加入' if trusted else '移出'}可信下载域名")
        domain_parser.add_argument("domain")
        domain_parser.add_argument("--confirm", action="store_true", help="确认修改可信域名")
        domain_parser.set_defaults(handler=lambda args, flag=trusted: cmd_trust_domain(args, flag))

    for command, enabled in (("enable", True), ("disable", False)):
        toggle = subparsers.add_parser(command, help=f"{'启用' if enabled else '停用'}邮箱账号")
        toggle.add_argument("--email", required=True)
        toggle.set_defaults(handler=lambda args, flag=enabled: cmd_toggle(args, flag))

    remove = subparsers.add_parser("remove-account", help="删除邮箱账号配置和凭据")
    remove.add_argument("--email", required=True)
    remove.set_defaults(handler=cmd_remove)

    run = subparsers.add_parser("run", help="执行一次扫描、下载和归档")
    run.add_argument("--account", help="只运行指定邮箱")
    run.add_argument("--from-date", help="开始日期 YYYY-MM-DD")
    run.add_argument("--to-date", help="结束日期 YYYY-MM-DD")
    run.add_argument("--confirm-first-run", action="store_true", help="确认首次邮箱读取和本地写入")
    run.set_defaults(handler=cmd_run)
    return parser


def main() -> int:
    try:
        args = build_parser().parse_args()
        return int(args.handler(args))
    except KeyboardInterrupt:
        print("操作已取消。", file=sys.stderr)
        return 130
    except SkillError as exc:
        print(f"错误 [{exc.code}]：{exc}", file=sys.stderr)
        return 2
    except Exception as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())

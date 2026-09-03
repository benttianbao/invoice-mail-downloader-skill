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
import html as html_module
from html.parser import HTMLParser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import imaplib
import ipaddress
import json
import os
from pathlib import Path, PurePosixPath
import re
import secrets
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
from http.cookiejar import CookieJar
import webbrowser
import zipfile


APP_NAME = "invoice-mail-downloader"
KEYRING_SERVICE = APP_NAME
MAX_FILE_BYTES = 50 * 1024 * 1024
MAX_ZIP_BYTES = 200 * 1024 * 1024
MAX_ZIP_FILES = 100
MAX_COMPRESSION_RATIO = 200
MAX_MESSAGE_BYTES = 70 * 1024 * 1024
IMAP_TIMEOUT_SECONDS = 20
KEYWORDS = ("发票", "电子票", "开票", "票据", "invoice", "e-invoice", "einvoice")
SUPPORTED_SUFFIXES = {".pdf", ".ofd", ".zip"}
KNOWN_PROVIDER_PAGE_HOSTS = {"nnfp.jss.com.cn", "pis.baiwang.com"}
REQUIRED_BUYER_NAME = "永赢金融租赁有限公司"
REQUIRED_BUYER_TAX_ID = "91330200316986507A"
PROVIDERS = {
    "163": {"domain": "163.com", "imap": "imap.163.com", "port": 993},
    "qq": {"domain": "qq.com", "imap": "imap.qq.com", "port": 993},
}
EXCLUDED_FLAGS = {b"\\Sent", b"\\Drafts", b"\\Trash", b"\\Junk", b"\\Noselect"}
EXCLUDED_NAMES = (
    "sent",
    "draft",
    "trash",
    "junk",
    "spam",
    "deleted messages",
    "deleted",
    "已发送",
    "草稿",
    "垃圾",
    "已删除",
)


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
        "version": 3,
        "invoice_root": str(default_invoice_root()),
        "trusted_domains": [],
        "accounts": [],
    }
    config = load_json(config_path(), defaults.copy())
    for key, value in defaults.items():
        config.setdefault(key, value)
    legacy_confirmed = bool(config.pop("first_run_confirmed", False))
    config.pop("initial_lookback_days", None)
    for account in config.get("accounts", []):
        account.setdefault("scan_confirmed", legacy_confirmed)
    config["version"] = 3
    return config


def load_state() -> dict[str, Any]:
    state = load_json(
        state_path(),
        {
            "version": 4,
            "folders": {},
            "files": {},
            "provenance": {},
            "invoice_formats": {},
            "invoice_format_index_built": False,
            "invoice_identity_version": 0,
        },
    )
    state.setdefault("folders", {})
    state.setdefault("files", {})
    state.setdefault("provenance", {})
    state.setdefault("invoice_formats", {})
    state.setdefault("invoice_format_index_built", False)
    state.setdefault("invoice_identity_version", 0)
    state["version"] = 4
    return state


def pid_is_running(pid: int) -> bool:
    if pid <= 0:
        return False
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    except OSError:
        return False
    return True


@contextlib.contextmanager
def exclusive_run_lock() -> Iterable[None]:
    lock = app_data_dir() / "run.lock"
    lock.parent.mkdir(parents=True, exist_ok=True)
    for _attempt in range(2):
        try:
            fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
            break
        except FileExistsError as exc:
            try:
                existing = json.loads(lock.read_text(encoding="utf-8"))
                existing_pid = int(existing.get("pid", 0))
            except (OSError, ValueError, TypeError, json.JSONDecodeError):
                existing_pid = 0
            if pid_is_running(existing_pid):
                raise SkillError("RUN_LOCKED", f"已有扫描正在运行（PID {existing_pid}）：{lock}") from exc
            with contextlib.suppress(FileNotFoundError):
                lock.unlink()
    else:
        raise SkillError("RUN_LOCKED", f"无法取得运行锁：{lock}")
    try:
        lock_value = json.dumps({"pid": os.getpid(), "created_at": dt.datetime.now().astimezone().isoformat()})
        os.write(fd, lock_value.encode("utf-8"))
        os.close(fd)
        yield
    finally:
        with contextlib.suppress(FileNotFoundError):
            lock.unlink()


def tls_context() -> ssl.SSLContext:
    import certifi

    return ssl.create_default_context(cafile=certifi.where())


def emit_progress(event: str, **values: Any) -> None:
    safe = {"event": event, **values}
    print(json.dumps(safe, ensure_ascii=False), file=sys.stderr, flush=True)


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
        parts = urllib.parse.urlsplit(url)
        if parts.scheme.lower() not in {"http", "https"}:
            continue
        if parts.hostname == "fp.nuonuo.com" and parts.path in {"", "/"} and not parts.query:
            continue
        suffix = Path(parts.path).suffix.lower()
        if suffix and suffix not in SUPPORTED_SUFFIXES:
            continue
        host = (parts.hostname or "").lower()
        if (
            has_invoice_keyword(label + " " + url)
            or (message_is_candidate and suffix in SUPPORTED_SUFFIXES)
            or (message_is_candidate and host in KNOWN_PROVIDER_PAGE_HOSTS)
        ):
            candidates.append(url)

    url_re = re.compile(r"https?://[^\s<>\"']+", re.IGNORECASE)
    for match in url_re.finditer(plain):
        url = match.group(0).rstrip(".,;:!?)]}，。；：！？）】")
        window = plain[max(0, match.start() - 100) : min(len(plain), match.end() + 100)]
        suffix = Path(urllib.parse.urlsplit(url).path).suffix.lower()
        if suffix and suffix not in SUPPORTED_SUFFIXES:
            continue
        if has_invoice_keyword(window) or (message_is_candidate and suffix in SUPPORTED_SUFFIXES):
            candidates.append(url)
    return list(dict.fromkeys(candidates))


def prioritized_download_links(links: Iterable[str]) -> list[str]:
    priority = {".pdf": 0, ".zip": 1, "": 2, ".ofd": 3}
    unique = list(dict.fromkeys(links))
    return sorted(unique, key=lambda url: priority.get(Path(urllib.parse.urlsplit(url).path).suffix.lower(), 2))


def provider_opener(trusted_domains: Iterable[str]) -> Any:
    return urllib.request.build_opener(
        urllib.request.HTTPCookieProcessor(CookieJar()),
        urllib.request.HTTPSHandler(context=tls_context()),
        SafeRedirectHandler(tuple(trusted_domains)),
    )


def read_limited(response: Any, limit: int) -> bytes:
    content = response.read(limit + 1)
    if len(content) > limit:
        raise SkillError("DOWNLOAD_PROVIDER_RESPONSE_TOO_LARGE", "发票平台接口响应超过安全限制")
    return content


def provider_download_links(page_url: str, trusted_domains: Iterable[str]) -> list[str]:
    """解析已知发票平台的只读详情接口，不执行网页 JavaScript。"""
    parts = urllib.parse.urlsplit(page_url)
    host = (parts.hostname or "").lower()
    query = urllib.parse.parse_qs(parts.query)
    if host == "pis.baiwang.com" and parts.path.rstrip("/") == "/smkp-vue/previewInvoiceAllEle":
        param = query.get("param", [""])[0]
        if not param:
            return []
        return [
            urllib.parse.urlunsplit(
                (
                    "https",
                    parts.netloc,
                    "/bwmg/mix/bw/downloadFormat",
                    urllib.parse.urlencode({"param": param, "formatType": "PDF"}),
                    "",
                )
            )
        ]
    if host != "nnfp.jss.com.cn":
        return []

    trusted = tuple(trusted_domains)
    validate_public_https(page_url, trusted)
    opener = provider_opener(trusted)
    request = urllib.request.Request(page_url, headers={"User-Agent": "invoice-mail-downloader/1.0"})
    with opener.open(request, timeout=30) as response:
        final_url = response.geturl()
        read_limited(response, 2 * 1024 * 1024)
    final_parts = urllib.parse.urlsplit(final_url)
    if final_parts.hostname != "nnfp.jss.com.cn" or not final_parts.path.endswith("/printQrcode"):
        return []
    final_query = urllib.parse.parse_qs(final_parts.query)
    param_list = final_query.get("paramList", [""])[0]
    if not param_list:
        return []
    api_url = urllib.parse.urlunsplit(("https", final_parts.netloc, "/scan2/getIvcDetailShow.do", "", ""))
    payload = urllib.parse.urlencode(
        {
            "paramList": param_list,
            "code": final_query.get("code", [""])[0],
            "aliView": final_query.get("aliView", [""])[0],
            "invoiceDetailMiddleUri": final_url,
            "shortLinkSource": final_query.get("shortLinkSource", [""])[0],
        }
    ).encode()
    api_request = urllib.request.Request(
        api_url,
        data=payload,
        headers={
            "User-Agent": "invoice-mail-downloader/1.0",
            "Referer": final_url,
            "Content-Type": "application/x-www-form-urlencoded",
        },
    )
    with opener.open(api_request, timeout=30) as response:
        body = read_limited(response, 2 * 1024 * 1024)
    try:
        result = json.loads(body.decode("utf-8"))
        pdf_url = result["data"]["invoiceSimpleVo"]["url"]
    except (KeyError, TypeError, UnicodeError, json.JSONDecodeError) as exc:
        raise SkillError("DOWNLOAD_PROVIDER_RESPONSE_INVALID", "诺诺发票详情接口未返回有效 PDF 地址") from exc
    pdf_parts = urllib.parse.urlsplit(str(pdf_url))
    if pdf_parts.scheme != "https" or Path(pdf_parts.path).suffix.lower() != ".pdf":
        raise SkillError("DOWNLOAD_PROVIDER_RESPONSE_INVALID", "诺诺发票详情接口返回的不是 HTTPS PDF 地址")
    return [str(pdf_url)]


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
        suffix = Path(parts.path).suffix.lower()
        marker = f"/<已脱敏>{suffix}" if suffix in SUPPORTED_SUFFIXES else "/<已脱敏>"
        return urllib.parse.urlunsplit((parts.scheme, parts.hostname or "", marker, "", ""))
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
        urllib.request.HTTPSHandler(context=tls_context()),
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


def unpack_zip(path: Path, target: Path) -> list[tuple[Path, str]]:
    output: list[tuple[Path, str]] = []
    members = sorted(
        safe_zip_members(path),
        key=lambda info: ({".pdf": 0, ".ofd": 1}.get(PurePosixPath(info.filename).suffix.lower(), 2), info.filename),
    )
    with zipfile.ZipFile(path) as archive:
        for index, info in enumerate(members, start=1):
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
                output.append((destination, info.filename))
            else:
                destination.unlink(missing_ok=True)
    return output


def pdf_text(path: Path) -> str:
    texts: list[str] = []
    errors: list[Exception] = []
    try:
        import pdfplumber

        with pdfplumber.open(path) as document:
            layout = "\n".join(
                page.extract_text(layout=True, x_tolerance=2, y_tolerance=3) or "" for page in document.pages[:5]
            )
        if layout.strip():
            texts.append(layout)
    except Exception as exc:
        errors.append(exc)
    try:
        from pypdf import PdfReader

        reader = PdfReader(str(path), strict=False)
        if reader.is_encrypted:
            with contextlib.suppress(Exception):
                reader.decrypt("")
        plain = "\n".join(page.extract_text() or "" for page in reader.pages[:5])
        if plain.strip() and plain not in texts:
            texts.append(plain)
    except Exception as exc:
        errors.append(exc)
    if texts:
        return "\n".join(texts)
    raise SkillError("PDF_TEXT_EXTRACTION_FAILED", f"PDF 文字提取失败：{type(errors[-1]).__name__ if errors else 'UnknownError'}")


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

    if not date_value and "开票日期" in flat:
        tail = flat.split("开票日期", 1)[1][:800]
        if match := re.search(r"(20\d{2})年(\d{1,2})月(\d{1,2})日", tail):
            with contextlib.suppress(ValueError):
                date_value = dt.date(*map(int, match.groups())).isoformat()
    if not date_value and "开票日期" in flat:
        if match := re.search(r"(20\d{2})年(\d{1,2})月(\d{1,2})日", flat):
            with contextlib.suppress(ValueError):
                date_value = dt.date(*map(int, match.groups())).isoformat()

    seller_patterns = (
        r"收到一张【([^】]{2,80})】开具",
        r"销售方(?:信息)?\s*(?:名称)?\s*[：:]?\s*([^：:]{2,80}?)(?=\s*(?:统一社会信用代码|纳税人识别号|地址|开户行|购买方|发票))",
        r"销方名称\s*[：:]?\s*([^：:]{2,80}?)(?=\s*(?:统一社会信用代码|纳税人识别号|地址|开户行|$))",
    )
    for pattern in seller_patterns:
        if match := re.search(pattern, flat):
            seller = normalize_text(match.group(1))
            break

    if not seller:
        for line in text.splitlines():
            names = [normalize_text(value) for value in re.findall(r"名称\s*[：:]\s*(.*?)(?=\s+名称\s*[：:]|$)", line)]
            names = [value for value in names if len(value) >= 2]
            if len(names) >= 2:
                seller = names[-1]
                break
    if not seller and "铁路电子客票" in flat:
        seller = "中国铁路"
    if not seller and re.search(r"销\s*售\s*方", flat):
        identity_pattern = r"20\d{2}年\d{1,2}月\d{1,2}日\s+(.{2,80}?)\s+[0-9A-Z]{15,20}\s+(.{2,100}?)\s+[0-9A-Z]{15,20}"
        if match := re.search(identity_pattern, flat):
            seller = normalize_text(match.group(2))

    amount_patterns = (
        r"价税合计[^\n]{0,160}?[（(]小写[)）]\s*[¥￥]?\s*([0-9][0-9,]*\.\d{2})",
        r"价税合计.{0,40}?(?:小写)?\s*[（(]?(?:小写)?[)）]?\s*[：:]?\s*[¥￥]?\s*([0-9][0-9,]*\.\d{2})",
        r"[（(]小写[)）]\s*[¥￥]?\s*([0-9][0-9,]*\.\d{2})",
        r"(?:票价|退票费)\s*[：:]?\s*[¥￥]?\s*([0-9][0-9,]*\.\d{2})",
    )
    for pattern in amount_patterns:
        if match := re.search(pattern, flat, re.I):
            amount = match.group(1).replace(",", "")
            break
    if not amount and "铁路电子客票" in flat:
        railway_amounts = re.findall(r"[¥￥]\s*([0-9][0-9,]*\.\d{2})", flat)
        if railway_amounts:
            amount = railway_amounts[-1].replace(",", "")
    if not amount and "价税合计" in flat:
        totals = re.findall(r"[¥￥]\s*([0-9][0-9,]*\.\d{2})", flat)
        if totals:
            amount = totals[-1].replace(",", "")
    return {"date": date_value, "seller": seller, "amount": amount}


def extract_buyer_fields(text: str) -> dict[str, str]:
    """提取票面上位于购买方一侧的名称和纳税人识别号。"""
    flat = normalize_text(text)
    compact = re.sub(r"\s+", "", flat)
    name = ""
    tax_id = ""

    name_label = r"(?:购\s*方\s*名\s*称|名\s*称)"
    next_label = (
        r"(?:名\s*称|统\s*一\s*社\s*会\s*信\s*用\s*代\s*码|"
        r"纳\s*税\s*人\s*识\s*别\s*号|地\s*址|开\s*户\s*行|"
        r"销\s*售\s*方|价\s*税\s*合\s*计)"
    )
    name_pattern = rf"{name_label}\s*[：:]\s*(.{{2,80}}?)(?=\s*{next_label}\s*[：:]?|$)"
    if match := re.search(name_pattern, flat, re.I):
        name = normalize_text(match.group(1))

    tax_label = (
        r"(?:统\s*一\s*社\s*会\s*信\s*用\s*代\s*码"
        r"(?:\s*/\s*纳\s*税\s*人\s*识\s*别\s*号)?|"
        r"纳\s*税\s*人\s*识\s*别\s*号|购\s*方\s*税\s*号)"
    )
    tax_pattern = rf"{tax_label}\s*[：:]?\s*([0-9A-Z](?:\s*[0-9A-Z]){{14,19}})"
    if match := re.search(tax_pattern, flat.upper()):
        tax_id = re.sub(r"[^0-9A-Z]", "", match.group(1).upper())

    # 数电发票的双栏文字层可能先输出两侧标签、再输出两侧值。
    # 此时用票面中第一个独立 18 位身份代码作为购买方代码；20 位发票号码不会误入。
    identity_codes = [
        re.sub(r"[^0-9A-Z]", "", value)
        for value in re.findall(
            r"(?<![0-9A-Z])([0-9A-Z](?:[ \t]*[0-9A-Z]){17})(?![0-9A-Z])",
            flat.upper(),
        )
    ]
    if identity_codes:
        tax_id = identity_codes[0]

    # 仅当购买方代码已经精确匹配时，才跨双栏文字层确认目标名称，
    # 避免目标公司仅出现在销售方一侧时产生误判。
    if tax_id == REQUIRED_BUYER_TAX_ID and re.sub(r"\s+", "", REQUIRED_BUYER_NAME) in compact:
        name = REQUIRED_BUYER_NAME

    return {"name": name, "tax_id": tax_id}


def is_railway_ticket(text: str) -> bool:
    compact = re.sub(r"\s+", "", normalize_text(text))
    return "铁路电子客票" in compact or ("中国铁路" in compact and "电子客票" in compact)


def buyer_validation_issues(text: str) -> list[str]:
    """返回购买方校验问题；铁路电子客票按业务规则豁免。"""
    if is_railway_ticket(text):
        return []
    buyer = extract_buyer_fields(text)
    actual_name = re.sub(r"\s+", "", buyer["name"])
    expected_name = re.sub(r"\s+", "", REQUIRED_BUYER_NAME)
    actual_tax_id = re.sub(r"[^0-9A-Z]", "", buyer["tax_id"].upper())
    issues: list[str] = []
    if not actual_name:
        issues.append("购买方名称缺失")
    elif actual_name != expected_name:
        issues.append("购买方名称不匹配")
    if not actual_tax_id:
        issues.append("购买方纳税人识别号缺失")
    elif actual_tax_id != REQUIRED_BUYER_TAX_ID:
        issues.append("购买方纳税人识别号不匹配")
    return issues


def invoice_identity(text: str, filename: str = "") -> str:
    """提取可用于跨格式去重的可靠票号；无法可靠识别时返回空字符串。"""
    flat = normalize_text(text)
    patterns = (
        r"发票号码?\s*[：:]?\s*([0-9][0-9\s]{6,30}[0-9])",
        r"(?:电子)?客票(?:号码?|号)\s*[：:]?\s*([0-9][0-9\s]{6,30}[0-9])",
    )
    for pattern in patterns:
        if match := re.search(pattern, flat, re.I):
            number = re.sub(r"\s+", "", match.group(1))
            if 8 <= len(number) <= 30:
                return f"number:{number}"

    # 仅当文件名明确标注票号时才降级使用文件名，避免把日期误当发票号码。
    if match := re.search(r"(?:dzfp[_-]?|发票|票号|invoice)[^0-9]{0,12}([0-9]{8,30})", filename, re.I):
        return f"number:{match.group(1)}"
    return ""


def is_invoice_document(text: str, filename: str = "") -> bool:
    flat = normalize_text(text)
    lowered_name = filename.lower()
    excluded_titles = ("境外汇款申请书", "application for funds transfers", "借款合同", "劳动合同")
    if any(value in flat.lower() for value in excluded_titles):
        return False
    strong_titles = ("电子发票", "数电发票", "增值税发票", "铁路电子客票", "航空运输电子客票")
    if any(value in flat for value in strong_titles):
        return True
    if has_invoice_keyword(lowered_name):
        return True
    signals = sum(
        bool(re.search(pattern, flat, re.I))
        for pattern in (
            r"发票号码?\s*[：:]?\s*\d{8,}",
            r"开票日期\s*[：:]?",
            r"价税合计|[（(]小写[)）]",
            r"销售方(?:信息)?|销方名称",
        )
    )
    return signals >= 3


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


def archive_invoice(
    source: Path,
    suffix: str,
    root: Path,
    state: dict[str, Any],
    text: str | None = None,
    provenance: dict[str, Any] | None = None,
) -> tuple[str, str, list[str]]:
    if text is None:
        text = pdf_text(source) if suffix == ".pdf" else ofd_text(source)
    fields = extract_invoice_fields(text)
    review_reasons = [key for key in ("date", "seller", "amount") if not fields[key]]
    review_reasons.extend(buyer_validation_issues(text))
    date_label = fields["date"] or "未知日期"
    seller_label = safe_component(fields["seller"], "未知销售方")
    amount_label = fields["amount"] or "未知金额"
    filename = f"{date_label}_{seller_label}_¥{amount_label}{suffix}"

    if review_reasons:
        destination_dir = root / "待确认"
    else:
        parsed_date = dt.date.fromisoformat(fields["date"])
        destination_dir = root / f"{parsed_date.year:04d}" / f"{parsed_date.month:02d}-{parsed_date.day:02d}"
    destination_dir.mkdir(parents=True, exist_ok=True)
    destination = destination_dir / filename

    digest = sha256_file(source)
    existing = state.setdefault("files", {}).get(digest)
    if existing and Path(existing).exists():
        existing_path = Path(existing)
        if existing_path != destination and not destination.exists():
            try:
                existing_path.resolve().relative_to(root.resolve())
                existing_path.replace(destination)
                for formats in state.setdefault("invoice_formats", {}).values():
                    for format_name, value in list(formats.items()):
                        if Path(value) == existing_path:
                            formats[format_name] = str(destination)
                existing = str(destination)
                state["files"][digest] = existing
            except (OSError, ValueError):
                pass
        if provenance:
            state.setdefault("provenance", {})[digest] = provenance
        return "skipped", existing, review_reasons

    if destination.exists():
        if sha256_file(destination) == digest:
            state["files"][digest] = str(destination)
            if provenance:
                state.setdefault("provenance", {})[digest] = provenance
            return "skipped", str(destination), review_reasons
        destination = destination.with_name(f"{destination.stem}_{digest[:8]}{suffix}")
    shutil.copy2(source, destination)
    state["files"][digest] = str(destination)
    if provenance:
        state.setdefault("provenance", {})[digest] = provenance
    return "success", str(destination), review_reasons


def recorded_format_path(state: dict[str, Any], invoice_key: str, suffix: str) -> Path | None:
    value = state.setdefault("invoice_formats", {}).get(invoice_key, {}).get(suffix.lstrip("."))
    if not value:
        return None
    path = Path(value)
    if path.exists():
        return path
    state["invoice_formats"][invoice_key].pop(suffix.lstrip("."), None)
    return None


def remove_recorded_ofd(state: dict[str, Any], invoice_key: str, root: Path) -> str:
    """PDF 成功归档后，仅删除本技能状态中记录且位于归档根目录内的同票号 OFD。"""
    ofd_path = recorded_format_path(state, invoice_key, ".ofd")
    if ofd_path is None:
        return ""
    try:
        resolved = ofd_path.resolve()
        resolved.relative_to(root.resolve())
    except (OSError, ValueError):
        return ""
    if resolved.suffix.lower() != ".ofd":
        return ""
    resolved.unlink(missing_ok=True)
    removed_digests = [digest for digest, value in state.setdefault("files", {}).items() if Path(value) == ofd_path]
    for digest in removed_digests:
        state["files"].pop(digest, None)
        state.setdefault("provenance", {}).pop(digest, None)
    state["invoice_formats"][invoice_key].pop("ofd", None)
    return str(ofd_path)


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


def folder_uidvalidity(client: imaplib.IMAP4_SSL, wire_name: str) -> int | None:
    value = response_number(client, "UIDVALIDITY")
    if value is not None:
        return value
    try:
        status, values = client.status(wire_name, "(UIDVALIDITY)")
    except imaplib.IMAP4.error:
        return None
    if status == "OK":
        for item in values or []:
            if isinstance(item, bytes) and (match := re.search(rb"UIDVALIDITY\s+(\d+)", item, re.I)):
                return int(match.group(1))
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
    provenance: dict[str, Any] | None = None,
    seen_pdf_keys: set[str] | None = None,
    trusted_invoice_source: bool = False,
) -> list[tuple[str, str, str]]:
    if seen_pdf_keys is None:
        seen_pdf_keys = set()
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
            return [("skipped", hinted_name, "ZIP 中没有有效 PDF/OFD，原 ZIP 不保留")]
        results: list[tuple[str, str, str]] = []
        for extracted_file, member_name in extracted:
            child_provenance = dict(provenance or {})
            child_provenance.update({"container": hinted_name, "member": member_name})
            results.extend(
                process_file(
                    extracted_file,
                    member_name,
                    root,
                    state,
                    temp_dir,
                    child_provenance,
                    seen_pdf_keys,
                    trusted_invoice_source,
                )
            )
        return results
    extraction_warning = ""
    try:
        text = pdf_text(path) if suffix == ".pdf" else ofd_text(path)
    except Exception as exc:
        if not trusted_invoice_source:
            return [("incomplete", hinted_name, str(exc))]
        text = ""
        extraction_warning = f"；可信发票平台文件已保留，但文字提取失败：{error_code(exc, 'PDF_TEXT_EXTRACTION_FAILED')}"
    trusted_context = str((provenance or {}).get("subject", "")) if trusted_invoice_source else ""
    invoice_text = f"{text}\n{trusted_context}".strip()
    if not is_invoice_document(invoice_text, hinted_name) and not trusted_invoice_source:
        return [("skipped", hinted_name, "附件内容不是可识别的发票，未归档")]
    invoice_key = invoice_identity(invoice_text, hinted_name)
    if suffix == ".ofd" and invoice_key:
        existing_pdf = recorded_format_path(state, invoice_key, ".pdf")
        if invoice_key in seen_pdf_keys or existing_pdf is not None:
            detail = str(existing_pdf) if existing_pdf is not None else "本批次已处理同一发票的 PDF"
            return [("skipped", hinted_name, f"同一发票已有 PDF，未保留 OFD；{detail}")]
    status, destination, review_reasons = archive_invoice(path, suffix, root, state, invoice_text, provenance)
    removed_ofd = ""
    if invoice_key:
        formats = state.setdefault("invoice_formats", {}).setdefault(invoice_key, {})
        formats[suffix.lstrip(".")] = destination
        if suffix == ".pdf":
            seen_pdf_keys.add(invoice_key)
            removed_ofd = remove_recorded_ofd(state, invoice_key, root)
    detail = destination
    if review_reasons:
        detail += "；待确认原因：" + "、".join(review_reasons)
    detail += extraction_warning
    if removed_ofd:
        detail += f"；已移除同一发票的 OFD：{removed_ofd}"
    return [(status, hinted_name, detail)]


def process_message(
    raw: bytes,
    root: Path,
    state: dict[str, Any],
    temp_dir: Path,
    trusted_domains: Iterable[str] = (),
    provenance_base: dict[str, Any] | None = None,
) -> tuple[list[dict[str, str]], bool]:
    message = BytesParser(policy=policy.default).parsebytes(raw)
    subject = decode_mime(message.get("Subject"))
    sender = parseaddr(decode_mime(message.get("From")))[1]
    message_date = decode_mime(message.get("Date"))
    plain, html, attachments = message_content(message)
    candidate = has_invoice_keyword(" ".join((subject, plain, re.sub(r"<[^>]+>", " ", html))))
    candidate = candidate or any(has_invoice_keyword(name) for name, _ in attachments)
    if not candidate:
        return [], False

    items: list[dict[str, str]] = []
    seen_pdf_keys: set[str] = set()
    with tempfile.TemporaryDirectory(dir=temp_dir) as message_tmp_name:
        message_tmp = Path(message_tmp_name)
        prioritized_attachments = sorted(
            enumerate(attachments, start=1),
            key=lambda item: ({".pdf": 0, ".zip": 1, ".ofd": 2}.get(Path(item[1][0]).suffix.lower(), 3), item[0]),
        )
        for index, (filename, payload) in prioritized_attachments:
            suffix = Path(filename).suffix.lower()
            if suffix not in SUPPORTED_SUFFIXES:
                continue
            if len(payload) > MAX_FILE_BYTES:
                items.append(report_item("incomplete", f"附件:{filename}", subject, sender, "文件超过脚本限制", "ATTACHMENT_TOO_LARGE"))
                continue
            attachment_path = message_tmp / f"attachment-{index}{suffix}"
            attachment_path.write_bytes(payload)
            provenance = dict(provenance_base or {})
            provenance.update(
                {
                    "message_date": message_date,
                    "subject": subject,
                    "sender": sender,
                    "source_type": "attachment",
                    "original_filename": filename,
                }
            )
            for status, source, detail in process_file(
                attachment_path, filename, root, state, message_tmp, provenance, seen_pdf_keys
            ):
                items.append(report_item(status, f"附件:{source}", subject, sender, detail))

        links = prioritized_download_links(extract_candidate_links(plain, html, candidate))
        for index, url in enumerate(links, start=1):
            target = message_tmp / f"download-{index}"
            redacted = redact_url(url)
            try:
                content_type, filename = download_http(url, target, trusted_domains)
                if content_type in {"text/html", "application/xhtml+xml"} or target.read_bytes()[:256].lstrip().lower().startswith(b"<"):
                    page_html = target.read_text(encoding="utf-8", errors="replace")
                    page_links = provider_download_links(url, trusted_domains)
                    page_links.extend([
                        urllib.parse.urljoin(url, child)
                        for child in extract_candidate_links("", page_html, True)
                        if urllib.parse.urljoin(url, child) != url
                    ])
                    page_links = prioritized_download_links(page_links)
                    target.unlink(missing_ok=True)
                    if page_links:
                        for child_index, child_url in enumerate(dict.fromkeys(page_links), start=1):
                            child_target = message_tmp / f"download-{index}-{child_index}"
                            child_source = redact_url(child_url)
                            try:
                                child_type, child_name = download_http(child_url, child_target, trusted_domains)
                                if child_type in {"text/html", "application/xhtml+xml"}:
                                    raise SkillError("DOWNLOAD_DYNAMIC_PAGE", "下载链接仍返回网页；脚本不会执行 JavaScript 或点击按钮")
                                provenance = dict(provenance_base or {})
                                provenance.update(
                                    {
                                        "message_date": message_date,
                                        "subject": subject,
                                        "sender": sender,
                                        "source_type": "link",
                                        "source_url": child_source,
                                    }
                                )
                                for status, _source, detail in process_file(
                                    child_target,
                                    child_name,
                                    root,
                                    state,
                                    message_tmp,
                                    provenance,
                                    seen_pdf_keys,
                                    True,
                                ):
                                    items.append(report_item(status, child_source, subject, sender, detail))
                            except Exception as exc:
                                items.append(report_item("incomplete", child_source, subject, sender, str(exc), error_code(exc, "DOWNLOAD_FAILED")))
                        continue
                    raise SkillError("DOWNLOAD_DYNAMIC_PAGE", "页面没有静态发票文件链接；脚本不会执行 JavaScript 或点击按钮")
                provenance = dict(provenance_base or {})
                provenance.update(
                    {
                        "message_date": message_date,
                        "subject": subject,
                        "sender": sender,
                        "source_type": "link",
                        "source_url": redacted,
                    }
                )
                for status, _source, detail in process_file(
                    target, filename, root, state, message_tmp, provenance, seen_pdf_keys
                ):
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
    status, metadata = client.uid("fetch", str(uid), "(RFC822.SIZE)")
    if status != "OK":
        raise RuntimeError("无法读取邮件大小")
    joined = b" ".join(value for value in metadata or [] if isinstance(value, bytes))
    if match := re.search(rb"RFC822\.SIZE\s+(\d+)", joined, re.I):
        if int(match.group(1)) > MAX_MESSAGE_BYTES:
            raise SkillError("MESSAGE_TOO_LARGE", "邮件超过 70 MB 安全限制，已保留待重试状态")
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
    start_from_now: bool = False,
    state_saver: Any | None = None,
) -> list[dict[str, str]]:
    email = account["email"]
    provider_name = account["provider"]
    provider = PROVIDERS[provider_name]
    output: list[dict[str, str]] = []
    context = tls_context()
    with imaplib.IMAP4_SSL(
        provider["imap"], provider["port"], ssl_context=context, timeout=IMAP_TIMEOUT_SECONDS
    ) as client:
        client.login(email, secret)
        send_imap_identity(client, provider_name)
        for wire_name, display_name in folder_entries(client):
            emit_progress("folder_start", account=email, folder=display_name)
            try:
                status, select_response = client.select(wire_name, readonly=True)
            except imaplib.IMAP4.error as exc:
                output.append(report_item("error", f"文件夹:{display_name}", "", email, str(exc), "FOLDER_SELECT_FAILED"))
                continue
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
            uidvalidity = folder_uidvalidity(client, wire_name)
            if uidvalidity is None:
                item_status = "skipped" if explicit_range else "error"
                detail = (
                    "服务器未返回 UIDVALIDITY；指定日期运行已安全跳过该文件夹"
                    if explicit_range
                    else "服务器未返回 UIDVALIDITY"
                )
                output.append(
                    report_item(item_status, f"文件夹:{display_name}", "", email, detail, "UIDVALIDITY_MISSING")
                )
                continue
            if folder_state.get("uidvalidity") != uidvalidity:
                folder_state.clear()
                folder_state.update({"uidvalidity": uidvalidity, "initialized": False, "last_uid": 0, "pending_uids": []})

            previous_pending = {int(value) for value in folder_state.get("pending_uids", [])}
            last_uid = int(folder_state.get("last_uid", 0))
            initialize_from_range = False
            try:
                if start_from_now:
                    checkpoint = current_last_uid(client)
                    folder_state.update(
                        {"uidvalidity": uidvalidity, "initialized": True, "last_uid": checkpoint, "pending_uids": []}
                    )
                    if state_saver:
                        state_saver()
                    emit_progress("folder_initialized", account=email, folder=display_name, last_uid=checkpoint)
                    continue
                if explicit_range:
                    uids = search_date_uids(client, from_date, to_date)
                    # 指定日期范围是一次独立扫描：不得把任何历史待处理 UID 混入本次候选。
                    pending: set[int] = set()
                    if folder_state.get("initialized", False):
                        checkpoint = last_uid
                    else:
                        checkpoint = current_last_uid(client)
                        uids = [uid for uid in uids if uid <= checkpoint]
                        initialize_from_range = True
                elif not folder_state.get("initialized", False):
                    raise SkillError("FIRST_RUN_MODE_REQUIRED", "首次扫描必须选择明确日期范围或从现在开始增量")
                else:
                    uids = sorted(set(search_new_uids(client, last_uid)) | previous_pending)
                    pending = set(uids)
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

            folder_state["pending_uids"] = sorted(pending)
            if not explicit_range or initialize_from_range:
                folder_state["last_uid"] = checkpoint
                folder_state["initialized"] = True
            if state_saver:
                state_saver()
            emit_progress("folder_candidates", account=email, folder=display_name, count=len(uids))

            for index, uid in enumerate(uids, start=1):
                emit_progress("message_start", account=email, folder=display_name, uid=uid, index=index, total=len(uids))
                try:
                    raw = fetch_message(client, uid)
                    items, incomplete = process_message(
                        raw,
                        root,
                        state,
                        temp_dir,
                        trusted_domains,
                        {"account": email, "folder": display_name, "folder_state_key": key, "uid": uid},
                    )
                    output.extend(items)
                    if incomplete:
                        pending.add(uid)
                    else:
                        pending.discard(uid)
                except (socket.timeout, imaplib.IMAP4.abort, OSError, ssl.SSLError) as exc:
                    pending.add(uid)
                    output.append(
                        report_item(
                            "error",
                            f"邮件UID:{uid}",
                            "",
                            email,
                            f"IMAP 命令超时或连接中断：{type(exc).__name__}",
                            "IMAP_COMMAND_TIMEOUT",
                        )
                    )
                    folder_state["pending_uids"] = sorted(pending)
                    if state_saver:
                        state_saver()
                    return output
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
                if state_saver:
                    state_saver()
                emit_progress("message_done", account=email, folder=display_name, uid=uid)
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
    start_from_now: bool = False,
    state_saver: Any | None = None,
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
                start_from_now,
                state_saver,
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


def validate_credentials(provider_name: str, email: str, secret: str) -> None:
    provider = PROVIDERS[provider_name]
    try:
        with imaplib.IMAP4_SSL(
            provider["imap"], provider["port"], ssl_context=tls_context(), timeout=IMAP_TIMEOUT_SECONDS
        ) as client:
            client.login(email, secret)
            send_imap_identity(client, provider_name)
    except imaplib.IMAP4.error as exc:
        raise SkillError("IMAP_AUTH_FAILED", "IMAP 登录失败，请检查服务开关和客户端授权码") from exc
    except (OSError, ssl.SSLError) as exc:
        raise SkillError("IMAP_NETWORK_FAILED", f"IMAP TLS 连接失败：{type(exc).__name__}") from exc


def save_account(provider: str, email: str, secret: str) -> None:
    try:
        credential_set(email, secret)
    except Exception as exc:
        raise SkillError("KEYRING_ERROR", f"系统凭据库写入失败：{type(exc).__name__}") from exc
    config = load_config()
    accounts = config.setdefault("accounts", [])
    existing = next((item for item in accounts if item.get("email", "").lower() == email), None)
    if existing:
        existing.update({"provider": provider, "enabled": True})
        existing.setdefault("scan_confirmed", False)
    else:
        accounts.append({"email": email, "provider": provider, "enabled": True, "scan_confirmed": False})
    atomic_json_write(config_path(), config)


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
    validate_credentials(args.provider, email, secret)
    save_account(args.provider, email, secret)
    print(f"已安全配置：{email}")
    return 0


def configuration_page(provider: str, email: str, token: str, error: str = "") -> str:
    error_html = f'<p class="error">{html_module.escape(error)}</p>' if error else ""
    return f"""<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>邮箱发票下载器配置</title><style>
body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;background:#f5f7fb;margin:0;padding:32px}}
main{{max-width:520px;margin:auto;background:white;padding:28px;border-radius:14px;box-shadow:0 8px 30px #0001}}
label{{display:block;margin:16px 0 6px}} input{{box-sizing:border-box;width:100%;padding:12px;border:1px solid #ccd3df;border-radius:8px}}
button{{margin-top:22px;width:100%;padding:12px;border:0;border-radius:8px;background:#1769e0;color:white;font-size:16px}}
.note{{color:#596579;font-size:14px;line-height:1.6}} .error{{color:#b42318;background:#fee4e2;padding:10px;border-radius:8px}}
</style></head><body><main><h1>配置 {html_module.escape(provider.upper())} 邮箱</h1>
<p class="note">授权码只会提交给本机 127.0.0.1，并直接写入系统凭据库。页面不会联网发送、记录或回显授权码。</p>
{error_html}<form method="post" autocomplete="off">
<input type="hidden" name="token" value="{html_module.escape(token)}">
<input type="hidden" name="provider" value="{html_module.escape(provider)}">
<label>邮箱地址</label><input name="email" type="email" readonly value="{html_module.escape(email)}">
<label>客户端授权码</label><input name="secret" type="password" required autofocus autocomplete="new-password">
<button type="submit">验证并安全保存</button></form></main></body></html>"""


def cmd_configure_ui(args: argparse.Namespace) -> int:
    provider = args.provider
    email = validate_account(provider, args.email)
    token = secrets.token_urlsafe(24)

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, _format: str, *_values: Any) -> None:
            return

        def send_html(self, value: str, status: int = 200) -> None:
            encoded = value.encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Length", str(len(encoded)))
            self.end_headers()
            self.wfile.write(encoded)

        def do_GET(self) -> None:  # noqa: N802
            query = urllib.parse.parse_qs(urllib.parse.urlsplit(self.path).query)
            if query.get("token", [""])[0] != token:
                self.send_html("<h1>链接无效或已过期</h1>", 403)
                return
            self.send_html(configuration_page(provider, email, token))

        def do_POST(self) -> None:  # noqa: N802
            length = min(int(self.headers.get("Content-Length", "0")), 65536)
            values = urllib.parse.parse_qs(self.rfile.read(length).decode("utf-8", errors="replace"))
            if values.get("token", [""])[0] != token:
                self.send_html("<h1>请求无效或已过期</h1>", 403)
                return
            secret = values.get("secret", [""])[0].strip()
            try:
                if not secret:
                    raise ValueError("授权码不能为空")
                validate_credentials(provider, email, secret)
                save_account(provider, email, secret)
            except Exception as exc:
                self.send_html(configuration_page(provider, email, token, str(exc)), 400)
                return
            self.server.configuration_complete = True  # type: ignore[attr-defined]
            self.send_html("<h1>配置成功</h1><p>账号已验证并保存到系统凭据库，可以关闭此页面并返回 Agent。</p>")

    with ThreadingHTTPServer(("127.0.0.1", 0), Handler) as server:
        server.configuration_complete = False  # type: ignore[attr-defined]
        server.timeout = 1
        url = f"http://127.0.0.1:{server.server_port}/?token={token}"
        print(json.dumps({"status": "waiting_for_user", "url": url, "email": email}, ensure_ascii=False), flush=True)
        webbrowser.open(url)
        deadline = time.monotonic() + 600
        while not server.configuration_complete and time.monotonic() < deadline:  # type: ignore[attr-defined]
            server.handle_request()
        if not server.configuration_complete:  # type: ignore[attr-defined]
            raise SkillError("CONFIGURATION_TIMEOUT", "本地配置页面等待超时，请重新启动")
    print(json.dumps({"status": "configured", "email": email}, ensure_ascii=False))
    return 0


def cmd_accounts(_args: argparse.Namespace) -> int:
    config = load_config()
    print(
        json.dumps(
            {
                "invoice_root": config["invoice_root"],
                "trusted_domains": config["trusted_domains"],
                "accounts": config["accounts"],
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


def cmd_preflight(_args: argparse.Namespace) -> int:
    import certifi
    import keyring
    import pdfplumber
    import pypdf

    ca_path = Path(certifi.where())
    result = {
        "ok": sys.version_info >= (3, 10) and ca_path.is_file(),
        "python": sys.version.split()[0],
        "platform": sys.platform,
        "ca_bundle": str(ca_path),
        "ca_bundle_exists": ca_path.is_file(),
        "keyring_backend": type(keyring.get_keyring()).__name__,
        "dependencies": {"pypdf": pypdf.__version__, "pdfplumber": pdfplumber.__version__},
        "data_dir": str(app_data_dir()),
        "invoice_root": load_config()["invoice_root"],
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["ok"] else 2


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
    start_from_now: bool = False,
    initializing: bool = False,
) -> dict[str, Any]:
    if start_from_now:
        return {
            "mode": "incremental_from_now_initialized",
            "description": "已记录当前 UID；本次不读取历史邮件，后续只处理新邮件",
        }
    if explicit_range:
        return {
            "mode": "date_range_initial" if initializing else "date_range_rescan",
            "from": str(from_date),
            "to": str(to_date),
            "last_uid_advanced": initializing,
            "pending_uids_updated": True,
            "pending_scope": "current_date_range_only",
        }
    return {
        "mode": "uid_incremental",
        "description": "扫描 last_uid 之后及 pending_uids 中的邮件，不使用隐含日期窗口",
    }


def rebuild_file_index(root: Path, state: dict[str, Any]) -> None:
    files = state.setdefault("files", {})
    previously_recorded = {Path(value) for value in files.values()}
    if not root.exists():
        return
    for path in root.rglob("*"):
        if path.is_file() and path.suffix.lower() in {".pdf", ".ofd"}:
            digest = sha256_file(path)
            recorded = files.get(digest)
            if not recorded or not Path(recorded).exists():
                files[digest] = str(path)
    if state.get("invoice_format_index_built"):
        return

    # 只迁移旧状态已登记的文件；不把用户手工放入根目录的文件纳入自动删除范围。
    for path in previously_recorded:
        if not path.exists() or path.suffix.lower() not in {".pdf", ".ofd"}:
            continue
        try:
            text = pdf_text(path) if path.suffix.lower() == ".pdf" else ofd_text(path)
        except Exception:
            continue
        invoice_key = invoice_identity(text, path.name)
        if invoice_key:
            state.setdefault("invoice_formats", {}).setdefault(invoice_key, {})[path.suffix.lower().lstrip(".")] = str(path)
    for invoice_key, formats in list(state.setdefault("invoice_formats", {}).items()):
        if formats.get("pdf") and formats.get("ofd"):
            remove_recorded_ofd(state, invoice_key, root)
    state["invoice_format_index_built"] = True


def provenance_filename(provenance: dict[str, Any], fallback: str) -> str:
    if provenance.get("original_filename"):
        return str(provenance["original_filename"])
    if provenance.get("source_url"):
        path = urllib.parse.urlsplit(str(provenance["source_url"])).path
        return urllib.parse.unquote(Path(path).name)
    return fallback


def rebuild_invoice_identity_index(root: Path, state: dict[str, Any]) -> None:
    """票号规则升级后，使用已记录来源重建格式索引并清理同票号 OFD。"""
    if int(state.get("invoice_identity_version", 0)) >= 2:
        return
    formats = state.setdefault("invoice_formats", {})
    provenance_by_digest = state.setdefault("provenance", {})
    for digest, value in list(state.setdefault("files", {}).items()):
        path = Path(value)
        if not path.exists() or path.suffix.lower() not in {".pdf", ".ofd"}:
            continue
        provenance = provenance_by_digest.get(digest, {})
        hint = provenance_filename(provenance, path.name)
        invoice_key = invoice_identity("", hint)
        if not invoice_key:
            try:
                text = pdf_text(path) if path.suffix.lower() == ".pdf" else ofd_text(path)
            except Exception:
                continue
            invoice_key = invoice_identity(text, hint)
        if invoice_key:
            formats.setdefault(invoice_key, {})[path.suffix.lower().lstrip(".")] = str(path)
    for invoice_key, available in list(formats.items()):
        if available.get("pdf") and available.get("ofd"):
            remove_recorded_ofd(state, invoice_key, root)
    state["invoice_identity_version"] = 2


def folder_state_key_from_provenance(state: dict[str, Any], provenance: dict[str, Any]) -> str:
    explicit = str(provenance.get("folder_state_key", ""))
    if explicit in state.setdefault("folders", {}):
        return explicit
    account = str(provenance.get("account", "")).lower()
    display_name = str(provenance.get("folder", ""))
    prefix = f"{account}::"
    for key in state["folders"]:
        if not key.startswith(prefix):
            continue
        wire_name = key[len(prefix) :].strip('"')
        if wire_name == display_name or decode_modified_utf7(wire_name) == display_name:
            return key
    return ""


def requeue_missing_archives(state: dict[str, Any]) -> list[dict[str, str]]:
    """把已登记但已丢失的归档文件对应邮件重新放入增量待处理队列。"""
    unresolved: list[dict[str, str]] = []
    missing_paths = {
        digest: Path(value)
        for digest, value in state.setdefault("files", {}).items()
        if not Path(value).exists()
    }
    for digest, missing_path in missing_paths.items():
        for formats in state.setdefault("invoice_formats", {}).values():
            for format_name, value in list(formats.items()):
                if Path(value) == missing_path:
                    formats.pop(format_name, None)
        provenance = state.setdefault("provenance", {}).get(digest, {})
        uid_value = provenance.get("uid")
        key = folder_state_key_from_provenance(state, provenance)
        try:
            uid = int(uid_value)
        except (TypeError, ValueError):
            uid = 0
        if key and uid > 0:
            folder_state = state["folders"][key]
            pending = {int(value) for value in folder_state.get("pending_uids", [])}
            pending.add(uid)
            folder_state["pending_uids"] = sorted(pending)
            continue
        # 旧版本没有记录来源的失效索引无法恢复；清除索引，避免每次运行重复报警。
        state.setdefault("files", {}).pop(digest, None)
        state.setdefault("provenance", {}).pop(digest, None)
    return unresolved


def repair_state(state: dict[str, Any], account: str | None = None) -> dict[str, int]:
    """丢弃未完成扫描队列并清理磁盘上已不存在的归档索引，不修改增量游标。"""
    account_prefix = f"{account.lower()}::" if account else ""
    discarded_pending = 0
    for key, folder_state in state.setdefault("folders", {}).items():
        if account_prefix and not key.lower().startswith(account_prefix):
            continue
        pending = folder_state.get("pending_uids", [])
        discarded_pending += len(pending)
        folder_state["pending_uids"] = []

    missing_digests = [
        digest for digest, value in state.setdefault("files", {}).items() if not Path(value).exists()
    ]
    missing_paths = {Path(state["files"][digest]) for digest in missing_digests}
    for digest in missing_digests:
        state["files"].pop(digest, None)
        state.setdefault("provenance", {}).pop(digest, None)
    for formats in state.setdefault("invoice_formats", {}).values():
        for format_name, value in list(formats.items()):
            if Path(value) in missing_paths:
                formats.pop(format_name, None)
    return {"discarded_pending_uids": discarded_pending, "pruned_missing_archives": len(missing_digests)}


def cmd_repair_state(args: argparse.Namespace) -> int:
    if not args.confirm:
        raise SkillError("CONFIRMATION_REQUIRED", "丢弃待处理扫描和清理失效索引需要用户确认后添加 --confirm")
    state = load_state()
    result = repair_state(state, args.account)
    atomic_json_write(state_path(), state)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


def cmd_run(args: argparse.Namespace) -> int:
    config = load_config()
    accounts = [item for item in config.get("accounts", []) if item.get("enabled", True)]
    if args.account:
        accounts = [item for item in accounts if item.get("email", "").lower() == args.account.lower()]
    if not accounts:
        raise ValueError("没有可运行的已启用邮箱账号")
    explicit_from = parse_iso_date(args.from_date, "--from-date")
    to_date = parse_iso_date(args.to_date, "--to-date")
    if bool(explicit_from) != bool(to_date):
        raise ValueError("指定日期范围时必须同时提供 --from-date 和 --to-date")
    explicit_range = explicit_from is not None and to_date is not None
    start_from_now = bool(getattr(args, "start_from_now", False))
    if explicit_range and start_from_now:
        raise ValueError("日期范围与 --start-from-now 不能同时使用")

    unconfirmed = [item for item in accounts if not item.get("scan_confirmed", False)]
    if unconfirmed and not args.confirm_first_run:
        emails = "、".join(item["email"] for item in unconfirmed)
        raise SkillError(
            "FIRST_RUN_CONFIRMATION_REQUIRED",
            f"首次运行需要用户选择日期范围或从现在开始增量，并确认邮箱：{emails}；写入目录：{config['invoice_root']}",
        )
    if unconfirmed and not (explicit_range or start_from_now):
        raise SkillError(
            "FIRST_RUN_MODE_REQUIRED",
            "首次运行不再默认回溯30天；请选择 --from-date/--to-date，或使用 --start-from-now",
        )
    today = dt.date.today()
    from_date = explicit_from or today
    if to_date and from_date > to_date:
        raise ValueError("开始日期不能晚于结束日期")
    root = Path(config["invoice_root"]).expanduser().resolve()
    root.mkdir(parents=True, exist_ok=True)
    if unconfirmed:
        for account in unconfirmed:
            account["scan_confirmed"] = True
        atomic_json_write(config_path(), config)
    state = load_state()
    rebuild_file_index(root, state)
    rebuild_invoice_identity_index(root, state)
    # 日期范围和“从现在开始”模式不得继承历史待处理扫描。
    results = [] if explicit_range or start_from_now else requeue_missing_archives(state)
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
                    start_from_now,
                    lambda: atomic_json_write(state_path(), state),
                )
            )
            atomic_json_write(state_path(), state)
    counts = {status: sum(item["status"] == status for item in results) for status in ("success", "skipped", "incomplete", "error")}
    print(
        json.dumps(
            {
                "invoice_root": str(root),
                "scan": scan_metadata(explicit_range, from_date, to_date, start_from_now, bool(unconfirmed)),
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

    configure_ui = subparsers.add_parser("configure-ui", help="在本机浏览器中安全配置邮箱账号")
    configure_ui.add_argument("--provider", choices=sorted(PROVIDERS), required=True)
    configure_ui.add_argument("--email", required=True)
    configure_ui.set_defaults(handler=cmd_configure_ui)

    accounts = subparsers.add_parser("accounts", help="查看非敏感配置")
    accounts.set_defaults(handler=cmd_accounts)

    preflight = subparsers.add_parser("preflight", help="检查 Python、CA 证书、凭据库和依赖")
    preflight.set_defaults(handler=cmd_preflight)

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

    repair = subparsers.add_parser("repair-state", help="丢弃未完成扫描并清理失效归档索引")
    repair.add_argument("--account", help="只清理指定邮箱的待处理 UID；失效归档索引始终全局清理")
    repair.add_argument("--confirm", action="store_true", help="确认丢弃待处理扫描和失效索引")
    repair.set_defaults(handler=cmd_repair_state)

    run = subparsers.add_parser("run", help="执行一次扫描、下载和归档")
    run.add_argument("--account", help="只运行指定邮箱")
    run.add_argument("--from-date", help="邮件接收开始日期 YYYY-MM-DD，必须与 --to-date 同时使用")
    run.add_argument("--to-date", help="邮件接收结束日期 YYYY-MM-DD，必须与 --from-date 同时使用")
    run.add_argument("--start-from-now", action="store_true", help="首次不扫描历史邮件，只建立增量 UID 起点")
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

#!/usr/bin/env python3
"""从 163/QQ 邮箱只读获取并归档电子发票。"""

from __future__ import annotations

import argparse
import base64
import contextlib
import contextvars
import datetime as dt
from decimal import Decimal, InvalidOperation
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
import subprocess
import sys
import tempfile
import time
from typing import Any, Iterable
import unicodedata
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
INVOICE_REGISTER_FILENAME = "发票登记.xlsx"
RECEIPT_INBOX_NAME = "报销凭证待匹配"
RECEIPT_SUFFIXES = {".jpg", ".jpeg", ".png", ".heic", ".webp", ".pdf"}
INBOX_INVOICE_SUFFIXES = {".pdf", ".ofd"}
XPARSE_FREE_MAX_BYTES = 10 * 1024 * 1024
XPARSE_TIMEOUT_SECONDS = 180
RECEIPT_DATE_BEFORE_DAYS = 90
RECEIPT_DATE_AFTER_DAYS = 7
PROVIDERS = {
    "163": {"domain": "163.com", "imap": "imap.163.com", "port": 993},
    "qq": {"domain": "qq.com", "imap": "imap.qq.com", "port": 993},
}
EXCLUDED_FLAGS = {b"\\Sent", b"\\Drafts", b"\\Trash", b"\\Junk", b"\\Noselect"}
EXCLUDED_FOLDER_NAMES = {
    "sent",
    "draft",
    "drafts",
    "trash",
    "junk",
    "spam",
    "deleted",
    "deleted messages",
    "已发送",
    "草稿",
    "垃圾",
    "垃圾箱",
    "垃圾邮件",
    "已删除",
}
EXCLUDED_FOLDER_TOKENS = {"sent", "draft", "drafts", "trash", "junk", "spam", "deleted"}
IMAP_MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
XPARSE_DETERMINISTIC_CODES = {"XPARSE_FILE_TOO_LARGE", "RECEIPT_IMAGE_INVALID"}
XPARSE_QUOTA_CODES = {"XPARSE_QUOTA_EXCEEDED"}
_buyer_name = contextvars.ContextVar("buyer_name", default=REQUIRED_BUYER_NAME)
_buyer_tax_id = contextvars.ContextVar("buyer_tax_id", default=REQUIRED_BUYER_TAX_ID)


class SkillError(RuntimeError):
    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code


def error_code(exc: BaseException, fallback: str) -> str:
    return exc.code if isinstance(exc, SkillError) else fallback


def current_buyer() -> tuple[str, str]:
    return _buyer_name.get(), _buyer_tax_id.get()


def apply_buyer_from_config(config: dict[str, Any]) -> None:
    _buyer_name.set(str(config.get("required_buyer_name") or REQUIRED_BUYER_NAME).strip() or REQUIRED_BUYER_NAME)
    _buyer_tax_id.set(
        re.sub(r"[^0-9A-Z]", "", str(config.get("required_buyer_tax_id") or REQUIRED_BUYER_TAX_ID).upper())
        or REQUIRED_BUYER_TAX_ID
    )


def imap_date(value: dt.date) -> str:
    return f"{value.day:02d}-{IMAP_MONTHS[value.month - 1]}-{value.year}"


def folder_is_excluded(display_name: str, flags: set[bytes], delimiter: str = "") -> bool:
    if flags & EXCLUDED_FLAGS:
        return True
    stripped = display_name.strip()
    lowered = stripped.lower()
    if lowered in EXCLUDED_FOLDER_NAMES or stripped in EXCLUDED_FOLDER_NAMES:
        return True
    separators = {" ", "/", "_", "-", "."}
    if delimiter:
        separators.add(delimiter)
    tokens = [token for token in re.split("[" + re.escape("".join(separators)) + "]+", stripped) if token]
    lowered_tokens = [token.lower() for token in tokens]
    if any(token in EXCLUDED_FOLDER_NAMES or token.lower() in EXCLUDED_FOLDER_NAMES for token in tokens):
        return True
    if any(token in EXCLUDED_FOLDER_TOKENS for token in lowered_tokens):
        return True
    last = tokens[-1] if tokens else stripped
    return last in EXCLUDED_FOLDER_NAMES or last.lower() in EXCLUDED_FOLDER_NAMES


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
        "required_buyer_name": REQUIRED_BUYER_NAME,
        "required_buyer_tax_id": REQUIRED_BUYER_TAX_ID,
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
            "version": 6,
            "folders": {},
            "files": {},
            "provenance": {},
            "invoice_formats": {},
            "invoice_records": {},
            "receipt_records": {},
            "invoice_format_index_built": False,
            "invoice_identity_version": 0,
        },
    )
    state.setdefault("folders", {})
    state.setdefault("files", {})
    state.setdefault("provenance", {})
    state.setdefault("invoice_formats", {})
    state.setdefault("invoice_records", {})
    state.setdefault("receipt_records", {})
    state.setdefault("invoice_format_index_built", False)
    state.setdefault("invoice_identity_version", 0)
    state["version"] = 6
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
    expected_name, expected_tax_id = current_buyer()
    if tax_id == expected_tax_id and re.sub(r"\s+", "", expected_name) in compact:
        name = expected_name

    return {"name": name, "tax_id": tax_id}


def is_railway_ticket(text: str) -> bool:
    compact = re.sub(r"\s+", "", normalize_text(text))
    return "铁路电子客票" in compact or ("中国铁路" in compact and "电子客票" in compact)


def buyer_validation_issues(text: str) -> list[str]:
    """返回购买方校验问题；铁路电子客票按业务规则豁免。"""
    if is_railway_ticket(text):
        return []
    buyer = extract_buyer_fields(text)
    expected_name, expected_tax_id = current_buyer()
    actual_name = re.sub(r"\s+", "", buyer["name"])
    expected_name = re.sub(r"\s+", "", expected_name)
    actual_tax_id = re.sub(r"[^0-9A-Z]", "", buyer["tax_id"].upper())
    issues: list[str] = []
    if not actual_name:
        issues.append("购买方名称缺失")
    elif actual_name != expected_name:
        issues.append("购买方名称不匹配")
    if not actual_tax_id:
        issues.append("购买方纳税人识别号缺失")
    elif actual_tax_id != expected_tax_id:
        issues.append("购买方纳税人识别号不匹配")
    return issues


def extract_invoice_number(text: str, filename: str = "") -> str:
    """提取票面或明确文件名中的可靠票号；无法可靠识别时返回空字符串。"""
    flat = normalize_text(text)
    patterns = (
        r"发票号码?\s*[：:]?\s*([0-9][0-9\s]{6,30}[0-9])",
        r"(?:电子)?客票(?:号码?|号)\s*[：:]?\s*([0-9][0-9\s]{6,30}[0-9])",
    )
    for pattern in patterns:
        if match := re.search(pattern, flat, re.I):
            number = re.sub(r"\s+", "", match.group(1))
            if 8 <= len(number) <= 30:
                return number

    # 仅当文件名明确标注票号时才降级使用文件名，避免把日期误当发票号码。
    if match := re.search(r"(?:dzfp[_-]?|发票|票号|invoice)[^0-9]{0,12}([0-9]{8,30})", filename, re.I):
        return match.group(1)
    return ""


def invoice_identity(text: str, filename: str = "") -> str:
    """提取可用于跨格式去重的可靠票号；无法可靠识别时返回空字符串。"""
    number = extract_invoice_number(text, filename)
    return f"number:{number}" if number else ""


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


def receipt_cache_path(digest: str) -> Path:
    return app_data_dir() / "receipt-cache" / f"{digest}.json"


def find_xparse_cli() -> Path:
    """定位 TextIn 官方 CLI；免费模式不读取或保存 API 密钥。"""
    configured = os.environ.get("INVOICE_XPARSE_CLI", "").strip()
    local_bin = app_data_dir() / "xparse-cli" / "node_modules" / ".bin"
    candidates = [
        Path(configured).expanduser() if configured else None,
        local_bin / ("xparse-cli.cmd" if os.name == "nt" else "xparse-cli"),
    ]
    for command in ("xparse-cli", "xparse-cli.exe"):
        located = shutil.which(command)
        if located:
            candidates.append(Path(located))
    executable = next((path for path in candidates if path and path.is_file()), None)
    if executable is None:
        raise SkillError(
            "XPARSE_CLI_MISSING",
            "未找到 TextIn xparse-cli；报销凭证未上传，请重新运行 bootstrap.py 后重试",
        )
    return executable


def prepare_receipt_upload(source: Path, temp_dir: Path) -> Path:
    """把图片转换为不含 EXIF 的兼容副本；原始凭证保持不变。"""
    suffix = source.suffix.lower()
    if source.stat().st_size > XPARSE_FREE_MAX_BYTES and suffix == ".pdf":
        raise SkillError("XPARSE_FILE_TOO_LARGE", "免费 XParse 单文件不得超过 10 MB")
    if suffix == ".pdf":
        return source
    try:
        from PIL import Image, ImageOps
        if suffix == ".heic":
            from pillow_heif import register_heif_opener

            register_heif_opener()
    except ImportError as exc:
        raise SkillError("RECEIPT_IMAGE_RUNTIME_MISSING", "缺少 Pillow/pillow-heif 图片处理依赖") from exc

    target = temp_dir / f"{sha256_file(source)}.jpg"
    try:
        with Image.open(source) as opened:
            image = ImageOps.exif_transpose(opened).convert("RGB")
            longest = max(image.size)
            if longest > 20000:
                scale = 20000 / longest
                image = image.resize((max(1, int(image.width * scale)), max(1, int(image.height * scale))))
            quality = 92
            while True:
                image.save(target, format="JPEG", quality=quality, optimize=True)
                if target.stat().st_size <= XPARSE_FREE_MAX_BYTES or quality <= 60:
                    break
                quality -= 8
            while target.stat().st_size > XPARSE_FREE_MAX_BYTES and max(image.size) > 1200:
                image = image.resize((max(1, int(image.width * 0.85)), max(1, int(image.height * 0.85))))
                image.save(target, format="JPEG", quality=76, optimize=True)
    except SkillError:
        raise
    except Exception as exc:
        raise SkillError("RECEIPT_IMAGE_INVALID", f"凭证图片无法读取：{type(exc).__name__}") from exc
    if target.stat().st_size > XPARSE_FREE_MAX_BYTES:
        raise SkillError("XPARSE_FILE_TOO_LARGE", "图片压缩后仍超过免费 XParse 10 MB 限制")
    return target


def _xparse_text_nodes(value: Any, output: list[dict[str, Any]]) -> None:
    if isinstance(value, dict):
        text_value = value.get("text")
        if isinstance(text_value, str) and text_value.strip():
            output.append(
                {
                    "text": text_value.strip(),
                    "page_number": value.get("page_number"),
                    "coordinates": value.get("coordinates") if isinstance(value.get("coordinates"), list) else None,
                }
            )
        for key, child in value.items():
            if key not in {"image", "image_data", "base64", "url", "page_image_url"}:
                _xparse_text_nodes(child, output)
    elif isinstance(value, list):
        for child in value:
            _xparse_text_nodes(child, output)


def simplify_xparse_result(value: dict[str, Any]) -> dict[str, Any]:
    """只保留字段解析需要的文本和位置，不落盘图片或远程 URL。"""
    data = value.get("data") if isinstance(value.get("data"), dict) else value
    markdown = data.get("markdown", "") if isinstance(data, dict) else ""
    nodes: list[dict[str, Any]] = []
    _xparse_text_nodes(data, nodes)
    unique: list[dict[str, Any]] = []
    seen: set[tuple[str, Any]] = set()
    for node in nodes:
        key = (node["text"], node.get("page_number"))
        if key not in seen:
            seen.add(key)
            unique.append(node)
    text = str(markdown).strip() or "\n".join(node["text"] for node in unique)
    return {
        "status": "success",
        "text": text,
        "elements": unique,
        "engine": "textin-xparse-free",
        "cached_at": dt.datetime.now().astimezone().replace(microsecond=0).isoformat(),
    }


def parse_json_output(output: str) -> dict[str, Any]:
    stripped = output.strip()
    try:
        value = json.loads(stripped)
    except json.JSONDecodeError:
        start = stripped.find("{")
        end = stripped.rfind("}")
        if start < 0 or end <= start:
            raise SkillError("XPARSE_RESPONSE_INVALID", "XParse 未返回 JSON")
        try:
            value = json.loads(stripped[start : end + 1])
        except json.JSONDecodeError as exc:
            raise SkillError("XPARSE_RESPONSE_INVALID", "XParse 返回的 JSON 无法解析") from exc
    if not isinstance(value, dict):
        raise SkillError("XPARSE_RESPONSE_INVALID", "XParse 返回结果不是 JSON 对象")
    return value


def run_xparse(source: Path, temp_dir: Path) -> dict[str, Any]:
    upload = prepare_receipt_upload(source, temp_dir)
    command = [
        str(find_xparse_cli()), "parse", str(upload), "--api", "free", "--view", "json",
        "--include-image-data=false", "--include-inline-objects=false",
        "--include-table-structure=false", "--include-title-tree=false",
    ]
    try:
        completed = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=XPARSE_TIMEOUT_SECONDS,
        )
    except subprocess.TimeoutExpired as exc:
        raise SkillError("XPARSE_TIMEOUT", "XParse 识别超过 180 秒") from exc
    except OSError as exc:
        raise SkillError("XPARSE_RUN_FAILED", f"XParse 无法启动：{type(exc).__name__}") from exc
    if completed.returncode != 0:
        detail = completed.stderr.strip().splitlines()[-1] if completed.stderr.strip() else "未知错误"
        code = "XPARSE_QUOTA_EXCEEDED" if re.search(r"1000|quota|额度", detail, re.I) else "XPARSE_RUN_FAILED"
        raise SkillError(code, f"XParse 免费识别失败：{detail[:300]}")
    return simplify_xparse_result(parse_json_output(completed.stdout))


def xparse_cache_usable(result: dict[str, Any]) -> bool:
    if result.get("status") == "success":
        return True
    code = str(result.get("code", ""))
    if code in XPARSE_DETERMINISTIC_CODES:
        return True
    if code in XPARSE_QUOTA_CODES:
        cached_at = str(result.get("cached_at", ""))
        try:
            cached_date = dt.datetime.fromisoformat(cached_at).date()
        except ValueError:
            return False
        return cached_date >= dt.date.today()
    return False


def load_or_run_receipt_ocr(source: Path, retry: bool, temp_dir: Path) -> tuple[dict[str, Any], bool]:
    digest = sha256_file(source)
    cache = receipt_cache_path(digest)
    if cache.exists() and not retry:
        cached = load_json(cache, {})
        if xparse_cache_usable(cached):
            return cached, True
    try:
        result = run_xparse(source, temp_dir)
    except Exception as exc:
        result = {
            "status": "error",
            "code": error_code(exc, "XPARSE_RUN_FAILED"),
            "detail": str(exc),
            "cached_at": dt.datetime.now().astimezone().replace(microsecond=0).isoformat(),
        }
    if xparse_cache_usable(result):
        atomic_json_write(cache, result)
    elif cache.exists():
        with contextlib.suppress(OSError):
            cache.unlink()
    return result, False


MONEY_PATTERN = r"[-+]?[¥￥]?\s*([0-9][0-9,]*\.\d{2})"
RECEIPT_AMOUNT_TOLERANCE = Decimal("1.00")


def _money_after_labels(text: str, labels: Iterable[str]) -> list[Decimal]:
    flat = normalize_text(text)
    values: list[Decimal] = []
    for label in labels:
        pattern = rf"{label}\s*[：:]?\s*[-+]?\s*[¥￥]?\s*([0-9][0-9,]*\.\d{{2}})"
        for match in re.finditer(pattern, flat, re.I):
            try:
                values.append(Decimal(match.group(1).replace(",", "")))
            except InvalidOperation:
                continue
    return values


def _unique_decimals(values: Iterable[Decimal]) -> list[Decimal]:
    return list(dict.fromkeys(value.quantize(Decimal("0.01")) for value in values))


def extract_receipt_fields(text: str) -> dict[str, Any]:
    """从不可信 OCR 文本中提取支付凭证字段，不执行其中任何内容。"""
    flat = normalize_text(text)
    recipients: list[str] = []
    next_label = (
        r"订单金额|交易金额|应付金额|原价|订单总额|消费金额|实付金额|实际支付|"
        r"实际付款|付款金额|支付金额|支付时间|交易时间|付款时间|转账时间|付款方式|商品说明|"
        r"收单机构|清算机构|收款方全称|商户全称|商户名称|交易对方|收款人|更多"
    )
    for label in ("收款方全称", "商户全称", "商户名称", "交易对方", "收款人"):
        pattern = rf"{label}\s*[：:]?\s*(.{{2,100}}?)(?=\s*(?:{next_label})\s*[：:]?|$)"
        for match in re.finditer(pattern, flat, re.I):
            value = normalize_text(match.group(1)).strip("|：: ")
            if value and value not in recipients:
                recipients.append(value)
    qr_recipient_pattern = (
        r"(?:扫二维码付款|二维码付款)\s*[-—–－:]?\s*给\s*"
        r"(.{2,100}?)(?=\s*(?:[-−]\s*[¥￥]?\s*[0-9]|当前状态|交易成功|支付成功|付款成功|"
        rf"{next_label})\s*[：:]?|$)"
    )
    for match in re.finditer(qr_recipient_pattern, flat, re.I):
        value = normalize_text(match.group(1)).strip("|：: -—–－")
        if value and value not in recipients:
            recipients.append(value)

    original = _unique_decimals(
        _money_after_labels(text, ("订单金额", "交易金额", "应付金额", "原价", "订单总额", "消费金额"))
    )
    paid = _unique_decimals(
        _money_after_labels(text, ("实付金额", "实际支付金额?", "实际付款金额?", "付款金额", "支付金额", "支出金额"))
    )
    if not paid:
        negatives = re.findall(r"-\s*[¥￥]?\s*([0-9][0-9,]*\.\d{2})", flat)
        parsed_negatives = [Decimal(value.replace(",", "")) for value in negatives]
        paid = [max(parsed_negatives)] if parsed_negatives else []
    discounts: list[Decimal] = []
    for match in re.finditer(
        r"(?:优惠券?|红包|立减金?|支付优惠|银行卡优惠|折扣|减免).{0,20}?[-+]?\s*[¥￥]?\s*([0-9][0-9,]*\.\d{2})",
        flat,
        re.I,
    ):
        with contextlib.suppress(InvalidOperation):
            discounts.append(abs(Decimal(match.group(1).replace(",", ""))))

    payment_time = ""
    time_match = re.search(
        r"(?:支付时间|交易时间|付款时间|转账时间)\s*[：:]?\s*(20\d{2})[-年/.](\d{1,2})[-月/.](\d{1,2})日?\s*(\d{1,2})?[：:]?(\d{1,2})?[：:]?(\d{1,2})?",
        flat,
    )
    if time_match:
        parts = [int(value) if value else 0 for value in time_match.groups()]
        try:
            payment_time = dt.datetime(*parts[:3], *parts[3:]).isoformat(timespec="seconds")
        except ValueError:
            payment_time = ""

    transaction_ids: list[str] = []
    for match in re.finditer(
        r"(?:交易单号|商户订单号|支付流水号|交易流水号|转账单号)\s*[：:]?\s*([0-9A-Za-z_-]{8,64})",
        flat,
        re.I,
    ):
        if match.group(1) not in transaction_ids:
            transaction_ids.append(match.group(1))

    if re.search(r"退款|已撤销|交易关闭|冲正|退回", flat):
        transaction_status = "refund"
    elif re.search(r"处理中|待入账|待支付|支付失败|交易失败|已关闭|已取消", flat):
        transaction_status = "invalid"
    elif re.search(r"交易成功|支付成功|付款成功|交易完成|支付完成|已完成", flat):
        transaction_status = "success"
    else:
        transaction_status = "missing"
    return {
        "merchant_names": recipients,
        "original_amounts": [str(value) for value in original],
        "paid_amounts": [str(value) for value in paid],
        "discount_amount": str(sum(discounts, Decimal("0")).quantize(Decimal("0.01"))) if discounts else "",
        "payment_time": payment_time,
        "transaction_ids": transaction_ids,
        "transaction_status": transaction_status,
    }


def normalized_party(value: str) -> str:
    value = unicodedata.normalize("NFKC", value).lower()
    return re.sub(r"[^0-9a-z\u4e00-\u9fff]", "", value)


def seller_matches_receipt(seller: str, names: Iterable[str]) -> bool:
    expected = normalized_party(seller)
    return bool(expected) and any(expected in normalized_party(name) for name in names)


def seller_matches_receipt_full_text(seller: str, text: str) -> bool:
    """在 OCR 全文中查找完整销方名称，但排除明显的付款方上下文。"""
    expected = normalized_party(seller)
    haystack = normalized_party(text)
    if not expected or not haystack:
        return False
    blocked_prefixes = tuple(
        normalized_party(label)
        for label in (
            "付款方全称", "付款方名称", "付款方", "付款人", "支付方",
            "付款户名", "付款账户名", "付款账户", "备注",
        )
    )
    start = 0
    while True:
        index = haystack.find(expected, start)
        if index < 0:
            return False
        prefix = haystack[max(0, index - 16):index]
        if not any(prefix.endswith(label) for label in blocked_prefixes):
            return True
        start = index + len(expected)


def receipt_seller_match_source(seller: str, fields: dict[str, Any], text: str = "") -> str:
    names = fields.get("merchant_names") or []
    if names:
        return "structured" if seller_matches_receipt(seller, names) else ""
    return "full_text" if seller_matches_receipt_full_text(seller, text) else ""


def receipt_transaction_fingerprint(fields: dict[str, Any]) -> str:
    transaction_ids = fields.get("transaction_ids") or []
    if len(transaction_ids) == 1:
        return f"id:{transaction_ids[0]}"
    payload = "|".join(
        (
            ",".join(normalized_party(value) for value in fields.get("merchant_names", [])),
            ",".join(fields.get("original_amounts", [])),
            ",".join(fields.get("paid_amounts", [])),
            str(fields.get("payment_time", "")),
        )
    )
    return "fields:" + hashlib.sha256(payload.encode("utf-8")).hexdigest()


def receipt_candidate(invoice: dict[str, Any], fields: dict[str, Any], text: str = "") -> tuple[bool, str]:
    if invoice.get("validation_status") != "通过":
        return False, "发票自身待确认"
    if not all(invoice.get(key) for key in ("invoice_number", "invoice_date", "invoice_amount", "seller")):
        return False, "发票关键字段缺失"
    if invoice.get("has_receipt") == "是":
        return False, "发票已有报销凭证"
    status = fields.get("transaction_status")
    if status == "refund":
        return False, "存在退款、撤销或冲正"
    if status == "invalid":
        return False, "交易未成功"
    if status != "success":
        return False, "交易状态缺失"
    names = fields.get("merchant_names") or []
    seller = str(invoice["seller"])
    if names and not seller_matches_receipt(seller, names):
        return False, "收款方与开票方不一致"
    if not names and not seller_matches_receipt_full_text(seller, text):
        return False, "收款方全称缺失"
    if (
        len(fields.get("transaction_ids") or []) > 1
        or len(fields.get("original_amounts") or []) > 1
        or len(fields.get("paid_amounts") or []) > 1
    ):
        return False, "凭证包含多笔交易"
    try:
        invoice_amount = Decimal(str(invoice["invoice_amount"])).quantize(Decimal("0.01"))
        original_values = [Decimal(value) for value in fields.get("original_amounts") or []]
        paid_values = [Decimal(value) for value in fields.get("paid_amounts") or []]
    except InvalidOperation:
        return False, "金额无法解析"
    original = original_values[0].quantize(Decimal("0.01")) if len(original_values) == 1 else None
    paid = paid_values[0].quantize(Decimal("0.01")) if len(paid_values) == 1 else None
    if original is not None:
        if original != invoice_amount:
            return False, "优惠前订单金额与发票金额不一致"
        if paid is not None and paid > original:
            return False, "实付金额高于订单金额"
        if paid is not None and paid < original:
            difference = original - paid
            if difference <= RECEIPT_AMOUNT_TOLERANCE:
                discount = fields.get("discount_amount")
                if discount and abs(difference - Decimal(str(discount))) > Decimal("0.01"):
                    return False, "优惠明细与差额不一致"
            else:
                discount = fields.get("discount_amount")
                if not discount:
                    return False, "优惠差额无法解释"
                if abs(difference - Decimal(str(discount))) > Decimal("0.01"):
                    return False, "优惠明细与差额不一致"
    elif paid is None:
        return False, "支付金额缺失"
    elif paid > invoice_amount:
        return False, "实付金额高于发票金额"
    elif invoice_amount - paid > RECEIPT_AMOUNT_TOLERANCE:
        return False, "实付金额与发票金额不一致"

    payment_time = str(fields.get("payment_time", ""))
    if payment_time:
        try:
            payment_date = dt.date.fromisoformat(payment_time[:10])
            invoice_date = dt.date.fromisoformat(str(invoice["invoice_date"]))
            days = (payment_date - invoice_date).days
            if days < -RECEIPT_DATE_BEFORE_DAYS or days > RECEIPT_DATE_AFTER_DAYS:
                return False, "支付日期超出允许范围"
        except ValueError:
            return False, "支付日期无法解析"
    return True, "开票方、金额和交易状态一致"


def receipt_amount_difference(invoice_amount: Any, fields: dict[str, Any]) -> str:
    paid_values = fields.get("paid_amounts") or []
    if len(paid_values) != 1:
        return ""
    try:
        difference = Decimal(str(invoice_amount)) - Decimal(str(paid_values[0]))
    except InvalidOperation:
        return ""
    return str(difference.quantize(Decimal("0.01")))


def receipt_match_detail(
    fields: dict[str, Any], invoice_amount: Any = "", seller: str = "", text: str = ""
) -> str:
    original = (fields.get("original_amounts") or [""])[0]
    paid = (fields.get("paid_amounts") or [""])[0]
    discount = fields.get("discount_amount", "")
    source = receipt_seller_match_source(seller, fields, text) if seller else "structured"
    parts = ["OCR 全文包含完整开票方" if source == "full_text" else "开票方与收款方一致"]
    if original:
        parts.append(f"订单金额 {original} 元")
    if paid:
        parts.append(f"实际支付 {paid} 元")
    if discount:
        parts.append(f"优惠 {discount} 元")
    difference = receipt_amount_difference(invoice_amount, fields)
    if difference:
        parts.append(f"发票与实付差额 {difference} 元")
    return "；".join(parts)


def receipt_destination(invoice: dict[str, Any], source: Path, digest: str) -> Path:
    invoice_path = Path(str(invoice["archive_path"]))
    destination = invoice_path.with_name(f"{invoice_path.stem}_报销凭证{source.suffix.lower()}")
    if destination.exists() and sha256_file(destination) != digest:
        destination = destination.with_name(f"{destination.stem}_{digest[:8]}{destination.suffix}")
    return destination


def _receipt_result(status: str, source: Path, detail: str, code: str = "") -> dict[str, Any]:
    value = {"status": status, "source": str(source), "detail": detail}
    if code:
        value["code"] = code
    return value


def inbox_invoice_handled(status: str, detail: str) -> bool:
    if status == "success":
        return True
    return status == "skipped" and "不是可识别的发票" not in detail


def inbox_invoice_code(status: str, detail: str) -> str:
    if not inbox_invoice_handled(status, detail):
        return "INBOX_NOT_INVOICE"
    if "文字提取失败" in detail:
        return "INBOX_INVOICE_UNREADABLE"
    if "未能识别为发票" in detail:
        return "INBOX_OFD_UNRECOGNIZED"
    return "INBOX_INVOICE_ARCHIVED"


def inbox_invoice_summary(items: list[dict[str, Any]]) -> dict[str, int]:
    archived = 0
    pending_review = 0
    left_for_receipts = 0
    for item in items:
        code = str(item.get("code", ""))
        if code == "INBOX_NOT_INVOICE":
            left_for_receipts += 1
            continue
        archived += 1
        if "待确认" in str(item.get("detail", "")) or code in {"INBOX_INVOICE_UNREADABLE", "INBOX_OFD_UNRECOGNIZED"}:
            pending_review += 1
    return {
        "archived": archived,
        "pending_review": pending_review,
        "left_for_receipts": left_for_receipts,
    }


def archive_dropped_invoices(root: Path, state: dict[str, Any]) -> dict[str, Any]:
    """投放目录中的 PDF/OFD 先按本地文字提取归档，不调用 XParse。"""
    inbox = root / RECEIPT_INBOX_NAME
    inbox.mkdir(parents=True, exist_ok=True)
    candidates = sorted(
        (path for path in inbox.iterdir() if path.is_file() and path.suffix.lower() in INBOX_INVOICE_SUFFIXES),
        key=lambda path: ({".pdf": 0, ".ofd": 1}.get(path.suffix.lower(), 9), path.name.lower()),
    )
    items: list[dict[str, Any]] = []
    seen_pdf_keys: set[str] = set()
    with tempfile.TemporaryDirectory(prefix="inbox-invoice-") as temp_name:
        temp_dir = Path(temp_name)
        for source in candidates:
            provenance = {"source_type": "inbox", "original_filename": source.name}
            results = process_file(
                source, source.name, root, state, temp_dir, provenance, seen_pdf_keys, False, True
            )
            handled = False
            destinations: list[Path] = []
            for status, _name, detail in results:
                item = _receipt_result(status, source, detail)
                item["code"] = inbox_invoice_code(status, detail)
                if inbox_invoice_handled(status, detail):
                    handled = True
                    destinations.append(Path(detail.split("；", 1)[0]))
                items.append(item)
            if not handled or not source.exists():
                continue
            if any(path.exists() and path.resolve() == source.resolve() for path in destinations):
                continue
            source.unlink(missing_ok=True)
    counts = {status: sum(item["status"] == status for item in items) for status in ("success", "skipped", "incomplete", "error")}
    summary = inbox_invoice_summary(items)
    return {"status": "success", "inbox": str(inbox), "counts": counts, "summary": summary, "items": items}


def discover_manually_placed_receipts(root: Path, state: dict[str, Any], invoices: list[dict[str, Any]]) -> None:
    """识别日期目录中用户手工放入的凭证，不扫描任意根目录文件。"""
    invoice_paths = {Path(value).resolve() for value in state.setdefault("files", {}).values() if Path(value).exists()}
    records = state.setdefault("receipt_records", {})
    invoices_by_dir: dict[Path, list[dict[str, Any]]] = {}
    for invoice in invoices:
        invoices_by_dir.setdefault(Path(str(invoice["archive_path"])).parent.resolve(), []).append(invoice)
    for directory, directory_invoices in invoices_by_dir.items():
        if not directory.exists():
            continue
        for path in directory.iterdir():
            if not path.is_file() or path.suffix.lower() not in RECEIPT_SUFFIXES or path.resolve() in invoice_paths:
                continue
            digest = sha256_file(path)
            record = records.setdefault(digest, {})
            record["archive_path"] = str(path)
            record["source_path"] = str(path)
            matched_invoice = directory_invoices[0] if len(directory_invoices) == 1 else next(
                (
                    invoice
                    for invoice in directory_invoices
                    if path.stem.startswith(Path(str(invoice["archive_path"])).stem + "_报销凭证")
                ),
                None,
            )
            if matched_invoice:
                record["invoice_key"] = matched_invoice["record_key"]
                record.setdefault("match_status", "待确认：人工放置，尚未校验")
                record.setdefault("match_detail", "凭证文件已存在于发票目录")


def match_receipts(root: Path, state: dict[str, Any], retry: bool = False) -> dict[str, Any]:
    inbox = root / RECEIPT_INBOX_NAME
    inbox.mkdir(parents=True, exist_ok=True)
    inbox_invoices = archive_dropped_invoices(root, state)
    invoices = rebuild_invoice_records(root, state)
    discover_manually_placed_receipts(root, state, invoices)
    sources = sorted(
        path for path in inbox.iterdir() if path.is_file() and path.suffix.lower() in RECEIPT_SUFFIXES
    )
    if not sources:
        return {
            "status": "success",
            "inbox": str(inbox),
            "counts": {},
            "items": [],
            "inbox_invoices": inbox_invoices,
            "inbox_invoice_summary": inbox_invoices.get("summary", inbox_invoice_summary([])),
        }
    records = state.setdefault("receipt_records", {})
    prepared: list[tuple[Path, str, dict[str, Any]]] = []
    ocr_texts: dict[str, str] = {}
    items: list[dict[str, Any]] = []
    with tempfile.TemporaryDirectory(prefix="receipt-xparse-") as temp_name:
        temp_dir = Path(temp_name)
        for source in sources:
            digest = sha256_file(source)
            ocr, cached = load_or_run_receipt_ocr(source, retry, temp_dir)
            record = records.setdefault(digest, {})
            record.update({"source_path": str(source), "original_filename": source.name, "ocr_cached": cached})
            if ocr.get("status") != "success":
                record.update(
                    {
                        "match_status": "识别失败",
                        "match_detail": str(ocr.get("detail", "XParse 识别失败")),
                        "processed_at": dt.datetime.now().astimezone().replace(microsecond=0).isoformat(),
                    }
                )
                items.append(
                    _receipt_result("incomplete", source, record["match_detail"], str(ocr.get("code", "XPARSE_RUN_FAILED")))
                )
                continue
            ocr_text = str(ocr.get("text", ""))
            fields = extract_receipt_fields(ocr_text)
            record["fields"] = fields
            ocr_texts[digest] = ocr_text
            prepared.append((source, digest, fields))

    candidate_map: dict[str, list[dict[str, Any]]] = {}
    failure_reasons: dict[str, list[str]] = {}
    for source, digest, fields in prepared:
        candidates: list[dict[str, Any]] = []
        reasons: list[str] = []
        for invoice in invoices:
            matched, reason = receipt_candidate(invoice, fields, ocr_texts.get(digest, ""))
            if matched:
                candidates.append(invoice)
            else:
                reasons.append(reason)
        candidate_map[digest] = candidates
        failure_reasons[digest] = reasons

    invoice_receipts: dict[str, list[tuple[Path, str, dict[str, Any]]]] = {}
    for entry in prepared:
        candidates = candidate_map[entry[1]]
        if len(candidates) == 1:
            invoice_receipts.setdefault(candidates[0]["record_key"], []).append(entry)

    selected: dict[str, tuple[Path, str, dict[str, Any]]] = {}
    duplicates: set[str] = set()
    conflicts: set[str] = set()
    for invoice_key, entries in invoice_receipts.items():
        if len(entries) == 1:
            selected[invoice_key] = entries[0]
            continue
        fingerprints = {receipt_transaction_fingerprint(entry[2]) for entry in entries}
        if len(fingerprints) == 1:
            choice = max(entries, key=lambda entry: (entry[0].stat().st_size, entry[0].name))
            selected[invoice_key] = choice
            duplicates.update(entry[1] for entry in entries if entry[1] != choice[1])
        else:
            conflicts.update(entry[1] for entry in entries)

    invoices_by_key = {invoice["record_key"]: invoice for invoice in invoices}
    for source, digest, fields in prepared:
        record = records[digest]
        candidates = candidate_map[digest]
        now = dt.datetime.now().astimezone().replace(microsecond=0).isoformat()
        if digest in duplicates:
            record.update({"match_status": "重复凭证", "match_detail": "同一交易已有字段更完整的凭证", "processed_at": now})
            items.append(_receipt_result("skipped", source, record["match_detail"], "RECEIPT_DUPLICATE"))
            continue
        if digest in conflicts or len(candidates) > 1:
            record.update({"match_status": "匹配冲突", "match_detail": "存在多个同等可信候选", "processed_at": now})
            items.append(_receipt_result("incomplete", source, record["match_detail"], "RECEIPT_MATCH_CONFLICT"))
            continue
        if not candidates:
            reason_order = [
                "存在退款、撤销或冲正", "交易未成功", "交易状态缺失", "收款方全称缺失",
                "凭证包含多笔交易", "优惠差额无法解释", "优惠明细与差额不一致",
                "支付金额缺失", "支付日期超出允许范围", "收款方与开票方不一致",
                "优惠前订单金额与发票金额不一致", "实付金额与发票金额不一致",
                "实付金额高于订单金额", "实付金额高于发票金额",
                "发票已有报销凭证",
            ]
            all_reasons = failure_reasons.get(digest, [])
            detail = next((reason for reason in reason_order if reason in all_reasons), "没有唯一匹配的有效发票")
            record.update({"match_status": "未匹配", "match_detail": detail, "processed_at": now})
            items.append(_receipt_result("incomplete", source, detail, "RECEIPT_NOT_MATCHED"))
            continue
        invoice = candidates[0]
        if selected.get(invoice["record_key"], (None, "", {}))[1] != digest:
            record.update({"match_status": "匹配冲突", "match_detail": "同一发票存在多份不同交易凭证", "processed_at": now})
            items.append(_receipt_result("incomplete", source, record["match_detail"], "RECEIPT_MATCH_CONFLICT"))
            continue
        destination = receipt_destination(invoice, source, digest)
        if destination.exists() and sha256_file(destination) == digest:
            record.update({"match_status": "已匹配", "invoice_key": invoice["record_key"], "archive_path": str(destination)})
            items.append(_receipt_result("skipped", source, "相同凭证已归档"))
            continue
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(source), str(destination))
        record.update(
            {
                "match_status": "已匹配",
                "match_detail": receipt_match_detail(
                    fields,
                    invoice["invoice_amount"],
                    str(invoice["seller"]),
                    ocr_texts.get(digest, ""),
                ),
                "invoice_key": invoice["record_key"],
                "archive_path": str(destination),
                "payment_date": str(fields.get("payment_time", ""))[:10],
                "amount_difference": receipt_amount_difference(invoice["invoice_amount"], fields),
                "reimbursement_amount": invoice["invoice_amount"],
                "processed_at": now,
            }
        )
        items.append(_receipt_result("success", destination, record["match_detail"]))
    counts = {status: sum(item["status"] == status for item in items) for status in ("success", "skipped", "incomplete", "error")}
    return {
        "status": "success",
        "inbox": str(inbox),
        "counts": counts,
        "items": items,
        "inbox_invoices": inbox_invoices,
        "inbox_invoice_summary": inbox_invoices.get("summary", inbox_invoice_summary([])),
    }


def archive_invoice(
    source: Path,
    suffix: str,
    root: Path,
    state: dict[str, Any],
    text: str | None = None,
    provenance: dict[str, Any] | None = None,
    extra_reasons: Iterable[str] | None = None,
) -> tuple[str, str, list[str]]:
    if text is None:
        text = pdf_text(source) if suffix == ".pdf" else ofd_text(source)
    fields = extract_invoice_fields(text)
    review_reasons = [key for key in ("date", "seller", "amount") if not fields[key]]
    review_reasons.extend(buyer_validation_issues(text))
    review_reasons.extend(extra_reasons or [])
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
    pattern = re.compile(rb"^\((.*?)\)\s+(\"[^\"]*\"|NIL)\s+(.+)$")
    for line in lines or []:
        if not isinstance(line, bytes) or not (match := pattern.match(line)):
            continue
        flags = set(match.group(1).split())
        delimiter_raw = match.group(2)
        delimiter = "" if delimiter_raw == b"NIL" else delimiter_raw.strip(b'"').decode("ascii", errors="replace")
        raw_name = match.group(3).strip()
        if raw_name.startswith(b'"') and raw_name.endswith(b'"'):
            raw_name = raw_name[1:-1].replace(b'\\"', b'"').replace(b"\\\\", b"\\")
        wire_name = raw_name.decode("ascii", errors="replace")
        display_name = decode_modified_utf7(wire_name)
        if folder_is_excluded(display_name, flags, delimiter):
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
    criteria: list[str] = ["SINCE", imap_date(from_date)]
    if to_date:
        criteria.extend(["BEFORE", imap_date(to_date + dt.timedelta(days=1))])
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
    inbox_drop: bool = False,
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
                    inbox_drop,
                )
            )
        return results
    extraction_warning = ""
    extra_reasons: list[str] = []
    try:
        text = pdf_text(path) if suffix == ".pdf" else ofd_text(path)
    except Exception as exc:
        if not trusted_invoice_source and not inbox_drop:
            return [("incomplete", hinted_name, str(exc))]
        text = ""
        extra_reasons.append("文字提取失败")
        extraction_warning = f"；文字提取失败：{error_code(exc, 'PDF_TEXT_EXTRACTION_FAILED')}"
    trusted_context = str((provenance or {}).get("subject", "")) if trusted_invoice_source else ""
    invoice_text = f"{text}\n{trusted_context}".strip()
    recognized = is_invoice_document(invoice_text, hinted_name) or trusted_invoice_source
    if not recognized:
        if inbox_drop and suffix == ".ofd":
            extra_reasons.append("文字提取失败" if not invoice_text else "未能识别为发票")
        elif inbox_drop and not invoice_text:
            extra_reasons.append("文字提取失败")
        else:
            return [("skipped", hinted_name, "附件内容不是可识别的发票，未归档")]
    invoice_key = invoice_identity(invoice_text, hinted_name)
    if suffix == ".ofd" and invoice_key:
        existing_pdf = recorded_format_path(state, invoice_key, ".pdf")
        if invoice_key in seen_pdf_keys or existing_pdf is not None:
            detail = str(existing_pdf) if existing_pdf is not None else "本批次已处理同一发票的 PDF"
            return [("skipped", hinted_name, f"同一发票已有 PDF，未保留 OFD；{detail}")]
    status, destination, review_reasons = archive_invoice(
        path, suffix, root, state, invoice_text, provenance, extra_reasons
    )
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
                    if folder_state.get("initialized", False):
                        emit_progress(
                            "folder_start_from_now_skipped",
                            account=email,
                            folder=display_name,
                            last_uid=last_uid,
                        )
                        output.append(
                            report_item(
                                "skipped",
                                f"文件夹:{display_name}",
                                "",
                                email,
                                "已初始化文件夹未因 --start-from-now 重置游标或待重试队列",
                                "START_FROM_NOW_SKIPPED_INITIALIZED",
                            )
                        )
                        continue
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
                    # 本次只处理日期查询结果，不把历史 pending 混入候选。
                    # 历史 pending 先原样保留：范围内成功的会 discard，失败的会加回。
                    pending = set(previous_pending)
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
        opened = webbrowser.open(url)
        if not opened:
            raise SkillError(
                "CONFIGURATION_BROWSER_UNAVAILABLE",
                "无法打开本机浏览器配置页；请改用交互式 configure",
            )
        print(
            json.dumps(
                {"status": "waiting_for_user", "email": email, "bind": "127.0.0.1", "browser_opened": True},
                ensure_ascii=False,
            ),
            flush=True,
        )
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
                "required_buyer_name": config.get("required_buyer_name", REQUIRED_BUYER_NAME),
                "required_buyer_tax_id": config.get("required_buyer_tax_id", REQUIRED_BUYER_TAX_ID),
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
    try:
        node_path, artifact_path = find_excel_runtime()
        excel_runtime = {
            "available": True,
            "backend": "artifact-tool",
            "node": str(node_path),
            "artifact_tool": str(artifact_path),
        }
    except SkillError as exc:
        if openpyxl_available():
            excel_runtime = {"available": True, "backend": "openpyxl"}
        else:
            excel_runtime = {"available": False, "code": exc.code, "detail": str(exc)}
    try:
        xparse_path = find_xparse_cli()
        xparse_cli = {"available": True, "path": str(xparse_path), "mode": "free"}
    except SkillError as exc:
        xparse_cli = {"available": False, "code": exc.code, "detail": str(exc)}
    result = {
        "ok": sys.version_info >= (3, 10) and ca_path.is_file(),
        "python": sys.version.split()[0],
        "platform": sys.platform,
        "ca_bundle": str(ca_path),
        "ca_bundle_exists": ca_path.is_file(),
        "keyring_backend": type(keyring.get_keyring()).__name__,
        "dependencies": {"pypdf": pypdf.__version__, "pdfplumber": pdfplumber.__version__},
        "excel_runtime": excel_runtime,
        "receipt_ocr": xparse_cli,
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


def cmd_set_buyer(args: argparse.Namespace) -> int:
    if not args.confirm:
        raise SkillError("CONFIRMATION_REQUIRED", "修改购买方校验需要用户确认后添加 --confirm")
    name = str(args.name).strip()
    tax_id = re.sub(r"[^0-9A-Z]", "", str(args.tax_id).upper())
    if len(name) < 2:
        raise ValueError("购买方名称至少 2 个字符")
    if not re.fullmatch(r"[0-9A-Z]{15,20}", tax_id):
        raise ValueError("购买方纳税人识别号必须是 15 到 20 位数字或大写字母")
    config = load_config()
    config["required_buyer_name"] = name
    config["required_buyer_tax_id"] = tax_id
    atomic_json_write(config_path(), config)
    print(json.dumps({"required_buyer_name": name, "required_buyer_tax_id": tax_id}, ensure_ascii=False, indent=2))
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
            "description": "仅初始化尚未建立游标的文件夹；已初始化文件夹的 last_uid 与 pending_uids 保持不变",
        }
    if explicit_range:
        return {
            "mode": "date_range_initial" if initializing else "date_range_rescan",
            "from": str(from_date),
            "to": str(to_date),
            "last_uid_advanced": initializing,
            "pending_uids_updated": True,
            "pending_scope": "merge_outside_range_with_current_failures",
        }
    return {
        "mode": "uid_incremental",
        "description": "扫描 last_uid 之后及 pending_uids 中的邮件，不使用隐含日期窗口",
    }


def rebuild_file_index(root: Path, state: dict[str, Any]) -> None:
    files = state.setdefault("files", {})
    previously_recorded = {Path(value) for value in files.values()}
    receipt_paths = {
        Path(value).resolve()
        for record in state.setdefault("receipt_records", {}).values()
        for value in (record.get("source_path"), record.get("archive_path"))
        if value and Path(value).exists()
    }
    if not root.exists():
        return
    for path in root.rglob("*"):
        if path.is_file() and path.suffix.lower() in {".pdf", ".ofd"}:
            try:
                relative_parts = path.resolve().relative_to(root.resolve()).parts
            except (OSError, ValueError):
                continue
            if RECEIPT_INBOX_NAME in relative_parts or path.resolve() in receipt_paths or "_报销凭证" in path.stem:
                continue
            digest = sha256_file(path)
            recorded = files.get(digest)
            if not recorded and path.suffix.lower() == ".pdf" and not archived_filename_fields(path.name)["date"]:
                continue
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


def archived_filename_fields(filename: str) -> dict[str, str]:
    """读取本技能生成的“日期_开票方_¥金额”归档名，仅用于回填缺失自动字段。"""
    stem = Path(filename).stem
    match = re.match(
        r"^(20\d{2}-\d{2}-\d{2}|未知日期)_(.+)_¥(\d+(?:\.\d{1,2})?|未知金额)(?:_[0-9a-f]{8})?$",
        stem,
        re.I,
    )
    if not match:
        return {"date": "", "seller": "", "amount": ""}
    date_value, seller, amount = match.groups()
    return {
        "date": "" if date_value == "未知日期" else date_value,
        "seller": "" if seller == "未知销售方" else seller,
        "amount": "" if amount == "未知金额" else amount,
    }


def invoice_validation_status(review_reasons: Iterable[str]) -> str:
    labels = {
        "date": "发票日期缺失",
        "seller": "开票方缺失",
        "amount": "发票金额缺失",
        "文字提取失败": "文字提取失败",
    }
    reasons = [labels.get(value, value) for value in review_reasons]
    return "通过" if not reasons else "待确认：" + "、".join(dict.fromkeys(reasons))


def rebuild_invoice_records(root: Path, state: dict[str, Any]) -> list[dict[str, Any]]:
    """从已登记且仍存在的归档文件重建 Excel 自动字段，保留首次下载时间。"""
    previous = state.setdefault("invoice_records", {})
    current: dict[str, dict[str, Any]] = {}
    resolved_root = root.resolve()
    for digest, value in sorted(state.setdefault("files", {}).items()):
        path = Path(value)
        if not path.exists() or path.suffix.lower() not in {".pdf", ".ofd"}:
            continue
        try:
            path.resolve().relative_to(resolved_root)
        except (OSError, ValueError):
            continue
        provenance = state.setdefault("provenance", {}).get(digest, {})
        hint = provenance_filename(provenance, path.name)
        extraction_failed = False
        try:
            text = pdf_text(path) if path.suffix.lower() == ".pdf" else ofd_text(path)
        except Exception:
            text = ""
            extraction_failed = True
        fields = extract_invoice_fields(text)
        archived_fields = archived_filename_fields(path.name)
        for field_name in ("date", "seller", "amount"):
            if not fields[field_name] and archived_fields[field_name]:
                fields[field_name] = archived_fields[field_name]
        invoice_number = extract_invoice_number(text, hint)
        record_key = f"number:{invoice_number}" if invoice_number else f"sha256:{digest}"
        review_reasons = [key for key in ("date", "seller", "amount") if not fields[key]]
        if extraction_failed:
            review_reasons.append("文字提取失败")
        review_reasons.extend(buyer_validation_issues(text))
        old_record = previous.get(record_key, {})
        downloaded_at = old_record.get("downloaded_at")
        if not downloaded_at:
            downloaded_at = dt.datetime.fromtimestamp(path.stat().st_mtime).astimezone().replace(microsecond=0).isoformat()
        record = {
            "record_key": record_key,
            "invoice_date": fields["date"],
            "invoice_amount": fields["amount"],
            "seller": fields["seller"],
            "invoice_number": invoice_number,
            "downloaded_at": downloaded_at,
            "validation_status": invoice_validation_status(review_reasons),
            "archive_path": str(path),
            "has_receipt": "否",
            "receipt_path": "",
            "receipt_name": "",
            "receipt_status": "未匹配",
            "receipt_detail": "尚未找到对应报销凭证",
            "receipt_payment_date": "",
            "receipt_amount_difference": "",
        }
        existing = current.get(record_key)
        if existing and Path(existing["archive_path"]).suffix.lower() == ".pdf":
            continue
        current[record_key] = record
    for receipt in state.setdefault("receipt_records", {}).values():
        invoice_key = str(receipt.get("invoice_key", ""))
        invoice = current.get(invoice_key)
        receipt_value = receipt.get("archive_path")
        if not invoice or not receipt_value:
            continue
        receipt_path = Path(str(receipt_value))
        if not receipt_path.exists():
            continue
        try:
            relative = receipt_path.resolve().relative_to(resolved_root)
        except (OSError, ValueError):
            continue
        if RECEIPT_INBOX_NAME in relative.parts:
            continue
        invoice.update(
            {
                "has_receipt": "是",
                "receipt_path": relative.as_posix(),
                "receipt_name": receipt_path.name,
                "receipt_status": receipt.get("match_status", "待确认"),
                "receipt_detail": receipt.get("match_detail", "凭证文件已存在"),
                "receipt_payment_date": receipt.get("payment_date", ""),
                "receipt_amount_difference": receipt.get("amount_difference", ""),
            }
        )
    state["invoice_records"] = current
    return list(current.values())


def find_excel_runtime() -> tuple[Path, Path]:
    """定位 Codex 随附的 Node.js 与 @oai/artifact-tool。"""
    dependency_root = (
        Path.home()
        / ".cache"
        / "codex-runtimes"
        / "codex-primary-runtime"
        / "dependencies"
        / "node"
    )
    node_candidates = [
        Path(os.environ["INVOICE_NODE_PATH"]) if os.environ.get("INVOICE_NODE_PATH") else None,
        dependency_root / "bin" / ("node.exe" if os.name == "nt" else "node"),
        Path(shutil.which("node")) if shutil.which("node") else None,
    ]
    artifact_candidates = [
        Path(os.environ["INVOICE_ARTIFACT_TOOL_ENTRY"])
        if os.environ.get("INVOICE_ARTIFACT_TOOL_ENTRY")
        else None,
        dependency_root / "node_modules" / "@oai" / "artifact-tool" / "dist" / "artifact_tool.mjs",
    ]
    node = next((candidate for candidate in node_candidates if candidate and candidate.is_file()), None)
    artifact = next((candidate for candidate in artifact_candidates if candidate and candidate.is_file()), None)
    if not node or not artifact:
        raise SkillError(
            "EXCEL_RUNTIME_MISSING",
            "未找到 Codex 电子表格运行时；将尝试使用 openpyxl",
        )
    return node, artifact


def openpyxl_available() -> bool:
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        return False
    return True


def load_invoice_register_module() -> Any:
    import importlib.util

    path = Path(__file__).with_name("invoice_register.py")
    spec = importlib.util.spec_from_file_location("invoice_register", path)
    if spec is None or spec.loader is None:
        raise SkillError("EXCEL_RUNTIME_MISSING", "未找到 openpyxl 登记表模块")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def sync_invoice_workbook(
    root: Path,
    state: dict[str, Any],
    preview_path: Path | None = None,
) -> dict[str, Any]:
    """生成或更新发票登记表；现有人工字段由工作簿中的记录键保留。"""
    records = rebuild_invoice_records(root, state)
    output = root / INVOICE_REGISTER_FILENAME
    artifact_error: BaseException | None = None
    try:
        node, artifact = find_excel_runtime()
        helper = Path(__file__).with_name("invoice_excel.mjs")
        payload = {
            "generated_at": dt.datetime.now().astimezone().replace(microsecond=0).isoformat(),
            "records": records,
        }
        with tempfile.TemporaryDirectory(prefix="invoice-register-") as temp_name:
            payload_path = Path(temp_name) / "records.json"
            payload_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
            command = [str(node), str(helper), str(payload_path), str(output)]
            if preview_path:
                command.append(str(preview_path))
            environment = os.environ.copy()
            environment["INVOICE_ARTIFACT_TOOL_ENTRY"] = str(artifact)
            last_subprocess_error: BaseException | None = None
            for attempt in range(3):
                try:
                    completed = subprocess.run(
                        command,
                        check=False,
                        capture_output=True,
                        text=True,
                        encoding="utf-8",
                        errors="replace",
                        timeout=120,
                        env=environment,
                    )
                except (OSError, subprocess.SubprocessError) as exc:
                    last_subprocess_error = exc
                    time.sleep(2**attempt)
                    continue
                if completed.returncode == 0:
                    return {"status": "success", "path": str(output), "rows": len(records), "backend": "artifact-tool"}
                error_text = completed.stderr + completed.stdout
                if attempt < 2 and re.search(r"EPERM|EBUSY|EACCES|locked|being used", error_text, re.I):
                    time.sleep(2**attempt)
                    continue
                error_lines = [line.strip() for line in completed.stderr.splitlines() if line.strip()]
                detail = next(
                    (line for line in reversed(error_lines) if not re.match(r"^(Node\.js v|at |\^+$)", line)),
                    "未知错误",
                )
                artifact_error = SkillError("EXCEL_SYNC_FAILED", f"发票登记表同步失败：{detail[:500]}")
                break
            else:
                artifact_error = SkillError(
                    "EXCEL_SYNC_FAILED",
                    f"发票登记表同步失败：{type(last_subprocess_error).__name__ if last_subprocess_error else '未知错误'}",
                )
    except SkillError as exc:
        if exc.code != "EXCEL_RUNTIME_MISSING":
            raise
        artifact_error = exc
    if not openpyxl_available():
        raise SkillError(
            "EXCEL_RUNTIME_MISSING",
            "未找到 Excel 运行时（Codex artifact-tool 或 openpyxl）；发票已归档，安装依赖后执行 sync-excel",
        ) from artifact_error
    if preview_path:
        emit_progress("excel_preview_skipped", reason="openpyxl_backend")
    try:
        result = load_invoice_register_module().write_invoice_workbook(output, records)
    except Exception as exc:  # noqa: BLE001 - 保护人工字段
        raise SkillError("EXCEL_SYNC_FAILED", f"发票登记表同步失败：{exc}") from (artifact_error or exc)
    if artifact_error:
        result = dict(result)
        result["fallback"] = "openpyxl"
        result["artifact_error"] = error_code(artifact_error, "EXCEL_SYNC_FAILED")
    return result


def cmd_sync_excel(args: argparse.Namespace) -> int:
    config = load_config()
    apply_buyer_from_config(config)
    root = Path(config["invoice_root"]).expanduser().resolve()
    root.mkdir(parents=True, exist_ok=True)
    state = load_state()
    rebuild_file_index(root, state)
    rebuild_invoice_identity_index(root, state)
    discover_manually_placed_receipts(root, state, rebuild_invoice_records(root, state))
    preview_path = Path(args.preview).expanduser().resolve() if args.preview else None
    result = sync_invoice_workbook(root, state, preview_path)
    atomic_json_write(state_path(), state)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


def cmd_match_receipts(args: argparse.Namespace) -> int:
    config = load_config()
    apply_buyer_from_config(config)
    root = Path(config["invoice_root"]).expanduser().resolve()
    root.mkdir(parents=True, exist_ok=True)
    state = load_state()
    with exclusive_run_lock():
        rebuild_file_index(root, state)
        rebuild_invoice_identity_index(root, state)
        result = match_receipts(root, state, bool(args.retry))
        atomic_json_write(state_path(), state)
    try:
        invoice_register = sync_invoice_workbook(root, state)
    except Exception as exc:
        rebuild_invoice_records(root, state)
        invoice_register = {
            "status": "error",
            "code": error_code(exc, "EXCEL_SYNC_FAILED"),
            "path": str(root / INVOICE_REGISTER_FILENAME),
            "detail": str(exc),
        }
    atomic_json_write(state_path(), state)
    result["invoice_register"] = invoice_register
    print(json.dumps(result, ensure_ascii=False, indent=2))
    counts = result.get("counts", {})
    return 2 if invoice_register["status"] == "error" else (1 if counts.get("incomplete") else 0)


def _invoice_by_number(root: Path, state: dict[str, Any], invoice_number: str) -> dict[str, Any]:
    matches = [
        record for record in rebuild_invoice_records(root, state)
        if str(record.get("invoice_number", "")) == invoice_number.strip()
    ]
    if len(matches) != 1:
        raise SkillError("INVOICE_NOT_UNIQUE", "未找到唯一的目标发票号码")
    return matches[0]


def cmd_assign_receipt(args: argparse.Namespace) -> int:
    config = load_config()
    apply_buyer_from_config(config)
    root = Path(config["invoice_root"]).expanduser().resolve()
    source = Path(args.receipt).expanduser().resolve()
    try:
        source.relative_to(root)
    except ValueError as exc:
        raise SkillError("RECEIPT_OUTSIDE_ROOT", "人工指定的凭证必须位于发票根目录内") from exc
    if not source.is_file() or source.suffix.lower() not in RECEIPT_SUFFIXES:
        raise SkillError("RECEIPT_NOT_FOUND", "未找到可支持的报销凭证文件")
    state = load_state()
    with exclusive_run_lock():
        rebuild_file_index(root, state)
        invoice = _invoice_by_number(root, state, args.invoice_number)
        if invoice.get("has_receipt") == "是":
            raise SkillError("INVOICE_ALREADY_HAS_RECEIPT", "目标发票已经存在报销凭证")
        digest = sha256_file(source)
        destination = receipt_destination(invoice, source, digest)
        destination.parent.mkdir(parents=True, exist_ok=True)
        if source != destination:
            shutil.move(str(source), str(destination))
        previous = state.setdefault("receipt_records", {}).setdefault(digest, {})
        fields = previous.get("fields", {})
        previous.update(
            {
                "source_path": str(source),
                "archive_path": str(destination),
                "invoice_key": invoice["record_key"],
                "match_status": "人工匹配",
                "match_detail": "用户通过 Agent 人工确认",
                "payment_date": str(fields.get("payment_time", ""))[:10],
                "reimbursement_amount": invoice["invoice_amount"],
                "processed_at": dt.datetime.now().astimezone().replace(microsecond=0).isoformat(),
            }
        )
        atomic_json_write(state_path(), state)
    result = sync_invoice_workbook(root, state)
    atomic_json_write(state_path(), state)
    print(json.dumps({"status": "success", "receipt": str(destination), "invoice_register": result}, ensure_ascii=False, indent=2))
    return 0


def cmd_unmatch_receipt(args: argparse.Namespace) -> int:
    config = load_config()
    apply_buyer_from_config(config)
    root = Path(config["invoice_root"]).expanduser().resolve()
    state = load_state()
    invoice = _invoice_by_number(root, state, args.invoice_number)
    matches = [
        (digest, record) for digest, record in state.setdefault("receipt_records", {}).items()
        if record.get("invoice_key") == invoice["record_key"] and record.get("archive_path")
        and Path(str(record["archive_path"])).exists()
    ]
    if len(matches) != 1:
        raise SkillError("RECEIPT_NOT_UNIQUE", "目标发票没有唯一的已归档凭证")
    digest, record = matches[0]
    source = Path(str(record["archive_path"]))
    inbox = root / RECEIPT_INBOX_NAME
    inbox.mkdir(parents=True, exist_ok=True)
    original_name = str(record.get("original_filename") or source.name)
    destination = inbox / safe_component(Path(original_name).stem, "报销凭证", 100)
    destination = destination.with_suffix(Path(original_name).suffix.lower() or source.suffix.lower())
    if destination.exists() and sha256_file(destination) != digest:
        destination = destination.with_name(f"{destination.stem}_{digest[:8]}{destination.suffix}")
    with exclusive_run_lock():
        if source != destination:
            shutil.move(str(source), str(destination))
        record.update(
            {
                "source_path": str(destination),
                "archive_path": "",
                "invoice_key": "",
                "match_status": "已解除",
                "match_detail": "用户通过 Agent 解除关联",
                "payment_date": "",
                "processed_at": dt.datetime.now().astimezone().replace(microsecond=0).isoformat(),
            }
        )
        atomic_json_write(state_path(), state)
    result = sync_invoice_workbook(root, state)
    atomic_json_write(state_path(), state)
    print(json.dumps({"status": "success", "receipt": str(destination), "invoice_register": result}, ensure_ascii=False, indent=2))
    return 0


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
    apply_buyer_from_config(config)
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
    try:
        receipt_matching = match_receipts(root, state)
    except Exception as exc:
        receipt_matching = {
            "status": "error",
            "code": error_code(exc, "RECEIPT_MATCH_FAILED"),
            "inbox": str(root / RECEIPT_INBOX_NAME),
            "detail": str(exc),
            "counts": {"error": 1},
            "items": [],
        }
    try:
        invoice_register = sync_invoice_workbook(root, state)
    except Exception as exc:
        rebuild_invoice_records(root, state)
        invoice_register = {
            "status": "error",
            "code": error_code(exc, "EXCEL_SYNC_FAILED"),
            "path": str(root / INVOICE_REGISTER_FILENAME),
            "detail": str(exc),
        }
    atomic_json_write(state_path(), state)
    counts = {status: sum(item["status"] == status for item in results) for status in ("success", "skipped", "incomplete", "error")}
    print(
        json.dumps(
            {
                "invoice_root": str(root),
                "scan": scan_metadata(explicit_range, from_date, to_date, start_from_now, bool(unconfirmed)),
                "counts": counts,
                "inbox_invoice_summary": receipt_matching.get("inbox_invoice_summary", inbox_invoice_summary([])),
                "invoice_register": invoice_register,
                "receipt_matching": receipt_matching,
                "items": results,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    receipt_counts = receipt_matching.get("counts", {})
    return 2 if counts["error"] or invoice_register["status"] == "error" else (
        1 if counts["incomplete"] or receipt_counts.get("incomplete") or receipt_counts.get("error") else 0
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="从 163/QQ 邮箱下载并整理电子发票")
    public_commands = (
        "configure,configure-ui,accounts,preflight,set-root,set-buyer,trusted-domains,trust-domain,"
        "untrust-domain,enable,disable,remove-account,repair-state,sync-excel,match-receipts,run"
    )
    subparsers = parser.add_subparsers(dest="command", required=True, metavar=f"{{{public_commands}}}")

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

    set_buyer = subparsers.add_parser("set-buyer", help="设置普通发票购买方名称和税号")
    set_buyer.add_argument("--name", required=True, help="购买方名称")
    set_buyer.add_argument("--tax-id", required=True, help="购买方纳税人识别号")
    set_buyer.add_argument("--confirm", action="store_true", help="确认修改购买方校验")
    set_buyer.set_defaults(handler=cmd_set_buyer)

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

    sync_excel = subparsers.add_parser("sync-excel", help="从已归档发票补录或更新发票登记表")
    sync_excel.add_argument("--preview", help="可选：同时把登记表渲染为指定 PNG，供检查使用")
    sync_excel.set_defaults(handler=cmd_sync_excel)

    match_receipts_parser = subparsers.add_parser(
        "match-receipts", help="使用 TextIn XParse 免费模式识别并匹配报销凭证"
    )
    match_receipts_parser.add_argument(
        "--retry", action="store_true", help="忽略本地识别缓存，重新上传投放目录中的凭证"
    )
    match_receipts_parser.set_defaults(handler=cmd_match_receipts)

    assign_receipt = subparsers.add_parser("assign-receipt", help=argparse.SUPPRESS)
    assign_receipt.add_argument("--receipt", required=True)
    assign_receipt.add_argument("--invoice-number", required=True)
    assign_receipt.set_defaults(handler=cmd_assign_receipt)

    unmatch_receipt = subparsers.add_parser("unmatch-receipt", help=argparse.SUPPRESS)
    unmatch_receipt.add_argument("--invoice-number", required=True)
    unmatch_receipt.set_defaults(handler=cmd_unmatch_receipt)
    subparsers._choices_actions = [
        action for action in subparsers._choices_actions
        if action.dest not in {"assign-receipt", "unmatch-receipt"}
    ]

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

from __future__ import annotations

import importlib.util
from email.message import EmailMessage
from pathlib import Path
import tempfile
import unittest
from unittest import mock
import zipfile


MODULE_PATH = Path(__file__).resolve().parents[1] / "scripts" / "invoice_mail.py"
SPEC = importlib.util.spec_from_file_location("invoice_mail", MODULE_PATH)
assert SPEC and SPEC.loader
invoice_mail = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(invoice_mail)


INVOICE_TEXT = (
    "电子发票 发票号码：12345678901234567890 开票日期：2026年09月02日 "
    "销售方名称：测试公司 统一社会信用代码 1 价税合计（小写）：￥128.00"
)


def ofd_bytes(text: str = INVOICE_TEXT) -> bytes:
    from io import BytesIO

    output = BytesIO()
    with zipfile.ZipFile(output, "w") as archive:
        archive.writestr("OFD.xml", "<OFD></OFD>")
        archive.writestr("Doc_0/Pages/Page_0/Content.xml", f"<TextCode>{text}</TextCode>")
    return output.getvalue()


class InvoiceMailTests(unittest.TestCase):
    def test_extract_fields(self) -> None:
        text = "开票日期：2026年09月02日 销售方名称：上海某科技公司 统一社会信用代码 123 价税合计（小写）：￥128.00"
        self.assertEqual(
            invoice_mail.extract_invoice_fields(text),
            {"date": "2026-09-02", "seller": "上海某科技公司", "amount": "128.00"},
        )

    def test_extract_fields_from_layout_order(self) -> None:
        text = """电子发票（普通发票）
开票日期：
名称：购买方公司                 名称：销售方公司
价税合计（大写） 壹佰贰拾捌圆整 （小写）￥128.00
2026年09月02日
"""
        self.assertEqual(
            invoice_mail.extract_invoice_fields(text),
            {"date": "2026-09-02", "seller": "销售方公司", "amount": "128.00"},
        )

    def test_railway_invoice_fields(self) -> None:
        text = "电子发票（铁路电子客票） 开票日期:2026年09月01日 退票费: ￥26.00"
        self.assertEqual(
            invoice_mail.extract_invoice_fields(text),
            {"date": "2026-09-01", "seller": "中国铁路", "amount": "26.00"},
        )

    def test_bank_transfer_form_is_not_invoice(self) -> None:
        text = "境外汇款申请书 APPLICATION FOR FUNDS TRANSFERS 本笔款项是否为报税货物项下付款 发票号"
        self.assertFalse(invoice_mail.is_invoice_document(text, "境外汇款申请书.pdf"))

    def test_invoice_identity_uses_invoice_number_not_bare_date(self) -> None:
        self.assertEqual(invoice_mail.invoice_identity(INVOICE_TEXT), "number:12345678901234567890")
        self.assertEqual(invoice_mail.invoice_identity("电子发票 开票日期：2026-09-02", "20260902.pdf"), "")

    def test_candidate_links_require_context(self) -> None:
        html = '<a href="https://example.com/a">退订</a><a href="https://example.com/b">下载发票</a>'
        self.assertEqual(invoice_mail.extract_candidate_links("", html), ["https://example.com/b"])

    def test_redact_url(self) -> None:
        value = invoice_mail.redact_url("https://example.com/download/a.pdf?token=secret#part")
        self.assertEqual(value, "https://example.com/download/a.pdf")

    def test_ofd_text(self) -> None:
        with tempfile.TemporaryDirectory() as temp_name:
            path = Path(temp_name) / "sample.ofd"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("OFD.xml", "<OFD></OFD>")
                archive.writestr("Doc_0/Pages/Page_0/Content.xml", '<TextCode>开票日期：2026-09-02</TextCode>')
            self.assertIn("2026-09-02", invoice_mail.ofd_text(path))
            self.assertEqual(invoice_mail.detect_file_type(path), ".ofd")

    def test_zip_rejects_path_traversal(self) -> None:
        with tempfile.TemporaryDirectory() as temp_name:
            path = Path(temp_name) / "bad.zip"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("../invoice.pdf", b"%PDF-1.4")
            with self.assertRaisesRegex(ValueError, "路径穿越"):
                invoice_mail.safe_zip_members(path)

    def test_zip_ignores_nested_zip_and_other_files(self) -> None:
        with tempfile.TemporaryDirectory() as temp_name:
            path = Path(temp_name) / "mixed.zip"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("invoice.pdf", b"%PDF-1.4")
                archive.writestr("nested.zip", b"PK\x03\x04")
                archive.writestr("notes.txt", b"ignore")
            selected = invoice_mail.safe_zip_members(path)
            self.assertEqual([item.filename for item in selected], ["invoice.pdf"])

    def test_process_file_rejects_oversized_input(self) -> None:
        with tempfile.TemporaryDirectory() as temp_name:
            root = Path(temp_name)
            path = root / "large.pdf"
            with path.open("wb") as handle:
                handle.seek(invoice_mail.MAX_FILE_BYTES)
                handle.write(b"x")
            result = invoice_mail.process_file(path, path.name, root / "out", {"files": {}}, root)
            self.assertEqual(result[0][0], "incomplete")
            self.assertIn("50 MB", result[0][2])

    def test_same_invoice_prefers_pdf_over_ofd(self) -> None:
        with tempfile.TemporaryDirectory() as temp_name:
            root = Path(temp_name)
            pdf = root / "invoice.pdf"
            ofd = root / "invoice.ofd"
            pdf.write_bytes(b"%PDF-1.4\n%%EOF")
            ofd.write_bytes(ofd_bytes())
            state: dict[str, object] = {"files": {}, "provenance": {}, "invoice_formats": {}}
            seen: set[str] = set()
            with (
                mock.patch.object(invoice_mail, "pdf_text", return_value=INVOICE_TEXT),
                mock.patch.object(invoice_mail, "ofd_text", return_value=INVOICE_TEXT),
            ):
                pdf_result = invoice_mail.process_file(pdf, pdf.name, root / "out", state, root, seen_pdf_keys=seen)
                ofd_result = invoice_mail.process_file(ofd, ofd.name, root / "out", state, root, seen_pdf_keys=seen)
            self.assertEqual(pdf_result[0][0], "success")
            self.assertEqual(ofd_result[0][0], "skipped")
            self.assertIn("已有 PDF", ofd_result[0][2])
            self.assertEqual([path.suffix for path in (root / "out").rglob("*") if path.is_file()], [".pdf"])

    def test_later_pdf_removes_recorded_ofd(self) -> None:
        with tempfile.TemporaryDirectory() as temp_name:
            root = Path(temp_name)
            pdf = root / "invoice.pdf"
            ofd = root / "invoice.ofd"
            pdf.write_bytes(b"%PDF-1.4\n%%EOF")
            ofd.write_bytes(ofd_bytes())
            state: dict[str, object] = {"files": {}, "provenance": {}, "invoice_formats": {}}
            with (
                mock.patch.object(invoice_mail, "pdf_text", return_value=INVOICE_TEXT),
                mock.patch.object(invoice_mail, "ofd_text", return_value=INVOICE_TEXT),
            ):
                ofd_result = invoice_mail.process_file(ofd, ofd.name, root / "out", state, root)
                pdf_result = invoice_mail.process_file(pdf, pdf.name, root / "out", state, root)
            self.assertEqual(ofd_result[0][0], "success")
            self.assertEqual(pdf_result[0][0], "success")
            self.assertIn("已移除同一发票的 OFD", pdf_result[0][2])
            self.assertEqual([path.suffix for path in (root / "out").rglob("*") if path.is_file()], [".pdf"])

    def test_zip_with_same_invoice_keeps_only_pdf(self) -> None:
        with tempfile.TemporaryDirectory() as temp_name:
            root = Path(temp_name)
            package = root / "invoices.zip"
            with zipfile.ZipFile(package, "w") as archive:
                archive.writestr("same.ofd", ofd_bytes())
                archive.writestr("same.pdf", b"%PDF-1.4\n%%EOF")
            state: dict[str, object] = {"files": {}, "provenance": {}, "invoice_formats": {}}
            with (
                mock.patch.object(invoice_mail, "pdf_text", return_value=INVOICE_TEXT),
                mock.patch.object(invoice_mail, "ofd_text", return_value=INVOICE_TEXT),
            ):
                results = invoice_mail.process_file(package, package.name, root / "out", state, root)
            self.assertEqual([result[0] for result in results], ["success", "skipped"])
            self.assertEqual([path.suffix for path in (root / "out").rglob("*") if path.is_file()], [".pdf"])

    def test_old_recorded_formats_are_migrated_with_pdf_priority(self) -> None:
        with tempfile.TemporaryDirectory() as temp_name:
            root = Path(temp_name)
            pdf = root / "old.pdf"
            ofd = root / "old.ofd"
            pdf.write_bytes(b"%PDF-1.4\n%%EOF")
            ofd.write_bytes(ofd_bytes())
            state = {
                "files": {
                    invoice_mail.sha256_file(pdf): str(pdf),
                    invoice_mail.sha256_file(ofd): str(ofd),
                },
                "provenance": {},
                "invoice_formats": {},
                "invoice_format_index_built": False,
            }
            with (
                mock.patch.object(invoice_mail, "pdf_text", return_value=INVOICE_TEXT),
                mock.patch.object(invoice_mail, "ofd_text", return_value=INVOICE_TEXT),
            ):
                invoice_mail.rebuild_file_index(root, state)
            self.assertTrue(state["invoice_format_index_built"])
            self.assertTrue(pdf.exists())
            self.assertFalse(ofd.exists())
            self.assertEqual(state["invoice_formats"]["number:12345678901234567890"], {"pdf": str(pdf)})

    def test_message_attachment_is_archived_for_confirmation(self) -> None:
        message = EmailMessage()
        message["Subject"] = "您的电子发票"
        message["From"] = "billing@example.com"
        message.set_content("发票见附件")
        message.add_attachment(b"%PDF-1.4\n%%EOF", maintype="application", subtype="pdf", filename="invoice.pdf")
        with tempfile.TemporaryDirectory() as temp_name:
            root = Path(temp_name)
            state = {"files": {}}
            with mock.patch.object(
                invoice_mail,
                "pdf_text",
                return_value="电子发票 开票日期：2026年09月02日 销售方名称：测试公司 统一社会信用代码 1 价税合计（小写）：￥1.00",
            ):
                items, incomplete = invoice_mail.process_message(
                    message.as_bytes(),
                    root / "Invoices",
                    state,
                    root,
                    provenance_base={"account": "user@qq.com", "folder": "INBOX", "uid": 12},
                )
            self.assertFalse(incomplete)
            self.assertEqual(items[0]["status"], "success")
            self.assertIn("2026/09-02", items[0]["detail"])
            provenance = next(iter(state["provenance"].values()))
            self.assertEqual(provenance["uid"], 12)
            self.assertEqual(provenance["original_filename"], "invoice.pdf")

    def test_candidate_without_download_is_incomplete(self) -> None:
        message = EmailMessage()
        message["Subject"] = "电子发票通知"
        message["From"] = "billing@example.com"
        message.set_content("本邮件不包含附件或链接")
        with tempfile.TemporaryDirectory() as temp_name:
            root = Path(temp_name)
            items, incomplete = invoice_mail.process_message(message.as_bytes(), root / "Invoices", {"files": {}}, root)
            self.assertTrue(incomplete)
            self.assertIn("未发现", items[0]["detail"])

    def test_modified_utf7_folder_name_is_decoded(self) -> None:
        import base64

        encoded = base64.b64encode("发票归档".encode("utf-16-be")).decode("ascii").rstrip("=").replace("/", ",")
        self.assertEqual(invoice_mail.decode_modified_utf7(f"&{encoded}-"), "发票归档")

    def test_untrusted_domain_is_rejected_before_download(self) -> None:
        with self.assertRaises(invoice_mail.SkillError) as caught:
            invoice_mail.validate_public_https("https://evil.example/invoice.pdf", ["trusted.example"])
        self.assertEqual(caught.exception.code, "UNTRUSTED_DOMAIN")
        self.assertFalse(invoice_mail.domain_is_trusted("cdn.trusted.example", ["trusted.example"]))

    def test_configure_refuses_non_interactive_input(self) -> None:
        args = type("Args", (), {"provider": "163", "email": "user@163.com"})()
        with mock.patch.object(invoice_mail.sys.stdin, "isatty", return_value=False):
            with self.assertRaises(invoice_mail.SkillError) as caught:
                invoice_mail.cmd_configure(args)
        self.assertEqual(caught.exception.code, "INTERACTIVE_TERMINAL_REQUIRED")

    def test_configuration_page_never_contains_secret_value(self) -> None:
        page = invoice_mail.configuration_page("qq", "user@qq.com", "token")
        self.assertIn('type="password"', page)
        self.assertNotIn("客户端授权码值", page)

    def test_first_run_requires_explicit_confirmation(self) -> None:
        args = type(
            "Args",
            (),
            {"account": None, "confirm_first_run": False, "from_date": None, "to_date": None},
        )()
        config = {
            "invoice_root": "/tmp/invoices",
            "trusted_domains": [],
            "accounts": [{"email": "user@qq.com", "provider": "qq", "enabled": True, "scan_confirmed": False}],
        }
        with mock.patch.object(invoice_mail, "load_config", return_value=config):
            with self.assertRaises(invoice_mail.SkillError) as caught:
                invoice_mail.cmd_run(args)
        self.assertEqual(caught.exception.code, "FIRST_RUN_CONFIRMATION_REQUIRED")

    def test_first_run_requires_explicit_mode(self) -> None:
        args = type(
            "Args",
            (),
            {
                "account": None,
                "confirm_first_run": True,
                "from_date": None,
                "to_date": None,
                "start_from_now": False,
            },
        )()
        config = {
            "invoice_root": "/tmp/invoices",
            "trusted_domains": [],
            "accounts": [{"email": "user@qq.com", "provider": "qq", "enabled": True, "scan_confirmed": False}],
        }
        with mock.patch.object(invoice_mail, "load_config", return_value=config):
            with self.assertRaises(invoice_mail.SkillError) as caught:
                invoice_mail.cmd_run(args)
        self.assertEqual(caught.exception.code, "FIRST_RUN_MODE_REQUIRED")

    def test_invalid_date_does_not_persist_first_run_confirmation(self) -> None:
        args = type(
            "Args",
            (),
            {"account": None, "confirm_first_run": True, "from_date": "not-a-date", "to_date": None},
        )()
        config = {
            "invoice_root": "/tmp/invoices",
            "trusted_domains": [],
            "accounts": [{"email": "user@qq.com", "provider": "qq", "enabled": True, "scan_confirmed": False}],
        }
        with (
            mock.patch.object(invoice_mail, "load_config", return_value=config),
            mock.patch.object(invoice_mail, "atomic_json_write") as write_config,
        ):
            with self.assertRaisesRegex(ValueError, "YYYY-MM-DD"):
                invoice_mail.cmd_run(args)
        write_config.assert_not_called()

    def test_incremental_scan_metadata_has_no_fake_date_range(self) -> None:
        metadata = invoice_mail.scan_metadata(
            False,
            invoice_mail.dt.date(2026, 8, 3),
            None,
        )
        self.assertEqual(metadata["mode"], "uid_incremental")
        self.assertNotIn("from", metadata)
        self.assertNotIn("to", metadata)

    def test_start_from_now_metadata(self) -> None:
        metadata = invoice_mail.scan_metadata(False, invoice_mail.dt.date(2026, 9, 2), None, True)
        self.assertEqual(metadata["mode"], "incremental_from_now_initialized")

    def test_date_rescan_metadata_describes_pending_behavior(self) -> None:
        metadata = invoice_mail.scan_metadata(
            True,
            invoice_mail.dt.date(2026, 8, 1),
            invoice_mail.dt.date(2026, 8, 31),
        )
        self.assertEqual(metadata["mode"], "date_range_rescan")
        self.assertFalse(metadata["last_uid_advanced"])
        self.assertTrue(metadata["pending_uids_updated"])

    def test_163_identity_is_sent(self) -> None:
        client = mock.Mock()
        client._simple_command.return_value = ("OK", [b"accepted"])
        invoice_mail.send_imap_identity(client, "163")
        client._simple_command.assert_called_once()
        self.assertEqual(client._simple_command.call_args.args[0], "ID")

    def test_163_identity_is_sent_before_folder_access(self) -> None:
        events: list[str] = []
        client = mock.MagicMock()
        client.__enter__.return_value = client
        client.__exit__.return_value = False
        client.login.side_effect = lambda *_args: events.append("login") or ("OK", [])
        client._simple_command.side_effect = lambda *_args: events.append("id") or ("OK", [])

        def folders(_client: object) -> list[tuple[str, str]]:
            events.append("folders")
            return []

        with tempfile.TemporaryDirectory() as temp_name:
            with (
                mock.patch.object(invoice_mail, "tls_context", return_value=mock.Mock()),
                mock.patch.object(invoice_mail.imaplib, "IMAP4_SSL", return_value=client),
                mock.patch.object(invoice_mail, "folder_entries", side_effect=folders),
            ):
                invoice_mail.run_account_once(
                    {"email": "user@163.com", "provider": "163"},
                    "secret",
                    Path(temp_name),
                    {"folders": {}, "files": {}},
                    invoice_mail.dt.date(2026, 8, 1),
                    None,
                    Path(temp_name),
                    False,
                    [],
                    True,
                )
        self.assertEqual(events, ["login", "id", "folders"])

    def test_explicit_date_search_does_not_apply_uid_cursor(self) -> None:
        client = mock.Mock()
        client.uid.return_value = ("OK", [b"20 21"])
        values = invoice_mail.search_date_uids(
            client,
            invoice_mail.dt.date(2026, 8, 1),
            invoice_mail.dt.date(2026, 8, 31),
        )
        self.assertEqual(values, [20, 21])
        called = client.uid.call_args.args
        self.assertNotIn("UID", called)

    def test_incremental_search_uses_uid_without_date_window(self) -> None:
        client = mock.Mock()
        client.uid.return_value = ("OK", [b"101 102"])
        self.assertEqual(invoice_mail.search_new_uids(client, 100), [101, 102])
        self.assertEqual(client.uid.call_args.args, ("search", None, "UID", "101:*"))

    def test_uidvalidity_falls_back_to_status(self) -> None:
        client = mock.Mock()
        client.response.return_value = ("UIDVALIDITY", [None])
        client.status.return_value = ("OK", [b'"INBOX" (UIDVALIDITY 99)'])
        self.assertEqual(invoice_mail.folder_uidvalidity(client, "INBOX"), 99)

    def test_fetch_message_rejects_oversized_message_before_body(self) -> None:
        client = mock.Mock()
        client.uid.return_value = ("OK", [f"1 (RFC822.SIZE {invoice_mail.MAX_MESSAGE_BYTES + 1})".encode()])
        with self.assertRaises(invoice_mail.SkillError) as caught:
            invoice_mail.fetch_message(client, 1)
        self.assertEqual(caught.exception.code, "MESSAGE_TOO_LARGE")
        self.assertEqual(client.uid.call_count, 1)

    def test_explicit_range_does_not_advance_incremental_cursor(self) -> None:
        client = mock.MagicMock()
        client.__enter__.return_value = client
        client.__exit__.return_value = False
        client.login.return_value = ("OK", [])
        client.select.return_value = ("OK", [])
        client.response.side_effect = lambda name: (name, [b"7"] if name == "UIDVALIDITY" else [b"151"])
        state = {
            "folders": {
                "user@qq.com::INBOX": {
                    "uidvalidity": 7,
                    "initialized": True,
                    "last_uid": 100,
                    "pending_uids": [30],
                }
            },
            "files": {},
        }
        with tempfile.TemporaryDirectory() as temp_name:
            with (
                mock.patch.object(invoice_mail, "tls_context", return_value=mock.Mock()),
                mock.patch.object(invoice_mail.imaplib, "IMAP4_SSL", return_value=client),
                mock.patch.object(invoice_mail, "folder_entries", return_value=[("INBOX", "INBOX")]),
                mock.patch.object(invoice_mail, "search_date_uids", return_value=[20]),
                mock.patch.object(invoice_mail, "fetch_message", return_value=b"message"),
                mock.patch.object(invoice_mail, "process_message", return_value=([], False)),
            ):
                invoice_mail.run_account_once(
                    {"email": "user@qq.com", "provider": "qq"},
                    "secret",
                    Path(temp_name),
                    state,
                    invoice_mail.dt.date(2026, 8, 1),
                    invoice_mail.dt.date(2026, 8, 31),
                    Path(temp_name),
                    True,
                    [],
                )
        folder_state = state["folders"]["user@qq.com::INBOX"]
        self.assertEqual(folder_state["last_uid"], 100)
        self.assertTrue(folder_state["initialized"])
        self.assertEqual(folder_state["pending_uids"], [30])

    def test_uidvalidity_change_resets_cursor_before_initial_scan(self) -> None:
        client = mock.MagicMock()
        client.__enter__.return_value = client
        client.__exit__.return_value = False
        client.login.return_value = ("OK", [])
        client.select.return_value = ("OK", [])
        client.response.side_effect = lambda name: (name, [b"7"] if name == "UIDVALIDITY" else [b"51"])
        state = {
            "folders": {
                "user@qq.com::INBOX": {
                    "uidvalidity": 6,
                    "initialized": True,
                    "last_uid": 100,
                    "pending_uids": [90],
                }
            },
            "files": {},
        }
        with tempfile.TemporaryDirectory() as temp_name:
            with (
                mock.patch.object(invoice_mail, "tls_context", return_value=mock.Mock()),
                mock.patch.object(invoice_mail.imaplib, "IMAP4_SSL", return_value=client),
                mock.patch.object(invoice_mail, "folder_entries", return_value=[("INBOX", "INBOX")]),
                mock.patch.object(invoice_mail, "search_date_uids", return_value=[45]),
                mock.patch.object(invoice_mail, "fetch_message", return_value=b"message"),
                mock.patch.object(invoice_mail, "process_message", return_value=([], False)),
            ):
                invoice_mail.run_account_once(
                    {"email": "user@qq.com", "provider": "qq"},
                    "secret",
                    Path(temp_name),
                    state,
                    invoice_mail.dt.date(2026, 8, 1),
                    None,
                    Path(temp_name),
                    False,
                    [],
                    True,
                )
        folder_state = state["folders"]["user@qq.com::INBOX"]
        self.assertEqual(folder_state["uidvalidity"], 7)
        self.assertEqual(folder_state["last_uid"], 50)
        self.assertEqual(folder_state["pending_uids"], [])


if __name__ == "__main__":
    unittest.main()

"""Synthetic-only checks for local text preparation and IMAP draft safety."""
import copy
import io
import json
from pathlib import Path
import ssl
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from scripts import mail_drafts as m


def write_template(path, *, extra=""):
    ns = m.NS["w"]
    root = ET.Element("{" + ns + "}document")
    body = ET.SubElement(root, "{" + ns + "}body")
    for index, category in enumerate(m.CATEGORIES, 1):
        for text in (f"{index} 初次建联 {category}", "策略：仅供内部阅读", "收件人：{{邮箱}}",
                     "主题：Example x {{TikTok账号}}", "Hi {{邮箱前缀}},",
                     "Account @{{TikTok账号}}. " + category + extra, "Best,\nSara\nTailki"):
            paragraph = ET.SubElement(body, "{" + ns + "}p")
            for i, line in enumerate(text.split("\n")):
                run = ET.SubElement(paragraph, "{" + ns + "}r")
                if i:
                    ET.SubElement(run, "{" + ns + "}br")
                ET.SubElement(run, "{" + ns + "}t").text = line
    with ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", ET.tostring(root))


def creator(identifier="SIM-001", category="宠物剧情类", **changes):
    result = {"达人ID": identifier, "达人昵称": "模拟达人 " + identifier,
              "联系状态": "初次建联未回复", "首联内容分类": category,
              "Email": identifier.lower() + "@example.invalid",
              "账号链接": "https://www.tiktok.com/@" + identifier.lower().replace("-", "_"),
              "初次建联日期": "2026-10-20"}
    result.update(changes)
    return result


class FakeMailbox:
    account = "sara@example.invalid"

    def __init__(self):
        self.messages = {}
        self.sent_messages = {}
        self.histories = {}
        self.append_calls = 0
        self.timeout_after_append = False
        self.timeout_before_append = False
        self.next_uid = 1

    def history(self, recipient):
        return copy.deepcopy(self.histories.get(recipient, []))

    def existing_drafts(self, recipient):
        return [x["uid"] for x in self.messages.values() if recipient.lower() in x["remote"]["to"]]

    def find(self, item, *, sent=False):
        pool = self.sent_messages if sent else self.messages
        return [copy.deepcopy(x) for x in pool.values() if x["remote"]["draft_id"] == item["id"]]

    def append(self, raw):
        self.append_calls += 1
        if self.timeout_before_append:
            raise TimeoutError()
        remote = m.snapshot(raw)
        self.messages[remote["draft_id"]] = {
            "uid": str(self.next_uid), "uidvalidity": "42", "folder": "Drafts", "remote": remote}
        self.next_uid += 1
        if self.timeout_after_append:
            self.timeout_after_append = False
            raise TimeoutError()


class DraftTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.directory = Path(self.temp.name)
        self.book = self.directory / "source.xlsx"
        self.book.write_bytes(b"simulated read-only workbook")
        self.template = self.directory / "source.docx"
        write_template(self.template)
        self.state = m.State(self.directory / "private" / "state.sqlite3")
        self.rows = [creator(), creator("SIM-002", "宠物用品类")]
        self.mailbox = FakeMailbox()

    def tearDown(self):
        self.state.close()
        self.temp.cleanup()

    def prepare(self, rows=None, ids=None, batch="SIM-BATCH"):
        rows = self.rows if rows is None else rows
        ids = [r["达人ID"] for r in rows] if ids is None else ids
        workbook = {"sheets": [{"name": "达人开发主表", "rows": rows}]}
        with patch.object(m.workflow, "read_book", return_value=workbook):
            return m.prepare(self.state, self.book, self.template, ids,
                             batch_id=batch, planned_date="2026-10-20", expected_count=len(ids))

    def test_category_fill_preserves_source_signature_and_removes_internal_notes(self):
        originals = [p.read_bytes() for p in (self.book, self.template)]
        result = self.prepare()
        self.assertEqual(result["categories"], {"宠物剧情类": 1, "宠物用品类": 1})
        items = self.state.items("SIM-BATCH")
        self.assertEqual(items[0]["subject"], "Example x sim_001")
        self.assertEqual(items[0]["body"], "Hi sim-001,\n\nAccount @sim_001. 宠物剧情类\n\nBest,\nSara\nTailki")
        self.assertNotIn("策略", items[0]["body"])
        self.assertNotIn("收件人", items[0]["body"])
        self.assertEqual(originals, [p.read_bytes() for p in (self.book, self.template)])
        self.assertEqual(self.state.path.stat().st_mode & 0o777, 0o600)

    def test_explicit_selection_ignores_group_rows_and_other_creators(self):
        rows = [{"达人ID": "分组标题"}, creator(), creator("SIM-OTHER")]
        result = self.prepare(rows, ids=["SIM-001"])
        self.assertEqual(result["prepared"], 1)
        self.assertEqual(result["items"][0]["creator_id"], "SIM-001")

    def test_twenty_synthetic_rows_have_thirteen_and_seven_templates(self):
        rows = [creator(f"SIM-{i:03}", "宠物剧情类" if i < 13 else "宠物用品类") for i in range(20)]
        self.assertEqual(self.prepare(rows)["categories"], {"宠物剧情类": 13, "宠物用品类": 7})

    def test_all_three_document_categories_are_supported(self):
        templates = m.load_templates(self.template)
        item = m.fill_text(creator(category="非宠物垂类"), templates, "2026-10-20")
        self.assertIn("非宠物垂类", item["body"])

    def test_missing_fields_reply_and_history_date_are_blocked(self):
        cases = ({"Email": ""}, {"Email": "a@example.invalid,b@example.invalid"},
                 {"Email": "a@example.invalid\r\nBcc: b@example.invalid"},
                 {"首联内容分类": "未知"}, {"账号链接": "https://example.invalid/@wrong"},
                 {"账号链接": "https://tiktok.com/@first/video/123"},
                 {"联系状态": "我方第一次回复"}, {"对方第一次回信日期": "2026-10-12"},
                 {"对方第一次回复邮件原始文本": "Already replied"},
                 {"初次建联日期": "2026-10-09"})
        templates = m.load_templates(self.template)
        for changes in cases:
            with self.subTest(changes=changes), self.assertRaises(m.DraftError):
                m.fill_text(creator(**changes), templates, "2026-10-20")

    def test_unfilled_placeholder_is_never_guessed(self):
        write_template(self.template, extra=" {{missing_field}}")
        result = self.prepare()
        self.assertEqual(result["prepared"], 0)
        self.assertEqual(len(result["blocked"]), 2)

    def test_duplicate_ids_and_shared_email_are_not_silently_merged(self):
        with self.assertRaises(m.DraftError):
            self.prepare([creator(), creator()], ids=["SIM-001"])
        rows = [creator(), creator("SIM-002", Email="sim-001@example.invalid")]
        result = self.prepare(rows)
        self.assertEqual(result["prepared"], 0)
        self.assertEqual(len(result["blocked"]), 2)

    def test_prepare_is_idempotent_and_cross_batch_recreation_is_blocked(self):
        self.prepare()
        ids = [x["id"] for x in self.state.items("SIM-BATCH")]
        self.prepare()
        self.assertEqual(ids, [x["id"] for x in self.state.items("SIM-BATCH")])
        with self.assertRaises(m.DraftError):
            self.prepare(batch="SIM-ANOTHER")
        self.assertEqual(self.state.db.execute("SELECT count(*) FROM batches").fetchone()[0], 1)

    def test_workbook_rename_cannot_bypass_creator_deduplication(self):
        self.prepare()
        renamed = self.directory / "renamed.xlsx"
        renamed.write_bytes(self.book.read_bytes())
        with self.assertRaises(m.DraftError):
            m.prepare(self.state, renamed, self.template, ["SIM-001"], batch_id="SIM-NEW",
                      planned_date="2026-10-20")

    def test_source_refresh_accepts_rename_and_new_rows_preserving_existing_drafts(self):
        self.prepare()
        before = self.state.items("SIM-BATCH")
        renamed = self.directory / "renamed.xlsx"
        renamed.write_bytes(b"new rows added in source workbook")
        self.book.unlink()
        workbook = {"sheets": [{"name": "达人开发主表", "rows": self.rows + [creator("SIM-003")]}]}
        with patch.object(m.workflow, "read_book", return_value=workbook):
            result = m.refresh_source(self.state, "SIM-BATCH", renamed)
        self.assertTrue(result["source_refreshed"])
        self.assertEqual(self.state.items("SIM-BATCH"), before)
        self.assertEqual(self.state.batch("SIM-BATCH")["sources"]["book"], str(renamed.resolve()))
        m.source_check(self.state, "SIM-BATCH")
        self.assertFalse(m.refresh_source(self.state, "SIM-BATCH", renamed)["source_refreshed"])

    def test_source_refresh_rejects_changed_recipient_without_partial_updates(self):
        self.prepare()
        before = self.state.batch("SIM-BATCH")
        self.book.write_bytes(b"changed recipient in source")
        rows = copy.deepcopy(self.rows)
        rows[1]["Email"] = "changed@example.invalid"
        with patch.object(m.workflow, "read_book", return_value={"sheets": [{"name": "达人开发主表", "rows": rows}]}):
            with self.assertRaises(m.DraftError):
                m.refresh_source(self.state, "SIM-BATCH", self.book)
        self.assertEqual(self.state.batch("SIM-BATCH"), before)

    def test_source_refresh_preserves_remote_web_edits(self):
        self.prepare()
        m.save(self.state, "SIM-BATCH", self.mailbox)
        item = self.state.items("SIM-BATCH")[0]
        self.mailbox.messages[item["id"]]["remote"]["body"] = "Edited in webmail"
        m.verify(self.state, "SIM-BATCH", self.mailbox)
        before = self.state.items("SIM-BATCH")
        self.book.write_bytes(b"unrelated new rows")
        with patch.object(m.workflow, "read_book", return_value={"sheets": [{"name": "达人开发主表", "rows": self.rows}]}):
            m.refresh_source(self.state, "SIM-BATCH", self.book)
        self.assertEqual(self.state.items("SIM-BATCH"), before)

    def test_source_changes_do_not_overwrite_prepared_text_or_upload(self):
        self.prepare()
        self.book.write_bytes(b"user changed workbook")
        with self.assertRaises(m.DraftError):
            self.prepare()
        with self.assertRaises(m.DraftError):
            m.save(self.state, "SIM-BATCH", self.mailbox)
        self.assertEqual(self.mailbox.append_calls, 0)

    def test_existing_correspondence_and_unrelated_draft_are_preserved(self):
        self.prepare()
        items = self.state.items("SIM-BATCH")
        self.mailbox.histories[items[0]["recipient"]] = [{"folder": "收件箱", "uid": "17"}]
        other = dict(items[1], id="unrelated", message_id="<unrelated@example.invalid>")
        self.mailbox.append(m.make_message(other, self.mailbox.account))
        before = copy.deepcopy(self.mailbox.messages)
        result = m.save(self.state, "SIM-BATCH", self.mailbox)
        self.assertEqual(result["states"], {"skipped_history": 1, "skipped_existing_draft": 1})
        self.assertEqual(self.mailbox.messages, before)
        self.assertEqual(self.mailbox.append_calls, 1)

    def test_upload_readback_then_repeat_never_appends_twice(self):
        self.prepare()
        self.assertEqual(m.save(self.state, "SIM-BATCH", self.mailbox)["states"], {"saved": 2})
        self.assertEqual(self.mailbox.append_calls, 2)
        self.assertEqual(m.save(self.state, "SIM-BATCH", self.mailbox)["states"], {"saved": 2})
        self.assertEqual(self.mailbox.append_calls, 2)

    def test_timeout_after_acceptance_is_reconciled_without_duplicate(self):
        self.prepare()
        self.mailbox.timeout_after_append = True
        result = m.save(self.state, "SIM-BATCH", self.mailbox)
        self.assertTrue(result["halted"])
        self.assertEqual(result["states"], {"uncertain": 1, "cleared": 1})
        self.assertEqual(self.mailbox.append_calls, 1)
        self.assertEqual(m.save(self.state, "SIM-BATCH", self.mailbox)["states"], {"saved": 2})
        self.assertEqual(self.mailbox.append_calls, 2)

    def test_unknown_outcome_without_visible_copy_never_retries(self):
        self.prepare()
        self.mailbox.timeout_before_append = True
        m.save(self.state, "SIM-BATCH", self.mailbox)
        result = m.save(self.state, "SIM-BATCH", self.mailbox)
        self.assertTrue(result["halted"])
        self.assertEqual(self.mailbox.append_calls, 1)

    def test_web_edits_are_recorded_and_never_overwritten(self):
        self.prepare()
        m.save(self.state, "SIM-BATCH", self.mailbox)
        item = self.state.items("SIM-BATCH")[0]
        remote = self.mailbox.messages[item["id"]]["remote"]
        remote.update(subject="Human edited subject", body="Human edited full body")
        result = m.verify(self.state, "SIM-BATCH", self.mailbox)
        self.assertEqual(result["states"], {"edited": 1, "saved": 1})
        recorded = self.state.items("SIM-BATCH")[0]
        self.assertEqual(recorded["remote"]["body"], "Human edited full body")
        m.save(self.state, "SIM-BATCH", self.mailbox)
        self.assertEqual(self.mailbox.append_calls, 2)
        self.assertEqual(remote["subject"], "Human edited subject")

    def test_deleted_and_manually_sent_drafts_are_not_recreated(self):
        self.prepare()
        m.save(self.state, "SIM-BATCH", self.mailbox)
        first, second = self.state.items("SIM-BATCH")
        del self.mailbox.messages[first["id"]]
        self.mailbox.sent_messages[second["id"]] = self.mailbox.messages.pop(second["id"])
        result = m.verify(self.state, "SIM-BATCH", self.mailbox)
        self.assertEqual(result["states"], {"removed": 1, "sent_manually": 1})
        m.save(self.state, "SIM-BATCH", self.mailbox)
        self.assertEqual(self.mailbox.append_calls, 2)

    def test_multiple_matching_copies_halt_instead_of_modifying(self):
        self.prepare()
        m.save(self.state, "SIM-BATCH", self.mailbox)
        first = self.state.items("SIM-BATCH")[0]
        duplicate = copy.deepcopy(self.mailbox.messages[first["id"]])
        duplicate["uid"] = "999"
        self.mailbox.messages["another-copy"] = duplicate
        result = m.save(self.state, "SIM-BATCH", self.mailbox)
        self.assertTrue(result["halted"])
        self.assertEqual(self.mailbox.append_calls, 2)

    def test_account_binding_prevents_upload_to_another_mailbox(self):
        self.prepare()
        m.audit(self.state, "SIM-BATCH", self.mailbox)
        self.mailbox.account = "other@example.invalid"
        with self.assertRaises(m.DraftError):
            m.save(self.state, "SIM-BATCH", self.mailbox)

    def test_same_state_cannot_be_opened_by_two_writers(self):
        with self.assertRaises(m.DraftError):
            m.State(self.state.path)

    def test_private_state_cannot_be_misnamed_as_publishable_source(self):
        path = self.directory / "public.txt"
        with self.assertRaises(m.DraftError):
            m.State(path)
        self.assertFalse(path.exists())

    def test_source_change_during_mail_audit_stops_before_first_append(self):
        self.prepare()
        def history(recipient):
            self.book.write_bytes(b"changed while checking mailbox")
            return []
        self.mailbox.history = history
        with self.assertRaises(m.DraftError):
            m.save(self.state, "SIM-BATCH", self.mailbox)
        self.assertEqual(self.mailbox.append_calls, 0)

    def test_gui_retries_invalid_address_and_keeps_password_out_of_state(self):
        args = SimpleNamespace(account=None, gui=True, drafts_folder=None, sent_folder=None)
        with patch.object(m, "prompt", side_effect=["Sara", "sara@example.invalid", "simulated-private-value"]):
            with patch.object(m, "Mailbox") as transport, patch.object(m, "Keychain") as keychain:
                transport.return_value.drafts = {"name": "Drafts"}
                transport.return_value.sent = {"name": "Sent"}
                result = m.setup(self.state, args)
        self.assertTrue(result["configured"])
        keychain.return_value.set.assert_called_once_with("sara@example.invalid", "simulated-private-value")
        self.assertEqual(self.state.config()["account"], "sara@example.invalid")
        self.assertNotIn(b"simulated-private-value", self.state.path.read_bytes())

    def test_failed_login_does_not_save_password_or_profile(self):
        args = SimpleNamespace(account="sara@example.invalid", gui=True, drafts_folder=None, sent_folder=None)
        with patch.object(m, "prompt", return_value="simulated-private-value"):
            with patch.object(m, "Mailbox", side_effect=m.imaplib.IMAP4.error("login failed")):
                with patch.object(m, "Keychain") as keychain, self.assertRaises(m.imaplib.IMAP4.error):
                    m.setup(self.state, args)
        keychain.assert_not_called()
        self.assertEqual(self.state.config(), {})

    def test_private_preview_contains_complete_escaped_text_and_source_is_unchanged(self):
        self.prepare()
        item = self.state.items("SIM-BATCH")[0]
        self.state.update(item, body=item["body"] + "\n<script>untrusted</script>")
        before = self.book.read_bytes()
        output = Path(m.render_preview(self.state, "SIM-BATCH"))
        try:
            html = output.read_text()
            self.assertIn("sim-001@example.invalid", html)
            self.assertIn("Best,\nSara\nTailki", html)
            self.assertIn("&lt;script&gt;untrusted&lt;/script&gt;", html)
            self.assertNotIn("<script>", html)
            self.assertEqual(output.stat().st_mode & 0o777, 0o600)
            self.assertEqual(self.book.read_bytes(), before)
        finally:
            output.unlink()
            output.parent.rmdir()


class FakeIMAP:
    def __init__(self, host, port, *, ssl_context, timeout):
        assert host == "imap.qiye.aliyun.com" and port == 993
        assert ssl_context.verify_mode == ssl.CERT_REQUIRED
        self.data = {b"INBOX": {}, b"Drafts": {}, b"Sent": {}}
        self.selected = None
        self.calls = []

    def login(self, account, secret):
        self.calls.append("LOGIN")
        return "OK", []

    def logout(self):
        self.calls.append("LOGOUT")

    def list(self):
        return "OK", [b'(\\Inbox) "/" "INBOX"', b'(\\Drafts) "/" "Drafts"', b'(\\Sent) "/" "Sent"']

    def select(self, mailbox, *, readonly):
        assert readonly is True
        self.selected = mailbox.strip(b'"')
        self.calls.append("EXAMINE")
        return "OK", []

    def response(self, name):
        return name, [b"42"]

    def uid(self, operation, *arguments):
        self.calls.append(operation)
        if operation == "SEARCH":
            candidates = self.data[self.selected]
            if arguments[1] == "UID":
                uid = str(arguments[2]).encode()
                return "OK", [uid if uid in candidates else b""]
            return "OK", [b" ".join(candidates)]
        assert operation == "FETCH"
        uid, part = arguments
        assert "BODY.PEEK[" in part
        return "OK", [(b"FETCH", self.data[self.selected][uid]), b")"]

    def append(self, mailbox, flags, date, raw):
        assert mailbox == b'"Drafts"' and flags == r"(\Draft)"
        self.calls.append("APPEND")
        pool = self.data[b"Drafts"]
        pool[str(len(pool) + 1).encode()] = raw
        return "OK", []


class ProtocolTests(unittest.TestCase):
    def test_login_failure_explains_configuration_without_echoing_server_data(self):
        class Rejected(FakeIMAP):
            def login(self, account, secret):
                raise m.imaplib.IMAP4.error("do-not-echo-simulated-private-value")
        with self.assertRaises(m.DraftError) as caught:
            m.Mailbox("sara@example.invalid", "simulated-private-value", factory=Rejected)
        self.assertIn("IMAP 登录", str(caught.exception))
        self.assertNotIn("simulated-private-value", str(caught.exception))

    def test_special_use_and_localized_folder_names(self):
        folders = m.parse_folders([b'(\\Drafts) "/" "My Drafts"', b'(\\Sent) "/" "Sent Items"'])
        self.assertEqual(m.choose_folder(folders, "drafts")["raw"], b"My Drafts")
        self.assertEqual(m.choose_folder(folders, "sent")["name"], "Sent Items")
        encoded = __import__("base64").b64encode("草稿箱".encode("utf-16-be")).decode().rstrip("=").replace("/", ",")
        folders = m.parse_folders([('(\\HasNoChildren) "/" "&' + encoded + '-"').encode()])
        self.assertEqual(m.choose_folder(folders, "drafts")["name"], "草稿箱")

    def test_ambiguous_or_unselectable_folder_fails_closed(self):
        folders = m.parse_folders([b'(\\Drafts) "/" "One"', b'(\\Drafts) "/" "Two"'])
        with self.assertRaises(m.DraftError):
            m.choose_folder(folders, "drafts")
        self.assertEqual(m.choose_folder(folders, "drafts", "Two")["raw"], b"Two")
        with self.assertRaises(m.DraftError):
            m.choose_folder(m.parse_folders([b'(\\NoSelect \\Drafts) "/" "Drafts"']), "drafts")

    def test_transport_uses_only_readonly_fetch_and_append(self):
        mailbox = m.Mailbox("sara@example.invalid", "simulated-secret", factory=FakeIMAP)
        templates = {category: {"subject": "Hello {{TikTok账号}}", "body": "Hi {{邮箱前缀}}"}
                     for category in m.CATEGORIES}
        item = m.fill_text(creator(), templates, "2026-10-20")
        item.update(id="sim-draft-id", batch_id="SIM-BATCH", message_id="<sim@example.invalid>")
        self.assertEqual(mailbox.history(item["recipient"]), [])
        mailbox.append(m.make_message(item, mailbox.account))
        found = mailbox.find(item)
        self.assertEqual(len(found), 1)
        self.assertTrue(m.matches_original(item, found[0]["remote"]))
        item.update(uid="1", uidvalidity="42", id="web-stripped-id", message_id="<web-changed@example.invalid>")
        self.assertEqual(len(mailbox.find(item)), 1)
        self.assertFalse(set(mailbox.client.calls) & {"STORE", "EXPUNGE", "DELETE", "SEND", "SMTP"})
        mailbox.close()

    def test_history_search_substring_is_not_treated_as_exact_address(self):
        mailbox = m.Mailbox("sara@example.invalid", "simulated-secret", factory=FakeIMAP)
        message = m.EmailMessage()
        message["From"] = "other-sim-001@example.invalid"
        message["To"] = "sara@example.invalid"
        message.set_content("Unrelated mail")
        mailbox.client.data[b"INBOX"][b"1"] = message.as_bytes()
        self.assertEqual(mailbox.history("sim-001@example.invalid"), [])
        self.assertEqual(len(mailbox.history("other-sim-001@example.invalid")), 1)
        mailbox.close()

    def test_password_has_no_cli_argument_and_errors_do_not_echo_server_text(self):
        output = io.StringIO()
        with tempfile.TemporaryDirectory() as directory:
            with patch.object(m, "connect", side_effect=m.imaplib.IMAP4.error("secret-should-never-be-printed")):
                with patch("sys.stderr", output):
                    result = m.main(["--state", str(Path(directory) / "local.sqlite3"), "audit", "SIM-BATCH"])
        self.assertEqual(result, 1)
        self.assertNotIn("secret-should-never-be-printed", output.getvalue())


if __name__ == "__main__":
    unittest.main()

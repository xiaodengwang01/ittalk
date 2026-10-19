"""Simulated, offline checks for the text-only archive. No real contacts."""
from copy import deepcopy
from datetime import datetime
import json
import unittest

try:
    from . import mail_text as mail
except ImportError:
    import mail_text as mail


def creator(identifier="SIM-001", address="one@example.invalid"):
    return {"达人ID": identifier, "达人昵称": "模拟达人" + identifier, "平台": "YouTube",
            "账号链接": "https://youtube.com/@" + identifier.lower(), "Email": address,
            "联系状态": "待联系", "下一步动作": "首联草稿"}


def received(**changes):
    record = {"邮件记录ID": "MAIL-000001", "达人ID": "SIM-001", "方向": "收信",
              "邮件时间": "2026-10-19 10:00:00", "发件人": "one@example.invalid",
              "收件人": "team@example.invalid", "主题": "Simulated collaboration",
              "邮件正文": "Hello,\n\nPlease share the sample details.\n\nBest,\nSim",
              "正文完整性": "完整", "记录来源": "原邮件文本", "邮箱消息ID": "SIM-MESSAGE-1",
              "邮箱账户": "team@example.invalid", "处理状态": "待回复", "邮件用途": "收信"}
    record.update(changes)
    return record


def sent(**changes):
    record = received(**{"邮件记录ID": "MAIL-000002", "方向": "发信", "邮件用途": "回复",
                         "实际发送结果": "成功", "处理状态": "已发送", "邮箱消息ID": "SIM-MESSAGE-2",
                         "发件人": "team@example.invalid", "收件人": "one@example.invalid",
                         "邮件正文": "Hello,\n\nHere are the sample details.\n\nBest,\nSim team"})
    record.update(changes)
    return record


class QuoteTests(unittest.TestCase):
    def test_exact_prose_and_paragraphs_preserved(self):
        body = "  Hello,\r\n\r\nPrice is 10.  \r\n> A normal quote\r\n\r\n  "
        self.assertEqual(mail.strip_quoted_history(body), body)

    def test_explicit_english_and_chinese_chain_boundaries(self):
        for boundary in ("On Mon, 19 Oct 2026, Sim <one@example.invalid> wrote:",
                         "在2026年10月19日，模拟发件人写道：", "-----Original Message-----",
                         "---------- Forwarded message ---------"):
            with self.subTest(boundary=boundary):
                self.assertEqual(mail.strip_quoted_history("Thanks.  \n\n" + boundary + "\n> Old chain"), "Thanks.  ")

    def test_contiguous_complete_header_block_is_old_chain(self):
        tail = "From: one@example.invalid\nSent: 19 October 2026\nTo: team@example.invalid\nSubject: Sim\n\n> Old"
        self.assertEqual(mail.strip_quoted_history("New reply\n\n" + tail), "New reply")
        chinese = "发件人：one@example.invalid\n发送时间：2026-10-19\n收件人：team@example.invalid\n主题：Sim\n\n> 旧文"
        self.assertEqual(mail.strip_quoted_history(chinese), "")

    def test_isolated_headers_or_normal_quotes_are_not_deleted(self):
        bodies = ("From: our point of view\nThis is a normal sentence.\nDate: not a header block",
                  "I wrote:\n> It is a good idea\nOur answer stays here.",
                  "From: one@example.invalid\nSubject: we intentionally lack a date\nAn author sentence")
        for body in bodies:
            self.assertEqual(mail.strip_quoted_history(body), body)

    def test_mid_paragraph_wrote_does_not_become_boundary(self):
        body = "They said On Monday Sim wrote: something surprising.\nWe disagree."
        self.assertEqual(mail.strip_quoted_history(body), body)

    def test_inline_and_bottom_replies_keep_unquoted_new_text(self):
        body = "Thanks.\n\nOn Mon, Sim wrote:\n> Europe?\n\nYes, Europe.\n> Price?\n\nIt is 10."
        result = mail.extract_new_body(body)
        self.assertEqual(result["邮件正文"], "Thanks.\n\nYes, Europe.\n\nIt is 10.")
        self.assertTrue(result["可确认新增正文"])
        self.assertFalse(result["需核实"])
        bottom = "On Mon, Sim wrote:\n> Europe?\n\nYes, Europe."
        self.assertEqual(mail.strip_quoted_history(bottom), "Yes, Europe.")
        self.assertTrue(mail.is_archive_message(sent(**{"邮件正文": bottom})))

    def test_plain_unquoted_tail_is_preserved_and_marked_ambiguous(self):
        body = "On Mon, Sim wrote:\nOld prose\n\nMaybe a bottom reply."
        row = mail.canonical_message(received(**{"邮件正文": body}))
        self.assertEqual(row["邮件正文"], body)
        self.assertEqual(row["正文完整性"], "待核实")
        self.assertIn("已保留原文", row["核实说明"])
        self.assertFalse(mail.is_archive_message(sent(**{"邮件正文": body})))

    def test_normal_quote_containing_mail_marker_is_not_removed(self):
        body = "An example for discussion:\n> On Mon, Sim wrote:\n> This is my normal quotation.\nOur comment."
        self.assertEqual(mail.strip_quoted_history(body), body)

    def test_ambiguity_note_is_idempotent_and_header_only_still_excluded(self):
        row = mail.canonical_message(received(**{"邮件正文": "On Mon, Sim wrote:\nUnmarked old text"}))
        self.assertEqual(mail.canonical_message(row)["核实说明"], row["核实说明"])
        self.assertFalse(mail.is_archive_message(sent(**{"邮件正文": "On Mon, Sim wrote:"})))

    def test_partial_headers_are_not_mistaken_for_sent_reply(self):
        for body in ("From: Sim", "发件人：sim@example.invalid\nDate: 2026-10-19",
                     "> From: Sim\n> To: team@example.invalid",
                     "From: Sim\n    <sim@example.invalid>"):
            self.assertTrue(mail.is_header_only(body))
            self.assertFalse(mail.is_archive_message(sent(**{"邮件正文": body})))
        self.assertFalse(mail.is_header_only("From: our point of view\nThis is a normal reply."))
        self.assertTrue(mail.is_archive_message(sent(**{"邮件正文": "From: our point of view\nThis is a normal reply."})))

    def test_nested_unmarked_chain_is_ambiguous_not_confirmed_new_reply(self):
        body = "On Mon, Sim wrote:\n> Old mail\nOn Sun, Team wrote:\nOlder prose"
        result = mail.extract_new_body(body)
        self.assertTrue(result["需核实"])
        self.assertIn("Older prose", result["邮件正文"])
        self.assertFalse(mail.is_archive_message(sent(**{"邮件正文": body})))


class ArchivePolicyTests(unittest.TestCase):
    def test_canonical_body_wins_even_if_empty(self):
        record = mail.canonical_message({"邮件正文": "", "正文或草稿": "old draft"})
        self.assertEqual(record["邮件正文"], "")
        self.assertNotIn("正文或草稿", record)
        self.assertEqual(record["正文完整性"], "缺失")

    def test_old_body_alias_and_safe_field_allowlist(self):
        record = mail.canonical_message({"正文或草稿": "Hello\n\nParagraph", "_images": ["SIM-IMAGE"],
                                         "历史截图": "SIM-IMAGE", "草稿指纹": "SIM-DRAFT", "private": "SIM-SECRET"})
        self.assertEqual(record["邮件正文"], "Hello\n\nParagraph")
        self.assertFalse(set(record) & {"_images", "历史截图", "private", "正文或草稿", "草稿指纹"})

    def test_successful_reply_and_incoming_allowed(self):
        self.assertTrue(mail.is_archive_message(received()))
        self.assertTrue(mail.is_archive_message(sent()))

    def test_initial_draft_failed_unknown_and_quote_only_not_archived(self):
        invalid = [sent(**{"邮件用途": "首联"}), sent(**{"实际发送结果": "失败"}),
                   sent(**{"实际发送结果": "结果待核实"}), sent(**{"实际发送结果": ""}),
                   sent(**{"处理状态": "已确认待发送"}), sent(**{"邮件用途": "历史"}),
                   sent(**{"邮件正文": "On Mon, Sim wrote:\n> Old mail"})]
        creators, messages = [creator()], []
        for record in invalid:
            with self.subTest(record=record):
                self.assertFalse(mail.is_archive_message(record))
                result = mail.import_message(creators, messages, record)
                self.assertFalse(result["已导入"])
                self.assertIsNone(result["记录"])
        self.assertEqual(messages, [])

    def test_legacy_initial_template_and_draft_purpose_detected(self):
        for extra in ({"草稿用途": "首联"}, {"模板及版本": "首联/v1"}, {"邮件类型": "首联"}):
            row = sent(**{"邮件用途": ""})
            row.update(extra)
            self.assertFalse(mail.is_archive_message(row))

    def test_initial_metadata_has_no_body_subject_or_recipient(self):
        row = sent(**{"邮件用途": "首联", "主题": "SIM-SECRET-SUBJECT", "邮件正文": "SIM-SECRET-BODY"})
        result = mail.initial_contact_metadata(row)
        self.assertEqual(set(result), {"达人ID", "最近发信时间", "首次联系日期", "首联发送结果"})
        self.assertNotIn("SIM-SECRET", json.dumps(result))
        self.assertEqual(mail.initial_contact_metadata({**row, "实际发送结果": ""}), {})

    def test_screenshot_transcription_does_not_default_to_complete(self):
        row = mail.canonical_message({"邮件正文": "Partial evidence", "记录来源": "历史截图转录"})
        self.assertEqual(row["正文完整性"], "待核实")


class ImportTests(unittest.TestCase):
    def setUp(self):
        self.creators = [creator()]
        self.messages = []

    def test_incoming_updates_state_and_repeat_keeps_id(self):
        one = mail.import_message(self.creators, self.messages, received())
        two = mail.import_message(self.creators, self.messages, received())
        self.assertTrue(one["已导入"])
        self.assertFalse(two["已导入"])
        self.assertEqual(len(self.messages), 1)
        self.assertEqual(self.messages[0]["邮件记录ID"], "MAIL-000001")
        self.assertEqual(self.creators[0]["联系状态"], "已回复")
        self.assertEqual(self.creators[0]["下一步动作"], "回复草稿")

    def test_ambiguous_quote_chain_is_not_imported_until_manually_separated(self):
        body = "New reply.\n\nOn Mon, Sim wrote:\nUnmarked old chain or bottom reply."
        before = deepcopy(self.creators)
        for direction_record in (received(**{"邮件正文": body}), sent(**{"邮件正文": body})):
            result = mail.import_message(self.creators, self.messages, direction_record)
            self.assertFalse(result["已导入"])
            self.assertIn("需核实引用链", result["原因"])
            self.assertIsNone(result["记录"])
        self.assertEqual(self.messages, [])
        self.assertEqual(self.creators, before)

    def test_ambiguous_quote_chain_fill_preserves_original_record(self):
        mail.import_message(self.creators, self.messages, received(**{"邮件正文": "Partial text", "正文完整性": "部分"}))
        before = deepcopy(self.messages)
        body = "New reply.\n\nOn Mon, Sim wrote:\nUnmarked old chain or bottom reply."
        with self.assertRaisesRegex(ValueError, "需核实引用链"):
            mail.fill_message(self.creators, self.messages, "MAIL-000001", {"邮件正文": body, "正文完整性": "完整"})
        self.assertEqual(self.messages, before)

    def test_clear_inline_reply_can_be_imported_with_old_quotes_removed(self):
        body = "On Mon, Sim wrote:\n> Europe?\n\nYes, Europe.\n> Price?\n\nIt is 10."
        result = mail.import_message(self.creators, self.messages, received(**{"邮件正文": body}))
        self.assertTrue(result["已导入"])
        self.assertEqual(result["记录"]["邮件正文"], "Yes, Europe.\n\nIt is 10.")

    def test_new_inbound_preserves_declined_or_paused_contact_status(self):
        for state in ("拒绝", "已拒绝", "不合作", "勿联系", "已完成"):
            creators = [creator()]
            creators[0]["联系状态"] = state
            messages = []
            result = mail.import_message(creators, messages, received())
            self.assertTrue(result["已导入"])
            self.assertEqual(creators[0]["联系状态"], state)
            self.assertEqual(creators[0]["下一步动作"], "暂停联系")
            self.assertEqual(creators[0]["最近收信时间"], received()["邮件时间"])
        creators = [creator()]
        creators[0]["下一步动作"] = "暂停联系"
        mail.import_message(creators, [], received())
        self.assertEqual(creators[0]["联系状态"], "待联系")
        self.assertEqual(creators[0]["下一步动作"], "暂停联系")

    def test_distinct_provider_messages_never_collapse(self):
        mail.import_message(self.creators, self.messages, received())
        row = received(**{"邮箱消息ID": "SIM-MESSAGE-2", "邮件记录ID": ""})
        mail.import_message(self.creators, self.messages, row)
        self.assertEqual(len(self.messages), 2)

    def test_provider_ids_are_scoped_to_mailbox(self):
        mail.import_message(self.creators, self.messages, received())
        row = received(**{"邮箱账户": "other@example.invalid", "邮件记录ID": ""})
        mail.import_message(self.creators, self.messages, row)
        self.assertEqual(len(self.messages), 2)

    def test_single_message_locator_deduplicates_but_thread_locator_does_not(self):
        for url, expected in (("https://mail.example.invalid/message/SIM-1", 1),
                              ("https://mail.example.invalid/thread/SIM-THREAD", 2)):
            messages = []
            row = received(**{"邮件记录ID": "", "邮箱消息ID": "", "原邮件定位": url, "邮件时间": ""})
            mail.import_message(self.creators, messages, row)
            mail.import_message(self.creators, messages, {**row, "邮件记录ID": ""})
            self.assertEqual(len(messages), expected)

    def test_fallback_requires_full_time_parties_subject_and_body(self):
        row = received(**{"邮箱消息ID": "", "原邮件定位": ""})
        mail.import_message(self.creators, self.messages, row)
        mail.import_message(self.creators, self.messages, row)
        self.assertEqual(len(self.messages), 1)
        for missing in ("邮件时间", "发件人", "收件人", "主题", "邮件正文"):
            messages = []
            incomplete = {**row, missing: "", "邮件记录ID": ""}
            mail.import_message(self.creators, messages, incomplete)
            mail.import_message(self.creators, messages, incomplete)
            self.assertEqual(len(messages), 2, missing)

    def test_duplicate_complete_conflict_does_not_overwrite(self):
        mail.import_message(self.creators, self.messages, received())
        old = self.messages[0]["邮件正文"]
        mail.set_mail_summary(self.messages[0], "模拟摘要")
        candidate = received(**{"邮件正文": "SIM-CONFLICTING-BODY"})
        result = mail.import_message(self.creators, self.messages, candidate)
        self.assertTrue(result["补全待处理"])
        self.assertEqual(self.messages[0]["邮件正文"], old)
        self.assertEqual(self.messages[0]["摘要状态"], "待更新")
        self.assertIn("需核实", self.messages[0]["核实说明"])
        self.assertNotIn("SIM-CONFLICTING-BODY", self.messages[0]["核实说明"])

    def test_partial_completion_is_explicit_and_preserves_id(self):
        first = received(**{"邮件正文": "Hello", "正文完整性": "部分"})
        mail.import_message(self.creators, self.messages, first)
        mail.set_mail_summary(self.messages[0], "旧的模拟摘要")
        mail.set_creator_summary(self.creators[0], self.messages, "旧的模拟沟通摘要")
        result = mail.import_message(self.creators, self.messages, received())
        self.assertTrue(result["补全待处理"])
        self.assertEqual(self.messages[0]["邮件正文"], "Hello")
        result = mail.fill_message(self.creators, self.messages, "MAIL-000001", received())
        self.assertTrue(result["已补全"])
        self.assertEqual(len(self.messages), 1)
        self.assertEqual(self.messages[0]["邮件记录ID"], "MAIL-000001")
        self.assertEqual(self.messages[0]["邮件正文"], received()["邮件正文"])
        self.assertEqual(self.messages[0]["正文完整性"], "完整")
        self.assertEqual(self.messages[0]["摘要状态"], "待更新")
        self.assertEqual(self.creators[0]["沟通摘要状态"], "待更新")

    def test_explicit_fill_conflict_retains_complete_text(self):
        mail.import_message(self.creators, self.messages, received())
        result = mail.fill_message(self.creators, self.messages, "MAIL-000001", {"邮件正文": "SIM-WRONG"})
        self.assertTrue(result["存在冲突"])
        self.assertFalse(result["已补全"])
        self.assertEqual(result["记录"]["邮件正文"], received()["邮件正文"])

    def test_fill_cannot_reassign_or_change_provider_identity(self):
        mail.import_message(self.creators, self.messages, received())
        for changes in ({"达人ID": "SIM-002"}, {"邮件记录ID": "MAIL-000003"},
                        {"邮箱消息ID": "SIM-OTHER"}, {"方向": "发信"}):
            with self.assertRaises(ValueError):
                mail.fill_message(self.creators, self.messages, "MAIL-000001", changes)

    def test_missing_record_id_allocated_and_duplicate_id_fails(self):
        mail.import_message(self.creators, self.messages, received(**{"邮件记录ID": ""}))
        self.assertRegex(self.messages[0]["邮件记录ID"], r"^MAIL-[0-9a-f]{32}$")
        self.messages[0]["邮件记录ID"] = "MAIL-000001"
        with self.assertRaises(ValueError):
            mail.import_message(self.creators, self.messages, received(**{"邮箱消息ID": "SIM-DIFFERENT"}))

    def test_local_partial_replay_without_provider_id_uses_permanent_id(self):
        row = received(**{"邮箱消息ID": "", "邮件时间": "", "邮件正文": "Partial", "正文完整性": "部分"})
        mail.import_message(self.creators, self.messages, row)
        result = mail.import_message(self.creators, self.messages, row)
        self.assertFalse(result["已导入"])
        self.assertEqual(len(self.messages), 1)
        candidate = {**row, "邮件正文": "Partial with the rest of this message", "正文完整性": "完整"}
        result = mail.import_message(self.creators, self.messages, candidate)
        self.assertTrue(result["补全待处理"])
        self.assertEqual(self.messages[0]["邮件正文"], "Partial")
        mail.fill_message(self.creators, self.messages, row["邮件记录ID"], candidate)
        self.assertEqual(len(self.messages), 1)
        self.assertEqual(self.messages[0]["邮件正文"], candidate["邮件正文"])

    def test_shared_agency_address_is_not_implicit_ownership(self):
        self.creators = [creator("SIM-001", "agent@example.invalid"), creator("SIM-002", "agent@example.invalid")]
        result = mail.import_message(self.creators, self.messages,
                                     received(**{"达人ID": "", "发件人": "agent@example.invalid"}))
        self.assertTrue(result["需人工归属"])
        self.assertEqual(result["记录"]["达人ID"], "")
        self.assertEqual(result["记录"]["处理状态"], "待人工归属")

    def test_known_thread_resolves_shared_agency_conservatively(self):
        self.creators = [creator("SIM-001", "agent@example.invalid"), creator("SIM-002", "agent@example.invalid")]
        prior = received(**{"达人ID": "SIM-002", "会话ID": "SIM-THREAD", "发件人": "agent@example.invalid"})
        self.messages = [prior]
        incoming = received(**{"邮件记录ID": "", "达人ID": "", "会话ID": "SIM-THREAD",
                               "邮箱消息ID": "SIM-MESSAGE-NEW", "发件人": "agent@example.invalid"})
        result = mail.import_message(self.creators, self.messages, incoming)
        self.assertEqual(result["记录"]["达人ID"], "SIM-002")


class SummaryTests(unittest.TestCase):
    def setUp(self):
        self.creators, self.messages = [creator()], [received()]

    def set_summaries(self):
        mail.set_mail_summary(self.messages[0], "模拟邮件摘要")
        mail.set_creator_summary(self.creators[0], self.messages, "模拟沟通摘要")

    def test_current_main_draft_edits_and_removal_invalidate_summary(self):
        row=self.creators[0]
        row.update({'草稿状态':'未发送草稿','草稿收件人':'one@example.invalid','草稿主题':'Sim reply','当前回复草稿':'SIM-UNSENT-PRIVATE-TEXT','草稿来源':'模拟对话'})
        self.set_summaries()
        for field in ('草稿状态','草稿收件人','草稿主题','当前回复草稿','草稿来源'):
            with self.subTest(field=field):
                old=row[field];row[field]=old+' changed'
                mail.refresh_summary_states(self.creators,self.messages)
                self.assertEqual(row['沟通摘要状态'],'待更新')
                row[field]=old;self.set_summaries()
        for field in ('草稿状态','草稿收件人','草稿主题','当前回复草稿','草稿来源'):row[field]=''
        mail.refresh_summary_states(self.creators,self.messages)
        self.assertEqual(row['沟通摘要状态'],'待更新')

    def test_main_draft_body_never_enters_overview_or_mail_context(self):
        self.creators[0].update({'草稿状态':'未发送草稿','当前回复草稿':'SIM-UNSENT-PRIVATE-TEXT'})
        self.set_summaries()
        for output in (mail.overview(self.creators,self.messages),mail.context(self.creators,self.messages,'SIM-001')):
            self.assertNotIn('SIM-UNSENT-PRIVATE-TEXT',json.dumps(output,ensure_ascii=False))
        self.assertEqual(len(self.messages),1)

    def test_body_hash_does_not_strip(self):
        self.assertNotEqual(mail.body_hash("Text"), mail.body_hash("Text "))
        self.assertNotEqual(mail.body_hash("Text"), mail.body_hash("\nText"))

    def test_summary_hash_survives_none_empty_and_iso_datetime_serialization(self):
        row = received(**{"核实说明": None, "邮件时间": datetime(2026, 10, 19, 10)})
        expected = mail.message_fingerprint(row)
        row.update({"核实说明": "", "邮件时间": "2026-10-19T10:00:00"})
        self.assertEqual(expected, mail.message_fingerprint(row))
        row["邮件时间"] = "2026-10-19 10:00:00"
        self.assertEqual(expected, mail.message_fingerprint(row))
        self.creators[0]["最近收信时间"] = datetime(2026, 10, 19, 10)
        expected = mail.creator_fingerprint(self.creators[0], [row])
        self.creators[0]["最近收信时间"] = "2026-10-19T10:00:00"
        self.creators[0]["计划跟进日期"] = None
        self.assertEqual(expected, mail.creator_fingerprint(self.creators[0], [row]))

    def test_body_and_metadata_edits_invalidate_both_levels(self):
        for field, value in (("邮件正文", "Changed body"), ("邮件正文", received()["邮件正文"] + "\n"),
                             ("发件人", "second@example.invalid"), ("收件人", "other@example.invalid"),
                             ("主题", "Changed subject"), ("邮件时间", "2026-10-18"),
                             ("记录来源", "人工核对原邮件"), ("正文完整性", "部分"),
                             ("核实说明", "模拟更正")):
            with self.subTest(field=field):
                self.messages = [received()]
                self.set_summaries()
                self.messages[0][field] = value
                mail.refresh_summary_states(self.creators, self.messages)
                self.assertEqual(self.messages[0]["摘要状态"], "待更新")
                self.assertEqual(self.creators[0]["沟通摘要状态"], "待更新")

    def test_appended_quote_chain_invalidates_exact_body_hash(self):
        self.set_summaries()
        old = self.messages[0]["摘要依据指纹"]
        self.messages[0]["邮件正文"] += "\n\nOn Mon, Sim wrote:\nOld chain"
        self.assertNotEqual(old, mail.message_fingerprint(self.messages[0]))

    def test_contact_state_and_next_action_invalidate_creator_summary(self):
        for field, value in (("联系状态", "已拒绝"), ("下一步动作", "暂停联系"), ("计划跟进日期", "2026-10-20")):
            self.set_summaries()
            self.creators[0][field] = value
            mail.refresh_summary_states(self.creators, self.messages)
            self.assertEqual(self.creators[0]["沟通摘要状态"], "待更新")
            self.assertEqual(self.messages[0]["摘要状态"], "有效")

    def test_new_mail_invalidates_creator_summary(self):
        self.set_summaries()
        self.messages.append(sent())
        mail.refresh_summary_states(self.creators, self.messages)
        self.assertEqual(self.creators[0]["沟通摘要状态"], "待更新")

    def test_summary_update_rebinds_fingerprint_and_overview_reports_stale(self):
        self.set_summaries()
        self.messages[0]["主题"] = "Changed subject"
        self.assertEqual(mail.overview(self.creators, self.messages)["邮件索引"][0]["摘要状态"], "待更新")
        self.set_summaries()
        mail.refresh_summary_states(self.creators, self.messages)
        self.assertEqual(self.messages[0]["摘要状态"], "有效")
        self.assertEqual(self.creators[0]["沟通摘要状态"], "有效")


class ReadTests(unittest.TestCase):
    def setUp(self):
        self.creators = [creator(), creator("SIM-002", "two@example.invalid")]
        self.messages = [received(), sent(),
                         received(**{"邮件记录ID": "MAIL-000003", "邮件时间": "", "邮件正文": "SIM-UNKNOWN-ONE"}),
                         received(**{"邮件记录ID": "MAIL-000004", "邮件时间": "时间待核实", "邮件正文": "SIM-UNKNOWN-TWO"}),
                         received(**{"邮件记录ID": "MAIL-000005", "达人ID": "SIM-002", "邮件正文": "SIM-OTHER-CREATOR"}),
                         received(**{"邮件记录ID": "MAIL-000006", "达人ID": "", "邮件正文": "SIM-UNASSIGNED"}),
                         sent(**{"邮件记录ID": "MAIL-000007", "邮件用途": "首联", "邮件正文": "SIM-INITIAL-SECRET"}),
                         sent(**{"邮件记录ID": "MAIL-000008", "处理状态": "待确认草稿", "实际发送结果": "", "邮件正文": "SIM-DRAFT-SECRET"})]

    def test_overview_has_no_body_draft_initial_or_notes(self):
        view = mail.overview(self.creators, self.messages)
        serialized = json.dumps(view, ensure_ascii=False)
        for secret in ("邮件正文", "正文或草稿", "SIM-UNKNOWN-ONE", "SIM-OTHER-CREATOR", "SIM-UNASSIGNED",
                       "SIM-INITIAL-SECRET", "SIM-DRAFT-SECRET", "发件人", "核实说明"):
            self.assertNotIn(secret, serialized)
        self.assertEqual(len(view["邮件索引"]), 6)
        self.assertEqual(view["待人工归属数量"], 1)

    def test_context_all_own_valid_mail_without_cross_creator_or_unassigned_leak(self):
        result = mail.context(self.creators, self.messages, "SIM-001", budget=100000)
        serialized = json.dumps(result)
        for secret in ("SIM-OTHER-CREATOR", "SIM-UNASSIGNED", "SIM-INITIAL-SECRET", "SIM-DRAFT-SECRET"):
            self.assertNotIn(secret, serialized)
        self.assertEqual(result["全部邮件数量"], 4)
        self.assertEqual(len(result["时间明确邮件"]), 2)
        self.assertEqual([m["邮件正文"] for m in result["时间不明邮件"]], ["SIM-UNKNOWN-ONE", "SIM-UNKNOWN-TWO"])

    def test_known_dates_sorted_and_unknowns_keep_source_order(self):
        self.messages[0]["邮件时间"] = "2026-10-20 10:00:00"
        self.messages[1]["邮件时间"] = "2026-10-19 10:00:00"
        result = mail.context(self.creators, self.messages, "SIM-001", budget=100000)
        self.assertEqual([m["邮件记录ID"] for m in result["时间明确邮件"]], ["MAIL-000002", "MAIL-000001"])

    def test_whole_mail_pagination_and_range_exhaustively_return_every_mail(self):
        cursor, seen, ranges = 0, [], []
        while True:
            result = mail.context(self.creators, self.messages, "SIM-001", budget=1, cursor=cursor)
            page = result["时间明确邮件"] + result["时间不明邮件"]
            self.assertEqual(len(page), 1)
            seen.extend(m["邮件记录ID"] for m in page)
            self.assertEqual(len(result["单封超过预算"]), 1)
            self.assertIn("邮件正文", page[0])
            ranges.append(result["本页范围"])
            cursor = result["下一页游标"]
            if cursor is None:
                break
        self.assertEqual(seen, ["MAIL-000001", "MAIL-000002", "MAIL-000003", "MAIL-000004"])
        self.assertEqual(ranges, [{"起始": n, "结束": n} for n in range(1, 5)])

    def test_changed_archive_rejects_snapshot_cursor(self):
        first = mail.context(self.creators, self.messages, "SIM-001", budget=1)
        self.messages[1]["邮件正文"] += "new words"
        with self.assertRaises(ValueError):
            mail.context(self.creators, self.messages, "SIM-001", budget=1, cursor=first["下一页游标"])

    def test_cursor_cannot_be_reused_for_other_creator(self):
        first = mail.context(self.creators, self.messages, "SIM-001", budget=1)
        with self.assertRaises(ValueError):
            mail.context(self.creators, self.messages, "SIM-002", cursor=first["下一页游标"])

    def test_empty_unknown_creator_invalid_budget_and_invalid_cursor(self):
        for identifier in ("", "SIM-DOES-NOT-EXIST"):
            with self.assertRaises(ValueError):
                mail.context(self.creators, self.messages, identifier)
        for budget in (0, -1, True, "100"):
            with self.assertRaises(ValueError):
                mail.context(self.creators, self.messages, "SIM-001", budget=budget)
        with self.assertRaises(ValueError):
            mail.context(self.creators, self.messages, "SIM-001", cursor=99)
        result = mail.context(self.creators, [], "SIM-001")
        self.assertEqual(result["全部邮件数量"], 0)
        self.assertEqual(result["本页范围"], {"起始": 0, "结束": 0})
        self.assertTrue(result["已读完"])


if __name__ == "__main__":
    unittest.main()

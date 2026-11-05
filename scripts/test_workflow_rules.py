"""Simulated, offline acceptance tests. No real contacts or mailbox access."""

import unittest
from copy import deepcopy
from datetime import date, datetime

try:
    from . import workflow_rules as rules
except ImportError:
    import workflow_rules as rules


def creator(identifier="SIM-001", email="creator@example.invalid", account="sim-one"):
    return {"达人ID": identifier, "平台": "YouTube", "账号链接": "https://youtube.com/@" + account,
            "Email": email, "昵称": "模拟达人", "下一步动作": "首联草稿", "联系状态": "待联系",
            "首联内容分类": "宠物剧情类", "分类核实状态": "初判"}


def initial_contact(creators, messages, sent_at="2026-10-01 10:00:00"):
    draft = rules.create_draft(creators, messages, creators[0]["达人ID"],
                               "Simulated collaboration", "Simulated body", "首联/v1")
    rules.confirm_batch(messages, [draft["邮件记录ID"]], "SIM-BATCH-1", "2026-10-01 09:00:00")
    rules.record_send_result(creators, messages, draft["邮件记录ID"], "成功", sent_at)
    return draft


class IdentityTests(unittest.TestCase):
    def test_ids_survive_reordering_and_new_rows(self):
        creators = [{"编号": "SIM-001"}, {}, {}]
        rules.ensure_creator_ids(creators)
        self.assertEqual([item["达人ID"] for item in creators], ["SIM-001", "KOL-000001", "KOL-000002"])
        creators.reverse()
        creators.append({})
        rules.ensure_creator_ids(creators)
        self.assertEqual(creators[-1]["达人ID"], "KOL-000003")
        self.assertEqual(creators[0]["达人ID"], "KOL-000002")

    def test_duplicate_ids_fail_instead_of_reassigning(self):
        with self.assertRaises(ValueError):
            rules.ensure_creator_ids([{"达人ID": "SIM-001"}, {"达人ID": "SIM-001"}])

    def test_canonical_accounts_and_case_sensitive_channel_ids(self):
        self.assertEqual(rules.canonical_account_url("yt", "https://www.youtube.com/@SIM/videos?utm_source=sim"),
                         "https://youtube.com/@sim")
        self.assertEqual(rules.canonical_account_url("ig", "@Sim"), "https://instagram.com/sim")
        self.assertEqual(rules.canonical_account_url("YouTube", "youtube.com/channel/UCAbC"),
                         "https://youtube.com/channel/UCAbC")
        self.assertNotEqual(rules.canonical_account_url("YouTube", "youtube.com/channel/UCAbC"),
                            rules.canonical_account_url("YouTube", "youtube.com/channel/UCabc"))

    def test_duplicate_profiles_merge_but_shared_email_does_not(self):
        one = creator(email="agent@example.invalid")
        duplicate = creator("SIM-002", email="agent@example.invalid")
        duplicate["平台"] = "yt"
        duplicate["账号链接"] = "https://www.youtube.com/@SIM-ONE/videos"
        duplicate["备注"] = "模拟历史"
        independent = creator("SIM-003", "agent@example.invalid", "sim-two")
        creators = [one, duplicate, independent]
        messages = [{"达人ID": "SIM-002", "正文或草稿": "模拟历史内容"}]
        result = rules.merge_duplicate_creators(creators, messages)
        self.assertEqual(result["merged"], 1)
        self.assertEqual(len(creators), 2)
        self.assertEqual(messages[0]["达人ID"], "SIM-001")
        self.assertEqual(creators[0]["备注"], "模拟历史")

    def test_scrape_preserves_manual_fields_and_records_real_provenance_only(self):
        creators = [creator()]
        creators[0].update({"首联内容分类": "宠物剧情类", "分类核实状态": "已确认", "粉丝数": 10})
        candidate = {"平台": "yt", "账号链接": "youtube.com/@sim-one", "Email": "other@example.invalid",
                     "首联内容分类": "宠物用品类", "分类核实状态": "初判", "粉丝数": "8.5m",
                     "来源链接": "https://youtube.com/@sim-one/about", "最近抓取时间": "伪造时间"}
        result = rules.merge_creator_candidates(creators, [candidate])
        self.assertEqual(creators[0]["Email"], "creator@example.invalid")
        self.assertEqual(creators[0]["分类核实状态"], "已确认")
        self.assertEqual(creators[0]["首联内容分类"], "宠物剧情类")
        self.assertEqual(creators[0]["粉丝数"], 10)
        self.assertNotIn("最近抓取时间", creators[0])
        self.assertTrue(result["conflicts"])
        rules.merge_creator_candidates(creators, [candidate], "2026-10-19 10:00:00")
        self.assertEqual(creators[0]["粉丝数"], 8500000)
        self.assertEqual(creators[0]["最近抓取时间"], "2026-10-19 10:00:00")

    def test_contact_url_is_not_email_and_not_merge_key(self):
        creators = []
        candidate = {"平台": "Instagram", "账号链接": "instagram.com/sim", "Email": "https://example.invalid/contact"}
        rules.merge_creator_candidates(creators, [candidate])
        self.assertEqual(creators[0]["Email"], "")
        self.assertEqual(creators[0]["联系页链接"], "https://example.invalid/contact")
        self.assertEqual(creators[0]["下一步动作"], "补资料")

    def test_real_workbook_metric_names_refresh_and_new_candidates_are_uncontacted(self):
        creators = [creator()]
        creators[0].update({"近10条中位播放": 10, "近10条最高播放": 20})
        existing = {"平台": "YouTube", "账号链接": "youtube.com/@sim-one",
                    "近10条中位播放": "1.2万", "近10条最高播放": "2.5m"}
        new = {"平台": "Instagram", "账号链接": "instagram.com/sim-new", "Email": "new@example.invalid",
               "首联内容分类": "宠物剧情类", "分类核实状态": "待核实", "下一步动作": "首联草稿",
               "联系状态": "已联系"}
        rules.merge_creator_candidates(creators, [existing, new], "2026-10-19")
        self.assertEqual(creators[0]["近10条中位播放"], 12000)
        self.assertEqual(creators[0]["近10条最高播放"], 2500000)
        self.assertEqual(creators[1]["联系状态"], "待联系")
        self.assertEqual(creators[1]["下一步动作"], "核实分类")

    def test_invalid_candidate_rejects_whole_batch_before_any_mutation(self):
        creators = [{"平台": "YouTube", "账号链接": "youtube.com/@sim-old", "Email": "old@example.invalid"}]
        before = deepcopy(creators)
        valid = {"平台": "Instagram", "账号链接": "instagram.com/sim-new"}
        for invalid in ({"平台": "YouTube"}, {"账号链接": "youtube.com/@sim-new"},
                        {"平台": "YouTube", "账号链接": ""}):
            with self.assertRaises(ValueError):
                rules.merge_creator_candidates(creators, [valid, invalid], "2026-10-19")
            self.assertEqual(creators, before)


class DataTests(unittest.TestCase):
    def test_strict_email(self):
        for email in ("sim@example.invalid", "sim+tag@agent.example.invalid"):
            self.assertTrue(rules.is_valid_email(email))
        for value in (None, "", "https://example.invalid/contact", "https://sim@example.invalid",
                      "Sim <sim@example.invalid>", "sim@example.invalid;other@example.invalid",
                      "sim..name@example.invalid", ".sim@example.invalid", "sim@-example.invalid"):
            self.assertFalse(rules.is_valid_email(value), value)

    def test_metric_units(self):
        expected = {"8.5m": 8500000, "2K": 2000, "1.2b": 1200000000,
                    "1.5万": 15000, "0.2亿": 20000000, "1,234": 1234, "0": 0}
        for value, number in expected.items():
            self.assertEqual(rules.parse_metric(value), number)
        for value in (None, "", "约1万", "1-2万", True, -2, float("nan")):
            self.assertIsNone(rules.parse_metric(value))

    def test_workdays_and_empty_dates(self):
        self.assertEqual(rules.add_workdays("2026-10-02", 7), date(2026, 10, 13))
        self.assertEqual(rules.add_workdays(datetime(2026, 10, 1, 10), 7), date(2026, 10, 12))
        self.assertIsNone(rules.add_workdays(None))
        self.assertIsNone(rules.add_workdays(""))
        self.assertIsNone(rules.add_workdays("not-a-date"))


class MailImportTests(unittest.TestCase):
    def setUp(self):
        self.creators = [creator()]
        self.messages = []
        self.incoming = {"方向": "收信", "邮件时间": "2026-10-19 09:00:00",
                         "发件人": "creator@example.invalid", "收件人": "team@example.invalid",
                         "主题": "Simulation", "正文或草稿": "Simulation reply"}

    def test_provider_id_duplicate_replay(self):
        incoming = {**self.incoming, "provider_message_id": "SIM-MESSAGE-1"}
        self.assertTrue(rules.import_message(self.creators, self.messages, incoming)["imported"])
        replay = {**incoming, "主题": "Provider remains authoritative"}
        self.assertFalse(rules.import_message(self.creators, self.messages, replay)["imported"])
        self.assertEqual(len(self.messages), 1)
        # Distinct mailbox messages are retained, even with identical text.
        rules.import_message(self.creators, self.messages, {**incoming, "provider_message_id": "SIM-MESSAGE-2"})
        self.assertEqual(len(self.messages), 2)

    def test_locator_and_full_fingerprint_fallback(self):
        for extra in ({"原邮件定位": "https://mail.example.invalid/message/SIM-1"}, {}):
            messages = []
            incoming = {**self.incoming, **extra}
            rules.import_message(self.creators, messages, incoming)
            self.assertFalse(rules.import_message(self.creators, messages, incoming)["imported"])
            self.assertEqual(len(messages), 1)

    def test_incomplete_message_not_falsely_deduplicated(self):
        incoming = {**self.incoming, "邮件时间": "", "正文或草稿": ""}
        rules.import_message(self.creators, self.messages, incoming)
        rules.import_message(self.creators, self.messages, incoming)
        self.assertEqual(len(self.messages), 2)
        self.assertNotIn("最近收信时间", self.creators[0])

    def test_date_only_history_preserves_precision(self):
        result = rules.import_message(self.creators, self.messages,
                                      {**self.incoming, "邮件时间": "2026-10-19"})
        self.assertEqual(result["record"]["邮件时间"], "2026-10-19")
        self.assertEqual(self.creators[0]["最近收信时间"], "2026-10-19")

    def test_shared_agency_email_needs_thread_or_stable_id(self):
        self.creators = [creator("SIM-001", "agent@example.invalid", "sim-one"),
                         creator("SIM-002", "agent@example.invalid", "sim-two")]
        incoming = {**self.incoming, "发件人": "agent@example.invalid", "邮箱消息ID": "SIM-M1"}
        result = rules.import_message(self.creators, self.messages, incoming)
        self.assertTrue(result["ambiguous"])
        self.assertEqual(result["record"]["达人ID"], "")
        self.assertEqual(result["record"]["处理状态"], "待人工归属")
        known = {"达人ID": "SIM-002", "会话ID": "SIM-THREAD-2"}
        self.messages.append(known)
        result = rules.import_message(self.creators, self.messages,
                                      {**incoming, "邮箱消息ID": "SIM-M2", "会话ID": "SIM-THREAD-2"})
        self.assertEqual(result["record"]["达人ID"], "SIM-002")

    def test_thread_and_explicit_creator_conflict_stays_unassigned(self):
        self.creators.append(creator("SIM-002", "second@example.invalid", "sim-two"))
        self.messages.append({"达人ID": "SIM-002", "会话ID": "SIM-THREAD"})
        result = rules.import_message(self.creators, self.messages,
                                      {**self.incoming, "达人ID": "SIM-001", "会话ID": "SIM-THREAD"})
        self.assertTrue(result["ambiguous"])

    def test_reply_invalidates_all_unsent_followups_and_confirmation(self):
        initial_contact(self.creators, self.messages)
        followup = rules.generate_followup_drafts(self.creators, self.messages, "2026-10-19")[0]
        rules.confirm_batch(self.messages, [followup["邮件记录ID"]], "SIM-FOLLOWUP", "2026-10-19")
        self.assertTrue(rules.confirmation_valid(followup))
        rules.import_message(self.creators, self.messages, self.incoming)
        self.assertEqual(followup["处理状态"], "已作废")
        self.assertFalse(rules.confirmation_valid(followup))
        self.assertEqual(followup["发送确认批次"], "")
        self.assertEqual(self.creators[0]["下一步动作"], "回复草稿")
        self.assertEqual(self.creators[0]["计划跟进日期"], "")
        self.assertEqual(rules.generate_followup_drafts(self.creators, self.messages, "2026-11-01"), [])


class DraftTests(unittest.TestCase):
    def setUp(self):
        self.creators, self.messages = [creator()], []

    def draft(self):
        return rules.create_draft(self.creators, self.messages, "SIM-001", "Simulation", "Simulation body", "首联/v1")

    def test_draft_time_is_not_mail_time(self):
        draft = self.draft()
        self.assertEqual(draft["邮件时间"], "")
        self.assertTrue(draft["草稿生成时间"])
        self.assertEqual(draft["草稿用途"], "首联")
        self.assertFalse(rules.confirmation_valid(draft))

    def test_mutating_recipient_subject_or_body_invalidates_confirmation(self):
        draft = self.draft()
        for field, value in (("收件人", "changed@example.invalid"), ("主题", "Changed subject"),
                             ("正文或草稿", "Changed body")):
            rules.confirm_batch(self.messages, [draft["邮件记录ID"]], "SIM-BATCH", "2026-10-19")
            self.assertTrue(rules.confirmation_valid(draft))
            version = draft["草稿版本"]
            rules.update_draft(draft, **{field: value})
            self.assertFalse(rules.confirmation_valid(draft))
            self.assertEqual(draft["草稿版本"], version + 1)
            self.assertEqual(draft["发送确认批次"], "")

    def test_external_edit_detected_and_reconfirmation_versions_it(self):
        draft = self.draft()
        rules.confirm_batch(self.messages, [draft["邮件记录ID"]], "SIM-BATCH", "2026-10-19")
        draft["正文或草稿"] = "Edited in workbook"
        self.assertFalse(rules.confirmation_valid(draft))
        rules.confirm_batch(self.messages, [draft["邮件记录ID"]], "SIM-BATCH-2", "2026-10-19")
        self.assertEqual(draft["草稿版本"], 2)
        self.assertTrue(rules.confirmation_valid(draft))

    def test_confirm_batch_is_atomic_and_selected_ids_only(self):
        first = self.draft()
        self.creators.append(creator("SIM-002", "second@example.invalid", "sim-two"))
        second = rules.create_draft(self.creators, self.messages, "SIM-002", "Simulation", "Simulation body", "首联/v1")
        second["收件人"] = "https://example.invalid/contact"
        with self.assertRaises(ValueError):
            rules.confirm_batch(self.messages, [first["邮件记录ID"], second["邮件记录ID"]], "SIM-BATCH", "2026-10-19")
        self.assertEqual(first["发送确认批次"], "")
        rules.confirm_batch(self.messages, [first["邮件记录ID"]], "SIM-BATCH", "2026-10-19")
        self.assertTrue(rules.confirmation_valid(first))
        self.assertFalse(rules.confirmation_valid(second))

    def test_result_requires_current_confirmation(self):
        draft = self.draft()
        with self.assertRaises(ValueError):
            rules.record_send_result(self.creators, self.messages, draft["邮件记录ID"], "成功", "2026-10-19")
        rules.confirm_batch(self.messages, [draft["邮件记录ID"]], "SIM-BATCH", "2026-10-19")
        rules.record_send_result(self.creators, self.messages, draft["邮件记录ID"], "失败", "2026-10-19")
        self.assertEqual(draft["处理状态"], "发送失败")
        self.assertFalse(rules.confirmation_valid(draft))

    def test_unknown_result_blocks_retry_confirm_followup_until_resolved(self):
        draft = self.draft()
        rules.confirm_batch(self.messages, [draft["邮件记录ID"]], "SIM-BATCH", "2026-10-01")
        rules.record_send_result(self.creators, self.messages, draft["邮件记录ID"], "结果待核实", "2026-10-01")
        self.assertFalse(rules.confirmation_valid(draft))
        with self.assertRaises(ValueError):
            rules.confirm_batch(self.messages, [draft["邮件记录ID"]], "SIM-RETRY", "2026-10-19")
        with self.assertRaises(ValueError):
            self.draft()
        with self.assertRaises(ValueError):
            rules.update_draft(draft, **{"正文或草稿": "Retry"})
        self.assertEqual(rules.generate_followup_drafts(self.creators, self.messages, "2026-10-19"), [])
        rules.record_send_result(self.creators, self.messages, draft["邮件记录ID"], "成功", "2026-10-01")
        self.assertEqual(draft["实际发送结果"], "成功")

    def test_unknown_result_requires_original_confirmed_content_and_confirmation(self):
        for field, value in (("收件人", "changed@example.invalid"), ("主题", "Edited subject"),
                             ("正文或草稿", "Edited body"), ("发送确认批次", ""), ("确认时间", "")):
            with self.subTest(field=field):
                creators, messages = [creator()], []
                draft = rules.create_draft(creators, messages, "SIM-001", "Simulation", "Original body", "首联/v1")
                rules.confirm_batch(messages, [draft["邮件记录ID"]], "SIM-ORIGINAL", "2026-10-01")
                rules.record_send_result(creators, messages, draft["邮件记录ID"], "结果待核实", "2026-10-01")
                # Simulate a direct workbook edit bypassing update_draft.
                draft[field] = value
                for outcome in ("成功", "失败"):
                    with self.assertRaises(ValueError):
                        rules.record_send_result(creators, messages, draft["邮件记录ID"], outcome, "2026-10-02")
                    self.assertEqual(draft["实际发送结果"], "结果待核实")
                    self.assertEqual(draft["处理状态"], "发送结果待核实")
                    self.assertEqual(draft["邮件时间"], "")
                with self.assertRaises(ValueError):
                    rules.confirm_batch(messages, [draft["邮件记录ID"]], "SIM-RETRY", "2026-10-02")

    def test_day_seven_workday_and_once_only(self):
        initial_contact(self.creators, self.messages)
        self.assertEqual(rules.generate_followup_drafts(self.creators, self.messages, "2026-10-11"), [])
        generated = rules.generate_followup_drafts(self.creators, self.messages, "2026-10-12")
        self.assertEqual(len(generated), 1)
        self.assertEqual(generated[0]["邮件时间"], "")
        self.assertEqual(generated[0]["草稿用途"], "跟进")
        self.assertEqual(generated[0]["发送确认批次"], "")
        self.assertEqual(rules.generate_followup_drafts(self.creators, self.messages, "2026-10-19"), [])

    def test_unknown_history_and_second_outbound_do_not_trigger(self):
        self.messages.append({"邮件记录ID": "MAIL-000001", "达人ID": "SIM-001", "方向": "发信",
                              "邮件时间": "2026-10-01", "实际发送结果": "成功", "处理状态": "已发送"})
        self.assertEqual(rules.generate_followup_drafts(self.creators, self.messages, "2026-10-19"), [])
        self.messages.clear()
        initial_contact(self.creators, self.messages)
        self.messages.append({"邮件记录ID": "MAIL-000002", "达人ID": "SIM-001", "方向": "发信",
                              "邮件时间": "2026-10-05", "处理状态": "已发送"})
        self.assertEqual(rules.generate_followup_drafts(self.creators, self.messages, "2026-10-19"), [])

    def test_pause_invalid_email_missing_time_and_reply_block_followup(self):
        for change in ({"下一步动作": "暂停联系"}, {"Email": "https://example.invalid/contact"},
                       {"最近收信时间": "2026-10-02"}, {"联系状态": "已回复"}):
            creators, messages = [creator()], []
            initial_contact(creators, messages)
            creators[0].update(change)
            self.assertEqual(rules.generate_followup_drafts(creators, messages, "2026-10-19"), [])
        initial = initial_contact(self.creators, self.messages)
        initial["邮件时间"] = ""
        self.assertEqual(rules.generate_followup_drafts(self.creators, self.messages, "2026-10-19"), [])

    def test_undated_reply_blocks_followup_and_direct_followup_creation(self):
        initial_contact(self.creators, self.messages)
        rules.import_message(self.creators, self.messages,
                             {"方向": "收信", "邮件时间": "", "发件人": "creator@example.invalid",
                              "收件人": "team@example.invalid", "主题": "Simulation", "正文或草稿": "Reply"})
        self.assertEqual(rules.generate_followup_drafts(self.creators, self.messages, "2026-10-19"), [])
        with self.assertRaises(ValueError):
            rules.create_draft(self.creators, self.messages, "SIM-001", "Simulation", "Followup", "跟进/v1", purpose="跟进")

    def test_rejected_do_not_contact_and_completed_block_all_drafts(self):
        for status in ("拒绝", "勿联系", "已完成", "已拒绝", "不合作"):
            creators, messages = [creator()], []
            initial_contact(creators, messages)
            creators[0]["联系状态"] = status
            self.assertEqual(rules.generate_followup_drafts(creators, messages, "2026-10-19"), [])
            for purpose in ("首联", "回复", "跟进"):
                with self.assertRaises(ValueError):
                    rules.create_draft(creators, messages, "SIM-001", "Simulation", "Body", "模拟/v1", purpose=purpose)

    def test_initial_draft_requires_usable_classification_and_uncontacted_creator(self):
        for change in ({"分类核实状态": "待核实"}, {"首联内容分类": ""},
                       {"首联内容分类": "待核实"}, {"联系状态": "已联系"}, {"最近发信时间": "2026-10-01"}):
            creators, messages = [creator()], []
            creators[0].update(change)
            with self.assertRaises(ValueError):
                rules.create_draft(creators, messages, "SIM-001", "Simulation", "Body", "首联/v1")
            self.assertEqual(messages, [])
        # An initial classification can be reviewed together with the draft.
        self.assertEqual(self.draft()["处理状态"], "待确认草稿")

    def test_initial_draft_is_never_created_twice_or_after_unknown_outbound(self):
        self.draft()
        with self.assertRaises(ValueError):
            self.draft()
        self.assertEqual(len(self.messages), 1)
        self.messages[:] = [{"邮件记录ID": "MAIL-000001", "达人ID": "SIM-001", "方向": "发信",
                             "邮件时间": "2026-10-01", "处理状态": "待核实"}]
        with self.assertRaises(ValueError):
            self.draft()
        self.assertEqual(len(self.messages), 1)

    def test_reply_draft_is_idempotent_and_edited_in_place(self):
        incoming = {"邮件记录ID": "MAIL-000001", "达人ID": "SIM-001", "方向": "收信", "会话ID": "SIM-THREAD"}
        self.messages.append(incoming)
        first = rules.create_draft(self.creators, self.messages, "SIM-001", "Simulation", "Reply body", "回复/v1",
                                   purpose="回复", in_reply_to="MAIL-000001")
        duplicate = rules.create_draft(self.creators, self.messages, "SIM-001", "Simulation", "Reply body", "回复/v1",
                                       purpose="回复", in_reply_to="MAIL-000001")
        self.assertIs(first, duplicate)
        self.assertEqual(first["草稿版本"], 1)
        rules.confirm_batch(self.messages, [first["邮件记录ID"]], "SIM-REPLY-BATCH", "2026-10-19")
        changed = rules.create_draft(self.creators, self.messages, "SIM-001", "Changed subject", "Changed reply", "回复/v2",
                                     purpose="回复", in_reply_to="MAIL-000001")
        self.assertIs(first, changed)
        self.assertEqual(len(self.messages), 2)
        self.assertEqual(changed["草稿版本"], 2)
        self.assertFalse(rules.confirmation_valid(changed))


if __name__ == "__main__":
    unittest.main()

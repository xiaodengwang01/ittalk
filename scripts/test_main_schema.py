"""Synthetic checks for the simplified creator workspace."""
import copy
from datetime import date, datetime
import unittest

try:
    from . import main_schema as schema
except ImportError:
    import main_schema as schema


def creator(**values):
    return {"达人ID": "SIM-000001", "达人昵称": "模拟达人", **values}


def incoming(identifier="SIM-MAIL-1", timestamp=None, **values):
    return {"邮件记录ID": identifier, "达人ID": "SIM-000001", "方向": "收信",
            "邮件时间": timestamp, "邮件用途": "收信", "邮件正文": "Original\n\nNew text.",
            "处理状态": "待回复", **values}


def outgoing(identifier="SIM-MAIL-2", timestamp=None, **values):
    return incoming(identifier, timestamp, **{"方向": "发信", "邮件用途": "回复",
                                             "实际发送结果": "成功", "处理状态": "已发送", **values})


class ContactStages(unittest.TestCase):
    def test_exact_user_options(self):
        self.assertEqual(schema.CONTACT_STAGES, ("初次建联已发送", "对方第一次回复", "我方第一次回复",
                                                "对方第二次回复", "我方第二次回复", "长期有效沟通"))

    def test_legacy_following_does_not_claim_initial_send(self):
        row = creator(**{"联系状态": "跟进中", "首次联系日期": "2026-10-09"})
        schema.migrate_creator(row, [])
        self.assertEqual(row["联系状态"], "")
        self.assertIn("跟进中", row["_notes"]["备注"])
        self.assertNotIn("首联发送结果", row)

    def test_verified_initial_metadata_is_evidence(self):
        row = creator(**{"首联发送结果": "成功", "首次联系日期": "2026-10-09T08:00:00"})
        schema.migrate_creator(row, [])
        self.assertEqual(row["联系状态"], "初次建联已发送")
        self.assertEqual(row["首联发送时间"], "2026-10-09T08:00:00")

    def test_any_number_of_incoming_without_sent_is_first_turn(self):
        messages = [incoming("SIM-MAIL-" + str(i)) for i in range(3)]
        self.assertEqual(schema.derive_stage(creator(), messages), "对方第一次回复")

    def test_dated_alternating_direction_transitions(self):
        messages = [incoming(timestamp="2026-10-01T10:00:00"),
                    outgoing(timestamp="2026-10-02T10:00:00"),
                    incoming("SIM-MAIL-3", "2026-10-03T10:00:00"),
                    outgoing("SIM-MAIL-4", "2026-10-04T10:00:00"),
                    incoming("SIM-MAIL-5", "2026-10-05T10:00:00")]
        for length, expected in enumerate(schema.CONTACT_STAGES[1:5] + ("我方第二次回复",), 1):
            self.assertEqual(schema.derive_stage(creator(), list(reversed(messages[:length]))), expected)

    def test_further_mail_cannot_automatically_assert_long_term_communication(self):
        messages = [(incoming if i % 2 == 0 else outgoing)(
            "SIM-MAIL-" + str(i), "2026-10-{:02d}T10:00:00".format(i + 1)) for i in range(10)]
        self.assertEqual(schema.derive_stage(creator(), messages), "我方第二次回复")
        self.assertEqual(schema.derive_stage(creator(**{"联系状态": "长期有效沟通"}), messages), "长期有效沟通")

    def test_consecutive_messages_form_one_turn(self):
        messages = [incoming(timestamp="2026-10-01"), incoming("SIM-MAIL-3", "2026-10-02"),
                    outgoing(timestamp="2026-10-03"), outgoing("SIM-MAIL-4", "2026-10-04")]
        self.assertEqual(schema.derive_stage(creator(), messages), "我方第一次回复")

    def test_unknown_mixed_mail_does_not_guess_order(self):
        messages = [outgoing(), incoming()]
        self.assertEqual(schema.derive_stage(creator(), messages), "我方第一次回复")
        self.assertEqual(schema.derive_stage(creator(), list(reversed(messages))), "我方第一次回复")
        row = creator()
        schema.migrate_creator(row, messages)
        self.assertIn("未推定未知邮件的先后", row["_notes"]["联系状态"])

    def test_unknown_full_reply_chain_can_order(self):
        messages = [incoming(), outgoing(**{"回复邮件记录ID": "SIM-MAIL-1"}),
                    incoming("SIM-MAIL-3", **{"回复邮件记录ID": "SIM-MAIL-2"})]
        self.assertEqual(schema.derive_stage(creator(), list(reversed(messages))), "对方第二次回复")

    def test_branching_chain_does_not_claim_round_two(self):
        messages = [incoming(), outgoing(**{"回复邮件记录ID": "SIM-MAIL-1"}),
                    outgoing("SIM-MAIL-3", **{"回复邮件记录ID": "SIM-MAIL-1"})]
        self.assertEqual(schema.derive_stage(creator(), messages), "我方第一次回复")

    def test_incompatible_known_time_blocks_explicit_chain(self):
        messages = [incoming(timestamp="2026-10-03"), outgoing(**{"回复邮件记录ID": "SIM-MAIL-1"}),
                    incoming("SIM-MAIL-3", "2026-10-01", **{"回复邮件记录ID": "SIM-MAIL-2"})]
        self.assertEqual(schema.derive_stage(creator(), messages), "我方第一次回复")

    def test_same_day_opposing_mail_has_no_guessed_order(self):
        messages = [incoming(timestamp="2026-10-01"), outgoing(timestamp="2026-10-01")]
        self.assertEqual(schema.derive_stage(creator(), messages), "我方第一次回复")

    def test_later_manual_stage_survives_incomplete_history(self):
        for stage in schema.CONTACT_STAGES[2:]:
            self.assertEqual(schema.derive_stage(creator(**{"联系状态": stage}), [incoming()]), stage)

    def test_drafts_failed_first_contact_and_header_only_not_counted(self):
        rejected = [outgoing(**{"实际发送结果": ""}), outgoing(**{"实际发送结果": "失败"}),
                    outgoing(**{"处理状态": "已确认待发送"}), outgoing(**{"邮件用途": "首联"}),
                    outgoing(**{"邮件正文": "From: sim@example.invalid\nDate: unknown\nTo: x@example.invalid"})]
        self.assertEqual(schema.derive_stage(creator(), rejected), "")
        self.assertEqual(schema.derive_stage(creator(), [incoming()] + rejected), "对方第一次回复")

    def test_other_creators_and_unassigned_mail_do_not_count(self):
        messages = [incoming(**{"达人ID": ""}), incoming("SIM-MAIL-9", **{"达人ID": "SIM-000002"})]
        self.assertEqual(schema.derive_stage(creator(), messages), "")

    def test_standalone_refresh_preserves_do_not_contact_limit(self):
        row = creator(**{"联系状态": "拒绝"})
        schema.refresh_progress([row], [incoming()])
        self.assertEqual(row["联系状态"], "对方第一次回复")
        self.assertTrue(schema.contact_blocked(row))
        self.assertIn("拒绝", row["_notes"]["备注"])

    def test_paused_action_survives_removal_as_contact_limit(self):
        row = creator(**{"下一步动作": "暂停联系"})
        schema.migrate_creator(row, [])
        self.assertTrue(schema.contact_blocked(row))
        self.assertNotIn("下一步动作", row)


class ColumnMigration(unittest.TestCase):
    def test_removed_values_and_their_notes_are_rehomed(self):
        row = creator(**{key: "旧值-" + key for key in schema.REMOVED_MASTER_FIELDS})
        row["_notes"] = {key: "原注\n" + key for key in schema.REMOVED_MASTER_FIELDS}
        row["_notes"]["备注"] = "保留原注"
        schema.migrate_creator(row, [])
        for key in schema.REMOVED_MASTER_FIELDS:
            self.assertNotIn(key, row)
            self.assertNotIn(key, row["_notes"])
            self.assertTrue(any("原注\n" + key in value for value in row["_notes"].values()))
            if key != "沟通摘要状态":
                self.assertTrue(any("旧值-" + key in value for value in row["_notes"].values()))
        self.assertIn("保留原注", row["_notes"]["备注"])

    def test_legacy_dates_merged_and_exact_raw_values_preserved(self):
        row = creator(**{"首次联系日期": "2026-10-09T08:00:00", "历史跟进日期": "2026-10-12T08:00:00"})
        schema.migrate_creator(row, [])
        self.assertEqual(row["最近联系"], "2026-10-12")
        self.assertIn("2026-10-09T08:00:00", row["_notes"]["最近联系"])
        self.assertIn("2026-10-12T08:00:00", row["_notes"]["最近联系"])
        self.assertIn("不触发自动跟进", row["_notes"]["最近联系"])
        self.assertNotIn("首联发送结果", row)

    def test_latest_actual_mail_date_preserves_time_precision(self):
        row = creator(**{"首次联系日期": "2026-10-09", "历史跟进日期": "2026-10-12"})
        schema.migrate_creator(row, [incoming(timestamp="2026-10-13T16:28:43+08:00")])
        self.assertEqual(row["最近联系"], "2026-10-13T16:28:43+08:00")

    def test_latest_from_recent_recv_or_sent(self):
        row = creator(**{"最近收信时间": "2026-10-13T10:00:00", "最近发信时间": "2026-10-14T11:00:00"})
        schema.refresh_progress([row], [incoming(timestamp="2026-10-12T19:00:00")])
        self.assertEqual(row["最近联系"], "2026-10-14T11:00:00")

    def test_timezone_comparison_uses_one_clock(self):
        row = creator(**{"最近收信时间": "2026-10-13T23:00:00+00:00", "最近发信时间": "2026-10-14T06:30:00+08:00"})
        schema.refresh_progress([row], [])
        self.assertEqual(row["最近联系"], "2026-10-13T23:00:00+00:00")

    def test_unknown_time_not_filled_from_quoted_or_relative_date(self):
        row = creator()
        schema.migrate_creator(row, [incoming(**{"邮件原始时间": "10月13日", "核实说明": "引用发信头2026-10-13"})])
        self.assertFalse(row.get("最近联系"))

    def test_invalid_legacy_date_is_not_guessed(self):
        row = creator(**{"首次联系日期": "10月13日"})
        schema.migrate_creator(row, [])
        self.assertFalse(row.get("最近联系"))
        self.assertIn("10月13日", row["_notes"]["最近联系"])

    def test_date_and_datetime_objects_accepted(self):
        row = creator(**{"首次联系日期": date(2026, 10, 9)})
        schema.migrate_creator(row, [incoming(timestamp=datetime(2026, 10, 13, 9, 7, 6))])
        self.assertEqual(row["最近联系"], "2026-10-13T09:07:06")

    def test_drafts_cannot_move_last_contact_date(self):
        row = creator(**{"最近联系": "2026-10-01"})
        schema.refresh_progress([row], [outgoing(timestamp="2026-10-15", **{"实际发送结果": "", "处理状态": "草稿"})])
        self.assertEqual(row["最近联系"], "2026-10-01")

    def test_unknown_extra_columns_ids_and_mail_bodies_unchanged(self):
        row = creator(**{"自定义列": "保留", "草稿状态": "未发送草稿", "当前回复草稿": "Draft\n\nText"})
        mail = [incoming()]
        before = copy.deepcopy(mail)
        schema.migrate_creator(row, mail)
        self.assertEqual(row["达人ID"], "SIM-000001")
        self.assertEqual(row["自定义列"], "保留")
        self.assertEqual(row["当前回复草稿"], "Draft\n\nText")
        self.assertEqual(mail, before)

    def test_migrate_then_refresh_is_idempotent_and_no_deleted_columns_return(self):
        row = creator(**{"联系状态": "已回复", "首次联系日期": "2026-10-09T08:00:00",
                         "历史跟进日期": "2026-10-12T08:00:00", "来源": "旧来源"})
        messages = [incoming()]
        schema.migrate_creator(row, messages)
        before = copy.deepcopy(row)
        for _ in range(3):
            schema.migrate_creator(row, messages)
            schema.refresh_progress([row], messages)
        self.assertEqual(row, before)
        self.assertFalse(schema.REMOVED_MASTER_FIELDS.intersection(row))


if __name__ == "__main__":
    unittest.main()

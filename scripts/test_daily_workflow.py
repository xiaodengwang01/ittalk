"""Synthetic tests for the paused daily workflow and current workbook layout."""
import copy
from datetime import datetime
import json
import os
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import patch
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from scripts import daily_workflow as d
from scripts.test_mail_drafts import write_template

HEADERS = ["达人ID", "达人昵称", "联系状态", "沟通摘要", "Email", "首联内容分类", "初次建联日期",
           "对方第一次回信日期", "对方第一次回复邮件原始文本", "对方首次回信摘要",
           "对方回复风格判断和应对策略", "我方第一次回复邮件日期", "我们的第一次回复邮件原始文本",
           "我们的第一次回复邮件摘要", "备注"] + [""] * 19 + [
           "应对策略下拉选项", "账号链接", "平台", "宠物类型", "内容类型", "粉丝数",
           "近10条最高播放", "近10条中位播放", "内容能力(1-5)", "宠物人群匹配度(1-5)",
           "播放能力(1-5)", "联系方式完整度(1-5)", "综合评分", "是否低粉高播放", "是否带货类似产品"]
NOW = datetime.fromisoformat("2026-10-22T09:10:00+08:00")


def candidate(n=2):
    return {"达人昵称": f"模拟达人 {n}", "Email": f"sim{n}@example.invalid",
            "账号链接": f"https://www.tiktok.com/@sim_{n}?tracking=1", "首联内容分类": d.mail.CATEGORIES[n % 3],
            "宠物类型": "Dog", "内容类型": "狗狗剧情", "英语市场依据": "模拟英语内容",
            "内容判断": "模拟完整剧情", "宠物匹配判断": "模拟养狗观众",
            "播放判断": "模拟近期稳定播放", "活跃判断": "模拟本周更新", "合格": True,
            "发现路径": "TikTok", "邮箱公开来源": f"https://example.invalid/sim{n}/contact",
            "数据来源": f"https://example.invalid/sim{n}", "核实时间": "2026-10-22T09:05:00+08:00",
            "已联系": False, "历史核对依据": "模拟邮箱中搜索此地址及账号，无历史联系或草稿",
            "历史核对时间": "2026-10-22T09:05:00+08:00",
            "近期视频": [{"url": f"https://example.invalid/sim{n}/video{i}", "views": 1000+i*100,
                       "date": "2026-10-21"} for i in range(10)],
            "粉丝数": 500, "内容能力(1-5)": 4, "宠物人群匹配度(1-5)": 5,
            "播放能力(1-5)": 4, "联系方式完整度(1-5)": 5,
            "是否低粉高播放": "是", "是否低粉高播放依据": "模拟中位播放/粉丝比",
            "是否带货类似产品": ""}


def fixture(path):
    cells = {f"{d.workbook.workflow.col_name(i)}1": {"value": v} for i, v in enumerate(HEADERS, 1) if v}
    values = {"A2": "SIM-000001", "B2": "模拟达人 One", "C2": "初次建联未回复",
              "E2": "sim1@example.invalid", "F2": "宠物剧情类", "I2": "=literal@example.invalid",
              "AJ2": "https://www.tiktok.com/@sim_1", "AK2": "TikTok"}
    cells.update({k: {"value": v} for k, v in values.items()})
    spec = path.with_suffix(".json")
    spec.write_text(json.dumps(cells))
    node = Path.home() / ".cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node"
    subprocess.run([str(node), str(Path(d.__file__).with_name("daily_cells.mjs")), str(spec), str(path),
                    "--fixture"], check=True, capture_output=True, timeout=120)
    with ZipFile(path) as z:
        files = {i.filename: (i, z.read(i.filename)) for i in z.infolist()}
    name = "xl/worksheets/sheet1.xml"
    raw = ET.tostring(ET.fromstring(files[name][1]), encoding="unicode")
    raw = raw.replace("</sheetData>", '<row r="505"><c r="A505" s="0"/></row></sheetData>')
    raw = raw.replace("</worksheet>", '<dataValidations count="1"><dataValidation type="list" sqref="C2:C2">'
                      '<formula1>"初次建联未回复,对方第一次回复"</formula1></dataValidation></dataValidations></worksheet>')
    files[name] = (files[name][0], raw.encode())
    with ZipFile(path, "w") as z:
        for i, content in files.values():
            z.writestr(i, content)


class DailyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.shared = tempfile.TemporaryDirectory()
        cls.base = Path(cls.shared.name) / "fixture.xlsx"
        fixture(cls.base)

    @classmethod
    def tearDownClass(cls):
        cls.shared.cleanup()

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)
        self.book = self.folder / "book.xlsx"
        self.book.write_bytes(self.base.read_bytes())
        self.template = self.folder / "template.docx"
        write_template(self.template)
        self.state = d.mail.State(self.folder / "private/state.sqlite3")
        d.initialize(self.state)
        self.clock = patch.object(d, "current_time", return_value=NOW)
        self.clock.start()

    def tearDown(self):
        self.clock.stop()
        self.state.close()
        self.temp.cleanup()

    def active(self):
        d.put(self.state, "activation", {"status": "active", "first_run": "2026-10-22T09:00:00+08:00",
                                      "evidence": {"sent_contacts": []}})

    def add(self, creators=None):
        return d.add_candidates(self.state, "2026-10-22",
                                {"routes_checked": ["FastMoss", "TikTok"], "creators": creators or [candidate()]},
                                book=self.book)

    def test_paused_and_no_backfill(self):
        before = self.book.read_bytes()
        with self.assertRaises(d.mail.DraftError):
            self.add()
        self.active()
        with self.assertRaises(d.mail.DraftError):
            d.day_state(self.state, "2026-10-21")
        with patch.object(d, "current_time", return_value=NOW.replace(hour=8)):
            with self.assertRaises(d.mail.DraftError):
                d.day_state(self.state, "2026-10-22")
        self.assertEqual(before, self.book.read_bytes())

    def test_activation_requires_all_conditions_and_next_nine(self):
        paths = dict(book=self.book, template=self.template, strategy=self.template, reference=self.template)
        evidence = {**d.preflight(self.state, **paths), "pipeline_ready": True, "backlog_cleared": True,
                    "pipeline": "模拟发送链路", "account": "team@example.invalid", "unresolved": [],
                    "sent_contacts": [], "backlog_receipts": {}, "checked_at": NOW.isoformat(),
                    "user_instruction": "模拟：两项已完成，可以启动", "user_instruction_at": NOW.isoformat()}
        for key in ("pipeline_ready", "backlog_cleared", "user_instruction"):
            bad = dict(evidence, **{key: False})
            with self.assertRaises(d.mail.DraftError):
                d.activate(self.state, bad, **paths)
        with self.assertRaises(d.mail.DraftError):
            d.activate(self.state, {**evidence, "sources": {}}, **paths)
        result = d.activate(self.state, evidence, **paths)
        self.assertEqual(result["first_run"], "2026-10-23T09:00:00+08:00")

    def manual_evidence(self):
        self.state.set_config({"account": "team@example.invalid"})
        paths = dict(book=self.book, template=self.template, strategy=self.template, reference=self.template)
        evidence = {**d.preflight(self.state, **paths), "scope": "prepare_only",
                    "account": "team@example.invalid", "history_folders": ["INBOX", "Sent", "Drafts"],
                    "unresolved": [], "sent_contacts": [], "checked_at": NOW.isoformat(),
                    "user_instruction": "模拟：只执行今天的资料搜集和首联预览", "user_instruction_at": NOW.isoformat()}
        return evidence, paths

    def test_manual_day_preserves_pause_progress_and_send_approval(self):
        evidence, paths = self.manual_evidence()
        evidence["sent_contacts"] = [{"email": "sim3@example.invalid"}]
        d.start_today(self.state, "2026-10-22", evidence, **paths)
        result = self.add([candidate(2), candidate(3)])
        self.assertEqual(len(result["added"]), 1)
        self.assertEqual(len(result["rejected"]), 1)
        again = d.start_today(self.state, "2026-10-22", evidence, **paths)
        self.assertEqual(again["creators"], result["added"])
        self.assertIsNone(d.get(self.state, "activation"))
        prepared = d.prepare_day(self.state, "2026-10-22", book=self.book, template=self.template)
        with self.assertRaises(d.mail.DraftError):
            d.claim_send(self.state, prepared["batch"])
        with patch.object(d, "current_time", return_value=NOW.replace(day=23)):
            with self.assertRaises(d.mail.DraftError):
                d.day_state(self.state, "2026-10-23")

    def test_manual_day_rejects_missing_or_stale_authorization(self):
        evidence, paths = self.manual_evidence()
        for change in ({"scope": "send"}, {"account": "other@example.invalid"},
                       {"user_instruction": ""}, {"user_instruction_at": "2026-10-21T09:00:00+08:00"},
                       {"checked_at": "2026-10-21T09:00:00+08:00"}, {"history_folders": ["Sent"]},
                       {"sources": {}}, {"sent_contacts": None}):
            with self.subTest(change=change), self.assertRaises(d.mail.DraftError):
                d.start_today(self.state, "2026-10-22", {**evidence, **change}, **paths)
        self.assertIsNone(d.get(self.state, "day:2026-10-22"))
        d.start_today(self.state, "2026-10-22", evidence, **paths)
        with self.assertRaises(d.mail.DraftError):
            d.start_today(self.state, "2026-10-22", {**evidence, "user_instruction": "模拟替换"}, **paths)
        self.state.set_config({"account": "other@example.invalid"})
        with self.assertRaises(d.mail.DraftError):
            d.day_state(self.state, "2026-10-22")

    def test_relative_video_date_preserves_visible_precision(self):
        item = candidate()
        post = item["近期视频"][0]
        del post["date"]
        post.update(date_text="1 周前", observed_at=NOW.isoformat())
        result = d.validate_candidate(item, "2026-10-22")
        self.assertNotIn("date", result["近期视频"][0])
        for change in ({"date_text": "未知"}, {"observed_at": "2026-10-21T09:00:00+08:00"},
                       {"observed_at": "2026-10-23T09:00:00+08:00"}):
            invalid = copy.deepcopy(item)
            invalid["近期视频"][0].update(change)
            with self.subTest(change=change), self.assertRaises(d.mail.DraftError):
                d.validate_candidate(invalid, "2026-10-22")

    def test_explicit_supplement_preserves_sent_round_and_total_limit(self):
        evidence, paths = self.manual_evidence()
        d.start_today(self.state, "2026-10-22", evidence, **paths)
        self.add([candidate(n) for n in (2, 3, 4)])
        first = d.prepare_day(self.state, "2026-10-22", self.book, self.template)
        old_items = self.state.items(first["batch"])
        instruction = {"user_instruction": "模拟确认第一批", "user_instruction_at": NOW.isoformat(),
                       "content_hashes": first["content_hashes"]}
        d.approve_batch(self.state, first["batch"], instruction)
        d.claim_send(self.state, first["batch"])
        extra, paths = self.manual_evidence()
        extra["user_instruction"] = "模拟：继续搜集今天剩下五位，仅准备完整预览"
        with self.assertRaises(d.mail.DraftError):
            d.start_supplement(self.state, "2026-10-22", extra, **paths)
        receipts = {item["id"]: {"result": "success", "verification": "sent_folder",
                    "message_id": "SIM-SENT-" + item["id"], "source": "https://example.invalid/sent/" + item["id"],
                    "sent_at": NOW.isoformat(), "recipient": item["recipient"], "content_hash": item["content_hash"]}
                    for item in old_items}
        d.record_sent(self.state, first["batch"], receipts, self.book)
        frozen = self.state.items(first["batch"])
        extra, paths = self.manual_evidence()
        extra["user_instruction"] = "模拟：继续搜集今天剩下五位，仅准备完整预览"
        extra["sent_contacts"] = [{"email": "sim10@example.invalid"}]
        value = d.start_supplement(self.state, "2026-10-22", extra, **paths)
        self.assertEqual(value["round_start"], 3)
        self.assertEqual(value["batch"], "tailki-20261022-part2")
        self.assertEqual(d.prepare_day(self.state, "2026-10-22", self.book, self.template)["batch"], None)
        self.add([candidate(n) for n in range(5, 11)])
        self.assertEqual(len(d.start_supplement(self.state, "2026-10-22", extra, **paths)["creators"]), 8)
        second = d.prepare_day(self.state, "2026-10-22", self.book, self.template)
        self.assertEqual(second["shortfall"], 0)
        self.assertEqual(len(self.state.items(second["batch"])), 5)
        self.assertEqual(self.state.items(first["batch"]), frozen)
        self.assertIsNone(d.get(self.state, "approval:" + second["batch"]))
        self.assertIsNone(d.get(self.state, "activation"))
        with self.assertRaises(d.mail.DraftError):
            d.claim_send(self.state, second["batch"])
        with self.assertRaises(d.mail.DraftError):
            d.start_supplement(self.state, "2026-10-22", {**extra, "user_instruction": "模拟再次补充"}, **paths)

    def test_append_preserves_parts_literal_history_and_deduplicates(self):
        self.active()
        with ZipFile(self.book) as z:
            original = {n: z.read(n) for n in z.namelist()}
        result = self.add()
        self.assertEqual(len(result["added"]), 1)
        self.assertEqual(result["shortfall"], 7)
        sheet = d.workbook.read(self.book)
        self.assertEqual(sheet["creators"][-1]["_row"], 3)
        self.assertEqual(sheet["creators"][-1]["联系状态"], "待首联确认")
        self.assertFalse(sheet["creators"][-1].get("初次建联日期"))
        self.assertEqual(sheet["creators"][-1]["综合评分"], 4.5)
        self.assertEqual(sheet["creators"][0]["对方第一次回复邮件原始文本"], "=literal@example.invalid")
        with ZipFile(self.book) as z:
            for name, content in original.items():
                if name != "xl/worksheets/sheet1.xml":
                    self.assertEqual(content, z.read(name), name)
            text = z.read("xl/worksheets/sheet1.xml").decode()
            self.assertIn('r="505"', text)
            self.assertIn("待首联确认,", text)
        saved = self.book.read_bytes()
        self.add()
        self.assertEqual(saved, self.book.read_bytes())

    def test_only_research_columns_can_be_written_and_crash_recovery(self):
        self.active()
        person = candidate()
        person["对方第一次回复邮件原始文本"] = "must not write"
        self.add([person])
        row = d.workbook.read(self.book)["creators"][-1]
        self.assertFalse(row.get("对方第一次回复邮件原始文本"))
        value = d.get(self.state, "day:2026-10-22")
        reserved = {**d.validate_candidate(candidate(), "2026-10-22"), "达人ID": row["达人ID"]}
        value["reserved"], value["creators"] = [reserved], []
        d.put(self.state, "day:2026-10-22", value)
        saved = self.book.read_bytes()
        d.recover_reserved(self.state, value, self.book)
        self.assertEqual(saved, self.book.read_bytes())
        self.assertEqual(len(value["creators"]), 1)

    def test_quality_gaps_and_eight_limit(self):
        self.active()
        people = [candidate(i) for i in range(2, 12)]
        people[0]["Email"] = ""
        people[1]["内容判断"] = ""
        result = self.add(people)
        self.assertEqual(len(result["added"]), 8)
        self.assertEqual(len(result["rejected"]), 2)
        self.assertEqual(len(d.workbook.read(self.book)["creators"]), 9)
        self.assertEqual(len(self.add([candidate(20)])["added"]), 8)

    def test_preview_is_idempotent_and_send_receipt_is_required(self):
        self.active()
        self.add()
        result = d.prepare_day(self.state, "2026-10-22", self.book, self.template)
        batch = result["batch"]
        one = self.state.items(batch)[0]
        saved = self.book.read_bytes()
        before_note = d.workbook.read(self.book)["creators"][-1]["备注"]
        d.prepare_day(self.state, "2026-10-22", self.book, self.template)
        self.assertEqual(len(self.state.items(batch)), 1)
        self.assertEqual(saved, self.book.read_bytes())
        self.assertIn(one["body"], Path(result["preview"]).read_text())
        with self.assertRaises(d.mail.DraftError):
            d.record_sent(self.state, batch, {}, self.book)
        instruction = {"user_instruction": "模拟：确认此批发送", "user_instruction_at": NOW.isoformat(),
                       "content_hashes": {one["id"]: one["content_hash"]}}
        d.approve_batch(self.state, batch, instruction)
        self.assertTrue(d.claim_send(self.state, batch)["handoff_to_user_pipeline"])
        with self.assertRaises(d.mail.DraftError):
            d.claim_send(self.state, batch)
        receipt = {"result": "success", "verification": "sent_folder", "message_id": "SIM-SENT-1",
                   "source": "https://example.invalid/sent/SIM-SENT-1", "sent_at": NOW.isoformat(),
                   "recipient": one["recipient"], "content_hash": one["content_hash"]}
        for changes in ({"result": "unknown"}, {"content_hash": "changed"}, {"message_id": ""}):
            with self.assertRaises(d.mail.DraftError):
                d.record_sent(self.state, batch, {one["id"]: {**receipt, **changes}}, self.book)
        self.assertEqual(saved, self.book.read_bytes())
        d.record_sent(self.state, batch, {one["id"]: receipt}, self.book)
        row = d.workbook.read(self.book)["creators"][-1]
        self.assertEqual(row["初次建联日期"], "2026-10-22")
        self.assertEqual(row["联系状态"], "初次建联未回复")
        self.assertEqual(row["备注"], before_note.replace("首联未发送；", "首联已发送（2026-10-22）；", 1))
        saved = self.book.read_bytes()
        d.record_sent(self.state, batch, {one["id"]: receipt}, self.book)
        self.assertEqual(saved, self.book.read_bytes())

    def test_summary_has_no_workbook_or_draft_side_effects(self):
        before = self.book.read_bytes()
        message = {"account": "team@example.invalid", "message_id": "SIM-IN-1", "creator_id": "SIM-000001",
                   "received_at": NOW.isoformat(), "source": "https://example.invalid/mail/SIM-IN-1",
                   "raw_text": "Could you show the device?", "summary": "希望了解产品", "changes": [],
                   "strategy": "C", "reason": "先明确试用范围", "response_direction": "说明产品和试用选择",
                   "decisions": ["是否提供体验样品"]}
        self.assertEqual(len(d.summarize_replies(self.state, [message])["new_or_changed"]), 1)
        self.assertEqual(d.summarize_replies(self.state, [message])["new_or_changed"], [])
        self.assertEqual(len(d.pending_replies(self.state)), 1)
        for field in ("reply_body", "send", "archived"):
            with self.assertRaises(d.mail.DraftError):
                d.summarize_replies(self.state, [{**message, field: "forbidden"}])
        self.assertEqual(before, self.book.read_bytes())
        self.assertEqual(d.drafts_snapshot(self.state), [])
        self.assertNotIn("raw_text", d.pending_replies(self.state)[0])

    def test_stale_workbook_cannot_be_overwritten(self):
        before = self.book.read_bytes()
        with self.assertRaises(d.mail.DraftError):
            d.workbook.patch_cells(self.book, {"C2": {"value": "other"}}, expected_hash="old",
                                   backup_dir=self.folder / "backups", style_row=2)
        self.assertEqual(before, self.book.read_bytes())


if __name__ == "__main__":
    unittest.main()

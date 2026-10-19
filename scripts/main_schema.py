"""Conservative progress and lossless main-sheet column simplification.

Only main-sheet dictionaries are changed. Mail bodies, identifiers and ownership
are never edited here. Legacy dates remain useful display metadata, but are not
evidence of a successful send and must not trigger an automatic reminder.
"""
from __future__ import annotations

from datetime import date, datetime, time, timedelta, timezone
import re


CONTACT_STAGES = (
    "初次建联已发送", "对方第一次回复", "我方第一次回复",
    "对方第二次回复", "我方第二次回复", "长期有效沟通",
)
REMOVED_MASTER_FIELDS = frozenset({
    "沟通摘要状态", "下一步动作", "分类核实状态", "联系页链接", "来源链接", "来源",
    "首次联系日期", "历史跟进日期", "资料核实说明", "联系方式备注",
})
_SHANGHAI = timezone(timedelta(hours=8))
_BLOCKED = {"暂停联系", "已拒绝", "不合作", "拒绝", "勿联系", "已完成"}
_BLOCKED_NOTE = "历史联系限制："
_LEGACY_DATE_FIELDS = ("首次联系日期", "历史跟进日期")
_DATE_FIELDS = ("最近联系", "最近收信时间", "最近发信时间", "首联发送时间")
_REHOME = {
    "沟通摘要状态": "沟通摘要", "下一步动作": "备注", "分类核实状态": "首联内容分类",
    "联系页链接": "Email", "来源链接": "账号链接", "来源": "账号链接",
    "首次联系日期": "最近联系", "历史跟进日期": "最近联系",
    "资料核实说明": "备注", "联系方式备注": "Email",
}


def _text(value):
    return "" if value is None else str(value).strip()


def _raw(value):
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return "" if value is None else str(value)


def _append_note(creator, field, text):
    if not text:
        return False
    notes = creator.setdefault("_notes", {})
    old = _raw(notes.get(field))
    if text in old:
        return False
    notes[field] = old + ("\n" if old else "") + text
    return True


def _parsed(value):
    """Return a comparable Shanghai datetime; malformed/missing dates stay unknown."""
    if isinstance(value, datetime):
        parsed = value
    elif isinstance(value, date):
        parsed = datetime.combine(value, time())
    elif isinstance(value, str) and value.strip():
        try:
            parsed = datetime.fromisoformat(value.strip().replace("Z", "+00:00"))
        except ValueError:
            return None
    else:
        return None
    return parsed.replace(tzinfo=_SHANGHAI) if parsed.tzinfo is None else parsed.astimezone(_SHANGHAI)


def _eligible(record):
    # The archive adapter validates full text separately. This independent guard
    # also prevents a loose draft/initial-send dict from advancing progress.
    state = _text(record.get("处理状态"))
    if state in {"待确认草稿", "已确认待发送", "草稿", "已作废", "发送失败"}:
        return False
    if "草稿" in state and record.get("实际发送结果") != "成功":
        return False
    if record.get("方向") == "收信":
        return True
    if record.get("方向") != "发信" or record.get("实际发送结果") != "成功":
        return False
    purpose = _text(record.get("邮件用途")) or _text(record.get("草稿用途")) or _text(record.get("邮件类型"))
    if purpose in {"首联", "初次建联", "首次联系", "初次联系"}:
        return False
    if purpose not in {"回复", "后续回复", "跟进"} and not _text(record.get("回复邮件记录ID")):
        return False
    body = _raw(record.get("邮件正文") if "邮件正文" in record else record.get("正文或草稿"))
    lines = [line.strip().lstrip("> ") for line in body.splitlines() if line.strip()]
    header = re.compile(r"^(?:From|Sent|Date|To|Subject|Cc|发件人|发送时间|日期|收件人|主题|抄送)\s*[:：]|^On\s+.+wrote\s*[:：]$", re.I)
    return bool(lines) and not all(header.match(line) for line in lines)


def _related(creator, messages):
    identifier = _text(creator.get("达人ID"))
    if not identifier:
        return []
    return [m for m in messages if _text(m.get("达人ID")) == identifier and _eligible(m)]


def _chain_order(records):
    """Order a single complete explicit reply chain, without using row position."""
    by_id = {_text(m.get("邮件记录ID")): m for m in records}
    if "" in by_id or len(by_id) != len(records):
        return None
    children = {}
    roots = []
    for record in records:
        parent = _text(record.get("回复邮件记录ID"))
        if parent in by_id:
            if parent in children:
                return None
            children[parent] = record
        else:
            roots.append(record)
    if len(roots) != 1:
        return None
    result, seen, current = [], set(), roots[0]
    while current is not None:
        identifier = _text(current.get("邮件记录ID"))
        if identifier in seen:
            return None
        seen.add(identifier)
        result.append(current)
        current = children.get(identifier)
    if len(result) != len(records):
        return None
    # An explicit chain conflicting with known timestamps needs human review.
    dated = [_parsed(m.get("邮件时间")) for m in result]
    known = [d for d in dated if d is not None]
    return result if known == sorted(known) else None


def _ordered(records):
    if not records:
        return []
    dates = [_parsed(m.get("邮件时间")) for m in records]
    if all(d is not None for d in dates):
        by_time = sorted(zip(dates, records), key=lambda item: item[0])
        ambiguous = any(left[0] == right[0] and left[1].get("方向") != right[1].get("方向")
                        for left, right in zip(by_time, by_time[1:]))
        if not ambiguous:
            return [record for _, record in by_time]
    return _chain_order(records)


def derive_stage(creator, messages):
    """Derive six requested stages without inventing an order or losing history.

    Consecutive emails from one side constitute one turn. An incomplete archive
    cannot downgrade a valid manually recorded later stage. Where undated mail
    of both directions lacks an explicit chain, only the minimum certain stage
    is inferred; no second round is guessed.
    """
    manual = _text(creator.get("联系状态"))
    rank = CONTACT_STAGES.index(manual) if manual in CONTACT_STAGES else -1
    if creator.get("首联发送结果") == "成功":
        rank = max(rank, 0)
    records = _related(creator, messages)
    if records:
        ordered = _ordered(records)
        if ordered is None:
            # An archived successful subsequent reply proves at least our
            # first reply, even when its position among undated incoming mail
            # is unresolved. It does not establish a second round.
            inferred = 2 if any(m.get("方向") == "发信" for m in records) else 1
        else:
            turns = []
            for record in ordered:
                direction = record.get("方向")
                if not turns or direction != turns[-1]:
                    turns.append(direction)
            # The archive deliberately omits initial outreach. A first stored
            # sent reply is therefore at least the first reply from our side.
            inferred = (1 if turns[0] == "收信" else 2) + len(turns) - 1
        # Long-term effective communication is a human judgement, not a mail
        # count. Automatic progression stops at our second reply; preserve a
        # manually selected long-term stage through the existing rank.
        rank = max(rank, min(inferred, 4))
    return CONTACT_STAGES[rank] if rank >= 0 else ""


def contact_blocked(creator):
    """Retain historical do-not-contact controls after their columns disappear."""
    if _text(creator.get("联系状态")) in _BLOCKED or creator.get("下一步动作") == "暂停联系":
        return True
    note = _raw(creator.get("_notes", {}).get("备注"))
    return _BLOCKED_NOTE in note


def _preserve_controls(creator):
    """Standalone refresh/import must preserve a legacy contact restriction."""
    status = _text(creator.get("联系状态"))
    action = _text(creator.get("下一步动作"))
    if status in _BLOCKED or action == "暂停联系":
        _append_note(creator, "备注", _BLOCKED_NOTE +
                     (status if status in _BLOCKED else action) +
                     "；保留原联系限制，不得自动安排跟进或发送。")
    if creator.get("首联发送结果") == "成功" and not creator.get("首联发送时间"):
        previous = creator.get("首次联系日期")
        if _parsed(previous) is not None:
            creator["首联发送时间"] = _raw(previous)


def _latest_contact(creator, messages):
    candidates = []
    for field in _DATE_FIELDS + _LEGACY_DATE_FIELDS:
        value = creator.get(field)
        parsed = _parsed(value)
        if parsed is None:
            continue
        # Legacy main-sheet times came from hand-entered contact dates. Keep
        # their exact old values in a note, displaying only their known day.
        display = parsed.date().isoformat() if field in _LEGACY_DATE_FIELDS else _raw(value)
        comparison = datetime.combine(parsed.date(), time(), tzinfo=_SHANGHAI) if field in _LEGACY_DATE_FIELDS else parsed
        candidates.append((comparison, display))
    for record in _related(creator, messages):
        parsed = _parsed(record.get("邮件时间"))
        if parsed is not None:
            candidates.append((parsed, _raw(record["邮件时间"])))
    return max(candidates, key=lambda item: item[0])[1] if candidates else creator.get("最近联系")


def refresh_progress(creators, messages):
    """Refresh only current stage/contact date; never recreate deleted columns."""
    changed = 0
    for creator in creators:
        _preserve_controls(creator)
        stage = derive_stage(creator, messages)
        latest = _latest_contact(creator, messages)
        if creator.get("联系状态") != stage:
            creator["联系状态"] = stage
            changed += 1
        if latest is not None and creator.get("最近联系") != latest:
            creator["最近联系"] = latest
            changed += 1
    return {"updated_fields": changed}


def migrate_creator(creator, messages):
    """Rehome removed values/comments, then delete their keys exactly once."""
    old_status = _text(creator.get("联系状态"))
    old_action = _text(creator.get("下一步动作"))
    blocked = contact_blocked(creator)
    _preserve_controls(creator)
    stage = derive_stage(creator, messages)
    latest = _latest_contact(creator, messages)
    note_count = 0
    if blocked:
        note_count += _append_note(creator, "备注", _BLOCKED_NOTE +
                                   (old_status if old_status in _BLOCKED else old_action) +
                                   "；保留原联系限制，不得自动安排跟进或发送。")
    if old_status and old_status not in CONTACT_STAGES:
        note_count += _append_note(creator, "备注", "原联系状态：" + old_status +
                                   "；仅保留历史标记，不据此确认发送或推定往来轮次。")
    notes = creator.setdefault("_notes", {})
    removed_count = 0
    for field in sorted(REMOVED_MASTER_FIELDS):
        value = creator.pop(field, None)
        old_note = notes.pop(field, None)
        if value is not None or old_note is not None:
            removed_count += 1
        target = _REHOME[field]
        # Summary validity is recomputed by the text archive rather than kept
        # as a second stale marker. Its pre-existing comment still survives.
        if _text(value) and field != "沟通摘要状态":
            annotation = "原" + field + "：" + _raw(value)
            if field in _LEGACY_DATE_FIELDS:
                annotation += "（历史手填日期，未核实为实际发信时间；不触发自动跟进）"
            note_count += _append_note(creator, target, annotation)
        if _text(old_note):
            note_count += _append_note(creator, target, "原" + field + "批注：\n" + _raw(old_note))
    creator["联系状态"] = stage
    if latest is not None:
        creator["最近联系"] = latest
    records = _related(creator, messages)
    if records and len({m.get("方向") for m in records}) > 1 and _ordered(records) is None:
        note_count += _append_note(creator, "联系状态", "部分邮件时间或回复关联不明确；未推定未知邮件的先后和后续沟通轮次。")
    return {"removed_fields": removed_count, "rehomed_notes": note_count,
            "contact_stage": stage, "recent_contact": latest}

"""Offline audit interface. There are no send/approve endpoints or network requests."""
from __future__ import annotations

import base64
import hashlib
from html import escape
from pathlib import Path


def render_page(snapshot, document):
    import mail_send as sender
    def text(value): return escape(str(value or ""))
    def image_blocks(blocks):
        output = []
        for block in blocks:
            if "text" in block: output.append("<p class=mailtext>" + text(block["text"]) + "</p>")
            else:
                output.append('<div class="images">')
                for item in block.get("images", []):
                    path = Path(item["path"])
                    mime = "image/jpeg" if path.suffix.lower() in (".jpg", ".jpeg") else "image/" + path.suffix[1:].lower()
                    output.append('<img alt="Tailki 产品或佩戴示意" src="data:' + mime + ';base64,' + base64.b64encode(path.read_bytes()).decode() + '">')
                output.append("</div>")
        return "".join(output)
    def full_mime(raw):
        message = sender.parse(base64.b64decode(raw))
        headers = "".join("<p><b>" + label + "：</b>" + text(message.get(key)) + "</p>" for key, label in
                          (("To", "收件人"), ("Cc", "抄送"), ("Bcc", "密送"), ("Subject", "主题")) if message.get(key))
        body = message.get_body(preferencelist=("plain",))
        result = headers + ("<pre>" + text(body.get_content()) + "</pre>" if body else "")
        html = message.get_body(preferencelist=("html",))
        if html:
            content = html.get_content()
            for part in message.walk():
                cid = str(part.get("Content-ID", "")).strip("<>")
                if cid and part.get_content_maintype() == "image":
                    content = content.replace("cid:" + cid, "data:" + part.get_content_type() + ";base64," + base64.b64encode(part.get_payload(decode=True) or b"").decode())
            csp = "default-src 'none'; img-src data:; style-src 'unsafe-inline'; form-action 'none'; base-uri 'none'"
            source = '<meta http-equiv="Content-Security-Policy" content="' + escape(csp, quote=True) + '">' + content
            result += '<iframe title="邮件原始排版" sandbox="" referrerpolicy="no-referrer" srcdoc="' + escape(source, quote=True) + '"></iframe>'
        attachments = list(message.iter_attachments())
        if attachments:
            result += "<p>附件：</p><ul>"
            for part in attachments:
                payload = part.get_payload(decode=True) or part.as_bytes()
                filename = part.get_filename() or "未命名附件"
                uri = "data:application/octet-stream;base64," + base64.b64encode(payload).decode()
                result += '<li><a download="' + escape(filename, quote=True) + '" href="' + uri + '">' + text(filename) + "</a></li>"
            result += "</ul>"
        return result
    cards = snapshot["cards"]
    draft_count = len(snapshot["mails"]) + sum(bool(c["draft"]) for c in cards)
    mentors = [c for c in cards if c["kind"] == "mentor"]
    nav = '<nav><a href="#decisions">带教待决</a><a href="#mails">邮件完整预览</a><a href="#pool">候选池</a><a href="#creators">今日新增</a><a href="#issues">待核实</a></nav>'
    chunks = []
    chunks.append('<header><div class="eyebrow">TAILKI · ' + text(snapshot["date"]) + '</div><h1>今日审核</h1><p class="muted">集中查看合作进展与完整邮件，在聊天中确认最终发送版本。</p><div class="actions"><a class="primary" href="' + escape(Path(document).as_uri(), quote=True) + '">查看带教 DOCX</a><a class="secondary" href="' + escape(Path(snapshot["book"]).as_uri(), quote=True) + '">打开跟进表</a></div></header>')
    chunks.append('<div class="metrics">' + ''.join('<div><strong>' + str(n) + '</strong><span>' + label + '</span></div>' for n, label in
                  ((len(snapshot["added"]), "今日新增"), (len(mentors), "带教待决"), (draft_count, "邮件待审"), (snapshot["sent_today"], "今日已核验发送"))) + '</div>')
    active = snapshot["schedule"].get("status") == "active"
    chunks.append('<p class="notice">' + ("每日准备已开启 · 北京时间 10:00" if active else "每日定时准备未启用") + ' · 每天最多 ' + str(snapshot.get('daily_limit', 5)) + ' 位。每批邮件均需你确认，查看或复制审核意见不会发送邮件。</p>')
    chunks.append(nav)
    chunks.append('<section id="decisions"><h2>带教待决与来信进展</h2>')
    if not cards: chunks.append('<p class="empty">暂无已整理的待处理来信。</p>')
    for card in cards:
        chunks.append('<article data-search="' + escape(card["name"] + " " + card["email"], quote=True) + '"><div class="cardhead"><h3>' + text(card["title"]) + '</h3><span class="tag">' + ("待带教决策" if card["kind"] == "mentor" else "待审核回复") + '</span></div><p class="muted">' + text(card["email"]) + " · " + text(card["status"]) + '</p><p>' + text(card["summary"]) + '</p>')
        chunks.append('<p class="muted">已记录来信 ' + str(card["incoming"]) + ' 封 · 我方发信 ' + str(card["outgoing"]) + ' 封（仅统计已采集记录）' + (' · 最近来信 ' + text(card["received_at"][:10]) if card["received_at"] else '') + '</p>')
        chunks.append('<p><b>建议方向：</b>' + text(card["strategy"]) + '</p>')
        if card["decisions"]: chunks.append('<div class="decision"><b>需要决定</b><ul>' + ''.join('<li>' + text(x) + '</li>' for x in card["decisions"]) + '</ul></div>')
        if card["stale"]: chunks.append('<p class="warning">资料已有变化，旧汇报及回复稿已暂停展示。</p>')
        if card["details"]:
            chunks.append('<details><summary>查看详细沟通与决策背景</summary>' + ''.join(('<h4>' + text(p["text"]) + '</h4>') if p.get("heading") else ('<p>' + text(p["text"]) + '</p>') for p in card["details"]) + '</details>')
        if card["raw_text"]: chunks.append('<details><summary>查看已采集来信原文</summary><pre>' + text(card["raw_text"]) + '</pre></details>')
        chunks.append('</article>')
    chunks.append('</section><section id="mails"><h2>邮件完整预览</h2><p class="muted">核对每封邮件的正文、图片和附件。仅有缓存文本的稿件会单独标明；发送前仍会核对邮箱的最新版本。</p>')
    if not draft_count: chunks.append('<p class="empty">暂无待审核的完整邮件。已发送邮件不会再次进入待发名单。</p>')
    for index, card in enumerate(c for c in cards if c["draft"]):
        draft = card["draft"]
        headers = ''.join('<p><b>' + label + '：</b>' + text(draft[key]) + '</p>' for key, label in (("cc", "抄送"), ("bcc", "密送")) if draft.get(key))
        chunks.append('<article><div class="cardhead"><h3>' + text(card["name"]) + '</h3><span class="tag">回复稿 · 待确认</span></div><p><b>收件人：</b>' + text(draft["recipient"]) + '</p>' + headers + '<p><b>主题：</b>' + text(draft["subject"]) + '</p>' + image_blocks(draft["blocks"]) + '<p class="muted">该图文稿尚未发送。</p></article>')
    for item in snapshot["mails"]:
        chunks.append('<article><div class="cardhead"><h3>' + text(item["name"]) + '</h3><span class="tag">' + text(item["label"]) + '</span></div>')
        if item.get("raw_b64"): chunks.append(full_mime(item["raw_b64"]))
        else:
            chunks.append('<p><b>收件人：</b>' + text(item["recipient"]) + '</p><p><b>主题：</b>' + text(item["subject"]) + '</p>')
            for key, label in (("cc", "抄送"), ("bcc", "密送")):
                if item.get(key): chunks.append('<p>' + label + '：' + text(', '.join(item[key])) + '</p>')
            if item.get("attachments_incomplete"):
                chunks.append('<p class="warning">缓存记录包含附件，需先读取邮箱完整最终稿，本页尚不构成完整附件预览。</p>')
            chunks.append('<pre>' + text(item["body"]) + '</pre>')
        chunks.append('<p class="muted">批次：' + text(item["review_id"]) + '</p></article>')
    pool = snapshot.get('candidate_pool', {})
    chunks.append('</section><section id="pool"><h2>KOL 候选池</h2><p>池内共 ' + str(pool.get('total', 0)) + ' 位；其中 ' + str(pool.get('available_today', 0)) + ' 位今日资料齐全、待联系历史复核。每天最多入选 ' + str(pool.get('limit', 5)) + ' 位，未入选的继续保留。</p>')
    if pool.get('items'):
        chunks.append('<details><summary>查看候选名单和处理原因</summary><div class="tablewrap"><table><thead><tr><th>达人</th><th>合作邮箱</th><th>状态</th><th>核实日期</th><th>处理说明</th></tr></thead><tbody>')
        for item in pool['items']:
            chunks.append('<tr>' + ''.join('<td>' + text(value) + '</td>' for value in (item['name'], item['email'], item['label'], item['checked_at'][:10], item['reason'])) + '</tr>')
        chunks.append('</tbody></table></div></details>')
    else:
        chunks.append('<p class="empty">候选池尚待补充。后续搜集结果会先入池、复核并去重，合格不足时如实列出缺口。</p>')
    chunks.append('</section><section id="creators"><h2>今日新增达人</h2>')
    if not snapshot["added"]: chunks.append('<p class="empty">今天尚无已登记的新增达人；历史待审事项继续保留。</p>')
    else:
        chunks.append('<div class="tablewrap"><table><thead><tr><th>达人</th><th>合作邮箱</th><th>内容分类</th><th>当前状态</th></tr></thead><tbody>')
        for row in snapshot["added"]:
            chunks.append('<tr>' + ''.join('<td>' + text(row.get(k)) + '</td>' for k in ("达人昵称", "Email", "首联内容分类", "联系状态")) + '</tr>')
        chunks.append('</tbody></table></div>')
    chunks.append('</section><section id="issues"><h2>待核实事项</h2>')
    chunks.append('<ul>' + ''.join('<li>' + text(x) + '</li>' for x in snapshot["warnings"]) + '</ul>' if snapshot["warnings"] else '<p class="empty">本次整理未发现需要额外处理的资料冲突。</p>')
    chunks.append('</section><section><h2>审核意见</h2><p>填写带教决定、修改要求或希望推进的达人，复制后发到当前聊天。</p><textarea id="feedback" aria-label="审核意见" placeholder="例如：Henri 暂缓；Fillian 先确认付款条件；Louie 按以下意见修改……"></textarea><div class="actions"><button id="copy">复制审核意见</button><span id="copied" role="status"></span></div></section>')
    chunks.append('<footer>更新于 ' + text(snapshot["generated_at"].replace("T", " ")) + '。邮件内容依据最近已采集资料，本页刷新不重新收信。双击“今日审核”入口可重新整理。<br>累计已核验发送 ' + str(snapshot["verified_total"]) + ' 封。</footer>')
    script = """document.getElementById('copy').addEventListener('click',async()=>{const value=document.getElementById('feedback').value.trim();const note=document.getElementById('copied');if(!value){note.textContent='请先填写审核意见';return;}try{await navigator.clipboard.writeText(value);note.textContent='已复制，请粘贴到聊天';}catch{const area=document.getElementById('feedback');area.focus();area.select();note.textContent='已选中，请按 Command+C 复制';}});"""
    digest = base64.b64encode(hashlib.sha256(script.encode()).digest()).decode()
    csp = "default-src 'none'; img-src data:; style-src 'unsafe-inline'; script-src 'sha256-" + digest + "'; frame-src 'self' data:; form-action 'none'; base-uri 'none'"
    css = """*{box-sizing:border-box}html{scroll-behavior:smooth}body{margin:0;background:#f5f6f8;color:#202a35;font:15px/1.7 -apple-system,BlinkMacSystemFont,'PingFang SC',sans-serif}main{max-width:1120px;margin:auto;padding:40px 28px}header{padding:12px 0 26px}h1{font-size:36px;line-height:1.3;margin:8px 0}h2{font-size:23px;margin:30px 0 14px}h3{font-size:19px;margin:0}h4{font-size:16px;margin:18px 0 6px}p{overflow-wrap:anywhere}.eyebrow{letter-spacing:1.5px;font-size:12px;color:#647487}.muted,footer{color:#657180;font-size:13px}.actions{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin:16px 0}.primary,button{background:#244f69;color:white;border:0;cursor:pointer}.primary,.secondary,button{padding:10px 18px;border-radius:6px;font:inherit;text-decoration:none}.secondary{background:#e7ecf0;color:#244f69}a{color:#245b7a}.metrics{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}.metrics>div{background:white;padding:20px 24px;border:1px solid #dce2e8;border-radius:8px}.metrics strong{display:block;font-size:30px;line-height:1.3}.metrics span{color:#657180;font-size:13px}.notice{background:#eaf1f4;border-radius:6px;padding:12px 16px;font-size:13px}nav{position:sticky;top:0;z-index:2;display:flex;gap:24px;background:#f5f6f8ed;padding:14px 0;border-bottom:1px solid #dce2e8}nav a{text-decoration:none}section{scroll-margin-top:62px}article{background:white;border:1px solid #dce2e8;border-radius:8px;padding:24px 28px;margin:14px 0}.cardhead{display:flex;align-items:center;justify-content:space-between;gap:14px}.tag{font-size:12px;background:#edf1f4;color:#44576b;padding:4px 10px;border-radius:4px}.decision{background:#faf5e9;padding:14px 18px;border-radius:5px}.decision ul{margin:6px 0}.warning{color:#8b4d13;background:#fff3e4;padding:12px}.empty{color:#657180;padding:14px 0}details{margin:18px 0}summary{color:#245b7a;cursor:pointer}pre,.mailtext{font:inherit;line-height:1.8;white-space:pre-wrap;overflow-wrap:anywhere}.images{display:flex;align-items:center;justify-content:center;gap:12px;margin:22px 0}.images img{max-width:48%;max-height:290px;object-fit:contain}iframe{width:100%;height:650px;border:1px solid #dce2e8}.tablewrap{overflow:auto}table{width:100%;border-collapse:collapse;background:white;text-align:left}th,td{padding:12px;border-bottom:1px solid #dce2e8;overflow-wrap:anywhere}th{background:#edf1f4}textarea{width:100%;min-height:120px;padding:14px;font:inherit;resize:vertical;border:1px solid #bdc8d2;border-radius:6px}footer{padding:30px 0}@media(max-width:650px){main{padding:20px 14px}.metrics{grid-template-columns:repeat(2,1fr)}article{padding:18px}.cardhead{align-items:flex-start;flex-direction:column}nav{gap:14px;font-size:13px}h1{font-size:30px}}"""
    return '<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="Content-Security-Policy" content="' + escape(csp, quote=True) + '"><title>Tailki 今日审核</title><style>' + css + '</style></head><body><main>' + ''.join(chunks) + '</main><script>' + script + '</script></body></html>'

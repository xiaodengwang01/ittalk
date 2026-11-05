"""Mentor memo generated from the same snapshot as the private review page."""
from __future__ import annotations

from pathlib import Path
import re


def heading(text):
    return re.sub(r"\s+", " ", re.sub(r"[^\w\s\u4e00-\u9fff]", " ", str(text))).strip()


def write_document(snapshot, output):
    from docx import Document
    from docx.shared import Cm, Inches, Pt, RGBColor
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.top_margin = section.bottom_margin = Cm(1.65)
    section.left_margin = section.right_margin = Cm(1.8)
    for name in ("Normal", "Title", "Heading 1", "Heading 2", "Caption"):
        style = doc.styles[name]
        style.font.name = "Arial"
        style.font.color.rgb = RGBColor(0, 0, 0)
        fonts = style.element.get_or_add_rPr().rFonts
        for key in ("ascii", "hAnsi", "cs"): fonts.set(qn("w:" + key), "Arial")
        fonts.set(qn("w:eastAsia"), "Hiragino Sans GB")
        for key in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"): fonts.attrib.pop(qn("w:" + key), None)
        for border in style.element.findall(".//" + qn("w:pBdr")): border.getparent().remove(border)
    doc.styles["Normal"].font.size = Pt(10.5)
    doc.styles["Normal"].paragraph_format.line_spacing = 1.12
    doc.styles["Normal"].paragraph_format.space_after = Pt(6)
    doc.styles["Title"].font.size = Pt(21)
    doc.styles["Heading 1"].font.size = Pt(15)
    doc.styles["Heading 2"].font.size = Pt(11.5)
    for name in ("Heading 1", "Heading 2"):
        doc.styles[name].paragraph_format.keep_with_next = True
    title = "Tailki 达人回信审核"
    cards = snapshot["cards"]
    if not cards or any(not c.get("draft") for c in cards):
        raise ValueError("仅为实际完整回复稿生成审核文档。")
    doc.core_properties.title = title
    doc.core_properties.author = "Tailki"
    doc.add_paragraph(title, "Title")
    doc.add_paragraph(snapshot["date"] + "  带教审核")
    doc.add_paragraph("本轮按来信原文、来信要点、回复策略和完整正文逐位整理。带教修改后核对最终版本；发送须另行确认。")
    cards = snapshot["cards"]
    for index, card in enumerate(cards):
        paragraph = doc.add_heading(heading(card["title"]), 1)
        if index:
            paragraph.paragraph_format.page_break_before = True
        doc.add_paragraph("达人ID：" + card["creator_id"])
        doc.add_heading("对方来信原文", 2)
        doc.add_paragraph(card.get("raw_text") or "本轮来信原文尚未提供。")
        doc.add_heading("来信要点", 2)
        doc.add_paragraph(card.get("incoming_points") or card["summary"] or "待整理。")
        doc.add_heading("我方回复要点及应对策略", 2)
        doc.add_paragraph(card.get("reply_points") or card["strategy"] or "待结合来信整理。")
        for decision in card["decisions"]: doc.add_paragraph("待确认：" + str(decision))
        for part in card["details"]:
            doc.add_paragraph(part["text"])
        doc.add_heading("我方完整回复正文", 2)
        if card["draft"]:
            draft = card["draft"]
            doc.add_paragraph("收件人：" + draft["recipient"])
            for key, label in (("cc", "抄送："), ("bcc", "密送：")):
                if draft.get(key): doc.add_paragraph(label + draft[key])
            doc.add_paragraph("主题：" + draft["subject"])
            for block in draft["blocks"]:
                if "text" in block:
                    doc.add_paragraph(block["text"])
                else:
                    if doc.paragraphs and doc.paragraphs[-1].text:
                        doc.paragraphs[-1].paragraph_format.keep_with_next = True
                    paragraph = doc.add_paragraph()
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    paragraph.paragraph_format.keep_together = True
                    paragraph.paragraph_format.space_after = Pt(2)
                    images = block.get("images", [])
                    from PIL import Image
                    for image in images:
                        with Image.open(image["path"]) as picture:
                            w, h = picture.size
                        width = min(14.4 / max(len(images), 1), 8.0 * w / h)
                        shape = paragraph.add_run().add_picture(image["path"], width=Cm(width))
                        shape._inline.docPr.set("descr", image.get("alt", "Tailki 产品或风格参考，功能以核实信息为准"))
            paragraph = doc.add_paragraph("审核状态：待确认，未发送。")
            for run in paragraph.runs: run.font.size = Pt(9)
        else:
            doc.add_paragraph("合作邮箱：" + card["email"])
            doc.add_paragraph("本轮完整回复尚未整理。")
            if card["stale"]: doc.add_paragraph("此前汇报或回复稿的依据已有变化，需先核实最新内容。")
    Path(output).parent.mkdir(parents=True, exist_ok=True)
    doc.save(output)

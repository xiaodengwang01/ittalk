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
    doc.core_properties.title = "KOL 跟进与待决事项"
    doc.core_properties.author = "Tailki"
    doc.add_paragraph("KOL 跟进与待决事项", "Title")
    paragraph = doc.add_paragraph(snapshot["date"] + "  带教审核")
    paragraph.paragraph_format.space_after = Pt(4)
    doc.add_paragraph("带教，我整理了目前待决定的合作条件和已准备的回复稿。请您确认推进方向，邮件随后仍由我核对最终版本再发送。以下沿用已采集的沟通资料；新来信和条件变化会另行更新。")
    if snapshot["added"]:
        doc.add_heading("今日新增达人", 1)
        table = doc.add_table(rows=1, cols=3)
        table.autofit = False
        for cell, text, width in zip(table.rows[0].cells, ("达人", "合作邮箱", "内容分类"), (5, 7.4, 5)):
            cell.text = text; cell.width = Cm(width)
        for row in snapshot["added"]:
            for cell, value in zip(table.add_row().cells, (row.get("达人昵称", ""), row.get("Email", ""), row.get("首联内容分类", ""))):
                cell.text = str(value)
        for index, row in enumerate(table.rows):
            for cell in row.cells:
                properties = cell._tc.get_or_add_tcPr()
                borders = OxmlElement("w:tcBorders")
                for edge in ("top", "left", "bottom", "right"):
                    line = OxmlElement("w:" + edge)
                    for key, value in (("val", "single"), ("sz", "4"), ("color", "D9D9D9")): line.set(qn("w:" + key), value)
                    borders.append(line)
                properties.append(borders)
                if index == 0:
                    shade = OxmlElement("w:shd"); shade.set(qn("w:fill"), "E9EDF2"); properties.append(shade)
                    for run in cell.paragraphs[0].runs: run.bold = True
        doc.add_paragraph("首联邮件的全部收件人、主题和完整正文见今日审核入口。")
    cards = snapshot["cards"]
    for index, card in enumerate(cards):
        if index or snapshot["added"]: doc.add_page_break()
        doc.add_heading(heading(card["title"]), 1)
        if card["draft"]:
            # Detailed context precedes the exact reusable mail body.
            for part in card["details"]:
                doc.add_paragraph(heading(part["text"]) if part.get("heading") else part["text"], "Heading 2" if part.get("heading") else "Normal")
            if not card["details"]: doc.add_paragraph(card["summary"])
            draft = card["draft"]
            doc.add_paragraph("收件人：" + draft["recipient"])
            for key, label in (("cc", "抄送："), ("bcc", "密送：")):
                if draft.get(key): doc.add_paragraph(label + draft[key])
            doc.add_paragraph("主题：" + draft["subject"])
            for block in draft["blocks"]:
                if "text" in block:
                    doc.add_paragraph(block["text"])
                else:
                    paragraph = doc.add_paragraph()
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    paragraph.paragraph_format.space_after = Pt(2)
                    images = block.get("images", [])
                    for image in images:
                        width = 2.65 if len(images) == 2 and image == images[0] else 3.25 if len(images) == 2 else 5.7 / max(len(images), 1)
                        shape = paragraph.add_run().add_picture(image["path"], width=Inches(width))
                        shape._inline.docPr.set("descr", "Tailki 产品或佩戴示意")
            paragraph = doc.add_paragraph("审核状态：待确认，未发送。")
            for run in paragraph.runs: run.font.size = Pt(9)
        else:
            if card["details"]:
                for part in card["details"]:
                    doc.add_paragraph(heading(part["text"]) if part.get("heading") else part["text"], "Heading 2" if part.get("heading") else "Normal")
            else:
                doc.add_paragraph("合作邮箱：" + card["email"])
                doc.add_paragraph(card["summary"])
                if card["received_at"]: doc.add_paragraph("最近已采集来信时间：" + card["received_at"][:10])
                doc.add_heading("建议推进方向", 2)
                doc.add_paragraph(card["strategy"] or "待结合最新来信判断。")
                doc.add_heading("请您决定", 2)
                for decision in card["decisions"]: doc.add_paragraph(str(decision))
                if not card["decisions"]: doc.add_paragraph("请确认下一步推进方向。")
            if card["stale"]: doc.add_paragraph("此前汇报或回复稿的依据已有变化，需先核实最新内容。")
    if not cards and not snapshot["added"]: doc.add_paragraph("目前没有已整理的新增达人或带教待决事项。")
    if snapshot["warnings"]:
        doc.add_heading("待核实事项", 1)
        for warning in snapshot["warnings"]: doc.add_paragraph(warning)
    Path(output).parent.mkdir(parents=True, exist_ok=True)
    doc.save(output)

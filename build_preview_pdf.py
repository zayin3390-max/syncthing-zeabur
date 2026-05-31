#!/usr/bin/env python3
"""Render a readable PDF preview of the manuscript from review_manuscript.md."""
import re
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame,
                                Paragraph, Spacer, PageBreak)

SRC = "review_manuscript.md"
OUT = "review_manuscript_preview.pdf"

def esc(t):
    t = t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t)
    t = re.sub(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", r"<i>\1</i>", t)
    t = re.sub(r"\^(-?[A-Za-z0-9]+)", r"<super>\1</super>", t)
    return t

styles = {
    "title": ParagraphStyle("title", fontName="Times-Bold", fontSize=15,
                            leading=19, alignment=TA_CENTER, spaceAfter=16),
    "center": ParagraphStyle("center", fontName="Times-Roman", fontSize=11,
                             leading=15, alignment=TA_CENTER, spaceAfter=4),
    "centerI": ParagraphStyle("centerI", fontName="Times-Italic", fontSize=9,
                              leading=12, alignment=TA_CENTER, spaceAfter=2),
    "h1": ParagraphStyle("h1", fontName="Times-Bold", fontSize=13, leading=16,
                         spaceBefore=12, spaceAfter=4),
    "h2": ParagraphStyle("h2", fontName="Times-BoldItalic", fontSize=12,
                        leading=15, spaceBefore=8, spaceAfter=2),
    "body": ParagraphStyle("body", fontName="Times-Roman", fontSize=12,
                          leading=24, alignment=TA_JUSTIFY, firstLineIndent=18,
                          spaceAfter=0),
    "bullet": ParagraphStyle("bullet", fontName="Times-Roman", fontSize=12,
                            leading=16, leftIndent=18, spaceAfter=3),
    "kw": ParagraphStyle("kw", fontName="Times-Roman", fontSize=12, leading=16,
                        spaceBefore=6, spaceAfter=8),
    "abbr": ParagraphStyle("abbr", fontName="Times-Roman", fontSize=11,
                          leading=15, leftIndent=70, firstLineIndent=-70),
    "ref": ParagraphStyle("ref", fontName="Times-Roman", fontSize=11, leading=14,
                         leftIndent=22, firstLineIndent=-22, spaceAfter=4),
}

with open(SRC, encoding="utf-8") as f:
    raw = f.read().split("\n")

story = []
seen_first_rule = False
in_title_block = False
in_refs = False
for i, line in enumerate(raw):
    s = line.strip()
    if not s:
        continue
    if s.startswith("# "):
        story.append(Paragraph(esc(s[2:].strip()), styles["title"]))
        in_title_block = True
        continue
    if s == "---":
        if in_title_block and not seen_first_rule:
            seen_first_rule = True
            in_title_block = False
            story.append(PageBreak())
        else:
            story.append(PageBreak())
        continue
    if s.startswith("## "):
        h = s[3:].strip()
        in_refs = h.lower().startswith("references")
        story.append(Paragraph(esc(h), styles["h1"]))
        continue
    if s.startswith("### "):
        story.append(Paragraph(esc(s[4:].strip()), styles["h2"]))
        continue
    if s.startswith("- "):
        story.append(Paragraph("&bull;&nbsp;&nbsp;" + esc(s[2:]), styles["bullet"]))
        continue
    if s.startswith("**Keywords:**"):
        story.append(Paragraph(esc(s), styles["kw"]))
        continue
    if in_title_block:
        st = "centerI" if "Department" in s else "center"
        story.append(Paragraph(esc(s), styles[st]))
        continue
    if in_refs and re.match(r"^\[\d+\]", s):
        story.append(Paragraph(esc(s), styles["ref"]))
        continue
    mab = re.match(r"^\*\*(.+?)\*\* — (.+)$", s)  # abbreviation line
    if mab:
        story.append(Paragraph(
            f"<b>{esc(mab.group(1))}</b>&nbsp;&nbsp;&nbsp;&nbsp;{esc(mab.group(2))}",
            styles["abbr"]))
        continue
    story.append(Paragraph(esc(s), styles["body"]))

def footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Times-Roman", 10)
    canvas.drawCentredString(A4[0] / 2, 12 * mm, str(doc.page))
    canvas.restoreState()

frame = Frame(22 * mm, 18 * mm, A4[0] - 44 * mm, A4[1] - 36 * mm, id="f")
doc = BaseDocTemplate(OUT, pagesize=A4,
                      pageTemplates=[PageTemplate(id="all", frames=[frame],
                                                  onPage=footer)])
doc.build(story)
print("Saved", OUT)

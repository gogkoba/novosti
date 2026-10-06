"""Собирает Word-файл из речи ведущего (или статьи) — обычного текста с абзацами.

Запуск:
    python tools/speech2docx.py output/2026-09-09/02_speech.md
    python tools/speech2docx.py output/2026-09-09/03_column.md -t "Статья" -o output/2026-09-09/Статья_2026-09-09.docx

Оформление под чтение с листа или суфлёра: Times New Roman 14, интервал 1,5, поля 2 см.
Заголовок берётся из ключа -t (по умолчанию «Речь ведущего») плюс дата из имени папки.
Строки, начинающиеся с «#» (заголовок и подзаголовки статьи), выводятся жирным.
Фрагменты **так** внутри абзаца (ключевые слова статьи) выводятся жирным, звёздочки убираются.
Абзац, начинающийся с «▶», — пометка для вёрстки поста (какой ролик вставить в это место):
выводится мелким серым шрифтом, ссылки в нём кликабельные, в число слов не входит.
"""

import argparse
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from md2docx import add_hyperlink

FONT = "Times New Roman"
BODY_PT = 14
TITLE_PT = 16
NOTE_PT = 11
NOTE_GREY = RGBColor(0x59, 0x59, 0x59)


def build(text, title, out_path):
    doc = Document()
    for section in doc.sections:
        section.left_margin = section.right_margin = Cm(2)
        section.top_margin = section.bottom_margin = Cm(2)

    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(BODY_PT)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    normal.paragraph_format.space_after = Pt(10)
    normal.paragraph_format.line_spacing = 1.5

    p = doc.add_paragraph()
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(TITLE_PT)
    p.paragraph_format.space_after = Pt(16)

    for para in [x.strip() for x in re.split(r"\n\s*\n", text) if x.strip()]:
        if para.startswith("▶"):                       # пометка для вёрстки поста: какой ролик вставить сюда
            p = doc.add_paragraph()
            for i, part in enumerate(re.split(r"(https?://\S+)", para)):
                if not part:
                    continue
                if i % 2:
                    add_hyperlink(p, part, part, size=NOTE_PT)
                else:
                    run = p.add_run(part)
                    run.font.size = Pt(NOTE_PT)
                    run.font.color.rgb = NOTE_GREY
            continue
        heading = para.startswith("#")                 # заголовок и подзаголовки статьи — жирным
        para = re.sub(r"^#+\s*", "", para)
        p = doc.add_paragraph()
        # **ключевые слова** статьи — отдельными жирными кусками, остальное обычным
        for i, part in enumerate(re.split(r"\*\*(.+?)\*\*", para)):
            if part:
                run = p.add_run(part)
                run.bold = heading or i % 2 == 1
        if not heading:
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    doc.save(out_path)


def main():
    ap = argparse.ArgumentParser(description="Речь/статья -> Word")
    ap.add_argument("src", help="файл с текстом")
    ap.add_argument("-o", "--out", help="куда сохранить .docx (по умолчанию рядом с исходником)")
    ap.add_argument("-t", "--title", default="Речь ведущего", help="заголовок документа")
    args = ap.parse_args()

    src = Path(args.src)
    out = Path(args.out) if args.out else src.with_suffix(".docx")
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", str(src.resolve()))
    title = f"{args.title}, {m.group(3)}.{m.group(2)}.{m.group(1)}" if m else args.title
    text = src.read_text(encoding="utf-8")
    build(text, title, out)
    words = sum(len(x.split()) for x in re.split(r"\n\s*\n", text) if not x.strip().startswith("▶"))
    print(f"готово: {out} ({words} слов)")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()

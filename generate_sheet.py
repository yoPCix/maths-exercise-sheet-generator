"""Generate a single-page maths exercise sheet (.docx) from a CSV file.

CSV columns: Number,ExerciseName,VariableName,LaTeX
"""

import argparse
import csv
import math
import os
import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from latex2mathml.converter import convert as latex_to_mathml
from lxml import etree

FONT_NAME = "Arial"
TITLE_SIZE = 14
MAX_COLUMNS = 3
MAX_FONT_SIZE = 12
MIN_FONT_SIZE = 8
MIN_GAP_PT = 4
MAX_GAP_PT = 60

PAGE_WIDTH = Cm(21.0)
PAGE_HEIGHT = Cm(29.7)
MARGIN = Cm(1.5)

XSL_CANDIDATES = [
    r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files\Microsoft Office\Office16\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\Office16\MML2OMML.XSL",
    r"C:\Program Files\Microsoft Office\Office15\MML2OMML.XSL",
]


def find_xsl(explicit):
    candidates = [explicit] if explicit else [os.environ.get("MML2OMML_XSL"), *XSL_CANDIDATES]
    for path in candidates:
        if path and Path(path).is_file():
            return path
    sys.exit("MML2OMML.XSL not found. Pass --xsl <path> or set MML2OMML_XSL.")


def read_exercises(csv_path):
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        rows = []
        for row in csv.DictReader(f):
            row = {k.strip(): (v or "").strip() for k, v in row.items() if k}
            if not row.get("LaTeX") and not row.get("Number"):
                continue
            rows.append(row)
    return rows


def clean_latex(latex):
    latex = latex.strip()
    for delim in ("$$", "$"):
        if latex.startswith(delim) and latex.endswith(delim) and len(latex) > 2 * len(delim):
            return latex[len(delim):-len(delim)].strip()
    if latex.startswith(r"\(") and latex.endswith(r"\)"):
        return latex[2:-2].strip()
    return latex


def latex_to_omml(latex, transform):
    mathml = etree.fromstring(latex_to_mathml(clean_latex(latex)).encode("utf-8"))
    return transform(mathml).getroot()


def estimate_exercise_height(exercise, font_pt):
    line = font_pt * 1.2
    lines = 1 if exercise.get("ExerciseName") else 0
    latex = exercise.get("LaTeX", "")
    eq_lines = 2.2 if any(t in latex for t in (r"\frac", r"\dfrac", r"\sqrt", r"\sum", r"\int", "^", "_")) else 1.2
    return line * (lines + eq_lines + 2)


def choose_layout(columns, available_pt):
    """Pick the largest font size (and gap between exercises) that fits on one page."""
    for font_pt in range(MAX_FONT_SIZE, MIN_FONT_SIZE - 1, -1):
        worst = max(sum(estimate_exercise_height(e, font_pt) for e in col) for col in columns)
        rows = max(len(col) for col in columns)
        gap = (available_pt - worst) / rows
        if gap >= MIN_GAP_PT:
            return font_pt, min(gap, MAX_GAP_PT)
    sys.exit("Too many exercises to fit on a single page.")


def split_columns(exercises):
    rows = math.ceil(len(exercises) / MAX_COLUMNS)
    return [exercises[i:i + rows] for i in range(0, len(exercises), rows)]


def set_columns(section, count):
    sect_pr = section._sectPr
    cols = sect_pr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        sect_pr.append(cols)
    cols.set(qn("w:num"), str(count))
    cols.set(qn("w:space"), str(int(Cm(0.8).twips)))


def add_paragraph(doc, text="", space_after=0, column_break=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    if column_break:
        p.add_run().add_break(WD_BREAK.COLUMN)
    if text:
        p.add_run(text)
    return p


def build_document(title, exercises, transform, output):
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = FONT_NAME
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), FONT_NAME)

    section = doc.sections[0]
    section.page_width, section.page_height = PAGE_WIDTH, PAGE_HEIGHT
    section.top_margin = section.bottom_margin = MARGIN
    section.left_margin = section.right_margin = MARGIN

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_after = Pt(12)
    run = title_p.add_run(title)
    run.font.name = FONT_NAME
    run.font.size = Pt(TITLE_SIZE)
    run.bold = True

    body = doc.add_section(WD_SECTION.CONTINUOUS)
    columns = split_columns(exercises)
    set_columns(body, len(columns))

    available_pt = PAGE_HEIGHT.pt - 2 * MARGIN.pt - (TITLE_SIZE * 1.2 + 12) - 6
    font_pt, gap = choose_layout(columns, available_pt)
    normal.font.size = Pt(font_pt)

    for col_index, column in enumerate(columns):
        for ex_index, ex in enumerate(column):
            number, name, var = ex.get("Number", ""), ex.get("ExerciseName", ""), ex.get("VariableName", "")
            omml = latex_to_omml(ex.get("LaTeX", ""), transform)
            # Breaking at the start of the column's first paragraph avoids an empty line at the top.
            column_break = ex_index == 0 and col_index > 0

            if name:
                add_paragraph(doc, f"{number}) {name}", column_break=column_break).paragraph_format.keep_with_next = True
                eq_p = add_paragraph(doc)
            else:
                eq_p = add_paragraph(doc, f"{number}) ", column_break=column_break)
            eq_p._p.append(omml)
            eq_p.paragraph_format.keep_with_next = True

            add_paragraph(doc, f"{var}=").paragraph_format.keep_with_next = True
            add_paragraph(doc, f"{var}=", space_after=gap)

    doc.save(output)


def main():
    parser = argparse.ArgumentParser(description="Create a single-page maths exercise sheet (.docx) from a CSV file.")
    parser.add_argument("title", help="Sheet title")
    parser.add_argument("csv_file", help="CSV with columns: Number,ExerciseName,VariableName,LaTeX")
    parser.add_argument("-o", "--output", help="Output .docx path (default: CSV name with .docx extension)")
    parser.add_argument("--xsl", help="Path to MML2OMML.XSL (from Microsoft Office)")
    args = parser.parse_args()

    exercises = read_exercises(args.csv_file)
    if not exercises:
        sys.exit("No exercises found in CSV.")

    transform = etree.XSLT(etree.parse(find_xsl(args.xsl)))
    output = args.output or str(Path(args.csv_file).with_suffix(".docx"))
    build_document(args.title, exercises, transform, output)
    print(f"Created {output}")


if __name__ == "__main__":
    main()

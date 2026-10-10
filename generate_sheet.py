"""Generate a single-page maths exercise sheet (.docx) from a CSV file.

CSV columns: Number,ExerciseName,VariableName,LaTeX
"""

import argparse
import csv
import math
import os
import sys
from pathlib import Path
from typing import Any, cast

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from latex2mathml.converter import convert as latex_to_mathml
from lxml import etree  # type: ignore[import-not-found]

FONT_NAME = "Arial"
TITLE_SIZE = 14
MAX_COLUMNS = 3
MAX_FONT_SIZE = 12
MIN_FONT_SIZE = 8
MIN_GAP_PT = 4
MAX_GAP_PT = 20

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
    omml = transform(mathml).getroot()
    if omml.tag == qn("m:oMath"):
        math_paragraph = OxmlElement("m:oMathPara")
        math_paragraph.append(omml)
        omml = math_paragraph
    properties = omml.find(qn("m:oMathParaPr"))
    if properties is None:
        properties = OxmlElement("m:oMathParaPr")
        omml.insert(0, properties)
    justification = properties.find(qn("m:jc"))
    if justification is None:
        justification = OxmlElement("m:jc")
        properties.append(justification)
    justification.set(qn("m:val"), "left")
    return omml


def estimate_exercise_height(exercise, font_pt):
    line = font_pt * 1.2
    lines = 1 if exercise.get("ExerciseName") else 0
    latex = exercise.get("LaTeX", "")
    eq_lines = 2.2 if any(t in latex for t in (r"\frac", r"\dfrac", r"\sqrt", r"\sum", r"\int", "^", "_")) else 1.2
    return line * (lines + eq_lines + 2 * 1.2)


def choose_layout(columns, available_pt):
    """Pick the largest font size (and gap between exercises) that fits on one page."""
    for font_pt in range(MAX_FONT_SIZE, MIN_FONT_SIZE - 1, -1):
        worst = max(sum(estimate_exercise_height(e, font_pt) for e in col) for col in columns)
        rows = max(sum(3 + bool(ex.get("ExerciseName")) for ex in col) for col in columns)
        gap = (available_pt - worst) / rows
        if gap >= MIN_GAP_PT:
            return font_pt, min(gap, MAX_GAP_PT)
    sys.exit("Too many exercises to fit on a single page.")


def split_columns(exercises):
    rows = math.ceil(len(exercises) / MAX_COLUMNS)
    return [exercises[i:i + rows] for i in range(0, len(exercises), rows)]


def configure_table(table, widths):
    table.autofit = False
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{edge}")
        border.set(qn("w:val"), "nil")
        borders.append(border)
    table._tbl.tblPr.append(borders)
    margins = OxmlElement("w:tblCellMar")
    for edge in ("top", "left", "bottom", "right"):
        margin = OxmlElement(f"w:{edge}")
        margin.set(qn("w:w"), "0")
        margin.set(qn("w:type"), "dxa")
        margins.append(margin)
    table._tbl.tblPr.append(margins)
    for column, width in zip(table.columns, widths):
        column.width = width
    for row in table.rows:
        row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        for cell, width in zip(row.cells, widths):
            cell.width = width
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP


def build_document(title, exercises, transform, output):
    doc = Document()
    normal = cast(Any, doc.styles["Normal"])
    normal.font.name = FONT_NAME
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), FONT_NAME)

    section = doc.sections[0]
    section.page_width, section.page_height = PAGE_WIDTH, PAGE_HEIGHT
    section.top_margin = MARGIN
    section.bottom_margin = MARGIN - Pt(MAX_FONT_SIZE)
    section.left_margin = section.right_margin = MARGIN

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_after = Pt(12)
    run = title_p.add_run(title)
    run.font.name = FONT_NAME
    run.font.size = Pt(TITLE_SIZE)
    run.bold = True

    columns = split_columns(exercises)

    available_pt = PAGE_HEIGHT.pt - 2 * MARGIN.pt - (TITLE_SIZE * 1.2 + 12) - 6
    font_pt, gap = choose_layout(columns, available_pt)
    normal.font.size = Pt(font_pt)

    column_width = int((PAGE_WIDTH - 2 * MARGIN) / MAX_COLUMNS)
    outer_table = doc.add_table(rows=1, cols=MAX_COLUMNS)
    configure_table(outer_table, [column_width] * MAX_COLUMNS)
    for column, outer_cell in zip(columns, outer_table.rows[0].cells):
        table = outer_cell.add_table(rows=0, cols=2)
        outer_cell._tc.remove(outer_cell.paragraphs[0]._p)
        for ex in column:
            number, name, var = ex.get("Number", ""), ex.get("ExerciseName", ""), ex.get("VariableName", "")
            lines = []
            if name:
                lines.append((name, None))
            lines.extend([(None, ex.get("LaTeX", "")), (None, f"{var}="), (None, f"{var}=")])
            for line_index, (text, latex) in enumerate(lines):
                cells = table.add_row().cells
                if line_index == 0:
                    cells[0].paragraphs[0].add_run(f"{number})")
                paragraph = cells[1].paragraphs[0]
                if text is not None:
                    paragraph.add_run(text)
                else:
                    omml = latex_to_omml(latex, transform)
                    if line_index == bool(name):
                        for math_run in omml.iter(qn("m:r")):
                            properties = math_run.find(qn("m:rPr"))
                            if properties is None:
                                properties = OxmlElement("m:rPr")
                                math_run.insert(0, properties)
                            style = properties.find(qn("m:sty"))
                            if style is None:
                                style = OxmlElement("m:sty")
                                properties.append(style)
                            style.set(qn("m:val"), "b" if style.get(qn("m:val")) == "p" else "bi")
                        run_properties = OxmlElement("w:rPr")
                        run_properties.append(OxmlElement("w:b"))
                        run_properties.append(OxmlElement("w:bCs"))
                        paragraph._p.get_or_add_pPr().append(run_properties)
                        for properties in list(omml.iter()):
                            tag = etree.QName(properties)
                            if tag.namespace != etree.QName(qn("m:oMath")).namespace:
                                continue
                            if not tag.localname.endswith("Pr") or tag.localname in ("rPr", "ctrlPr", "oMathParaPr"):
                                continue
                            control = properties.find(qn("m:ctrlPr"))
                            if control is None:
                                control = OxmlElement("m:ctrlPr")
                                properties.append(control)
                            control_run = control.find(qn("w:rPr"))
                            if control_run is None:
                                control_run = OxmlElement("w:rPr")
                                control.append(control_run)
                            for property_name in ("b", "bCs"):
                                bold = control_run.find(qn(f"w:{property_name}"))
                                if bold is None:
                                    bold = OxmlElement(f"w:{property_name}")
                                    control_run.append(bold)
                                bold.set(qn("w:val"), "1")
                    paragraph._p.append(omml)
                for cell in cells:
                    paragraph = cell.paragraphs[0]
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    paragraph.paragraph_format.space_before = Pt(0)
                    paragraph.paragraph_format.space_after = Pt(gap)
                    paragraph.paragraph_format.keep_with_next = line_index < len(lines) - 1
        number_width = max(Cm(0.94), Pt((max(len(ex.get("Number", "")) for ex in column) + 1) * font_pt * 0.6))
        configure_table(table, [number_width, column_width - number_width])

    for cell in outer_table.rows[0].cells:
        paragraph = cell.paragraphs[-1]
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = Pt(1)
        paragraph.add_run().font.size = Pt(1)
    trailing_paragraph = doc.add_paragraph()
    trailing_paragraph.paragraph_format.space_after = Pt(0)
    trailing_paragraph.paragraph_format.line_spacing = Pt(1)
    trailing_paragraph.add_run().font.size = Pt(1)

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

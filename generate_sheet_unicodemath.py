"""Generate a maths exercise sheet (.docx) from a UnicodeMath CSV file.

CSV columns: Number,ExerciseName,UnicodeMath1,...,UnicodeMathN.
Each non-empty UnicodeMath cell creates a left-aligned equation paragraph.
Number,ExerciseName, and '--blank--' cells are ignored.
Microsoft Word is required to convert UnicodeMath expressions to OMML.
"""

import argparse
import csv
import os
import sys
from copy import deepcopy
from pathlib import Path
from typing import Any, cast

import win32com.client
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from lxml import etree  # type: ignore[import-not-found]

FONT_NAME = "Arial"
TITLE_SIZE = 14
EQUATION_SIZE = 12

PAGE_WIDTH = Cm(21.0)
PAGE_HEIGHT = Cm(29.7)
MARGIN = Cm(1.5)
MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def read_exercises(csv_path):
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        rows = []
        for row in csv.DictReader(f):
            row = {k.strip(): (v or "").strip() for k, v in row.items() if k}
            if not exercise_lines(row) and not row.get("Number"):
                continue
            rows.append(row)
    return rows


def exercise_lines(exercise):
    lines = []
    for key, value in exercise.items():
        suffix = key.removeprefix("UnicodeMath")
        if key.startswith("UnicodeMath") and suffix.isdecimal() and int(suffix) >= 1:
            unicode_math = value.strip()
            if unicode_math:
                lines.append((int(suffix), unicode_math))
    return sorted(lines)


class UnicodeMathConverter:
    def __init__(self):
        self.word: Any = None
        self.document: Any = None
        try:
            self.word = win32com.client.DispatchEx("Word.Application")
            self.word.Visible = False
            self.document = self.word.Documents.Add()
        except Exception as error:
            self.close()
            raise RuntimeError(
                "Microsoft Word is required to convert UnicodeMath expressions."
            ) from error

    def convert(self, unicode_math):
        self.document.Content.Text = ""
        paragraph = self.document.Paragraphs(1).Range
        paragraph.Text = unicode_math.strip()
        self.document.OMaths.Add(paragraph)
        equation = self.document.OMaths.Item(1)
        equation.Range.Select()
        self.word.CommandBars.ExecuteMso("EquationUnicodeFormat")
        self.word.CommandBars.ExecuteMso("EquationConvert")

        package_xml = equation.Range.WordOpenXML
        package = etree.fromstring(package_xml.encode("utf-8"))
        math_paragraph = package.find(f".//{{{MATH_NS}}}oMathPara")
        if math_paragraph is None:
            math = package.find(f".//{{{MATH_NS}}}oMath")
            if math is None:
                raise ValueError(f"Word did not produce math for: {unicode_math}")
            math_paragraph = OxmlElement("m:oMathPara")
            math_paragraph.append(deepcopy(math))
        else:
            math_paragraph = deepcopy(math_paragraph)

        properties = math_paragraph.find(qn("m:oMathParaPr"))
        if properties is None:
            properties = OxmlElement("m:oMathParaPr")
            math_paragraph.insert(0, properties)
        justification = properties.find(qn("m:jc"))
        if justification is None:
            justification = OxmlElement("m:jc")
            properties.append(justification)
        justification.set(qn("m:val"), "left")
        self.document.Content.Text = ""
        return math_paragraph

    def close(self):
        if self.document is not None:
            self.document.Close(False)
            self.document = None
        if self.word is not None:
            self.word.Quit()
            self.word = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.close()


def build_document(title, exercises, converter, output):
    doc = Document()
    normal = cast(Any, doc.styles["Normal"])
    normal.font.name = FONT_NAME
    normal.font.size = Pt(EQUATION_SIZE)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), FONT_NAME)

    section = doc.sections[0]
    section.page_width, section.page_height = PAGE_WIDTH, PAGE_HEIGHT
    section.top_margin = MARGIN
    section.bottom_margin = MARGIN
    section.left_margin = section.right_margin = MARGIN

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_after = Pt(12)
    run = title_p.add_run(title)
    run.font.name = FONT_NAME
    run.font.size = Pt(TITLE_SIZE)
    run.bold = True

    for exercise in exercises:
        for _, unicode_math in exercise_lines(exercise):
            if unicode_math == "--blank--":
                continue
            paragraph = doc.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(6)
            paragraph._p.append(converter.convert(unicode_math))

    doc.save(output)


def main():
    parser = argparse.ArgumentParser(description="Create a single-page maths exercise sheet (.docx) from a UnicodeMath CSV file.")
    parser.add_argument("title", help="Sheet title")
    parser.add_argument("csv_file", help="CSV with columns: Number,ExerciseName,UnicodeMath1,...,UnicodeMathN (use --blank-- for a blank line)")
    parser.add_argument("-o", "--output", help="Output .docx path (default: CSV name with .docx extension)")
    args = parser.parse_args()

    exercises = read_exercises(args.csv_file)
    if not exercises:
        sys.exit("No exercises found in CSV.")

    output = args.output or str(Path(args.csv_file).with_suffix(".docx"))
    try:
        with UnicodeMathConverter() as converter:
            build_document(args.title, exercises, converter, output)
    except RuntimeError as error:
        sys.exit(str(error))
    print(f"Created {output}")


if __name__ == "__main__":
    main()
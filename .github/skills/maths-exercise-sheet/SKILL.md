---
name: maths-exercise-sheet
description: 'Create maths exercise sheets from natural-language descriptions using the LLM to generate CSV exercises and generate_sheet.py to produce a Word document. Use for requests to create a maths exercise sheet, generate maths exercises, build an exercise CSV, or specify exercise counts, difficulty progression, and proportions.'
argument-hint: 'Describe the exercises, count, difficulty progression, and proportions'
user-invocable: true
---

# Maths Exercise Sheet

Use the current assistant's LLM to compose exercises from the user's description. No separate LLM API, credentials, or network service is needed. Create a CSV and, unless the user requests CSV only, use the existing Python generator to create the .docx.

## Inputs And Defaults

- Extract the topic, exercise count, difficulty progression, group counts or proportions, example equations, title, output paths, and any preferred column count from the request.
- Treat an explicit breakdown such as `5/5/5/5` as group counts when it sums to the requested total. Otherwise treat it as relative weights, allocate whole exercises by largest remainder, and break ties in the specified group order.
- Preserve the requested order, with easier exercises first within each group. Do not shuffle groups or distribute them across CSV columns.
- If the user requests a different page layout, honor it with the generator's `-c` / `--columns` option, defaulting to 3 columns when no value is supplied.
- Unless the user explicitly specifies a variable or variable set, randomly choose the unknown independently for each exercise from `a,b,c,d,k,m,n,t,v,z,x,y`. Repeated letters are allowed. A letter appearing only in an example equation does not override this default.
- Use a short topic-based title if none is supplied. Use a descriptive CSV basename in the current workspace and a matching .docx basename unless specified otherwise. Never overwrite user files without permission; choose a new basename if needed.
- Ask a concise question if the exercise count is missing, requirements conflict, or the target unknown is ambiguous. Do not ask for information that can be inferred from the supplied example.
- Resolve paths relative to the current workspace root. Use `./generate_sheet.py` and read `./sample.csv` if present before generating. If no generator is available, still create the CSV and report that document generation is unavailable.
- Store the CSV file in `./exercises` folder by default.

## CSV Contract

The exact header and field order are:

```csv
Number,ExerciseName,VariableName,LaTeX
1,,a,1+a=5
2,,x,\frac{x}{4}=5
3,"Solve, then check",y,y-2=7
```

- `Number`: consecutive integers starting at 1 by default, without a closing parenthesis. Use an alternative numbering scheme only when explicitly requested.
- `ExerciseName`: empty by default. Include names or instructions only when requested. Never invent group-heading rows.
- `VariableName`: the unknown being solved for in that row's equation. Use the randomly chosen letter from the default pool unless the user explicitly specifies otherwise, and use the same letter in `LaTeX`. If the user explicitly asks for a different answer label, honor it. An equation must have a clearly identifiable unknown; do not silently substitute a label for a missing unknown.
- `LaTeX`: mandatory and nonempty for every row. Use valid mathematical LaTeX, without `$`, `$$`, `\(`, or `\)` wrappers. Use single backslashes in the actual CSV, not JSON-escaped backslashes. Do not include solutions or explanatory prose in this field.
- Produce UTF-8, comma-delimited CSV, one exercise per row, exactly four fields. Quote any field containing a comma, double quote, or newline; double embedded double quotes. Use a CSV writer when available. Do not put Markdown fences in the saved file.
- For decimal numbers, use a decimal point by default. Honor an explicitly requested decimal comma using valid LaTeX and CSV quoting.
- Check that equations are mathematically valid, solvable at the requested level, and varied without unintended duplicates. Internally solve each equation to confirm its intended answer, but do not add an answer column or answer key unless requested.

## Document Layout

The CSV is a flat, ordered list, not a spreadsheet. Leave layout to `generate_sheet.py`:

- Centered Arial 14 title.
- Use a configurable number of exercise columns, with 3 as the default. The generator accepts `-c` / `--columns` and preserves the requested count when the user supplies one.
- Exercises are ordered down the first column and then down subsequent columns. With the default 3-column mode, the current generator uses up to `ceil(count / 3)` exercises per column: 8 exercises become 3/3/2; 20 become 7/7/6.
- Each exercise starts with its number. A nonempty exercise name precedes the equation on its own line; an empty name keeps the number and equation on the same line. Two `VariableName=` lines follow.
- Equations must be editable Microsoft Equation objects, not images or literal LaTeX text.
- Keep the document to one page. Never silently omit exercises to make it fit.

## Workflow

1. Resolve the requested counts and sequence. Use the user's supplied example as the model for notation. Apply the random-variable default unless the user explicitly specifies a variable or variable set.
2. Generate the equations with the current LLM, apply the CSV defaults, and write the CSV to the chosen output path.
3. Parse the saved CSV using a structured CSV parser. Verify the exact header, four fields per row, requested total, group counts and order, numeric sequential numbering by default, blank names by default, nonempty LaTeX, and matching answer variables. Check each equation's solution and difficulty.
4. If the user requests CSV only, stop after validation and report its path.
5. Otherwise configure the Python environment for the generator, use its selected interpreter, and install dependencies from its `requirements.txt` only if needed. Invoke the generator with a quoted title and CSV path, optionally adding `-c <n>` to override the default three-column layout, for example:

   ```powershell
   & <selected-python> <generator-path> "Addition and Subtraction" <csv-path> -o <docx-path> -c 2
   ```

6. If `MML2OMML.XSL` cannot be found, use a supplied `--xsl` path or `MML2OMML_XSL`; otherwise report the missing dependency. Do not replace real equations with text or images.
7. Verify generation succeeded and the .docx exists. Check its equation count against the CSV total. When Word is available, open only the generated document read-only, repaginate, and verify the rendered page count is exactly one, then close it and release the automation instance. The generator estimates heights, so file creation alone does not guarantee one-page rendering. If rendering cannot be checked, say that the page count is unverified. If it exceeds one page, report the failure rather than claiming success or removing exercises.
8. Return the CSV and .docx paths with a concise generation and verification result. If a dependency or layout check fails, retain the CSV and explain the blocker.

## Worked Example

User request:

> Create a maths exercise sheet for simple addition. 20 examples. For example '1 + a = 5'. Start with easier (addition), then subtraction, then addition with decimal fractions, then subtraction with decimal numbers. proportion 5/5/5/5.

Interpret this as 20 equations in four consecutive groups of five, with randomly chosen unknowns from the default pool, no exercise names, and decimal fractions written as decimal numbers. The example's `a` illustrates notation, not an explicit variable requirement. Use the title `Addition and Subtraction`. One valid randomized CSV is:

```csv
Number,ExerciseName,VariableName,LaTeX
1,,a,1+a=5
2,,t,2+t=8
3,,c,4+c=11
4,,y,7+y=16
5,,m,9+m=23
6,,k,k-1=4
7,,b,b-3=6
8,,z,z-5=8
9,,n,n-7=12
10,,d,d-9=17
11,,v,0.5+v=2.0
12,,x,1.2+x=3.8
13,,m,2.35+m=5.70
14,,a,4.75+a=10.20
15,,z,6.85+z=14.30
16,,c,c-0.5=1.5
17,,v,v-1.3=2.4
18,,k,k-2.45=3.80
19,,y,y-4.65=7.25
20,,b,b-6.85=9.75
```

Save only the CSV contents, then invoke the generator. Preserve the sequence: rows 1-5 integer addition, 6-10 integer subtraction, 11-15 decimal addition, and 16-20 decimal subtraction. The document's column starts are exercises 1, 8, and 15; topic boundaries do not need to coincide with column boundaries.
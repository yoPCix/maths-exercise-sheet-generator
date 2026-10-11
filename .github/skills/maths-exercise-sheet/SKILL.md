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

## LaTeX tips
- use `·` for multiplication instead of `*`

## Word Equation / LaTeX Conversion Guidelines

When generating or converting mathematical expressions intended for Microsoft Word’s equation editor (LaTeX input mode), follow these rules. Word’s LaTeX-to-OMML converter is limited and not a full TeX parser.

### Exponents and Subscripts
- Always brace multi-character, negative, or complex exponents/subscripts.
  - Correct: `a^{-2}`, `x^{10}`, `e^{-x^2}`, `a_{i,j}`
  - Avoid: `a^-2`, `x^10` (unbraced multi-digit), `a_i,j`
- Single positive digits may work unbraced (`a^2`), but bracing is preferred for consistency.

### Grouping for Nested Expressions
- Prefer explicit braces `{}` when a parenthesized expression (or any complex group) is the base of a superscript/subscript.
  - Preferred (Word-friendly): `({(a^2)^{-3}})^2` or `{(a^{2})^{-3}}^{2}`
  - May fail in Word: `((a^2)^{-3})^2`
- Use braces to force unambiguous grouping of the base before applying `^` or `_`.
- Nested bare parentheses + scripts (especially negative exponents) frequently confuse Word’s speculative delimiter matching.

### General Conversion Tips
- Keep expressions as simple as possible; excessive nesting of bare parentheses increases failure risk.
- After pasting into a Word equation field, convert to Professional format (Convert button or Ctrl+=).
- Negative signs in exponents must be inside braces.
- For troubleshooting, convert the equation to Linear (UnicodeMath) format to inspect how Word parsed the structure.
- Word supports most common LaTeX math commands and bare parentheses, but explicit braces are the most reliable grouping mechanism for scripts.

### Examples
| Intended | Word-safe LaTeX |
|----------|-----------------|
| \( a^{-2} \) | `a^{-2}` |
| \( ((a^2)^{-3})^2 \) | `({(a^2)^{-3}})^2` or `{(a^{2})^{-3}}^{2}` |
| \( (x+y)^{n+1} \) | `(x+y)^{n+1}` or `{(x+y)}^{n+1}` |
| \( e^{-x^2} \) | `e^{-x^2}` |
| \( (-2)^3 \) | `({-2})^3` |
| \( ((a^2)^-3)^2 \) | `({(a^2)^{-3}})^2` |

When outputting LaTeX that will be pasted into Word equations, apply these rules by default.

## CSV Contract

Use `Number,ExerciseName` followed by consecutive numbered LaTeX columns from `LaTeX1` through `LaTeXN`. Include as many columns as the longest exercise needs; shorter exercises leave unused fields empty. The default three-line format is:

```csv
Number,ExerciseName,LaTeX1,LaTeX2,LaTeX3
1,,1+a=5,a=,a=
2,,\frac{x}{4}=5,--blank--,--blank--
3,"Solve, then check",y-2=7,y=,
4,,3(t-2)=12,,
```

- `Number`: consecutive integers starting at 1 by default, without a closing parenthesis. Use an alternative numbering scheme only when explicitly requested.
- `ExerciseName`: empty by default. Include names or instructions only when requested. Never invent group-heading rows.
- `LaTeX1`: mandatory and nonempty for every generated exercise; contains the actual exercise and renders in bold. Use the randomly chosen unknown from the default pool unless the user explicitly specifies otherwise. An equation must have a clearly identifiable unknown; do not silently substitute a label for a missing unknown. Do not include solutions or explanatory prose in this field.
- `LaTeX2` through `LaTeXN`: subsequent lines in numeric column order, rendered without bold. By default, include two answer prompts in `LaTeX2` and `LaTeX3`, each containing `<unknown>=` with the same unknown as the exercise. Honor requested answer labels, line counts, and blank working space. The generator does not infer an unknown or automatically add answer lines.
- A value of `--blank--` creates an empty text line, not an equation. An empty value creates no line at all. Use the placeholder for each requested blank line; do not use an empty value to reserve working space.
- For all equation fields, use valid mathematical LaTeX without `$`, `$$`, `\(`, or `\)` wrappers. Use single backslashes in the actual CSV, not JSON-escaped backslashes.
- Produce UTF-8, comma-delimited CSV, one exercise per row, with the same number of fields as the header (`2 + N`). Quote any field containing a comma, double quote, or newline; double embedded double quotes. Use a CSV writer when available. Do not put Markdown fences in the saved file.
- For decimal numbers, use a decimal point by default. Honor an explicitly requested decimal comma using valid LaTeX and CSV quoting.
- Check that equations are mathematically valid, solvable at the requested level, and varied without unintended duplicates. Internally solve each equation to confirm its intended answer, but do not add an answer column or answer key unless requested.

## Document Layout

The CSV is a flat, ordered list, not a spreadsheet. Leave layout to `generate_sheet.py`:

- Centered Arial 14 title.
- Use a configurable number of exercise columns, with 3 as the default. The generator accepts `-c` / `--columns` and preserves the requested count when the user supplies one.
- Exercises are ordered down the first column and then down subsequent columns. With the default 3-column mode, the current generator uses up to `ceil(count / 3)` exercises per column: 8 exercises become 3/3/2; 20 become 7/7/6.
- Each exercise starts with its number. A nonempty exercise name precedes the equation on its own line; an empty name keeps the number and equation on the same line. `LaTeX1` is bold, including equation structures such as fractions and superscripts. Subsequent populated LaTeX fields supply the remaining lines; `--blank--` supplies an empty text line, and empty values are omitted.
- Equations must be editable Microsoft Equation objects, not images or literal LaTeX text.
- Keep the document to one page. The generator estimates font size and spacing from the actual populated lines, including blank text lines. Never silently omit exercises or requested working lines to make it fit.

## Workflow

1. Resolve the requested counts and sequence. Use the user's supplied example as the model for notation. Apply the random-variable default unless the user explicitly specifies a variable or variable set.
2. Generate the equations with the current LLM, apply the CSV defaults, and write the CSV to the chosen output path.
3. Parse the saved CSV using a structured CSV parser. Verify `Number,ExerciseName` followed by `LaTeX1` through `LaTeXN`, matching field counts per row, requested total, group counts and order, numeric sequential numbering by default, blank names by default, nonempty exercise equations in `LaTeX1`, and matching answer variables where prompts are supplied. Check the requested line counts, distinguishing `--blank--` from omitted empty fields. Check each exercise equation's solution and difficulty.
4. If the user requests CSV only, stop after validation and report its path.
5. Otherwise configure the Python environment for the generator, use its selected interpreter, and install dependencies from its `requirements.txt` only if needed. Invoke the generator with a quoted title and CSV path, optionally adding `-c <n>` to override the default three-column layout, for example:

   ```powershell
   & <selected-python> <generator-path> "Addition and Subtraction" <csv-path> -o <docx-path> -c 2
   ```

6. If `MML2OMML.XSL` cannot be found, use a supplied `--xsl` path or `MML2OMML_XSL`; otherwise report the missing dependency. Do not replace real equations with text or images.
7. Verify generation succeeded and the .docx exists. Check its equation count against the number of nonempty LaTeX values excluding `--blank--`, not just the exercise count. Check that placeholders create blank text lines, empty values create no lines, and only `LaTeX1` equations are bold. When Word is available, open only the generated document read-only, repaginate, and verify the rendered page count is exactly one, then close it and release the automation instance. The generator estimates heights, so file creation alone does not guarantee one-page rendering. If rendering cannot be checked, say that the page count is unverified. If it exceeds one page, report the failure rather than claiming success or removing exercises.
8. Return the CSV and .docx paths with a concise generation and verification result. If a dependency or layout check fails, retain the CSV and explain the blocker.

## Worked Example

User request:

> Create a maths exercise sheet for simple addition. 20 examples. For example '1 + a = 5'. Start with easier (addition), then subtraction, then addition with decimal fractions, then subtraction with decimal numbers. proportion 5/5/5/5.

Interpret this as 20 equations in four consecutive groups of five, with randomly chosen unknowns from the default pool, no exercise names, and decimal fractions written as decimal numbers. The example's `a` illustrates notation, not an explicit variable requirement. Use the title `Addition and Subtraction`. One valid randomized CSV is:

```csv
Number,ExerciseName,LaTeX1,LaTeX2,LaTeX3
1,,1+a=5,a=,a=
2,,2+t=8,t=,t=
3,,4+c=11,c=,c=
4,,7+y=16,y=,y=
5,,9+m=23,m=,m=
6,,k-1=4,k=,k=
7,,b-3=6,b=,b=
8,,z-5=8,z=,z=
9,,n-7=12,n=,n=
10,,d-9=17,d=,d=
11,,0.5+v=2.0,v=,v=
12,,1.2+x=3.8,x=,x=
13,,2.35+m=5.70,m=,m=
14,,4.75+a=10.20,a=,a=
15,,6.85+z=14.30,z=,z=
16,,c-0.5=1.5,c=,c=
17,,v-1.3=2.4,v=,v=
18,,k-2.45=3.80,k=,k=
19,,y-4.65=7.25,y=,y=
20,,b-6.85=9.75,b=,b=
```

Save only the CSV contents, then invoke the generator. Preserve the sequence: rows 1-5 integer addition, 6-10 integer subtraction, 11-15 decimal addition, and 16-20 decimal subtraction. The document's column starts are exercises 1, 8, and 15; topic boundaries do not need to coincide with column boundaries.
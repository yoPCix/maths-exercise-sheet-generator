---
name: maths-exercise-sheet
description: 'Create maths exercise sheets from natural-language descriptions using the LLM to generate CSV exercises and generate_sheet_unicodemath.py to produce a Word document. Use for requests to create a maths exercise sheet, generate maths exercises, build an exercise CSV, or specify exercise counts, difficulty progression, and proportions.'
argument-hint: 'Describe the exercises, count, difficulty progression, and proportions'
user-invocable: true
---

# Maths Exercise Sheet

Use the current assistant's LLM to compose exercises from the user's description. No separate LLM API, credentials, or network service is needed. Create a CSV and, unless the user requests CSV only, use the existing Python generator to create the .docx.

## Inputs And Defaults

- Extract the topic, exercise count, difficulty progression, group counts or proportions, example equations, title, output paths, and any preferred column count from the request.
- Treat an explicit breakdown such as `5/5/5/5` as group counts when it sums to the requested total. Otherwise treat it as relative weights, allocate whole exercises by largest remainder, and break ties in the specified group order.
- Preserve the requested order, with easier exercises first within each group. Do not shuffle groups or distribute them across CSV columns.
- Unless the user explicitly specifies a variable or variable set, randomly choose the unknown independently for each exercise from `a,b,c,d,k,m,n,t,v,z,x,y`. Repeated letters are allowed. A letter appearing only in an example equation does not override this default.
- Do not read existing CSV files unless explicitly requested by the user.
- Use a short topic-based title if none is supplied. Use a descriptive CSV basename in the current workspace and a matching .docx basename unless specified otherwise. Never overwrite user files without permission; choose a new basename if needed.
- Save the original user prompt as a Markdown file next to the CSV using the same basename and a `.md` extension, for example `exercise_name.csv` and `exercise_name.md` in the same folder. Keep the prompt verbatim or in a clean markdown summary while preserving the exact user request.
- Ask a concise question if the exercise count is missing, requirements conflict, or the target unknown is ambiguous. Do not ask for information that can be inferred from the supplied example.
- Resolve paths relative to the current workspace root. Use `./generate_sheet_unicodemath.py` and read `./sample_unicodemath.csv` if present before generating. If no generator is available, still create the CSV and report that document generation is unavailable.
- Store the CSV file in `./exercises` folder by default.

## UnicodeMath tips
- use `·` for multiplication instead of `*`
- use `,` as the decimal separator if needed

## CSV Contract

Use `Number,ExerciseName` followed by consecutive numbered UnicodeMath columns from `UnicodeMath1` through `UnicodeMathN`. Include as many columns as the longest exercise needs; shorter exercises leave unused fields empty. The default three-line format is:

```csv
Number,ExerciseName,UnicodeMath1,UnicodeMath2,UnicodeMath3
1,,1+a=5,a=,a=
2,,x/4=5,--blank--,--blank--
3,"Solve, then check",y-2=7,y=,
4,,3(t-2)=12,,
```

- `Number`: consecutive integers starting at 1 by default, without a closing parenthesis. Use an alternative numbering scheme only when explicitly requested.
- `ExerciseName`: empty by default. Include names or instructions only when requested. Never invent group-heading rows.
- `UnicodeMath1`: mandatory and nonempty for every generated exercise; contains the actual exercise and renders in bold. Use the randomly chosen unknown from the default pool unless the user explicitly specifies otherwise. An equation must have a clearly identifiable unknown; do not silently substitute a label for a missing unknown. Do not include solutions or explanatory prose in this field.
- `UnicodeMath2` through `UnicodeMathN`: subsequent lines in numeric column order, rendered without bold. For example, you can include two answer prompts in `UnicodeMath2` and `UnicodeMath3`, each containing `<unknown>=` with the same unknown as the exercise. Honor requested answer labels, line counts, and blank working space. The generator does not infer an unknown or automatically add answer lines. By default, leave these fields empty unless explicitly requested.
- A value of `--blank--` creates an empty text line, not an equation. An empty value creates no line at all. Use the placeholder for each requested blank line; do not use an empty value to reserve working space.
- For all equation fields, use valid UnicodeMath, not `$`, `$$`, `\(`, or `\)` wrappers, and not raw LaTeX commands such as `\frac` unless specifically required by the task. Keep the actual CSV values plain text with single backslashes only where needed for mathematical structure.
- Produce UTF-8, comma-delimited CSV, one exercise per row, with the same number of fields as the header (`2 + N`). Quote any field containing a comma, double quote, or newline; double embedded double quotes. Use a CSV writer when available. Do not put Markdown fences in the saved file.
- For decimal numbers, use a decimal point by default. Honor an explicitly requested decimal comma using valid UnicodeMath and CSV quoting.
- Check that equations are mathematically valid, solvable at the requested level, and varied without unintended duplicates. Internally solve each equation to confirm its intended answer, but do not add an answer column or answer key unless requested.

## Workflow

1. Resolve the requested counts and sequence. Use the user's supplied example as the model for notation. Apply the random-variable default unless the user explicitly specifies a variable or variable set.
2. Generate the equations with the current LLM, apply the CSV defaults, save the original prompt as a Markdown file with the same basename and `.md` extension, and write the CSV to the chosen output path.
3. Parse the saved CSV using a structured CSV parser. Verify `Number,ExerciseName` followed by `UnicodeMath1` through `UnicodeMathN`, matching field counts per row, requested total, group counts and order, numeric sequential numbering by default, blank names by default, nonempty exercise equations in `UnicodeMath1`, and matching answer variables where prompts are supplied. Check the requested line counts, distinguishing `--blank--` from omitted empty fields. Check each exercise equation's solution and difficulty.
4. If the user requests CSV only, stop after validation and report the CSV and prompt markdown paths.
5. Otherwise configure the Python environment for the generator, use its selected interpreter, and install dependencies from its `requirements.txt` only if needed. Invoke the UnicodeMath generator with a quoted title and CSV path, for example:

   ```powershell
   & <selected-python> <generator-path> "Addition and Subtraction" <csv-path> -o <docx-path>
   ```

   Use the UnicodeMath generator path as `<generator-path>` and prefer `./generate_sheet_unicodemath.py` unless the user specifies a different script.
6. Verify generation succeeded and the .docx exists.
7. Return the CSV and .docx paths with a concise generation and verification result. If a dependency or layout check fails, retain the CSV and explain the blocker.

## Worked Example

User request:

> Create a maths exercise sheet for simple addition. 20 examples. For example '1 + a = 5'. Start with easier (addition), then subtraction, then addition with decimal fractions, then subtraction with decimal numbers. proportion 5/5/5/5.

Interpret this as 20 equations in four consecutive groups of five, with randomly chosen unknowns from the default pool, no exercise names, and decimal fractions written as decimal numbers. The example's `a` illustrates notation, not an explicit variable requirement. Use the title `Addition and Subtraction`. One valid randomized CSV is:

```csv
Number,ExerciseName,UnicodeMath1,UnicodeMath2,UnicodeMath3
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

Save only the CSV contents, then invoke the generator. Preserve the sequence: rows 1-5 integer addition, 6-10 integer subtraction, 11-15 decimal addition, and 16-20 decimal subtraction.
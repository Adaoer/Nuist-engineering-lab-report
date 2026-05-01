---
name: engineering-lab-report
description: Generate a realistic Chinese engineering lab report from experiment instructions, screenshots, result charts, waveforms, notes, and school formatting samples. Use when the task is to write or rewrite a Chinese engineering experiment report with strict no-fabrication constraints, explicit grouping by variable conditions such as resistor values and device sizes, evidence-aware image placement, or direct DOCX output that closely matches a provided Word template and cover-page style.
---

# Engineering Lab Report

Generate Chinese engineering lab reports that feel like they were organized by a real student after actually doing the experiment. Stay strict about evidence, avoid fabrication, and keep the prose natural and restrained.

This skill supports two deliverable modes:

- Markdown body only
- Submission-style `.docx` report that inherits a school template cover and layout

## Use This Skill As A Production Workflow

When the user wants a finished report, do not stop at drafting text. Carry the work through:

1. collect experiment materials
2. extract confirmed facts
3. draft the report body
4. place images near the text they prove
5. if Word output is requested, generate the final `.docx`

If the user provides a school template or asks for a Word deliverable, prefer DOCX mode.

## Expected Inputs

Expect one or more of the following:

1. experiment instruction or process document
2. experiment screenshots, charts, waveforms, or result images
3. user notes
4. an existing draft
5. a school cover/template document
6. a directory containing report-related assets

Use only the information supported by those materials.

## Absolute Content Rules

Keep these boundaries strict:

- `实验结果` contains only directly visible or explicitly stated facts
- `实验结果分析` contains interpretation, explanation, comparison, and error discussion
- unsupported guesses must not appear anywhere

Never invent:

- experiment steps
- parameters
- numeric values not shown in the materials
- instrument models
- personal identity fields on the cover page
- conclusions that are not supported by the evidence

If information is missing, prefer omission over invention.

## Body Structure

Unless the user provides a required alternative, keep the body in this order:

```markdown
## 实验目的
1. ...
2. ...

## 实验过程
...

## 实验结果
...

## 实验结果分析
...
```

When generating DOCX, convert the top-level headings into numbered Chinese section headings:

- `一、实验目的`
- `二、实验过程`
- `三、实验结果`
- `四、实验结果分析`

## Grouping Rules

When the materials show multiple comparison groups, organize both `实验结果` and `实验结果分析` by the explicit conditions shown in the materials.

Typical grouping keys:

- resistor values such as `R0 = 30 kΩ`
- device sizes such as `NM0 与 NM1 宽度均为 440 nm`
- structure variants such as `普通电流镜` and `共源共栅电流镜`

Do not collapse clear parameter labels into vague phrases like `不同参数条件下`.

## Image Placement Rules

Treat images as evidence, not as an appendix by default.

Preferred placement:

- put the schematic directly after the process paragraph that says the circuit was built
- put each result screenshot directly after the subgroup it supports
- reuse the exact condition labels immediately above the image group

Preferred pattern for repeated groups:

- subgroup heading
- one short result paragraph
- matching image or images

Only collect all figures at the end when the user explicitly asks for that layout or when inline placement would make the document unreadable.

## Writing Style

Aim for:

- rigorous
- restrained
- realistic
- student-like rather than polished-paper-like

Avoid obvious AI phrasing and inflated conclusions.

Good style:

- `从图上看`
- `整体上看`
- `和理想情况相比`
- `这里还是有一定偏差`
- `基本符合`

Use those sparingly.

## DOCX Workflow

If the user wants a Word file, follow this workflow.

### 1. Draft the body first

First produce or update the Markdown report body. Keep the body content clean and complete before exporting to DOCX.

### 2. Prefer a real template over imitation

If a template is provided, use it as the DOCX base instead of trying to visually recreate it from scratch.

Look for a sibling template in the report folder such as:

- `实验模板/实验模板.docx`

If no user template is provided, use the bundled default template in:

- `assets/default-template/实验模板.docx`

### 3. Generate the final DOCX with the script

Use the bundled script:

- `scripts/generate_lab_report_docx.py`

Usage:

```powershell
python scripts/generate_lab_report_docx.py <report.md>
python scripts/generate_lab_report_docx.py <report.md> --output <report.docx>
python scripts/generate_lab_report_docx.py <report.md> --template <template.docx>
python scripts/generate_lab_report_docx.py <report.md> --title "实验五：电流镜"
```

What the script does:

- loads the template cover page
- rewrites the top cover title to `（）实验报告`
- rewrites the experiment title line and adds a fillable underline
- sets the date to the generation date
- appends a page break
- converts the Markdown body into a formatted Word report
- converts top-level body headings into `一、二、三、四` style
- keeps images inline with their supporting paragraphs

### 4. Leave unknown cover fields blank

If the template contains fields such as:

- 姓名
- 学号
- 班级
- 学院
- 指导老师

leave them blank unless the user explicitly provides them.

### 5. Match layout expectations

For the default bundled template and the school sample used to build this skill, preserve these expectations:

- A4 page size
- top and bottom margins around `2.54 cm`
- left and right margins around `3.17 cm`
- body line spacing around `1.25`
- body text in `宋体`
- main section headings in `黑体`
- figure captions centered below the figure

## Missing Information Handling

If the user only provides the instruction document:

- write `实验目的` and `实验过程`
- keep `实验结果` and `实验结果分析` conservative

If the user only provides screenshots or result charts:

- write only the parts supported by those results
- do not invent the process

If parameter labels are visible but the mapping between screenshots and conditions is unclear:

- prefer filenames or the user's own grouping notes
- if that still does not disambiguate safely, ask one short question

## Final Self-Check

Before returning the report or the generated file, silently verify:

1. no unsupported information was added
2. `实验结果` and `实验结果分析` are still separated
3. all parameter-group comparisons are traceable to the materials
4. inline images appear close to the text they support
5. DOCX output keeps the cover, date, heading numbering, and template structure intact
6. the overall tone still sounds like a normal student's organized lab report

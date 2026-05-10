# Engineering Lab Report Skill

一个用于生成中文工科实验报告的 Codex Skill。它面向实验指导书、实验截图、结果曲线、波形图、学生笔记和学校 Word 模板，帮助 Codex 生成更接近真实提交稿的实验报告。

本项目的核心原则是：只写材料能够支持的内容，不编造实验数据，不把实验结果和结果分析混在一起，并将图片放在能够支撑对应文字的位置。

## 功能特点

- 根据实验指导书、截图、图表、波形和笔记整理中文实验报告
- 支持 Markdown 报告正文和提交版 `.docx` 两种输出
- 在 Word 输出中继承学校模板封面和基础页面格式
- 自动将正文一级标题转换为 `一、实验目的`、`二、实验过程`、`三、实验结果`、`四、实验结果分析`
- 将实验结果截图、原理图、波形图插入到正文相关段落附近
- 按明确实验条件分组整理结果，例如电阻值、器件尺寸、结构类型等
- 严格区分 `实验结果` 和 `实验结果分析`
- 对缺失信息保持保守，避免补造步骤、参数、仪器型号、数据和个人信息

## 适用场景

适合用于：

- 根据实验指导书和结果截图生成实验报告
- 将零散实验记录整理成可提交的中文报告
- 将已有 Markdown 报告导出为带学校模板封面的 Word 文件
- 按不同参数条件整理电路、电子、通信、自动化等工科实验结果
- 将实验图片嵌入正文对应位置，并自动添加图注

不适合用于：

- 在没有任何实验材料时凭空生成完整报告
- 伪造实验数据、实验结论或个人身份信息
- 替代人工核对关键参数、实验事实和最终提交格式

## 目录结构

```text
engineering-lab-report/
├── SKILL.md
├── agents/
│   └── openai.yaml
├── assets/
│   └── default-template/
│       ├── 实验模板.docx
│       └── 校徽.png
└── scripts/
    └── generate_lab_report_docx.py
```

## 主要文件

| 文件 | 说明 |
| --- | --- |
| `SKILL.md` | Codex 触发该 Skill 后读取的核心工作流和写作约束 |
| `agents/openai.yaml` | 面向 Codex UI 的展示名、简介和默认提示词 |
| `assets/default-template/实验模板.docx` | 未提供用户模板时使用的默认 Word 模板 |
| `scripts/generate_lab_report_docx.py` | 将 Markdown 报告正文转换为 `.docx` 的生成脚本 |

## 安装

将整个目录复制到 Codex skills 目录中：

```powershell
Copy-Item -Recurse engineering-lab-report "$env:USERPROFILE\.codex\skills\engineering-lab-report"
```

安装后重启 Codex，或重新加载 skills，使 `engineering-lab-report` 出现在可用 Skill 列表中。

## 使用方式

在 Codex 中说明要使用该 Skill，并提供实验材料。

英文示例：

```text
Use $engineering-lab-report to generate a Chinese engineering lab report from this experiment instruction, screenshots, result charts, and notes. Keep it evidence-based, place figures near the text they support, and output a DOCX file using the provided template.
```

中文示例：

```text
使用 $engineering-lab-report，根据这个实验指导书、截图和结果曲线写一份中文实验报告。不要编造数据，结果和分析分开写，图片放在对应段落附近。如果可以，请输出 Word 文件。
```

## 输入材料

可以提供一种或多种材料：

- 实验指导书、实验步骤或实验要求文档
- 实验截图、结果曲线、波形图或仿真图
- 学生自己的补充笔记
- 已有实验报告草稿
- 学校封面、格式样例或 Word 模板
- 包含报告图片和材料的文件夹

材料越完整，输出越接近可提交成品。缺失的信息不会被自动补造。

## 输出模式

### Markdown 正文

当用户只需要报告文字、轻量草稿或 `.md` 文件时，Skill 会优先输出 Markdown 正文。默认结构为：

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

### Word 文档

当用户要求 Word 文件、提交版报告，或提供了学校模板时，Skill 会进入 DOCX 模式。DOCX 输出会：

- 使用用户提供的模板；没有模板时使用 `assets/default-template/实验模板.docx`
- 保留封面和基础版式
- 设置 A4 页面、页边距、正文行距、宋体正文和黑体标题
- 将图片插入正文对应位置
- 在图片下方生成居中图注
- 将生成日期写入模板日期位置
- 未提供的姓名、学号、班级、学院、指导老师等字段保持空白

## 手动生成 DOCX

如果已经有整理好的 Markdown 报告正文，可以直接运行脚本生成 Word 文件：

```powershell
python scripts/generate_lab_report_docx.py report.md
```

指定输出路径：

```powershell
python scripts/generate_lab_report_docx.py report.md --output report.docx
```

指定学校模板：

```powershell
python scripts/generate_lab_report_docx.py report.md --template 实验模板.docx
```

指定实验标题：

```powershell
python scripts/generate_lab_report_docx.py report.md --title "实验五：电流镜"
```

## 依赖

DOCX 生成脚本需要 Python 3 和 `python-docx`：

```powershell
pip install python-docx
```

如果只使用 Skill 的写作工作流，不手动运行 DOCX 脚本，则不需要单独安装依赖。

## 写作约束

该 Skill 特别强调证据边界：

- `实验结果` 只写图中可见、材料中明确给出的事实
- `实验结果分析` 才写解释、原因、偏差和理论对照
- 不生成材料中没有的实验步骤、参数、数值、仪器型号或结论
- 不编造封面上的姓名、学号、班级、学院和指导老师
- 当截图和条件对应关系不清楚时，优先根据文件名和用户说明判断；仍不明确时应询问用户

## 图片放置规则

图片默认作为正文证据使用，而不是统一放到文末。

推荐结构：

```markdown
### R0 = 30 kΩ，NM0 与 NM1 宽度均为 440 nm

从图中可以看出，输出曲线整体保持在相对稳定的范围内，随输入条件变化出现了对应变化。

![R0 = 30 kΩ 的输出曲线](images/r0-30k-output.png)
```

同一组参数下的结果图应放在该组结果描述之后。原理图应放在说明电路搭建完成的段落之后。

## 开源前检查

如果要将该 Skill 发布到公开仓库，建议在提交前检查：

- 默认模板和校徽是否允许公开分发
- 示例报告中是否包含真实姓名、学号、学校内部信息或未授权截图
- `assets/default-template/实验模板.docx` 是否需要替换成脱敏模板
- 仓库根目录是否添加了合适的 `LICENSE`
- README 中的安装路径是否符合目标用户使用的 Codex 版本

## License

请在开源前添加许可证文件，例如 `LICENSE`。如果还没有确定许可证，建议先不要在 README 中声明具体授权类型。

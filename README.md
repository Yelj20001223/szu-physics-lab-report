# 深大大学物理实验报告 Skill

> 把「课件 PPT + 原始数据 Excel + 预习试卷选择题截图 + 手写原始数据照片」变成一份符合深圳大学《大学物理实验》课程批注要求的 `.docx` 报告（并自动导出一份**同名 `.pdf`**）。

一个给 [WorkBuddy](https://www.workbuddy.cn) / Claude Code 这类 AI 编码助手用的 **Skill（技能包）**：喂给它课件和原始数据，它按老师批注版的模板规范，把预习报告或完整报告写出来，并自动做排版与合规自检。

> 非官方项目，与深圳大学及任何任课教师无关，纯属学生自己写的效率工具。**模板与课件版权归学校 / 课程组所有**，本项目只做格式自动化，不替代你自己做实验、记数据、想误差来源。

---

## 它能做什么

| 交付物 | 内容 |
|---|---|
| **预习报告** | 封面 + 预习试卷选择题截图页 + `一、实验目的` `二、实验原理` `三、实验仪器` `四、实验内容与步骤` |
| **完整报告** | 上述内容 + `五、数据处理`（表格 + 公式 + 作图）+ `六、结果陈述` + `七、思考题` |

生成过程中自动完成的事情：

- 从课件 PPT 里**按阅读顺序**提取正文与插图（表格转 Markdown、图片带页码与尺寸）；
- 把课件里的**公式截图改写成真正的 Word 公式**（OMML，可双击编辑的那种），并自动编号；
- 把 Excel 原始数据**原样**搬进报告表格（不改数值、位数、单位）；
- 从 Excel 导出**你自己做好的图表**（不替你重画）；
- 把你发来的**预习试卷选择题截图自动排进第 2 页**、**手写原始数据记录表照片自动排到最后一页**（都放大到接近整页；多张截图纵向堆叠，永远只占一页，不会把正文挤到后面）；
- 每次产出 docx 后，**再导出一份同名 `.pdf`**（`光拍法测量光速.docx` → `光拍法测量光速.pdf`，同目录同主名）；
- 按老师 24 条批注做**自动自检**：章节是否齐、图题表题位置、图号表号有没有被正文引用、字号字体、封面日期是否超期、公式是否都带编号……

## 快速开始

```bash
# 1) 提取课件 → outline.md + figures/
python scripts/extract_ppt.py --pptx "光拍法测光速.pptx" --out-dir work/ppt

# 2) 自己写 content.md（报告正文，语法见下）

# 3) 生成 docx
#    预习试卷截图 → 自动排进第 2 页（可重复传多张，纵向堆叠只占一页）
#    手写原始数据照片 → 自动排到最后一页（可重复）
python scripts/report_builder.py \
    --content content.md \
    --out "光拍法测量光速.docx" \
    --kind preview \
    --experiment "光拍法测量光速" --location "某某楼210" \
    --date 2026-10-10 --submit-date 2026-10-13 \
    --quiz "figures/quiz-1.png" \
    --rawdata-image "figures/rawdata.jpg"

# 4) 自检（ERROR 必须为 0）
python scripts/check_report.py "光拍法测量光速.docx" --kind preview

# 5) 导出交付 PDF —— 与 docx 同名同目录。用 PowerShell 执行（细节见 SKILL.md 第六节）：
#    Copy-Item "scripts\export_pdf.ps1" "$env:TEMP\export_pdf.ps1" -Force
#    & "$env:TEMP\export_pdf.ps1" -Docx "光拍法测量光速.docx"
#    → 打印 PDF=... / PAGES=...

# 6) 完整报告：在预习报告的 docx 上增量补 五/六/七 节
python scripts/patch_sections.py \
    --docx "光拍法测量光速.docx" \
    --content content-完整报告.md \
    --sections data,conclusion,questions

# 7) 改过 docx 就重跑第 5 步，让 PDF 跟上
```

## 目录结构

```
.
├── SKILL.md                        总纲：硬规矩、流程、环境限制、交付清单
├── reference/
│   ├── format-rules.md             老师 24 条批注 + 评分表 + 交报告前清单
│   └── content-guide.md            content.md 语法 + 每节该写什么（含完整范例）
├── assets/
│   ├── base.docx                   学校模板（封面与各节已清空，可回炉重做）
│   └── profile.json                长期个人信息（默认留空，用前填自己的）
└── scripts/
    ├── extract_ppt.py              课件 → outline.md + figures/
    ├── extract_excel_charts.py     Excel 里做好的图表 → PNG
    ├── plot_chart.py               自己画 PNG（仅在明确要求"你帮我画"时用）
    ├── report_builder.py           content.md → 报告 .docx
    ├── patch_sections.py           在已有 docx 上只重渲染指定小节
    ├── check_report.py             规范自检（对照老师批注）
    ├── export_pdf.ps1              用 Word COM 把 docx 导成同名同目录的 .pdf
    ├── make_base.py                拿写好的报告回炉空白模板 → base.docx
    ├── scrub_docx_props.py         清 docx 的遗留文档属性（分享前用）
    └── latex2omml.py               公式转换（被 report_builder 调用）
```

## 安装

把整个目录放进宿主的技能目录即可：

```
WorkBuddy / DSH :  ~/.workbuddy/skills/szu-physics-lab-report/
Claude Code 等   :  ~/.claude/skills/szu-physics-lab-report/
```

然后**先填 `assets/profile.json`**（姓名、学号、组号、学院、教师、默认地点）——公开版这里是空的，别把个人信息提交回仓库。

## content.md 语法

| 写法 | 生成效果 |
|---|---|
| `# 一、实验目的` | 开一个小节。编号可省，脚本统一写成 `一、…七、`；七节连续排版，**任何一节都不独占一页** |
| `## 2.1 光速测量的历史。` | 二级小标题：加粗、不缩进、**必须句号结尾** |
| 普通一行 | 正文段落：首行缩进两字、单倍行距、两端对齐 |
| `$$h\nu=E_m-E_n$$ (1)` | 居中 Word 公式 + 编号（编号前自动空四格） |
| `![图1 能级跃迁图。](figures/x.png){w=8}` | 居中插图 + 图下方题注；`{w=8}` 限宽 8 cm |
| `表1 峰值电压表` + 管道表 | 表题写在表**上方**，表格带边框 |

公式是 LaTeX 子集：支持 `\frac` `\sqrt` `^{}` `_{}` `\sum` `\bar` `\overline` `\pm` `\times` `\cdot` `\le` `\approx` 与常用希腊字母；**不要用 `\begin{}` 环境**。详见 `reference/content-guide.md`。

## 硬规矩（踩过的坑，都写成规则了）

这些不是"建议"，是 `check_report.py` 会直接报 ERROR 的硬性要求：

- **交付物是 `<实验名称>.docx` + 同名 `<实验名称>.pdf` 两个文件**（主名逐字一致，不要"（预览）"这类后缀）；docx 改过就要重导 PDF。
- **分页只有三处**：封面第 1 页、预习试卷截图第 2 页、末尾原始数据页（默认不排）。正文七节连续。
- **试卷截图与手写数据照片是脚本自动排进去的**（`--quiz` / `--rawdata-image`，可重复）——不用自己在 Word 里粘。
- **图题在图下方、表题在表上方**，且必须含「图」/「表」字。
- **每张图 / 每个表 / 每条公式的前一段**都要有「…如图N所示 / 如表N所示 / 如式（N）所示」。
- **公式全部住在五节**：求平均值、求斜率、求相对误差……每条都按「**公式 → 代数据 → 结果**」写全。
- **六、结果陈述不出现任何公式**，只写文字 + 最终数值（相对误差的算式留在五节）。
- **图里有文字/数字的要抄进正文**——PPT 里那些部件标注、编号、轴符号是矢量文本框，导出成栅格图会整块丢失。
- 字号：正文五号、题注小五；**正文出现 >14 pt 就是错的**。

## 环境依赖与已知限制

- **Windows + Microsoft Word**。导交付 PDF、核对分页都走 Word COM（`ExportAsFixedFormat` 导 PDF、`ComputeStatistics(2)` 数页、`Range.Information(3)` 取页码），以**只读**方式打开源 docx。
- **PDF 导出用 PowerShell**：Word COM 只能在 PowerShell 里调（`scripts/export_pdf.ps1`）。若技能目录所在盘被执行策略拦（报 `PSSecurityException: 未对文件进行数字签名`），**先把脚本复制到系统临时目录再执行**——`SKILL.md` 第六节有可直接抄的命令。
- 只用已装好的 `python-docx` / `python-pptx` / `Pillow` / `numpy` / `openpyxl` / `PyMuPDF`，**不需要额外 pip 安装**。
- **本机可能没有 LibreOffice**：`.doc → .docx`、课件整页渲染、Excel 图表导出这几步依赖外部转换器，缺了就换条路（见 `SKILL.md` 第六节）。
- 想肉眼核对版式时，用 **PyMuPDF**（`import pymupdf`）把交付 PDF 逐页渲染成 PNG 看：`pymupdf.open(pdf)[i].get_pixmap(dpi=110).save(png)`。PNG 放系统临时目录、看完即删，**实验目录里只留 docx + 同名 pdf**。
- `check_report.py` 的分页判断是**近似**的（靠显式分页符与渲染标记），最终仍建议在 Word 里过一眼。

## 常见问题

**「分享 / 公开这份模板前，怎么确认没夹带别人的信息？」**
Office 文档的属性里常年躺着模板原作者的痕迹：`Company`、累计编辑时长 `TotalTime`、上次打印时间，以及修订与批注的作者标记。清掉：

```bash
python scripts/scrub_docx_props.py assets/base.docx        # 先 --dry-run 看要改什么
```

本仓库里的 `assets/base.docx` 已经清过了（`dc:creator` 署名为 `szu-physics-lab-report`，无 `Company`、无 `custom.xml`）。注意删部件时要同时清掉 `[Content_Types].xml` 与 `_rels/.rels` 里的引用，否则 Word 会报「文件已损坏」——脚本已经处理了这一点。

**「文档里多了个黑方块 / 怪符号？」**
先分清是**内容**还是**编辑标记**。解包 docx 搜方块字符 `■`(U+25A0)、`wingdings` 字体、`<w:numPr>`；如果 XML 干净、导 PDF 渲染也没有，那它在文档里根本不存在，只是 Word/WPS 开着「显示编辑标记」。关掉即可：Word 按 `Ctrl+Shift+8` 或点 `¶`；WPS 在「视图 → 显示/隐藏段落标记」。

**「我发的试卷截图 / 手写数据照片，报告里怎么没看到？」**
这两张图都是**自动排**的，不需要你手动粘：试卷截图传给 `--quiz`（落在第 2 页），手写照片传给 `--rawdata-image`（落在最后一页）。忘了传的话，试卷页会留一整页空白、手写照片则完全不会有。多张试卷截图会纵向堆叠在同一个第 2 页里、整体等比缩小，不会把正文挤到后面。路径相对 `content.md` 所在目录解析。

**「我要交 PDF，PDF 在哪？」**
每次生成 docx 都会顺手导一份**同名同目录**的 PDF：`光拍法测量光速.docx` → `光拍法测量光速.pdf`，主名逐字一致。用 `patch_sections.py` 改过 docx 之后记得重导，否则 PDF 还是旧内容。导出走 Word COM（`scripts/export_pdf.ps1`），命令见 `SKILL.md` 第六节；若报 `PSSecurityException: 未对文件进行数字签名`，把脚本复制到系统临时目录再跑就行。

**「第 2 页粘了试卷截图，正文前凭空多一页？」**
老写法往正文前插分页空段落，上一页被占满时它会被挤到下一页。现在改用 `pageBreakBefore` + 把占位段落标记压到 1 pt，实测 16.5×20.6 / 16.5×24.0 / 16.5×25.5 cm 三种手动粘图都不再多页；用 `--quiz` 自动排时，两张竖图（958×1243、1279×1706）排成 8.66×11.24 + 8.66×11.56 cm，仍只占第 2 页。

**「我在 Word 里改过报告了，还能补五/六/七节吗？」**
能，**别整体重生成**（会冲掉手改）。用 `patch_sections.py` 增量补节，它只重渲染指定单元格，打完会打印**指纹对比**（封面段落、正文各行、内嵌图片 md5），确认除目标小节外全是 SAME。

## 变更记录

| 版本 | 内容 |
|---|---|
| v2 | 匿名化：清空 `profile.json` 与文档属性中的个人信息 |
| v3 | 修复「第 2 页粘图后凭空多一页」（改用 `pageBreakBefore`） |
| v4 | 修复 `check_report` 的 `equation.consecutive` 误报（未判断公式是否真相邻） |
| v5 | 只交付 `.docx`，不再生成 PDF 预览 |
| v6 | 新增「黑方块 = 编辑标记」的判别法 |
| v7 | 新增 `patch_sections.py`：在已有 docx 上增量补节 |
| v8 | 公式全归五节、六节不出公式；图内文字要写进正文（并新增自检规则 `section.equation-in-conclusion`） |
| v9 | 预习试卷截图与手写原始数据照片**自动排版**：`--quiz` 支持多张纵向堆叠、只占第 2 页；`--rawdata-image` 自动排到最后一页；`--quiz` 相对路径改为按 `content.md` 所在目录解析 |
| v10 | **恢复导出 PDF**：每次产出 docx 后自动导一份同名同目录的 `.pdf`（`scripts/export_pdf.ps1`，v5「不导 PDF」作废）；版式核对改用已装的 PyMuPDF |

## 使用与版权

- 本项目是**格式自动化工具**，帮你省掉排版和自查的时间，**不代替**动手做实验、记录原始数据、分析误差的功夫。
- `assets/base.docx` 来自课程提供的报告模板；课件、模板、批注的版权属深圳大学及课程组。**请勿用于商业用途**，仅供个人学习参考。
- 请遵守你所在课程关于 AI 工具使用的规定；用它写作业之前，先确认老师允不允许。

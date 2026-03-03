# pubtab

<div align="center">

  <p>
    <a href="https://pypi.org/project/pubtab/"><img src="https://img.shields.io/pypi/v/pubtab?style=flat-square&color=blue" alt="PyPI Version"/></a>
    <a href="https://pypi.org/project/pubtab/"><img src="https://img.shields.io/pypi/pyversions/pubtab?style=flat-square" alt="Python Versions"/></a>
    <img src="https://img.shields.io/badge/License-MIT-green?style=flat-square" alt="License"/>
    <a href="https://pypi.org/project/pubtab/"><img src="https://img.shields.io/pypi/dm/pubtab?style=flat-square&color=orange" alt="Downloads"/></a>
  </p>

  <strong>Language</strong>: <a href="README.md">English</a> | <a href="README.zh-CN.md">中文</a>

</div>

> Excel 到出版级 LaTeX 表格 — 双向转换，完整保留样式。

## 亮点

- **双向转换** — Excel ↔ LaTeX 互转（`.xlsx`/`.xls` ↔ `.tex`）
- **样式保真** — 颜色、粗体、斜体、合并单元格、旋转、对角线表头完整保留
- **零配置预览** — 首次使用自动安装 TinyTeX，一条命令生成 PNG 预览
- **学术优化** — 前导零剥离、`\diagbox` 对角线表头、section 行自动检测、`\cmidrule` 自动生成

## 快速开始

```bash
pip install pubtab
```

**命令行：**

```bash
# Excel → LaTeX
pubtab xlsx2tex table.xlsx -o output.tex

# 带选项
pubtab xlsx2tex table.xlsx -o output.tex --theme three_line --caption "Results" --label "tab:results" --preview

# LaTeX → Excel（反向转换）
pubtab tex2xlsx paper_table.tex -o recovered.xlsx

# 从 .tex 生成 PNG 预览
pubtab preview output.tex -o preview.png --dpi 300
```

**Python API：**

```python
import pubtab

# Excel → LaTeX
pubtab.xlsx2tex("table.xlsx", output="table.tex", theme="three_line",
                caption="Experimental Results", label="tab:results")

# LaTeX → PNG 预览
pubtab.preview("table.tex", dpi=300)

# LaTeX → Excel
pubtab.tex_to_excel("table.tex", "output.xlsx")
```

## 功能

### Excel → LaTeX 转换

读取 `.xlsx`（openpyxl）和 `.xls`（xlrd）文件，通过 Jinja2 模板生成出版级 LaTeX。

**支持的单元格特性：**
| 特性 | LaTeX 输出 |
|------|-----------|
| 粗体 / 斜体 / 下划线 | `\textbf{}`、`\textit{}`、`\underline{}` |
| 字体颜色 | `\textcolor[RGB]{r,g,b}{}` |
| 背景色 | `\cellcolor[RGB]{r,g,b}` |
| 水平合并 | `\multicolumn{n}{c}{}` |
| 垂直合并 | `\multirow{n}{*}{}` |
| 文字旋转 | `\rotatebox[origin=c]{angle}{}` |
| 对角线表头 | `\diagbox{Row}{Col}` |
| 多行内容 | `\makecell{...\\\\...}` |
| 富文本（逐段样式） | 逐段颜色/粗体/斜体 |

**表格级特性：**
- 从合并表头自动生成 `\cmidrule`
- Section 行检测（首列跨全宽 → 自动添加 `\midrule`）
- 通过 `group_separators` 参数设置行组分隔
- 可配置间距（`tabcolsep`、`arraystretch`、线宽）
- `resizebox` 和 `font_size` 覆盖
- `table*` 双栏跨页（`span_columns=True`）

### LaTeX → Excel 转换

将 LaTeX 表格解析回 Excel，支持丰富的命令：

- `\multicolumn`、`\multirow`（含负值）
- `\textbf`、`\textit`、`\underline`、`\emph`
- `\textcolor`、`\cellcolor`、`\rowcolor`、`\rowcolors`
- `\diagbox`、`\makecell`、`\rotatebox`
- `\newcommand`/`\renewcommand` 展开（最多 10 轮）
- `\definecolor` 自定义颜色解析
- 80+ LaTeX 符号 → Unicode 映射（±、×、→、✓、α-ω 等）
- 嵌套 tabular → `\makecell` 转换

### PNG 预览

直接从 `.tex` 文件生成出版级 PNG 预览：

```bash
pubtab preview table.tex --dpi 300
```

**TinyTeX 自动安装：** 若系统未找到 `pdflatex`，pubtab 会自动下载安装 TinyTeX（约 90 MB）到 `~/.pubtab/TinyTeX/`，并安装所需 LaTeX 包（booktabs、multirow、xcolor 等）。仅需一次。

**PDF → PNG 管线：** `pdf2image` → `qlmanage`（macOS）→ `convert`（ImageMagick），使用第一个可用工具。

安装可选依赖以获得最佳质量：

```bash
pip install pubtab[preview]  # 安装 pdf2image
```

## 命令行参考

| 命令 | 说明 |
|------|------|
| `pubtab xlsx2tex` | Excel 转 LaTeX |
| `pubtab tex2xlsx` | LaTeX 转 Excel |
| `pubtab preview` | 从 .tex 生成 PNG |
| `pubtab themes` | 列出可用主题 |

<details>
<summary>完整 <code>xlsx2tex</code> 选项</summary>

```
pubtab xlsx2tex INPUT -o OUTPUT [OPTIONS]

Options:
  -o, --output TEXT          输出 .tex 文件（必填）
  -c, --config TEXT          YAML 配置文件
  --sheet TEXT               工作表名称或索引（从 0 开始）
  --theme TEXT               主题名称 [默认: three_line]
  --caption TEXT             表格标题
  --label TEXT               LaTeX 标签
  --header-rows INTEGER      表头行数
  --position TEXT            浮动位置 [默认: htbp]
  --font-size TEXT           字号（如 footnotesize）
  --resizebox TEXT           缩放宽度（如 0.8\textwidth）
  --col-spec TEXT            列格式（如 lccc）
  --span-columns            使用 table* 双栏跨页
  --preview                 生成 PNG 预览
  --dpi INTEGER             预览 DPI [默认: 300]
  --header-sep TEXT          自定义表头分隔符
```

</details>

## 配置文件

所有参数均可通过 YAML 配置文件设置，命令行参数优先级更高：

```yaml
theme: three_line
caption: "Experimental Results"
label: "tab:results"
header_rows: 2
span_columns: false
position: htbp
font_size: footnotesize
spacing:
  tabcolsep: "4pt"
  arraystretch: "1.2"
group_separators: [3, 6]
```

```bash
pubtab xlsx2tex table.xlsx -o output.tex -c config.yaml
```

## 主题系统

pubtab 使用基于 Jinja2 的主题系统。内置 `three_line` 主题生成经典 booktabs 三线表。

**自定义主题：** 在 `themes/` 下创建目录，包含 `config.yaml` + `template.tex`：

```
my_theme/
├── config.yaml    # packages, spacing, font_size, caption_position
└── template.tex   # Jinja2 模板
```

列出可用主题：

```bash
pubtab themes
```

## 项目结构

<details>
<summary>查看项目结构</summary>

```
pubtab/
├── pyproject.toml
├── README.md
├── README.zh-CN.md
├── LICENSE
└── src/pubtab/
    ├── __init__.py        # 公共 API：xlsx2tex, preview, tex_to_excel
    ├── cli.py             # 命令行（click）
    ├── models.py          # 数据模型（Cell, TableData, SpacingConfig, ThemeConfig）
    ├── reader.py          # Excel 读取器（.xlsx/.xls）
    ├── renderer.py        # LaTeX 渲染引擎（Jinja2）
    ├── tex_reader.py      # LaTeX 解析器（tex → TableData）
    ├── writer.py          # Excel 写入器
    ├── _preview.py        # PNG 预览（TinyTeX 自动安装）
    ├── config.py          # YAML 配置加载器
    ├── utils.py           # LaTeX 转义、颜色转换
    └── themes/
        └── three_line/
            ├── config.yaml
            └── template.tex
```

</details>

## 贡献

欢迎在 [GitHub](https://github.com/Galaxy-Dawn/pubtab) 提交 Issue 和 Pull Request。

## 许可证

[MIT](LICENSE)

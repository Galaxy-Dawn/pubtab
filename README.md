# pubtab

Excel to publication-ready LaTeX tables.

## Install

```bash
pip install pubtab
```

## Usage

```bash
pubtab convert table.xlsx -o output.tex --theme three_line --caption "Results" --label "tab:results"
```

```python
import pubtab
pubtab.convert("table.xlsx", output="table.tex", theme="three_line")
```

# ExcelColouredCells

Replace nonempty cells with visible direct or supported conditional background
fills with numeric zero, including colored formulas. Skip `Check` and `Sheet3`.
Leave blank cells and uncolored values/formulas alone.

## Run on Windows, macOS, or Linux

Requires Python 3.10 or newer. Download the script into the same folder as your
workbook, then open a terminal in that folder:

```sh
python -m pip install openpyxl==3.1.5
python zero_colored_cells.py "v005.xlsx"
```

This creates `v005_zeroed.xlsx`. For another destination:

```sh
python zero_colored_cells.py "v005.xlsx" --output "v005_edited.xlsx"
```

Existing files are never overwritten. Close the workbook in Excel before running.
The JSON report lists processed and ignored sheets, per-sheet counts, and the total.
Counts include selected cells that already contained zero.

## Conditional formatting

Supported rules:

- Formula/expression rules with relative, absolute, mixed, and cross-sheet references.
- Cell comparisons: greater/less than, greater/less than or equal, equal, not equal,
  between, and not between.
- Color scales applied to numeric cells.
- Rule priority and `stopIfTrue`, including rules that change fonts without fills.
- Background-only differential fills, which Excel commonly uses for conditional colors.

Supported expression syntax includes arithmetic, comparisons, parentheses, cell and
bounded range references, and `LN`, `LOG10`, `SQRT`, `ABS`, `EXP`, `SUM`, `MIN`,
`MAX`, `AND`, `OR`, `NOT`, `IF`, `IFERROR`, `ISNUMBER`, and `ISBLANK`.
This covers the conditional rules in the supplied laser-parameter workbook.
Defined names, external workbooks, table styles, structured references, and other
conditional rule types/functions are not supported; relevant unsupported rules cause
an explicit error rather than silently producing incomplete results. Excel's full
formula engine is not implemented. Formula errors make a rule inactive.

**Colors are selected using the original workbook's saved formula results.**
Open the original in Excel, recalculate (Ctrl+Alt+F9 on Windows), save, and close
before running, especially after changing inputs. openpyxl cannot recalculate formulas.
The tool reports missing saved formula results instead of guessing.

**Conditional formatting remains dynamic.** Its rules are preserved exactly, so a
cell's displayed color may change when Excel recalculates after the value becomes
zero. The tool does not freeze conditional colors or convert them to direct fills.

## Preservation and verification

The tool reads with openpyxl but edits only selected cell payloads inside the `.xlsx`
ZIP package. It does not resave the whole workbook through openpyxl. Excluded sheets
and all non-worksheet package parts remain byte-for-byte unchanged. Direct styles,
conditional formatting rules, sheet names/order, row/column dimensions, merged cells,
comments, and hyperlinks are preserved.

Shared formulas are expanded to equivalent individual formulas only where needed to
preserve uncolored formulas after replacing their shared master. Targets overlapping
array formulas are rejected, because replacing part of an array can affect other cells.

The output is reopened to verify numeric zeros and unchanged untargeted values/formulas.
The worksheet XML is also compared to verify all other content is unchanged, accounting
for equivalent shared-formula expansion. Verification failure deletes only the new output.
The source is never modified. Unsupported-feature warnings stop processing.

Run regression tests with:

```sh
python -m unittest discover -s tests -v
```

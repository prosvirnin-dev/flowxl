# FlowXL

**Language:** English · [Русский](usage.ru.md)

Predictable Excel writing over openpyxl.

Loud errors instead of silent data corruption. Fluent API for creating and editing `.xlsx` files.

## Install

```bash
pip install -e .
# Optional extras.
pip install -e '.[polars]'
pip install -e '.[pandas]'
```

Python 3.12+. The only required dependency is openpyxl.

## What you import

Everything public lives on `flowxl`. If it is not in this list, it may change in a patch.

```python
from flowxl import (
    Workbook,  # Create, open or stream a workbook.
    Sheet,  # Random-access sheet from Workbook.sheet.
    StreamingWorkbook,  # Write-only workbook from Workbook.stream.
    StreamingSheet,
    WorkbookSettings,  # Theme, sheet-name policy, autofit caps.
    CellRef,  # One cell, 1-based. Use parse('B12').
    CellRange,  # A rectangle. Use parse('A1:C10').
    ORIGIN,  # Cell A1.
    KEEP,  # Do not change the value.
    StyleSpec,  # A named cell style drawing.
    FontSpec,
    FillSpec,
    Color,
    DefaultTheme,
    Theme,
    to_excel_formula,  # Russian formula to en-US text stored in xlsx.
    FlowxlError,  # Catch anything this library raises.
    UsageError,  # The call is wrong.
)
```

## Create, open, save

```python
from flowxl import Workbook

with Workbook.new() as wb:  # Empty document, no sheets yet.
    wb.sheet('Sales')
    wb.save('sales.xlsx')  # You must call save. The with block only closes.

with Workbook.open('sales.xlsx') as wb:
    print(wb.sheet_names)
    wb.save('sales.xlsx')  # Save may be called more than once.

raw = Workbook.new().to_bytes()  # Same document in memory.
wb = Workbook.from_bytes(raw)
```

`with` **closes** the workbook. It does **not** save.

| Call | What it does |
|---|---|
| `Workbook.new()` | Empty workbook |
| `Workbook.open(path, formulas=True)` | Open a file. `formulas=False` reads cached values |
| `Workbook.from_bytes(data)` | Open from memory |
| `Workbook.stream()` | Write-only workbook for large tables |
| `wb.sheet(name)` | Get or create a sheet |
| `wb.has_sheet(name)` | Whether the sheet exists |
| `wb.remove(name)` | Delete a sheet |
| `wb.sheet_names` | Names in document order |
| `wb.save(path)` | Write to disk; workbook stays usable |
| `wb.to_bytes()` | Write to memory |
| `wb.close()` | Release resources |

## Cells

Coordinates are 1-based. `CellRef.parse('B12')` is row 12, column 2.

```python
from datetime import date
from flowxl import CellRef, FontSpec, StyleSpec, Workbook

bold = StyleSpec(name='bold', font=FontSpec(bold=True))

with Workbook.new() as wb:
    sheet = wb.sheet('Report')
    sheet.cell(CellRef.parse('A1'), 'Hello')
    sheet.cell(CellRef.parse('B1'), 42)
    sheet.cell(CellRef.parse('C1'), date(2024, 1, 2))
    sheet.cell(CellRef.parse('A1'), style=bold)  # Omit the value: style only.
    sheet.cell(CellRef.parse('B1'), None)  # None clears the cell.
    print(sheet.read(CellRef.parse('A1')))
    wb.save('report.xlsx')
```

| Call | What it does |
|---|---|
| `sheet.cell(ref, value=KEEP, style=None)` | Write a value and/or a style |
| `sheet.row(at, values, style=None)` | Write left to right |
| `sheet.column(at, values, style=None)` | Write top to bottom |
| `sheet.read(ref)` | Current stored value |
| `sheet.merge(area)` | Merge a rectangle |

`KEEP` means “do not change the value”. `None` means “clear the cell”. They are not the same.

## Formulas

The library stores the formula as text. Excel evaluates it when the file is opened.

xlsx always stores **English** names and **en-US** separators. Russian Excel only *displays* `СУММ` and `;`. FlowXL translates on write:

```python
sheet.formula(CellRef.parse('D1'), '=СУММ(B1;C1)')  # Stored as =SUM(B1,C1).
sheet.formula(CellRef.parse('E1'), '=A1+1,5')  # Stored as =A1+1.5.
sheet.formula(CellRef.parse('F1'), '=SUM(B1,C1)')  # Already English: unchanged.
```

`cell(ref, '=СРЗНАЧ(A1:A10)')` is translated the same way: any string that starts with `=` goes through `to_excel_formula`.

| Call | What it does |
|---|---|
| `sheet.formula(ref, expression, style=None)` | Write a formula. Must start with `=` |
| `to_excel_formula(expression)` | Translate without a workbook |

Unknown Cyrillic function names raise `UsageError`. Quoted text and sheet names are left alone.

## Rows, columns, ranges

`CellRange` is the rectangle for merge, autofilter, and dropdowns.

```python
from flowxl import CellRange, CellRef, Workbook

with Workbook.new() as wb:
    sheet = wb.sheet('Report')
    sheet.cell(CellRef.parse('A1'), 'Q1')
    sheet.merge(CellRange.parse('A1:C1'))
    sheet.row(CellRef.parse('A3'), ['Name', 'Qty', 'Total'])
    sheet.column(CellRef.parse('A4'), ['Ann', 'Bob'])
    sheet.widths({'A': 16, 'B': 8, 'C': 10})
    sheet.heights({1: 24})
    sheet.freeze(CellRef.parse('A4'))  # First unfrozen cell.
    sheet.pin_header()  # Freeze the first row.
    sheet.autofilter()  # Covers everything written through FlowXL.
    sheet.autofit()
    wb.save('report.xlsx')
```

| Call | What it does |
|---|---|
| `sheet.freeze(ref)` | Freeze everything above and left of `ref` |
| `sheet.pin_header(rows=1)` | Freeze the top rows |
| `sheet.autofilter(area=None)` | Filter. Omit `area` to use tracked bounds |
| `sheet.widths({1: 12})` / `sheet.widths({'A': 12})` | Column widths |
| `sheet.heights({1: 18})` | Row heights |
| `sheet.autofit(max_width=None, only=None)` | Size columns to written content |

After `sheet.raw_openpyxl()`, bounds are no longer trusted. Pass `area=` to `autofilter`.

## Tables

`frame()` writes a table and formats columns by logical type. A list of dicts, or a polars / pandas DataFrame.

```python
from datetime import date
from flowxl import CellRef, Workbook

records = [
    {'Name': 'Ann', 'Amount': 10, 'Day': date(2024, 1, 2)},
    {'Name': 'Bob', 'Amount': 20, 'Day': date(2024, 1, 3)},
]

with Workbook.new() as wb:
    wb.sheet('Sales').frame(records).pin_header().autofilter().autofit()
    wb.sheet('Offset').frame(records, at=CellRef.parse('B2'))
    wb.save('sales.xlsx')
```

| Call | What it does |
|---|---|
| `sheet.frame(data, at=ORIGIN, header=True, styles=None)` | Write a table. `styles` is `{column_name: StyleSpec}` |

## Styles

Dates, numbers, percents and currency use Excel's built-in styles (`Normal`, `Comma [0]`, `Comma`, `Percent`, `Currency`). The header style `sw_header` is the only custom named style FlowXL adds by default.

A `StyleSpec` needs a **name** when you define your own.

```python
from flowxl import (
    CellRef,
    Color,
    DefaultTheme,
    FillSpec,
    FontSpec,
    StyleSpec,
    Workbook,
    WorkbookSettings,
)

title = StyleSpec(
    name='title',
    font=FontSpec(size=16, bold=True, color=Color.parse('#FFFFFF')),
    fill=FillSpec.solid(Color.parse('#007CD6')),
)
percent = StyleSpec(name='pct', number_format='0.0%')
records = [{'Name': 'Ann', 'Amount': 0.12}, {'Name': 'Bob', 'Amount': 0.88}]

with Workbook.new() as wb:
    sheet = wb.sheet('Sales')
    sheet.cell(CellRef.parse('A1'), 'Q1', style=title)
    sheet.frame(records, at=CellRef.parse('A3'), styles={'Amount': percent})
    wb.save('sales.xlsx')

money = StyleSpec(name='usd', number_format='$#,##0.00')
settings = WorkbookSettings(theme=DefaultTheme().with_overrides(float=money))
with Workbook.new(settings=settings) as wb:
    wb.sheet('Sales').frame(records)
    wb.save('money.xlsx')
```

To reuse an Excel built-in style instead of creating a new named style:

```python
StyleSpec(name='my_percent', builtin='Percent')
StyleSpec(name='my_date', builtin='Normal', number_format='mm-dd-yy')
```

## Dropdowns

```python
sheet.dropdown(CellRef.parse('A4'), ['Ann', 'Bob'])
sheet.dropdown(CellRange.parse('B1:B10'), CellRange.parse('A1:A3'))  # Options already on the sheet.
```

A comma or a quote inside an option would split Excel's list. Put those options on the sheet and pass a `CellRange`. A single string `'Yes'` is rejected: Python would iterate characters.

| Call | What it does |
|---|---|
| `sheet.dropdown(area, options, allow_blank=True)` | Pick-list on a cell or a range |

## Sheet protection

```python
sheet.protect()  # Lock the sheet.
sheet.protect('secret')  # Lock with a password.
sheet.unprotect()
```

By default every cell is locked. After `protect()`, nothing is editable unless a style uses `ProtectionSpec(locked=False)`.

| Call | What it does |
|---|---|
| `sheet.protect(password=None)` | Enable sheet protection |
| `sheet.unprotect()` | Disable it |

## Stream

Write-only, for large tables. No `cell()`, `read()`, or `merge()`. Column widths and a frozen header are set when the sheet is created. `protect` and `dropdown` still work: they decorate the sheet, they do not write a cell.

```python
from flowxl import Workbook

rows = [{'Name': 'Ann', 'Amount': 10}, {'Name': 'Bob', 'Amount': 20}]

with Workbook.stream() as wb:
    sheet = wb.sheet('Events', pin_header=True, widths={'A': 24, 'B': 12})
    sheet.frame(rows)
    wb.save('events.xlsx')
```

Streaming `save` / `to_bytes` **close** the workbook. Create a new one to write again.

## Errors

Catch `FlowxlError` for any FlowXL failure.

- `UsageError` — the call is wrong (bad address, formula without `=`, unknown type). Fix the program.
- `BackendError` — the file, the disk, or openpyxl failed. Try another path.

Messages are English. The text always says what happened and what to try.

## What this library does not do

See [scope.md](scope.md). It does not evaluate formulas, read `.xls` / `.csv`, or replace openpyxl.

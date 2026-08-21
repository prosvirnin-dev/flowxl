# FlowXL

A wrapper over openpyxl: predictable `.xlsx` writing.

## What it does

- Create `.xlsx` from scratch with a declarative fluent API
- Open and edit existing `.xlsx` files in place
- Raise strict contextual errors instead of silently corrupting data
- Write tabular structures with type-based auto-formatting

## What it does not do

- Does not read `.xls`, `.xlsb`, `.csv`, or `.ods`
- Does not evaluate formulas
- Does not render charts or images
- Does not replace openpyxl
- Does not perform data analysis
- Does not provide an async API

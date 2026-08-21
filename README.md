# FlowXL

Predictable Excel writing over openpyxl.

Loud errors instead of silent data corruption. Fluent API for creating and editing `.xlsx` files.

Python 3.12+. The only required dependency is openpyxl.

**Docs:** [English](docs/usage.md) · [Русский](docs/usage.ru.md)

## Install

```bash
pip install -e .
pip install -e '.[polars]'
pip install -e '.[pandas]'
```

## Scope

See [docs/scope.md](docs/scope.md). FlowXL does not evaluate formulas, read `.xls` / `.csv`, or replace openpyxl.

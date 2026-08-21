"""Built-in Excel styles for primitive formats."""

from datetime import date
from pathlib import Path

from openpyxl import load_workbook

from flowxl import Workbook


def test_primitive_formats_use_excel_builtin_styles(tmp_path: Path) -> None:
    """Integer, date and header: Excel Comma/Normal, plus one custom header style."""
    path: Path = tmp_path / 'styled.xlsx'
    records: list[dict[str, object]] = [
        {'Name': 'Ann', 'Amount': 10, 'Day': date(2024, 1, 2)},
    ]
    with Workbook.new() as wb:
        wb.sheet('Sales').frame(records)
        wb.save(path)

    opened = load_workbook(path)
    names = set(opened.style_names)
    assert 'sw_integer' not in names
    assert 'sw_date' not in names
    assert 'sw_float' not in names
    assert 'sw_header' in names
    assert 'Comma [0]' in names
    sheet = opened['Sales']
    assert sheet['B2'].style == 'Comma [0]'
    assert sheet['C2'].style == 'Normal'
    assert sheet['C2'].number_format == 'mm-dd-yy'
    assert sheet['A1'].style == 'sw_header'

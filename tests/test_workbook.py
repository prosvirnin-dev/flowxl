"""Workbook write and frame() tests."""

from datetime import date
from pathlib import Path

from flowxl import CellRef, Workbook


def test_cell_round_trip(tmp_path: Path) -> None:
    """Write a cell, save, reopen and read the same value."""
    path: Path = tmp_path / 'out.xlsx'
    with Workbook.new() as wb:
        wb.sheet('Demo').cell(CellRef.parse('A1'), 'hello')
        wb.save(path)
    with Workbook.open(path) as wb:
        assert wb.sheet('Demo').read(CellRef.parse('A1')) == 'hello'


def test_frame_writes_header_and_rows() -> None:
    """Write a list of dicts and keep the header, values and bounds."""
    records: list[dict[str, object]] = [
        {'Name': 'Ann', 'Amount': 10, 'Day': date(2024, 1, 2)},
        {'Name': 'Bob', 'Amount': 20, 'Day': date(2024, 1, 3)},
    ]
    with Workbook.new() as wb:
        ws = wb.sheet('Sales').frame(records)
        assert ws.read(CellRef.parse('A1')) == 'Name'
        assert ws.read(CellRef.parse('B2')) == 10
        assert ws.bounds is not None
        assert ws.bounds.to_a1() == 'A1:C3'

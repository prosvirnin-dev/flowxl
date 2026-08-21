"""Dropdown list tests."""

import pytest

from flowxl import CellRange, CellRef, UsageError, Workbook


def test_dropdown_literal_list() -> None:
    """A list of strings becomes formula1 '"a,b"' on the cell."""
    with Workbook.new() as wb:
        sheet = wb.sheet('Demo')
        sheet.dropdown(CellRef.parse('A1'), ['Yes', 'No'])
        validations = sheet.raw_openpyxl().data_validations.dataValidation
        assert len(validations) == 1
        assert validations[0].formula1 == '"Yes,No"'
        assert validations[0].type == 'list'
        assert str(validations[0].sqref) == 'A1'


def test_dropdown_from_range() -> None:
    """A CellRange of options is stored as an absolute A1 formula."""
    with Workbook.new() as wb:
        sheet = wb.sheet('Demo')
        sheet.dropdown(CellRange.parse('B1:B3'), CellRange.parse('A1:A3'))
        validations = sheet.raw_openpyxl().data_validations.dataValidation
        assert validations[0].formula1 == '$A$1:$A$3'
        assert str(validations[0].sqref) == 'B1:B3'


def test_dropdown_rejects_a_bare_string() -> None:
    """A single string would be iterated as characters. Refuse it."""
    with Workbook.new() as wb:
        sheet = wb.sheet('Demo')
        with pytest.raises(UsageError, match='not one string'):
            sheet.dropdown(CellRef.parse('A1'), 'Yes')


def test_dropdown_rejects_comma_in_option() -> None:
    """A comma inside an option would split the Excel list. Point to a range."""
    with Workbook.new() as wb:
        sheet = wb.sheet('Demo')
        with pytest.raises(UsageError, match='comma or a quote'):
            sheet.dropdown(CellRef.parse('A1'), ['A, B'])

"""Sheet protection tests."""

from flowxl import CellRef, Workbook


def test_protect_and_unprotect() -> None:
    """Lock the sheet, then unlock it."""
    with Workbook.new() as wb:
        sheet = wb.sheet('Demo')
        sheet.cell(CellRef.parse('A1'), 1)
        sheet.protect()
        raw = sheet.raw_openpyxl()
        assert raw.protection.sheet is True
        sheet.unprotect()
        assert raw.protection.sheet is False


def test_protect_with_password() -> None:
    """A password enables protection and hashes the secret."""
    with Workbook.new() as wb:
        sheet = wb.sheet('Demo')
        sheet.protect('secret')
        raw = sheet.raw_openpyxl()
        assert raw.protection.sheet is True
        assert raw.protection.password

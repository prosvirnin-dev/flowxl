"""Russian formula translation tests."""

import pytest

from flowxl import CellRef, UsageError, Workbook, to_excel_formula


def test_english_formula_unchanged() -> None:
    """Leave an already English formula alone, including argument commas."""
    assert to_excel_formula('=SUM(A1,A2)') == '=SUM(A1,A2)'
    assert to_excel_formula('=IF(A1>0,A1,0)') == '=IF(A1>0,A1,0)'


def test_russian_function_and_semicolon() -> None:
    """Translate СУММ and turn the argument semicolon into a comma."""
    assert to_excel_formula('=СУММ(A1;A2)') == '=SUM(A1,A2)'


def test_decimal_comma_next_to_russian_function() -> None:
    """A comma between digits is a decimal when the formula is Russian."""
    assert to_excel_formula('=СУММ(1,5;2)') == '=SUM(1.5,2)'


def test_decimal_comma_without_function() -> None:
    """``=A1+1,5`` uses a decimal comma, not an English argument list."""
    assert to_excel_formula('=A1+1,5') == '=A1+1.5'


def test_english_sum_of_two_numbers_is_not_a_decimal() -> None:
    """``=SUM(1,5)`` is two arguments, not 1.5."""
    assert to_excel_formula('=SUM(1,5)') == '=SUM(1,5)'


def test_if_with_constants_and_decimal() -> None:
    """Translate ЕСЛИ, ИСТИНА, ЛОЖЬ and a decimal comma together."""
    assert to_excel_formula('=ЕСЛИ(A1>1,5;ИСТИНА;ЛОЖЬ)') == '=IF(A1>1.5,TRUE,FALSE)'


def test_yo_and_ye_fold() -> None:
    """СЧЁТ and СЧЕТ are the same function after folding yo."""
    assert to_excel_formula('=СЧЁТ(A1:A10)') == '=COUNT(A1:A10)'
    assert to_excel_formula('=СЧЕТ(A1:A10)') == '=COUNT(A1:A10)'


def test_string_is_left_alone() -> None:
    """A semicolon inside quotes is not an argument separator."""
    assert to_excel_formula('="a;b"') == '="a;b"'
    assert to_excel_formula('=СЦЕПИТЬ("а,б";"в")') == '=CONCATENATE("а,б","в")'


def test_sheet_name_is_left_alone() -> None:
    """Cyrillic sheet names in quotes are not function names."""
    assert to_excel_formula("='Лист1'!A1") == "='Лист1'!A1"
    assert to_excel_formula("=СУММ('Продажи'!A1:A10)") == "=SUM('Продажи'!A1:A10)"


def test_english_semicolon_still_converts() -> None:
    """English names with Russian separators are a common mixed form."""
    assert to_excel_formula('=SUM(A1;A2)') == '=SUM(A1,A2)'


def test_unknown_cyrillic_function_raises() -> None:
    """A Cyrillic call that is not in the table is a usage error, not a silent store."""
    with pytest.raises(UsageError, match='Unknown Russian Excel function'):
        to_excel_formula('=НЕИЗВЕСТНАЯФУНКЦИЯ(A1)')


def test_requires_leading_equals() -> None:
    """The same rule as Sheet.formula: no silent 'add the equals'."""
    with pytest.raises(UsageError, match='must start with "="'):
        to_excel_formula('СУММ(A1)')


def test_formula_write_translates_russian() -> None:
    """Store the en-US form, even when the caller wrote Russian."""
    with Workbook.new() as wb:
        sheet = wb.sheet('Demo')
        sheet.formula(CellRef.parse('A1'), '=СУММ(B1;C1)')
        assert sheet.read(CellRef.parse('A1')) == '=SUM(B1,C1)'


def test_cell_write_translates_a_formula_string() -> None:
    """cell() is another write path; it must not skip translation."""
    with Workbook.new() as wb:
        sheet = wb.sheet('Demo')
        sheet.cell(CellRef.parse('A1'), '=СРЗНАЧ(A2:A10)')
        assert sheet.read(CellRef.parse('A1')) == '=AVERAGE(A2:A10)'

"""Cell coordinate tests."""

import pytest

from flowxl import CellRef, InvalidCoordinateError
from flowxl.primitives.cell import index_to_letter


def test_a1_round_trip() -> None:
    """Parse an A1 address and return the same string."""
    assert CellRef.parse('B12').to_a1() == 'B12'


def test_absolute_a1() -> None:
    """Prefix both parts with $ when asked."""
    assert CellRef.parse('B12').to_a1(absolute=True) == '$B$12'


def test_rejects_row_zero() -> None:
    """Reject a zero row. Coordinates are counted from 1, as in Excel."""
    with pytest.raises(InvalidCoordinateError):
        CellRef(0, 1)


@pytest.mark.parametrize(
    ('index', 'letters'),
    [(1, 'A'), (26, 'Z'), (27, 'AA'), (52, 'AZ'), (53, 'BA')],
)
def test_column_letters(index: int, letters: str) -> None:
    """Check letter transitions, especially Z to AA."""
    assert index_to_letter(index) == letters

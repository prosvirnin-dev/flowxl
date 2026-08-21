"""Immutable cell and range coordinates.

Coordinates are 1-based, matching Excel and the xlsx format. The rest of the
library never asks whether an address fits on the sheet: an invalid
:class:`CellRef` cannot be constructed.
"""

from __future__ import annotations

import re
from collections.abc import Iterator
from dataclasses import dataclass
from typing import Final, Self

from flowxl.primitives.exceptions import InvalidCoordinateError, InvalidRangeError

__all__ = [
    'MAX_COL',
    'MAX_ROW',
    'ORIGIN',
    'CellRange',
    'CellRef',
    'index_to_letter',
    'letter_to_index',
]

# Последняя строка листа в формате xlsx. Ограничение Excel.
MAX_ROW: Final[int] = 1_048_576
# Последняя колонка листа (XFD). Ограничение Excel.
MAX_COL: Final[int] = 16_384

# Одна ячейка: необязательные знаки доллара, буквы колонки, цифры строки.
# Примеры, которые проходят: B12, $B$12, aa999.
_A1_RE: Final[re.Pattern[str]] = re.compile(r'^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$')
# Диапазон: два адреса через двоеточие, например A1:C10.
_RANGE_RE: Final[re.Pattern[str]] = re.compile(r'^(.+?):(.+)$')

# Размер английского алфавита. Нужен для перевода номера колонки в буквы.
_ALPHABET_SIZE: Final[int] = 26
# Код буквы A. Дальше смещение remainder дает нужную букву.
_LETTER_A: Final[int] = ord('A')


def index_to_letter(col: int) -> str:
    """Convert a 1-based column index into Excel column letters.

    Examples:
        1 becomes A, 26 becomes Z, 27 becomes AA.

    Args:
        col: Column index counted from 1. A is 1, not 0.

    Returns:
        Column letters such as ``B`` or ``AA``.

    Raises:
        InvalidCoordinateError: If ``col`` is outside ``1..MAX_COL``.
    """
    if not 1 <= col <= MAX_COL:
        raise InvalidCoordinateError(f'Column index must be in 1..{MAX_COL}, got {col}.')

    letters: list[str] = []
    # Excel считает колонки без нуля: после Z идет AA, а не BA. Поэтому перед делением вычитаем единицу.
    # Без этого переход Z в AA ломается.
    while col > 0:
        col, remainder = divmod(col - 1, _ALPHABET_SIZE)
        letters.append(chr(_LETTER_A + remainder))

    # Младшие разряды попадали в список первыми, поэтому буквы разворачиваем.
    return ''.join(reversed(letters))


def letter_to_index(letters: str) -> int:
    """Convert Excel column letters into a 1-based column index.

    Args:
        letters: Column letters, case-insensitive, e.g. ``ab``.

    Returns:
        Column index counted from 1.

    Raises:
        InvalidCoordinateError: If the string is empty, non-alphabetic, or past the last column.
    """
    if not letters or not letters.isalpha():
        raise InvalidCoordinateError(f'Column letters must be alphabetic, got {letters!r}.')

    index: int = 0
    # Обратная операция к index_to_letter. Каждая буква это разряд в системе с основанием 26, где A равно 1, а не 0.
    for char in letters.upper():
        index = index * _ALPHABET_SIZE + (ord(char) - _LETTER_A + 1)

    if index > MAX_COL:
        raise InvalidCoordinateError(f'Column {letters!r} resolves to {index}, past the last column ({MAX_COL}).')
    return index


@dataclass(frozen=True, order=True)
class CellRef:
    """A single cell address.

    An instance is always inside the xlsx grid and never changes after creation.
    Use :meth:`offset` to get a different address.

    Attributes:
        row: Row number counted from 1. The first row is 1, not 0.
        col: Column number counted from 1. Column A is 1.
    """

    row: int
    col: int

    def __post_init__(self) -> None:
        """Validate coordinates right after construction.

        Dataclass calls this method automatically after ``__init__``. A bad coordinate never becomes an object.
        The check lives in one place instead of being spread across the library.
        """
        if not 1 <= self.row <= MAX_ROW:
            raise InvalidCoordinateError(f'Row must be in 1..{MAX_ROW}, got {self.row}. Coordinates are 1-based.')
        if not 1 <= self.col <= MAX_COL:
            raise InvalidCoordinateError(f'Column must be in 1..{MAX_COL}, got {self.col}. Coordinates are 1-based.')

    @classmethod
    def parse(cls, a1: str) -> Self:
        """Build a reference from A1 notation.

        Args:
            a1: Address such as ``B12`` or ``$B$12``. Leading and trailing
                whitespace is stripped. Dollar signs are ignored.

        Returns:
            A validated cell reference.

        Raises:
            InvalidCoordinateError: If the string is not a valid A1 address.
        """
        match: re.Match[str] | None = _A1_RE.match(a1.strip())
        if match is None:
            raise InvalidCoordinateError(f'Cannot parse {a1!r} as an A1 cell address. Expected something like "B12".')

        letters: str
        digits: str
        letters, digits = match.groups()
        # Цифры становятся номером строки, буквы проходят через letter_to_index.
        # Вызов cls снова идет через __post_init__, это вторая страховка.
        return cls(row=int(digits), col=letter_to_index(letters))

    @property
    def column_letter(self) -> str:
        """Excel letters for this cell's column.

        Returns:
            Letters such as ``B``.
        """
        return index_to_letter(self.col)

    def offset(self, rows: int = 0, cols: int = 0) -> Self:
        """Return a new reference shifted by the given deltas.

        The original object is not modified: the dataclass is frozen.

        Args:
            rows: Vertical shift. Positive moves down, negative moves up.
            cols: Horizontal shift. Positive moves right, negative moves left.

        Returns:
            A new reference at the shifted position.

        Raises:
            InvalidCoordinateError: If the result falls outside the sheet.
        """
        # Создаем новый объект вместо изменения полей.
        # Старый адрес остается валидным снимком, на который можно опираться в кэше и в границах листа.
        return type(self)(row=self.row + rows, col=self.col + cols)

    def span(self, rows: int, cols: int) -> CellRange:
        """Return a range anchored at this cell.

        Args:
            rows: Number of rows in the range, at least 1.
            cols: Number of columns in the range, at least 1.

        Returns:
            A range of the given size, starting here.
        """
        return CellRange.from_size(self, rows=rows, cols=cols)

    def to_a1(self, *, absolute: bool = False) -> str:
        """Return A1 notation.

        Args:
            absolute: If True, prefix column and row with ``$``, as in ``$B$12``.

        Returns:
            An address such as ``B12`` or ``$B$12``.
        """
        if absolute:
            return f'${self.column_letter}${self.row}'
        return f'{self.column_letter}{self.row}'

    def __str__(self) -> str:
        """Return the same string as :meth:`to_a1`.

        Returns:
            An address such as ``B12``.
        """
        return self.to_a1()


# Ячейка A1. Самый частый старт для шапки таблицы и для frame().
ORIGIN: Final[CellRef] = CellRef(row=1, col=1)


@dataclass(frozen=True)
class CellRange:
    """A rectangular block of cells.

    The top-left corner is always above and left of the bottom-right.
    Both corners are validated :class:`CellRef` instances.

    Attributes:
        start: Top-left corner of the range.
        end: Bottom-right corner of the range.
    """

    start: CellRef
    end: CellRef

    def __post_init__(self) -> None:
        """Ensure start is above and left of end.

        Excel does not accept a flipped range. Failing here is more honest than silently swapping the corners.
        """
        if self.start.row > self.end.row or self.start.col > self.end.col:
            raise InvalidRangeError(
                f'Range start {self.start.to_a1()} must be above and left of end {self.end.to_a1()}.'
            )

    @classmethod
    def parse(cls, a1: str) -> Self:
        """Build a range from ``A1:C10`` notation.

        Args:
            a1: Range string with a colon between two cell addresses.

        Returns:
            A validated range.

        Raises:
            InvalidRangeError: If the string is not valid range notation.
            InvalidCoordinateError: If either corner is invalid.
        """
        match: re.Match[str] | None = _RANGE_RE.match(a1.strip())
        if match is None:
            raise InvalidRangeError(f'Cannot parse {a1!r} as a range. Expected "A1:C10".')

        left: str
        right: str
        left, right = match.groups()
        # Каждый угол разбирается как отдельная ячейка. Так мы не дублируем разбор A1.
        return cls(start=CellRef.parse(left), end=CellRef.parse(right))

    @classmethod
    def from_size(cls, at: CellRef, *, rows: int, cols: int) -> Self:
        """Build a range of the given size anchored at ``at``.

        Args:
            at: Top-left corner.
            rows: Number of rows, at least 1.
            cols: Number of columns, at least 1.

        Returns:
            A range of the given size.

        Raises:
            InvalidRangeError: If ``rows`` or ``cols`` is below 1.
            InvalidCoordinateError: If the computed bottom-right corner falls
                outside the sheet.
        """
        if rows < 1 or cols < 1:
            raise InvalidRangeError(f'A range needs at least one row and one column, got rows={rows}, cols={cols}.')

        # Смещение на rows минус один: диапазон 1 на 1 совпадает в start и end.
        # Три строки от A1 заканчиваются в A3, то есть плюс две строки.
        return cls(start=at, end=at.offset(rows=rows - 1, cols=cols - 1))

    @property
    def height(self) -> int:
        """Number of rows covered, inclusive of both corners.

        Returns:
            Height of the range in rows.
        """
        # Единица нужна потому что оба угла входят в диапазон.
        return self.end.row - self.start.row + 1

    @property
    def width(self) -> int:
        """Number of columns covered, inclusive of both corners.

        Returns:
            Width of the range in columns.
        """
        return self.end.col - self.start.col + 1

    def union(self, other: CellRange) -> CellRange:
        """Return the smallest range that covers both operands.

        The library uses this to remember everything written so far. Each new
        cell grows the box. Calling it twice for the same area is safe.

        Args:
            other: The other range.

        Returns:
            The minimal enclosing rectangle.
        """
        # Минимум по углам start и максимум по углам end дают наименьший прямоугольник.
        # Он накрывает оба диапазона.
        return CellRange(
            start=CellRef(
                row=min(self.start.row, other.start.row),
                col=min(self.start.col, other.start.col),
            ),
            end=CellRef(
                row=max(self.end.row, other.end.row),
                col=max(self.end.col, other.end.col),
            ),
        )

    def contains(self, ref: CellRef) -> bool:
        """Return whether a cell lies inside this range.

        Args:
            ref: Cell to test.

        Returns:
            True if the cell is inside, inclusive of the border.
        """
        return self.start.row <= ref.row <= self.end.row and self.start.col <= ref.col <= self.end.col

    def iter_refs(self) -> Iterator[CellRef]:
        """Iterate every cell in the range, row by row, left to right.

        Yields:
            Each reference inside the range.
        """
        # Внешний цикл идет по строкам, внутренний по колонкам. Так читает Excel.
        for row in range(self.start.row, self.end.row + 1):
            for col in range(self.start.col, self.end.col + 1):
                yield CellRef(row=row, col=col)

    def to_a1(self, *, absolute: bool = False) -> str:
        """Return range notation.

        Args:
            absolute: If True, both corners use ``$``, as in ``$A$1:$C$10``.

        Returns:
            A string such as ``A1:C10`` or ``$A$1:$C$10``.
        """
        return f'{self.start.to_a1(absolute=absolute)}:{self.end.to_a1(absolute=absolute)}'

    def __str__(self) -> str:
        """Return the same string as :meth:`to_a1`.

        Returns:
            A string such as ``A1:C10``.
        """
        return self.to_a1()

"""Shared sheet behaviour: the frame() algorithm and bounds tracking."""

from __future__ import annotations

from abc import ABC, abstractmethod
from collections.abc import Mapping, Sequence
from typing import Final, Self

from flowxl.api.settings import WorkbookSettings
from flowxl.backend.openpyxl_backend import OpenpyxlSheetBase
from flowxl.integrations.protocol import TabularSource
from flowxl.integrations.registry import adapt
from flowxl.primitives.cell import ORIGIN, CellRange, CellRef
from flowxl.primitives.exceptions import UnsupportedTypeError, UntrackedBoundsError, UsageError
from flowxl.primitives.sheet_name import SheetName
from flowxl.styles.spec import StyleSpec

__all__ = ['BaseSheet']

_DROPDOWN_MAX_CHARS: Final[int] = 255


def _dropdown_source(options: Sequence[str] | CellRange) -> str:
    """Build the list-validation formula Excel stores.

    Args:
        options: Literal values or a range that already holds them.

    Returns:
        ``'"a,b"'`` or an absolute range such as ``$A$1:$A$3``.

    Raises:
        UsageError: If the options cannot be stored as a list.
    """
    if isinstance(options, CellRange):
        return options.to_a1(absolute=True)
    if isinstance(options, (str, bytes)):
        raise UsageError(
            'dropdown options must be a sequence of strings or a CellRange, not one string. '
            'Pass a list such as ["Yes", "No"].'
        )
    items: list[str] = list(options)
    if not items:
        raise UsageError('dropdown needs at least one option.')
    for item in items:
        if not isinstance(item, str):
            raise UsageError(f'dropdown options must be strings, got {type(item).__name__}.')
        if ',' in item or '"' in item:
            raise UsageError(
                f'dropdown option {item!r} contains a comma or a quote. '
                f'Put the options on the sheet and pass a CellRange instead.'
            )
    formula: str = '"' + ','.join(items) + '"'
    if len(formula) > _DROPDOWN_MAX_CHARS:
        raise UsageError(
            f'dropdown list is {len(formula)} characters; Excel allows {_DROPDOWN_MAX_CHARS}. '
            f'Put the options on the sheet and pass a CellRange instead.'
        )
    return formula


class BaseSheet(ABC):
    """Common behaviour of every sheet.

    :attr:`bounds` always covers every cell written through this object,
    unless :meth:`~flowxl.api.sheet.Sheet.raw_openpyxl` was used.

    The ``frame`` algorithm lives here once. Subclasses supply the two steps
    that actually differ: how to write the header and how to write the body.
    """

    def __init__(
        self,
        name: SheetName,
        backend: OpenpyxlSheetBase,
        settings: WorkbookSettings,
    ) -> None:
        """Store the name, backend and settings shared by every sheet.

        Args:
            name: Validated sheet name.
            backend: openpyxl adapter.
            settings: Workbook-wide configuration.
        """
        self._name: SheetName = name
        self._backend: OpenpyxlSheetBase = backend
        self._settings: WorkbookSettings = settings
        self._bounds: CellRange | None = None
        self._bounds_trusted: bool = True
        self._content_width: dict[int, int] = {}

    @property
    def name(self) -> SheetName:
        """Return the sheet's name.

        Returns:
            Validated sheet name.
        """
        return self._name

    @property
    def bounds(self) -> CellRange | None:
        """Return the area of everything written so far, if known.

        Returns:
            Bounding range, or ``None`` if nothing has been written yet.
        """
        return self._bounds

    def frame(
        self,
        data: object,
        *,
        at: CellRef = ORIGIN,
        header: bool = True,
        styles: Mapping[str, StyleSpec] | None = None,
    ) -> Self:
        """Write a table, formatting columns by logical type.

        Args:
            data: A polars or pandas frame, a list of records, or any object
                implementing :class:`~flowxl.integrations.protocol.TabularSource`.
            at: Top-left corner of the block.
            header: Whether to write column names first.
            styles: Per-column style overrides, keyed by column name.

        Returns:
            This sheet, for chaining.

        Raises:
            UnsupportedTypeError: If a column type has no style in the theme.
        """
        # Функция adapt смотрит корневой модуль типа: список, polars или pandas.
        source: TabularSource = adapt(data)
        if not source.columns:
            return self

        column_styles: list[StyleSpec] = self._resolve_column_styles(source, styles)
        cursor: CellRef = at

        if header:
            # Как именно чиркнуть шапку, решает наследник: по клеткам или append.
            self._write_header(cursor, source.columns)
            cursor = cursor.offset(rows=1)

        written: int = self._write_body(cursor, source, column_styles)
        total: int = written + (1 if header else 0)
        if total:
            # Границы расширяем одним прямоугольником, а не по каждой ячейке.
            self._extend_bounds(CellRange.from_size(at, rows=total, cols=len(source.columns)))
        return self

    def autofilter(self, area: CellRange | None = None) -> Self:
        """Enable the autofilter over ``area`` or over everything written.

        Args:
            area: Explicit range. If omitted, tracked bounds are used.

        Returns:
            This sheet, for chaining.

        Raises:
            UntrackedBoundsError: If no range is given and bounds are unknown
                or no longer trusted.
        """
        if area is not None:
            self._backend.set_autofilter(area)
            return self
        if self._bounds is None:
            raise UntrackedBoundsError(
                f'Sheet {self._name.value!r} has no data yet, so autofilter() '
                f'has nothing to cover. Write data first or pass area=.'
            )
        if not self._bounds_trusted:
            raise UntrackedBoundsError(
                f'Sheet {self._name.value!r} was modified through raw_openpyxl(), '
                f'so its bounds are no longer tracked. Pass area= explicitly.'
            )
        self._backend.set_autofilter(self._bounds)
        return self

    def protect(self, password: str | None = None) -> Self:
        """Lock the sheet against edits.

        Cells stay editable when their style uses ``ProtectionSpec(locked=False)``.
        By default every cell is locked, so this call freezes the whole sheet.

        Args:
            password: Optional password. Omit it to lock without one.

        Returns:
            This sheet, for chaining.
        """
        self._backend.protect_sheet(password)
        return self

    def unprotect(self) -> Self:
        """Unlock the sheet.

        Returns:
            This sheet, for chaining.
        """
        self._backend.unprotect_sheet()
        return self

    def dropdown(
        self,
        area: CellRef | CellRange,
        options: Sequence[str] | CellRange,
        *,
        allow_blank: bool = True,
    ) -> Self:
        """Add a dropdown list to one cell or a rectangle.

        Args:
            area: Cell or range that shows the list.
            options: Values to pick, or a range that already holds them.
            allow_blank: Whether an empty cell is accepted.

        Returns:
            This sheet, for chaining.

        Raises:
            UsageError: If ``options`` is empty, is a single string, contains
                a comma or a quote, or is longer than Excel allows.
        """
        target: CellRange = area if isinstance(area, CellRange) else CellRange(area, area)
        source: str = _dropdown_source(options)
        self._backend.add_dropdown(target, source, allow_blank=allow_blank)
        return self

    @abstractmethod
    def _write_header(self, at: CellRef, columns: Sequence[str]) -> None:
        """Write the header row starting at ``at``.

        Args:
            at: First header cell.
            columns: Column names left to right.
        """

    @abstractmethod
    def _write_body(
        self,
        at: CellRef,
        source: TabularSource,
        styles: Sequence[StyleSpec],
    ) -> int:
        """Write the data rows and return how many were written.

        Args:
            at: First data cell.
            source: Tabular data.
            styles: One style per column.

        Returns:
            Number of data rows written.
        """

    def _resolve_column_styles(
        self,
        source: TabularSource,
        overrides: Mapping[str, StyleSpec] | None,
    ) -> list[StyleSpec]:
        """Pick one style per column and add context to any failure.

        Args:
            source: Table with names and types.
            overrides: Optional style map keyed by column name.

        Returns:
            One style per column, left to right.

        Raises:
            UnsupportedTypeError: If the theme has no style for a column type.
        """
        theme = self._settings.theme
        resolved: list[StyleSpec] = []
        pairs = zip(source.columns, source.logical_types, strict=True)
        for index, (column, logical_type) in enumerate(pairs, start=1):
            if overrides and column in overrides:
                resolved.append(overrides[column])
                continue
            try:
                resolved.append(theme.style_for(logical_type))
            except UnsupportedTypeError as exc:
                raise UnsupportedTypeError(f'Sheet {self._name.value!r}, column {column!r} (#{index}): {exc}') from exc
        return resolved

    def _extend_bounds(self, area: CellRange) -> None:
        """Grow the tracked area. Safe to call twice for the same cells.

        Args:
            area: Rectangle just written.
        """
        self._bounds = area if self._bounds is None else self._bounds.union(area)

    def _track_width(self, col: int, value: object) -> None:
        """Remember the longest text value seen in a column.

        Args:
            col: Column index counted from 1.
            value: Written value. ``None`` counts as zero width.
        """
        length: int = len(str(value)) if value is not None else 0
        if length > self._content_width.get(col, 0):
            self._content_width[col] = length

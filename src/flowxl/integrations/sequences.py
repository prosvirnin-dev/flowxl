"""Adapter for a list of dicts and a list of tuples. No third-party dependencies."""

from __future__ import annotations

from collections.abc import Iterator, Sequence
from dataclasses import dataclass
from typing import Self

from flowxl.primitives.exceptions import UnsupportedTypeError
from flowxl.primitives.types import LogicalType, infer_logical_type

__all__ = ['SequenceSource']


@dataclass(frozen=True)
class SequenceSource:
    """A table from ``list[dict]`` or ``list[tuple]`` plus column names.

    Types are taken from the first non-empty value of each column.
    """

    _columns: tuple[str, ...]
    _rows: tuple[Sequence[object], ...]
    _types: tuple[LogicalType, ...]

    @classmethod
    def from_records(cls, records: object, columns: Sequence[str] | None = None) -> Self:
        """Build a source from a list of dicts or a list of tuples.

        Args:
            records: The data.
            columns: Required if rows are tuples. For dicts you may omit this
                unless you want to set the order explicitly.

        Returns:
            A tabular source.

        Raises:
            UnsupportedTypeError: If the data shape is not recognised.
        """
        if not isinstance(records, (list, tuple)):
            raise UnsupportedTypeError(f'Expected a list or tuple of rows, got {type(records).__name__!r}.')
        if not records:
            return cls(tuple(columns or ()), (), ())

        first: object = records[0]
        names: tuple[str, ...]
        rows: tuple[Sequence[object], ...]
        if isinstance(first, dict):
            # Порядок колонок берём из первого словаря, если вызывающий его не задал.
            names = tuple(columns) if columns else tuple(first)
            rows = tuple(tuple(row.get(name) for name in names) for row in records)
        elif isinstance(first, (list, tuple)):
            if columns is None:
                raise UnsupportedTypeError('Rows are tuples, so column names must be supplied via columns=.')
            names = tuple(columns)
            rows = tuple(tuple(row) for row in records)
        else:
            raise UnsupportedTypeError(f'Rows must be dicts, lists or tuples, got {type(first).__name__!r}.')
        return cls(names, rows, cls._infer_types(names, rows))

    @staticmethod
    def _infer_types(
        columns: Sequence[str],
        rows: Sequence[Sequence[object]],
    ) -> tuple[LogicalType, ...]:
        """Infer one type per column from the first non-empty value.

        Args:
            columns: Column names, needed only for length.
            rows: Data rows.

        Returns:
            One logical type per column. Empty columns become text.
        """
        inferred: list[LogicalType] = []
        for index in range(len(columns)):
            found: LogicalType = LogicalType.TEXT
            for row in rows:
                value: object = row[index] if index < len(row) else None
                if value is not None:
                    found = infer_logical_type(value)
                    break
            inferred.append(found)
        return tuple(inferred)

    @property
    def columns(self) -> Sequence[str]:
        """Return column names.

        Returns:
            Column names left to right.
        """
        return self._columns

    @property
    def logical_types(self) -> Sequence[LogicalType]:
        """Return one logical type per column.

        Returns:
            Logical types in :attr:`columns` order.
        """
        return self._types

    @property
    def row_count(self) -> int | None:
        """Return the number of rows.

        Returns:
            Row count. For an in-memory sequence this is never ``None``.
        """
        return len(self._rows)

    def iter_rows(self) -> Iterator[Sequence[object]]:
        """Iterate the stored rows.

        Yields:
            One row of values.
        """
        yield from self._rows

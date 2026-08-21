"""polars DataFrame adapter. Imported only on demand."""

from __future__ import annotations

from collections.abc import Iterator, Sequence
from typing import Final

import polars as pl

from flowxl.primitives.types import LogicalType

__all__ = ['PolarsSource']

# Ключ это базовый тип, не type(dtype).
# Datetime('us', 'UTC') иначе не совпадёт с pl.Datetime и колонка уедет в неизвестный тип.
_BASE_MAP: Final[dict[type, LogicalType]] = {
    pl.Int8: LogicalType.INTEGER,
    pl.Int16: LogicalType.INTEGER,
    pl.Int32: LogicalType.INTEGER,
    pl.Int64: LogicalType.INTEGER,
    pl.UInt8: LogicalType.INTEGER,
    pl.UInt16: LogicalType.INTEGER,
    pl.UInt32: LogicalType.INTEGER,
    pl.UInt64: LogicalType.INTEGER,
    pl.Float32: LogicalType.FLOAT,
    pl.Float64: LogicalType.FLOAT,
    pl.Decimal: LogicalType.FLOAT,
    pl.Date: LogicalType.DATE,
    pl.Datetime: LogicalType.DATETIME,
    pl.Time: LogicalType.TIME,
    pl.Duration: LogicalType.DURATION,
    pl.Boolean: LogicalType.BOOLEAN,
    pl.String: LogicalType.TEXT,
    pl.Categorical: LogicalType.TEXT,
    pl.Enum: LogicalType.TEXT,
}


def to_logical_type(dtype: pl.DataType) -> LogicalType:
    """Classify a polars dtype.

    Uses ``base_type()``, not ``type(dtype)``, so parameterised types such as
    ``Datetime('us', 'UTC')`` are found correctly.

    Args:
        dtype: A polars data type.

    Returns:
        The matching logical type, or :attr:`LogicalType.UNKNOWN`.
    """
    mapped: LogicalType | None = _BASE_MAP.get(dtype.base_type())
    if mapped is not None:
        return mapped
    if dtype.is_integer():
        return LogicalType.INTEGER
    if dtype.is_float():
        return LogicalType.FLOAT
    if dtype.is_temporal():
        return LogicalType.DATETIME
    return LogicalType.UNKNOWN


class PolarsSource:
    """Present a polars DataFrame as a :class:`~flowxl.integrations.protocol.TabularSource`."""

    def __init__(self, frame: object) -> None:
        """Wrap a polars DataFrame.

        Args:
            frame: Must be a ``polars.DataFrame``.

        Raises:
            TypeError: If ``frame`` is not a DataFrame.
        """
        if not isinstance(frame, pl.DataFrame):
            raise TypeError(f'PolarsSource needs a polars.DataFrame, got {type(frame).__name__!r}.')
        self._frame: pl.DataFrame = frame

    @property
    def columns(self) -> Sequence[str]:
        """Return column names.

        Returns:
            Frame column names.
        """
        return self._frame.columns

    @property
    def logical_types(self) -> Sequence[LogicalType]:
        """Return one logical type per column.

        Returns:
            Logical types in :attr:`columns` order.
        """
        return [to_logical_type(dtype) for dtype in self._frame.dtypes]

    @property
    def row_count(self) -> int | None:
        """Return the number of rows.

        Returns:
            Frame height.
        """
        return self._frame.height

    def iter_rows(self) -> Iterator[Sequence[object]]:
        """Iterate rows as tuples.

        Yields:
            One row of values.
        """
        yield from self._frame.iter_rows()

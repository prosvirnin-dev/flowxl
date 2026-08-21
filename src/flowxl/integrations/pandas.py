"""pandas DataFrame adapter. Imported only on demand."""

from __future__ import annotations

from collections.abc import Iterator, Sequence

import pandas as pd
from pandas.api import types as pdt

from flowxl.primitives.types import LogicalType

__all__ = ['PandasSource']


def to_logical_type(dtype: object) -> LogicalType:
    """Classify a pandas dtype through the public ``pandas.api.types`` helpers.

    Comparing against a concrete class breaks across pandas versions. These
    helpers are the supported path.

    Args:
        dtype: A pandas data type.

    Returns:
        The matching logical type, or :attr:`LogicalType.UNKNOWN`.
    """
    if pdt.is_bool_dtype(dtype):
        return LogicalType.BOOLEAN
    if pdt.is_integer_dtype(dtype):
        return LogicalType.INTEGER
    if pdt.is_float_dtype(dtype):
        return LogicalType.FLOAT
    if pdt.is_datetime64_any_dtype(dtype):
        return LogicalType.DATETIME
    if pdt.is_timedelta64_dtype(dtype):
        return LogicalType.DURATION
    if pdt.is_string_dtype(dtype):
        return LogicalType.TEXT
    return LogicalType.UNKNOWN


class PandasSource:
    """Present a pandas DataFrame as a :class:`~flowxl.integrations.protocol.TabularSource`."""

    def __init__(self, frame: object) -> None:
        """Wrap a pandas DataFrame.

        Args:
            frame: Must be a ``pandas.DataFrame``.

        Raises:
            TypeError: If ``frame`` is not a DataFrame.
        """
        if not isinstance(frame, pd.DataFrame):
            raise TypeError(f'PandasSource needs a pandas.DataFrame, got {type(frame).__name__!r}.')
        self._frame: pd.DataFrame = frame

    @property
    def columns(self) -> Sequence[str]:
        """Return column names as strings.

        Returns:
            Column names. Non-string names go through ``str``.
        """
        return [str(name) for name in self._frame.columns]

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
            Index length.
        """
        return len(self._frame.index)

    def iter_rows(self) -> Iterator[Sequence[object]]:
        """Iterate rows as tuples without the index.

        Yields:
            One row of values.
        """
        yield from self._frame.itertuples(index=False, name=None)

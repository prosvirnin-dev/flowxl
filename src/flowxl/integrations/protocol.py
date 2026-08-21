"""The only contract the API layer knows about tabular data."""

from __future__ import annotations

from collections.abc import Iterator, Sequence
from typing import Protocol, runtime_checkable

from flowxl.primitives.types import LogicalType

__all__ = ['TabularSource']


@runtime_checkable
class TabularSource(Protocol):
    """A rectangular typed collection of rows.

    runtime_checkable is required so ``isinstance(obj, TabularSource)`` works.
    Without it, isinstance against a Protocol is forbidden.
    """

    @property
    def columns(self) -> Sequence[str]:
        """Return column names left to right.

        Returns:
            Column names.
        """
        ...

    @property
    def logical_types(self) -> Sequence[LogicalType]:
        """Return one logical type per column, in the same order as :attr:`columns`.

        Returns:
            Logical types.
        """
        ...

    @property
    def row_count(self) -> int | None:
        """Return the number of rows, or None if the source is lazy.

        Returns:
            Row count, or ``None``.
        """
        ...

    def iter_rows(self) -> Iterator[Sequence[object]]:
        """Iterate rows as sequences in :attr:`columns` order.

        Yields:
            One row of values.
        """
        ...

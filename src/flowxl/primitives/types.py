"""Logical value types shared by data sources and themes.

A theme maps :class:`LogicalType` to a style. Adapters translate polars or
pandas types into :class:`LogicalType`. Neither side knows about the other,
so the style layer stays free of dataframe libraries.
"""

from __future__ import annotations

import datetime as dt
from decimal import Decimal
from enum import StrEnum

__all__ = ['LogicalType', 'infer_logical_type']


class LogicalType(StrEnum):
    """What a value means, independent of how it is stored.

    Attributes:
        INTEGER: Whole numbers.
        FLOAT: Real numbers, including Decimal.
        TEXT: Strings and empty cells.
        DATE: Calendar dates without a time of day.
        DATETIME: Date and time together.
        TIME: Time of day without a date.
        DURATION: Time spans.
        BOOLEAN: True or False.
        PERCENT: Values that should be shown as percentages.
        CURRENCY: Values that should be shown as money.
        UNKNOWN: Anything the library cannot classify.
    """

    INTEGER = 'integer'
    FLOAT = 'float'
    TEXT = 'text'
    DATE = 'date'
    DATETIME = 'datetime'
    TIME = 'time'
    DURATION = 'duration'
    BOOLEAN = 'boolean'
    PERCENT = 'percent'
    CURRENCY = 'currency'
    UNKNOWN = 'unknown'


def infer_logical_type(value: object) -> LogicalType:
    """Classify a single Python value.

    Two orderings are load-bearing. ``bool`` is a subclass of ``int``, so
    ``True`` would become an integer if ``int`` were tested first.
    ``datetime`` is a subclass of ``date``, so a timestamp would lose its
    clock if ``date`` were tested first.

    Args:
        value: Any scalar written into a cell.

    Returns:
        The matching logical type, or :attr:`LogicalType.UNKNOWN`.
    """
    if value is None:
        # Пустая ячейка не должна внезапно получить числовой формат.
        return LogicalType.TEXT
    if isinstance(value, bool):
        # Эта проверка стоит раньше int: в Python bool это подкласс int.
        return LogicalType.BOOLEAN
    if isinstance(value, int):
        return LogicalType.INTEGER
    if isinstance(value, (float, Decimal)):
        return LogicalType.FLOAT
    if isinstance(value, dt.datetime):
        # Эта проверка стоит раньше date: datetime это подкласс date.
        return LogicalType.DATETIME
    if isinstance(value, dt.date):
        return LogicalType.DATE
    if isinstance(value, dt.time):
        return LogicalType.TIME
    if isinstance(value, dt.timedelta):
        return LogicalType.DURATION
    if isinstance(value, str):
        return LogicalType.TEXT
    return LogicalType.UNKNOWN

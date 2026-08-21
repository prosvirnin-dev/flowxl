"""Logical type tests."""

from datetime import date, datetime

from flowxl import LogicalType
from flowxl.primitives.types import infer_logical_type


def test_bool_is_not_int() -> None:
    """Treat True as boolean. In Python bool is a subclass of int."""
    assert infer_logical_type(True) is LogicalType.BOOLEAN
    assert infer_logical_type(1) is LogicalType.INTEGER


def test_datetime_is_not_date() -> None:
    """Check datetime before date. datetime is a subclass of date."""
    assert infer_logical_type(datetime(2024, 1, 1)) is LogicalType.DATETIME
    assert infer_logical_type(date(2024, 1, 1)) is LogicalType.DATE

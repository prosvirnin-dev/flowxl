"""Workbook-wide configuration."""

from __future__ import annotations

from dataclasses import dataclass, field

from flowxl.primitives.sheet_name import OnInvalid
from flowxl.styles.theme import DefaultTheme, Theme

__all__ = ['WorkbookSettings']


@dataclass(frozen=True)
class WorkbookSettings:
    """Everything that applies to a whole workbook.

    A frozen settings object replaces a constructor with many keyword
    arguments. Copy it with :func:`dataclasses.replace`.

    Attributes:
        theme: Rule that maps a logical type to a style.
        on_invalid_name: What to do with a sheet name Excel would reject.
        autofit_max_width: Cap for :meth:`~flowxl.api.sheet.Sheet.autofit`.
        autofit_padding: Extra width beyond the longest seen string.
    """

    theme: Theme = field(default_factory=DefaultTheme)
    on_invalid_name: OnInvalid = OnInvalid.RAISE
    autofit_max_width: float = 60.0
    autofit_padding: float = 2.0

"""Themes decide which style a logical type receives."""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field
from typing import Protocol, Self

from flowxl.primitives.exceptions import UnsupportedTypeError
from flowxl.primitives.types import LogicalType
from flowxl.styles import presets
from flowxl.styles.spec import StyleSpec

__all__ = ['DefaultTheme', 'Theme']


class Theme(Protocol):
    """A theme is anything with these three methods.

    Inheritance is not required. A class that implements the methods is a theme.
    """

    def style_for(self, logical_type: LogicalType) -> StyleSpec:
        """Return the style for a column or scalar of the given logical type.

        Args:
            logical_type: Meaning of the value, independent of storage.

        Returns:
            Style to apply.
        """
        ...

    def header_style(self) -> StyleSpec:
        """Return the style for column headers.

        Returns:
            Header style.
        """
        ...

    def default_style(self) -> StyleSpec:
        """Return the style used when the caller asked for none.

        Returns:
            Neutral style.
        """
        ...


def _default_mapping() -> dict[LogicalType, StyleSpec]:
    """Build the built-in mapping from logical type to preset style.

    Returns:
        A fresh dictionary so each theme instance gets its own mapping.
    """
    return {
        LogicalType.INTEGER: presets.INTEGER,
        LogicalType.FLOAT: presets.FLOAT,
        LogicalType.PERCENT: presets.PERCENT,
        LogicalType.CURRENCY: presets.CURRENCY,
        LogicalType.DATE: presets.DATE,
        LogicalType.DATETIME: presets.DATETIME,
        LogicalType.TIME: presets.TIME,
        LogicalType.DURATION: presets.DURATION,
        LogicalType.BOOLEAN: presets.BOOLEAN,
        LogicalType.TEXT: presets.TEXT,
    }


@dataclass(frozen=True)
class DefaultTheme:
    """The built-in theme.

    An unknown logical type raises unless ``fallback`` is set. A silently
    mis-formatted column is worse than a loud failure.

    Attributes:
        mapping: Logical type to style.
        header: Style for column headers.
        fallback: Style for unmapped types. ``None`` means raise instead.
        default: Style used when the caller asked for none.
    """

    mapping: Mapping[LogicalType, StyleSpec] = field(default_factory=_default_mapping)
    header: StyleSpec = presets.HEADER
    fallback: StyleSpec | None = None
    default: StyleSpec = presets.DEFAULT

    def style_for(self, logical_type: LogicalType) -> StyleSpec:
        """Return the style for a logical type.

        Args:
            logical_type: Meaning of the value.

        Returns:
            Style from the mapping, or the fallback.

        Raises:
            UnsupportedTypeError: If the type is unmapped and there is no fallback.
        """
        style: StyleSpec | None = self.mapping.get(logical_type)
        if style is not None:
            return style
        if self.fallback is not None:
            return self.fallback
        raise UnsupportedTypeError(
            f'No style defined for logical type {logical_type.value!r}. '
            f'Extend the theme mapping or set a fallback style.'
        )

    def header_style(self) -> StyleSpec:
        """Return the header style.

        Returns:
            Header style.
        """
        return self.header

    def default_style(self) -> StyleSpec:
        """Return the default style.

        Returns:
            Neutral style.
        """
        return self.default

    def with_overrides(self, **by_type: StyleSpec) -> Self:
        """Return a copy with some logical types remapped.

        Args:
            **by_type: Keyword per :class:`LogicalType` value, e.g.
                ``float=my_style``.

        Returns:
            A new theme. The original is not modified.
        """
        merged: dict[LogicalType, StyleSpec] = dict(self.mapping)
        for key, style in by_type.items():
            merged[LogicalType(key)] = style
        return type(self)(
            mapping=merged,
            header=self.header,
            fallback=self.fallback,
            default=self.default,
        )

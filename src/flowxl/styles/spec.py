"""Immutable style descriptions. Engine-agnostic: no openpyxl here.

Every spec is frozen and hashable because :class:`~flowxl.backend.registry.StyleRegistry`
uses them as cache keys. A style is data until a workbook exists.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, replace
from enum import StrEnum
from typing import Any, Final, Self

from flowxl.primitives.exceptions import InvalidColorError

__all__ = [
    'BLACK',
    'GREY',
    'WHITE',
    'AlignmentSpec',
    'BorderSpec',
    'BorderStyle',
    'Color',
    'FillPattern',
    'FillSpec',
    'FontSpec',
    'HAlign',
    'ProtectionSpec',
    'SideSpec',
    'StyleSpec',
    'VAlign',
]

_HEX_RE: Final[re.Pattern[str]] = re.compile(r'^[0-9A-F]+$')


class FillPattern(StrEnum):
    """Closed vocabulary of cell fill patterns.

    Attributes:
        NONE: No fill.
        SOLID: Solid fill.
        GRAY125: Excel light-gray pattern.
        LIGHT_GRAY: Named light-gray pattern.
    """

    NONE = 'none'
    SOLID = 'solid'
    GRAY125 = 'gray125'
    LIGHT_GRAY = 'lightGray'


class BorderStyle(StrEnum):
    """Closed vocabulary of one border-edge stroke."""

    NONE = 'none'
    HAIR = 'hair'
    THIN = 'thin'
    MEDIUM = 'medium'
    THICK = 'thick'
    DASHED = 'dashed'
    DOTTED = 'dotted'
    DOUBLE = 'double'


class HAlign(StrEnum):
    """Horizontal placement of text inside a cell."""

    GENERAL = 'general'
    LEFT = 'left'
    CENTER = 'center'
    RIGHT = 'right'
    JUSTIFY = 'justify'


class VAlign(StrEnum):
    """Vertical placement of text inside a cell."""

    TOP = 'top'
    CENTER = 'center'
    BOTTOM = 'bottom'


@dataclass(frozen=True)
class Color:
    """An ARGB colour.

    The stored value is always eight uppercase hex digits. Human forms such as ``#f00`` go through :meth:`parse`.

    Attributes:
        argb: Eight uppercase hex digits in AARRGGBB order.
    """

    argb: str

    def __post_init__(self) -> None:
        """Reject anything that is not eight uppercase hex digits.

        The constructor does not guess formats. That is :meth:`parse`, so a
        :class:`Color` in memory is always canonical.
        """
        if len(self.argb) != 8 or not _HEX_RE.match(self.argb):
            raise InvalidColorError(
                f'Expected eight uppercase hex digits (AARRGGBB), got {self.argb!r}. '
                f'Use Color.parse() for other formats.'
            )

    @classmethod
    def parse(cls, raw: str) -> Self:
        """Build a colour from ``#RGB``, ``#RRGGBB`` or ``AARRGGBB``.

        Args:
            raw: Colour literal, with or without a leading ``#``.

        Returns:
            A colour whose ``argb`` field is eight uppercase hex digits.

        Raises:
            InvalidColorError: If the literal is not a recognised shape.
        """
        text: str = raw.strip().removeprefix('#').upper()
        if len(text) == 3:
            # Цвет #f00 разворачивается в FF0000: каждую цифру дублируем.
            text = ''.join(char * 2 for char in text)
        if len(text) == 6:
            # Шесть цифр это RGB без альфы. Непрозрачность добавляем сами.
            text = f'FF{text}'
        return cls(text)

    @classmethod
    def rgb(cls, red: int, green: int, blue: int, alpha: int = 255) -> Self:
        """Build a colour from channel values in ``0..255``.

        Args:
            red: Red channel.
            green: Green channel.
            blue: Blue channel.
            alpha: Opacity. 255 is fully opaque.

        Returns:
            A colour with those channels.

        Raises:
            InvalidColorError: If any channel is outside ``0..255``.
        """
        channels: tuple[int, int, int, int] = (alpha, red, green, blue)
        if any(not 0 <= channel <= 255 for channel in channels):
            raise InvalidColorError(f'Color channels must be in 0..255, got {channels}.')
        return cls(''.join(f'{channel:02X}' for channel in channels))

    def __str__(self) -> str:
        """Return a CSS-like hex string without the alpha channel.

        Returns:
            A string such as ``#007CD6``.
        """
        return f'#{self.argb[2:]}'


# Готовые константы безопасны как значения по умолчанию, потому что Color - неизменяемый объект.
# Изменяемый объект на этом месте в Python делили бы все вызовы.
BLACK: Final[Color] = Color('FF000000')
WHITE: Final[Color] = Color('FFFFFFFF')
GREY: Final[Color] = Color('FF808080')


@dataclass(frozen=True)
class FontSpec:
    """Typeface settings for a cell.

    Attributes:
        name: Font family.
        size: Size in points.
        bold: Whether the glyphs are bold.
        italic: Whether the glyphs are italic.
        underline: Whether the text is underlined.
        color: Glyph colour.
    """

    name: str = 'Aptos Narrow'
    size: float = 12
    bold: bool = False
    italic: bool = False
    underline: bool = False
    color: Color = BLACK


@dataclass(frozen=True)
class FillSpec:
    """Cell background.

    Attributes:
        pattern: Fill pattern.
        start: Start colour of the fill, if any.
        end: End colour of the fill, if any.
    """

    pattern: FillPattern = FillPattern.NONE
    start: Color | None = None
    end: Color | None = None

    @classmethod
    def solid(cls, color: Color) -> Self:
        """Build the solid fill that almost every report actually uses.

        Args:
            color: Background colour.

        Returns:
            A fill with that colour at both ends.
        """
        return cls(pattern=FillPattern.SOLID, start=color, end=color)


@dataclass(frozen=True)
class SideSpec:
    """One edge of a cell border.

    Attributes:
        style: Stroke kind.
        color: Stroke colour.
    """

    style: BorderStyle = BorderStyle.NONE
    color: Color = BLACK


@dataclass(frozen=True)
class BorderSpec:
    """All four edges of a cell.

    Attributes:
        left: Left edge.
        right: Right edge.
        top: Top edge.
        bottom: Bottom edge.
    """

    left: SideSpec = SideSpec()
    right: SideSpec = SideSpec()
    top: SideSpec = SideSpec()
    bottom: SideSpec = SideSpec()

    @classmethod
    def uniform(cls, style: BorderStyle, color: Color = BLACK) -> Self:
        """Use the same edge on all four sides.

        Args:
            style: Stroke kind for every side.
            color: Stroke colour for every side.

        Returns:
            A border made of four identical sides.
        """
        side: SideSpec = SideSpec(style=style, color=color)
        return cls(left=side, right=side, top=side, bottom=side)


@dataclass(frozen=True)
class AlignmentSpec:
    """Text placement inside the cell.

    Attributes:
        horizontal: Horizontal alignment.
        vertical: Vertical alignment.
        wrap_text: Whether long text wraps inside the cell.
        indent: Indentation level.
    """

    horizontal: HAlign = HAlign.GENERAL
    vertical: VAlign = VAlign.CENTER
    wrap_text: bool = False
    indent: int = 0


@dataclass(frozen=True)
class ProtectionSpec:
    """Cell locking. Effective only when the sheet is protected.

    Attributes:
        locked: Whether the cell is locked.
        hidden: Whether the formula is hidden.
    """

    locked: bool = True
    hidden: bool = False


@dataclass(frozen=True)
class StyleSpec:
    """A complete named cell style.

    The name matters. xlsx stores a named style once and cells keep a
    reference, so a document with named styles is much smaller than one
    with inline formatting on every cell.

    Attributes:
        name: Unique style name inside one workbook.
        font: Typeface settings.
        fill: Cell background.
        border: Four edges.
        alignment: Text placement.
        protection: Cell locking.
        number_format: Excel number-format string. Overlaid after a
            built-in style when it is not ``General``.
        builtin: Excel built-in style name such as ``Percent`` or ``Normal``.
            When set, the workbook does not gain a new named style; ``name``
            is only our cache key.
    """

    name: str
    font: FontSpec = FontSpec()
    fill: FillSpec = FillSpec()
    border: BorderSpec = BorderSpec()
    alignment: AlignmentSpec = AlignmentSpec()
    protection: ProtectionSpec = ProtectionSpec()
    number_format: str = 'General'
    builtin: str | None = None

    def __post_init__(self) -> None:
        """Require a non-empty name.

        Excel identifies a named style by this string. An empty name would collide with any other empty name.
        """
        if not self.name:
            raise ValueError('A style specification requires a non-empty name.')
        if self.builtin is not None and not self.builtin:
            raise ValueError('A built-in Excel style name must be non-empty when provided.')

    def with_(self, name: str, **overrides: Any) -> StyleSpec:
        """Return a modified copy under a new name.

        The name is a required positional argument on purpose. Forgetting it
        would create two different styles with the same name, and half the
        sheet would pick up the wrong formatting.

        Args:
            name: Name of the derived style.
            **overrides: Fields to replace.

        Returns:
            A new specification.
        """
        return replace(self, name=name, **overrides)

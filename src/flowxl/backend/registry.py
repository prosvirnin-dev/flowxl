"""Style registration inside a single workbook.

Named styles are registered lazily and exactly once. Primitive formats
reuse Excel's built-in styles (``Percent``, ``Comma``, ``Normal``) and
never call ``add_named_style`` for a new name. Custom drawings such as
the header still become named styles. The registry returns an
:class:`AppliedStyle`: that is enough for ``cell.style = ...`` and an
optional number-format overlay. The openpyxl object is not returned, so
this module never reaches into a private API.
"""

from __future__ import annotations

from dataclasses import dataclass

from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Protection, Side
from openpyxl.styles.builtins import styles as EXCEL_BUILTINS
from openpyxl.workbook.workbook import Workbook as OpxWorkbook

from flowxl.primitives.exceptions import StyleConflictError, UsageError
from flowxl.styles.spec import (
    BorderStyle,
    FillPattern,
    HAlign,
    SideSpec,
    StyleSpec,
)

__all__ = ['AppliedStyle', 'StyleRegistry', 'build_named_style']


@dataclass(frozen=True)
class AppliedStyle:
    """How to paint one cell after resolving a :class:`StyleSpec`.

    Attributes:
        named_style: Name to assign to ``cell.style``. A built-in or one we
            registered.
        number_format: Overlay after the named style, or ``None`` to keep
            the named style's format.
    """

    named_style: str
    number_format: str | None = None


def _side(spec: SideSpec) -> Side:
    """Translate one border side into an openpyxl Side.

    Args:
        spec: Side description in our language.

    Returns:
        An openpyxl side. ``NONE`` becomes a side with no line.
    """
    if spec.style is BorderStyle.NONE:
        return Side(border_style=None)
    return Side(border_style=spec.style.value, color=spec.color.argb)


def build_named_style(spec: StyleSpec) -> NamedStyle:
    """Translate our drawing into an openpyxl named style.

    Args:
        spec: Style description in our language.

    Returns:
        An openpyxl named style, ready to register on a workbook.
    """
    style: NamedStyle = NamedStyle(name=spec.name)
    style.font = Font(
        name=spec.font.name,
        size=spec.font.size,
        bold=spec.font.bold,
        italic=spec.font.italic,
        underline='single' if spec.font.underline else None,
        color=spec.font.color.argb,
    )
    style.fill = PatternFill(
        fill_type=(None if spec.fill.pattern is FillPattern.NONE else spec.fill.pattern.value),
        start_color=spec.fill.start.argb if spec.fill.start else None,
        end_color=spec.fill.end.argb if spec.fill.end else None,
    )
    style.border = Border(
        left=_side(spec.border.left),
        right=_side(spec.border.right),
        top=_side(spec.border.top),
        bottom=_side(spec.border.bottom),
    )
    style.alignment = Alignment(
        horizontal=(None if spec.alignment.horizontal is HAlign.GENERAL else spec.alignment.horizontal.value),
        vertical=spec.alignment.vertical.value,
        wrap_text=spec.alignment.wrap_text,
        indent=spec.alignment.indent,
    )
    style.protection = Protection(locked=spec.protection.locked, hidden=spec.protection.hidden)
    style.number_format = spec.number_format
    return style


class StyleRegistry:
    """Turn drawings into named styles of one workbook.

    One drawing is registered at most once. Two different drawings cannot
    share a name. Each workbook gets its own registry on purpose:
    ``add_named_style`` mutates the object it is given, so a process-wide
    shared style was the race this library exists to prevent.

    A specification with :attr:`~flowxl.styles.spec.StyleSpec.builtin` set
    does not create a custom named style. Excel already has that style.
    """

    def __init__(self, workbook: OpxWorkbook) -> None:
        """Bind the registry to one openpyxl workbook.

        Args:
            workbook: Workbook that will own the named styles.
        """
        self._workbook: OpxWorkbook = workbook
        # Ключ это весь чертёж, поэтому StyleSpec заморожен и хешируем.
        self._registered: dict[StyleSpec, AppliedStyle] = {}
        self._by_name: dict[str, StyleSpec] = {}

    def resolve(self, spec: StyleSpec) -> AppliedStyle:
        """Make sure the style can be applied and return how to paint the cell.

        Args:
            spec: Style drawing.

        Returns:
            Named style plus optional number-format overlay.

        Raises:
            StyleConflictError: If a different specification already owns ``spec.name``.
            UsageError: If ``spec.builtin`` is not an Excel built-in style.
        """
        cached: AppliedStyle | None = self._registered.get(spec)
        if cached is not None:
            return cached

        owner: StyleSpec | None = self._by_name.get(spec.name)
        if owner is not None and owner != spec:
            raise StyleConflictError(
                f'Style name {spec.name!r} is already used by a different specification. '
                f'Derive a new name with StyleSpec.with_().'
            )

        applied: AppliedStyle = self._apply(spec)
        self._registered[spec] = applied
        self._by_name[spec.name] = spec
        return applied

    def _apply(self, spec: StyleSpec) -> AppliedStyle:
        """Register a custom style or point at an Excel built-in.

        Args:
            spec: Style drawing.

        Returns:
            How to paint the cell.
        """
        if spec.builtin is not None:
            if spec.builtin not in EXCEL_BUILTINS:
                raise UsageError(
                    f'Unknown Excel built-in style {spec.builtin!r}. '
                    f'Use a name such as "Normal", "Percent", "Comma" or "Currency".'
                )
            overlay: str | None = spec.number_format if spec.number_format != 'General' else None
            return AppliedStyle(named_style=spec.builtin, number_format=overlay)

        # Имя уже может быть в книге, если файл открыли, а не создали с нуля.
        if spec.name not in self._workbook.style_names:
            self._workbook.add_named_style(build_named_style(spec))
        return AppliedStyle(named_style=spec.name)

"""Ready-made styles.

Primitive number formats reuse Excel's built-in named styles (Normal,
Comma, Percent, Currency) so the workbook is not filled with custom
``sw_`` styles that only change how a date or a number looks. ``HEADER``
is still ours: the accent fill is not a built-in.
"""

from __future__ import annotations

from typing import Final

from flowxl.styles.spec import (
    WHITE,
    AlignmentSpec,
    Color,
    FillSpec,
    FontSpec,
    HAlign,
    StyleSpec,
    VAlign,
)

__all__ = [
    'BOOLEAN',
    'CURRENCY',
    'DATE',
    'DATETIME',
    'DEFAULT',
    'DURATION',
    'FLOAT',
    'HEADER',
    'INTEGER',
    'PERCENT',
    'TEXT',
    'TIME',
]

ACCENT: Final[Color] = Color.parse('#007CD6')
_LEFT: Final[AlignmentSpec] = AlignmentSpec(horizontal=HAlign.LEFT, vertical=VAlign.CENTER)

# Нейтральное оформление. Ниже публичного API дырок None нет.
# Если человек ничего не указал, подставляем DEFAULT, а не «отсутствие стиля».
DEFAULT: Final[StyleSpec] = StyleSpec(name='sw_default', builtin='Normal')

HEADER: Final[StyleSpec] = StyleSpec(
    name='sw_header',
    font=FontSpec(color=WHITE),
    fill=FillSpec.solid(ACCENT),
    alignment=_LEFT,
)

TEXT: Final[StyleSpec] = StyleSpec(name='sw_text', builtin='Normal')
BOOLEAN: Final[StyleSpec] = StyleSpec(name='sw_boolean', builtin='Normal')
INTEGER: Final[StyleSpec] = StyleSpec(name='sw_integer', builtin='Comma [0]')
FLOAT: Final[StyleSpec] = StyleSpec(name='sw_float', builtin='Comma')
PERCENT: Final[StyleSpec] = StyleSpec(name='sw_percent', builtin='Percent')
CURRENCY: Final[StyleSpec] = StyleSpec(name='sw_currency', builtin='Currency')
DATE: Final[StyleSpec] = StyleSpec(
    name='sw_date',
    builtin='Normal',
    number_format='mm-dd-yy',
)
DATETIME: Final[StyleSpec] = StyleSpec(
    name='sw_datetime',
    builtin='Normal',
    number_format='m/d/yy h:mm',
)
TIME: Final[StyleSpec] = StyleSpec(
    name='sw_time',
    builtin='Normal',
    number_format='h:mm:ss',
)
DURATION: Final[StyleSpec] = StyleSpec(
    name='sw_duration',
    builtin='Normal',
    number_format='[h]:mm:ss',
)

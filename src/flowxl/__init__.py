"""FlowXL: predictable Excel writing over openpyxl.

Everything exported here is public API. Anything else may change in a patch
without a major version bump.
"""

from __future__ import annotations

from typing import Final

from flowxl.api.settings import WorkbookSettings
from flowxl.api.sheet import KEEP, Sheet
from flowxl.api.streaming import StreamingSheet, StreamingWorkbook
from flowxl.api.workbook import Workbook
from flowxl.primitives.cell import ORIGIN, CellRange, CellRef
from flowxl.primitives.exceptions import (
    BackendError,
    CorruptedWorkbookError,
    FlowxlError,
    InvalidColorError,
    InvalidCoordinateError,
    InvalidRangeError,
    InvalidSheetNameError,
    StyleConflictError,
    UnsupportedTypeError,
    UntrackedBoundsError,
    UsageError,
    WorkbookClosedError,
    WorkbookOpenError,
    WorkbookSaveError,
)
from flowxl.primitives.formula import to_excel_formula
from flowxl.primitives.sheet_name import OnInvalid, SheetName
from flowxl.primitives.types import LogicalType
from flowxl.styles.spec import (
    BLACK,
    GREY,
    WHITE,
    AlignmentSpec,
    BorderSpec,
    BorderStyle,
    Color,
    FillPattern,
    FillSpec,
    FontSpec,
    HAlign,
    ProtectionSpec,
    SideSpec,
    StyleSpec,
    VAlign,
)
from flowxl.styles.theme import DefaultTheme, Theme

__version__: Final[str] = '0.1.0'

__all__ = [
    'ORIGIN',
    'BLACK',
    'GREY',
    'KEEP',
    'WHITE',
    'AlignmentSpec',
    'BackendError',
    'BorderSpec',
    'BorderStyle',
    'CellRange',
    'CellRef',
    'Color',
    'CorruptedWorkbookError',
    'DefaultTheme',
    'FillPattern',
    'FillSpec',
    'FlowxlError',
    'FontSpec',
    'HAlign',
    'InvalidColorError',
    'InvalidCoordinateError',
    'InvalidRangeError',
    'InvalidSheetNameError',
    'LogicalType',
    'OnInvalid',
    'ProtectionSpec',
    'Sheet',
    'SheetName',
    'SideSpec',
    'StreamingSheet',
    'StreamingWorkbook',
    'StyleConflictError',
    'StyleSpec',
    'Theme',
    'UnsupportedTypeError',
    'UntrackedBoundsError',
    'UsageError',
    'VAlign',
    'Workbook',
    'WorkbookClosedError',
    'WorkbookOpenError',
    'WorkbookSaveError',
    'WorkbookSettings',
    '__version__',
    'to_excel_formula',
]

"""Exception hierarchy for FlowXL.

Catch :class:`FlowxlError` to catch everything this library raises.

Two branches exist. :class:`UsageError` means the calling code is wrong and
the program must change. :class:`BackendError` means the outside world failed:
a file, the disk, or openpyxl.
"""

from __future__ import annotations

__all__ = [
    'BackendError',
    'CorruptedWorkbookError',
    'FlowxlError',
    'InvalidColorError',
    'InvalidCoordinateError',
    'InvalidRangeError',
    'InvalidSheetNameError',
    'StyleConflictError',
    'UnsupportedTypeError',
    'UntrackedBoundsError',
    'UsageError',
    'WorkbookClosedError',
    'WorkbookOpenError',
    'WorkbookSaveError',
]


class FlowxlError(Exception):
    """Root of every error raised by this library.

    Catch this class if you want one handler for any FlowXL failure.
    """


class UsageError(FlowxlError):
    """The calling code is wrong.

    Fixing this requires changing the program, not retrying the call.
    """


class InvalidCoordinateError(UsageError):
    """A cell coordinate is outside the limits of the xlsx format."""


class InvalidRangeError(UsageError):
    """A range is malformed.

    Typical case: the bottom-right corner is above or left of the top-left.
    """


class InvalidSheetNameError(UsageError):
    """A sheet name violates Excel's rules and no fixing policy was given."""


class InvalidColorError(UsageError):
    """A color literal cannot be parsed into ARGB."""


class UnsupportedTypeError(UsageError):
    """No style is defined for a value or column type."""


class StyleConflictError(UsageError):
    """Two different style specifications claim the same style name."""


class UntrackedBoundsError(UsageError):
    """A bounds-dependent operation was requested but bounds are unknown."""


class WorkbookClosedError(UsageError):
    """The workbook has been closed and can no longer be used."""


class BackendError(FlowxlError):
    """The outside world failed.

    Usually the filesystem, a broken file format, or openpyxl.
    """


class WorkbookOpenError(BackendError):
    """An existing workbook could not be opened."""


class WorkbookSaveError(BackendError):
    """A workbook could not be written to its target."""


class CorruptedWorkbookError(BackendError):
    """The file is not a readable xlsx document."""

"""The public entry point for a regular random-access workbook."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import BinaryIO

from flowxl.api.settings import WorkbookSettings
from flowxl.api.sheet import Sheet
from flowxl.api.streaming import StreamingWorkbook
from flowxl.backend.openpyxl_backend import OpenpyxlStreamingWorkbook, OpenpyxlWorkbook
from flowxl.primitives.exceptions import WorkbookClosedError
from flowxl.primitives.sheet_name import SheetName

__all__ = ['Workbook']


class Workbook:
    """An Excel document.

    Create instances through named constructors, not ``Workbook()``:

    * :meth:`new` creates an empty document.
    * :meth:`open` opens an existing file on disk.
    * :meth:`from_bytes` opens a document from memory.
    * :meth:`stream` returns a write-only :class:`StreamingWorkbook`.

    The context manager does not save. Call :meth:`save` explicitly.
    """

    def __init__(self, backend: OpenpyxlWorkbook, settings: WorkbookSettings) -> None:
        """Wrap a backend and workbook settings.

        Args:
            backend: openpyxl adapter.
            settings: Theme, name policy, autofit parameters.
        """
        self._backend: OpenpyxlWorkbook = backend
        self._settings: WorkbookSettings = settings
        self._sheets: dict[SheetName, Sheet] = {}
        self._closed: bool = False

    @classmethod
    def new(cls, *, settings: WorkbookSettings | None = None) -> Workbook:
        """Create an empty workbook with no sheets.

        Args:
            settings: Workbook configuration. Defaults to the built-in theme
                and a strict name policy.

        Returns:
            A new workbook.
        """
        return cls(OpenpyxlWorkbook.create(), settings or WorkbookSettings())

    @classmethod
    def open(
        cls,
        source: str | Path,
        *,
        formulas: bool = True,
        settings: WorkbookSettings | None = None,
    ) -> Workbook:
        """Open an existing document from disk.

        Args:
            source: Path to an ``.xlsx`` file.
            formulas: ``True`` keeps formulas as written. ``False`` reads the
                cached values.
            settings: Workbook configuration.

        Returns:
            An opened workbook.

        Raises:
            WorkbookOpenError: If the file cannot be read.
            CorruptedWorkbookError: If it is not a valid xlsx document.
        """
        return cls(
            OpenpyxlWorkbook.load(source, formulas=formulas),
            settings or WorkbookSettings(),
        )

    @classmethod
    def from_bytes(
        cls,
        data: bytes | BinaryIO,
        *,
        formulas: bool = True,
        settings: WorkbookSettings | None = None,
    ) -> Workbook:
        """Open a document held in memory.

        Args:
            data: Raw bytes or a binary stream.
            formulas: ``True`` keeps formulas as written.
            settings: Workbook configuration.

        Returns:
            An opened workbook.
        """
        return cls(
            OpenpyxlWorkbook.load(data, formulas=formulas),
            settings or WorkbookSettings(),
        )

    @classmethod
    def stream(cls, *, settings: WorkbookSettings | None = None) -> StreamingWorkbook:
        """Create a write-only workbook for large tables.

        Returns a different type on purpose. A factory method may return
        whatever best serves the caller.

        Args:
            settings: Workbook configuration.

        Returns:
            A streaming workbook, not a :class:`Workbook`.
        """
        return StreamingWorkbook(
            OpenpyxlStreamingWorkbook.create(),
            settings or WorkbookSettings(),
        )

    def sheet(self, name: str | SheetName) -> Sheet:
        """Return an existing sheet or create a new one.

        The same name always returns the same object, so bounds tracking
        survives repeated calls.

        Args:
            name: Sheet name, run through the workbook name policy.

        Returns:
            The sheet facade.
        """
        self._ensure_open()
        key: SheetName = self._coerce(name)
        cached: Sheet | None = self._sheets.get(key)
        if cached is not None:
            return cached
        backend = self._backend.get_sheet(key) if self._backend.has_sheet(key) else self._backend.create_sheet(key)
        sheet: Sheet = Sheet(key, backend, self._settings)
        self._sheets[key] = sheet
        return sheet

    def has_sheet(self, name: str | SheetName) -> bool:
        """Return whether a sheet with this name exists.

        Args:
            name: Sheet name.

        Returns:
            True if the sheet is in the workbook.
        """
        self._ensure_open()
        return self._backend.has_sheet(self._coerce(name))

    def remove(self, name: str | SheetName) -> None:
        """Delete a sheet.

        Args:
            name: Sheet name.

        Raises:
            UsageError: If no such sheet exists.
        """
        self._ensure_open()
        key: SheetName = self._coerce(name)
        self._backend.remove_sheet(key)
        self._sheets.pop(key, None)

    @property
    def sheet_names(self) -> tuple[SheetName, ...]:
        """Return the names of all sheets, in document order.

        Returns:
            Validated sheet names.
        """
        self._ensure_open()
        return self._backend.sheet_names()

    def save(self, target: str | Path) -> None:
        """Write the document to disk.

        May be called more than once. The workbook stays usable afterwards.

        Args:
            target: Destination path.
        """
        self._ensure_open()
        self._backend.save(target)

    def to_bytes(self) -> bytes:
        """Serialise the document to memory.

        Returns:
            The xlsx file as bytes.
        """
        self._ensure_open()
        with BytesIO() as buffer:
            self._backend.save(buffer)
            return buffer.getvalue()

    def close(self) -> None:
        """Release resources. Safe to call more than once. Does not save."""
        if not self._closed:
            self._backend.close()
            self._closed = True

    def __enter__(self) -> Workbook:
        """Enter the context manager without saving.

        Returns:
            This workbook.
        """
        return self

    def __exit__(self, *exc_info: object) -> None:
        """Close the workbook. Does not save."""
        self.close()

    def _coerce(self, name: str | SheetName) -> SheetName:
        """Turn a raw string into a validated sheet name.

        Args:
            name: An already validated name or a raw string.

        Returns:
            A :class:`SheetName`.
        """
        if isinstance(name, SheetName):
            return name
        return SheetName.create(name, on_invalid=self._settings.on_invalid_name)

    def _ensure_open(self) -> None:
        """Reject operations after :meth:`close`.

        Raises:
            WorkbookClosedError: If the workbook is already closed.
        """
        if self._closed:
            raise WorkbookClosedError('This workbook has been closed. Open a new one to continue.')

"""Write-only sheets and workbooks. There is no write_only flag anywhere else."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from io import BytesIO
from pathlib import Path

from flowxl.api.base_sheet import BaseSheet
from flowxl.api.settings import WorkbookSettings
from flowxl.backend.openpyxl_backend import OpenpyxlStreamingSheet, OpenpyxlStreamingWorkbook
from flowxl.integrations.protocol import TabularSource
from flowxl.primitives.cell import ORIGIN, CellRef, letter_to_index
from flowxl.primitives.exceptions import UsageError, WorkbookClosedError
from flowxl.primitives.sheet_name import SheetName
from flowxl.styles.spec import StyleSpec

__all__ = ['StreamingSheet', 'StreamingWorkbook']


class StreamingSheet(BaseSheet):
    """A sheet that only accepts appended rows.

    Methods ``cell()``, ``row()``, ``merge()``, ``freeze()`` and ``read()``
    do not exist here. A wrong call cannot even be written. ``protect`` and
    ``dropdown`` do: they decorate the sheet, they do not write a cell.
    """

    def __init__(
        self,
        name: SheetName,
        backend: OpenpyxlStreamingSheet,
        settings: WorkbookSettings,
    ) -> None:
        """Bind a validated name to a streaming sheet backend.

        Args:
            name: Validated sheet name.
            backend: Write-only openpyxl adapter.
            settings: Workbook settings.
        """
        super().__init__(name, backend, settings)
        self._backend: OpenpyxlStreamingSheet = backend

    def frame(
        self,
        data: object,
        *,
        at: CellRef = ORIGIN,
        header: bool = True,
        styles: Mapping[str, StyleSpec] | None = None,
    ) -> StreamingSheet:
        """Write a table from the top-left corner. Any other anchor is forbidden.

        A streaming sheet can only append rows. ``at=B2`` would be a lie.

        Args:
            data: Tabular source.
            at: Must be :data:`~flowxl.primitives.cell.ORIGIN`.
            header: Whether to write column names first.
            styles: Per-column style overrides.

        Returns:
            This sheet, for chaining.

        Raises:
            UsageError: If ``at`` is not the origin.
        """
        if at != ORIGIN:
            raise UsageError('Streaming sheets can only append from the top-left. Omit at= or pass at=ORIGIN.')
        return super().frame(data, at=at, header=header, styles=styles)

    def _write_header(self, at: CellRef, columns: Sequence[str]) -> None:
        """Append the header as one row.

        Args:
            at: Used only for width tracking. Always the origin.
            columns: Column names.
        """
        header_style: StyleSpec = self._settings.theme.header_style()
        self._backend.append_row(list(columns), [header_style] * len(columns))
        for offset, column in enumerate(columns):
            self._track_width(at.col + offset, column)

    def _write_body(
        self,
        at: CellRef,
        source: TabularSource,
        styles: Sequence[StyleSpec],
    ) -> int:
        """Append data rows.

        Args:
            at: Used only to know the starting column for width tracking.
            source: Tabular data.
            styles: One style per column.

        Returns:
            Number of data rows written.
        """
        written: int = 0
        style_list: list[StyleSpec] = list(styles)
        for values in source.iter_rows():
            self._backend.append_row(list(values), style_list)
            written += 1
            for offset, value in enumerate(values):
                self._track_width(at.col + offset, value)
        return written


class StreamingWorkbook:
    """A workbook written once, top to bottom, without holding it in memory.

    Decoration is set when the sheet is created, not afterwards: openpyxl
    serialises column widths and freeze before the first row.
    """

    def __init__(
        self,
        backend: OpenpyxlStreamingWorkbook,
        settings: WorkbookSettings,
    ) -> None:
        """Wrap a streaming backend.

        Args:
            backend: Write-only openpyxl adapter.
            settings: Workbook settings.
        """
        self._backend: OpenpyxlStreamingWorkbook = backend
        self._settings: WorkbookSettings = settings
        self._sheets: dict[SheetName, StreamingSheet] = {}
        self._closed: bool = False

    def sheet(
        self,
        name: str | SheetName,
        *,
        pin_header: bool = False,
        widths: Mapping[int | str, float] | None = None,
    ) -> StreamingSheet:
        """Create a streaming sheet.

        Args:
            name: Sheet name.
            pin_header: Freeze the first row.
            widths: Column widths, keyed by 1-based index or letter.

        Returns:
            A new streaming sheet.

        Raises:
            UsageError: If the sheet already exists. A streaming sheet cannot
                be reopened.
            WorkbookClosedError: If the workbook was already saved.
        """
        self._ensure_open()
        key: SheetName = self._coerce(name)
        if key in self._sheets:
            raise UsageError(
                f'Streaming sheet {key.value!r} already exists and cannot be reopened. '
                f'Keep the object returned by the first call.'
            )

        backend: OpenpyxlStreamingSheet = self._backend.create_sheet(key)
        # Openpyxl в write-only сериализует freeze и ширины до первой строки.
        # Поэтому украшения принимаем здесь, а не «потом».
        if pin_header:
            backend.set_freeze(CellRef(row=2, col=1))
        for column, width in (widths or {}).items():
            index: int = column if isinstance(column, int) else letter_to_index(column)
            backend.set_column_width(index, width)

        sheet: StreamingSheet = StreamingSheet(key, backend, self._settings)
        self._sheets[key] = sheet
        return sheet

    def save(self, target: str | Path) -> None:
        """Write the document and close the workbook.

        Unlike :meth:`~flowxl.api.workbook.Workbook.save`, this is a one-shot
        operation.

        Args:
            target: Destination path.
        """
        self._ensure_open()
        self._backend.save(target)
        self.close()

    def to_bytes(self) -> bytes:
        """Serialise to memory and close the workbook.

        Returns:
            The xlsx file as bytes.
        """
        self._ensure_open()
        with BytesIO() as buffer:
            self._backend.save(buffer)
            payload: bytes = buffer.getvalue()
        self.close()
        return payload

    def close(self) -> None:
        """Release resources. Safe to call more than once."""
        if not self._closed:
            self._backend.close()
            self._closed = True

    def __enter__(self) -> StreamingWorkbook:
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
        """Reject operations after save or close.

        Raises:
            WorkbookClosedError: If the workbook is already closed.
        """
        if self._closed:
            raise WorkbookClosedError(
                'This streaming workbook was already saved or closed. Create a new one with Workbook.stream().'
            )

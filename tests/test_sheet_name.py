"""Sheet name tests."""

import pytest

from flowxl import InvalidSheetNameError, OnInvalid, SheetName


def test_valid_name() -> None:
    """Accept a name that Excel already likes."""
    assert SheetName.create('Report').value == 'Report'


def test_strict_by_default() -> None:
    """Reject a bad name until the caller asks for a fix."""
    with pytest.raises(InvalidSheetNameError):
        SheetName.create('a/b')


def test_fix_is_opt_in() -> None:
    """Fix a name only under the FIX policy."""
    assert SheetName.create('a/b', on_invalid=OnInvalid.FIX).value == 'a_b'

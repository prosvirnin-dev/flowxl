"""Excel worksheet names and the policy for invalid ones.

Fixing is opt-in. By default the library raises, so a bad name never reaches Excel silently.
"""

from __future__ import annotations

import re
import warnings
from dataclasses import dataclass
from enum import StrEnum
from typing import ClassVar, Final, Self

from flowxl.primitives.exceptions import InvalidSheetNameError

__all__ = ['OnInvalid', 'SheetName']

# Символы, которые Excel запрещает в имени листа: слэши, звездочка, вопрос, двоеточие и квадратные скобки.
_FORBIDDEN_RE: Final[re.Pattern[str]] = re.compile(r'[\\/*?:\[\]]')


class OnInvalid(StrEnum):
    """What to do when a raw sheet name breaks Excel's rules.

    Attributes:
        RAISE: Reject the name. This is the default.
        FIX: Repair the name silently.
        WARN: Repair the name and emit a warning.
    """

    RAISE = 'raise'
    FIX = 'fix'
    WARN = 'warn'


@dataclass(frozen=True)
class SheetName:
    """A worksheet name that Excel is guaranteed to accept.

    An invalid instance cannot be constructed. The rest of the code may pass
    a :class:`SheetName` to Excel without checking again.

    Attributes:
        value: The validated name string.
    """

    value: str

    MAX_LENGTH: ClassVar[int] = 31
    FORBIDDEN: ClassVar[frozenset[str]] = frozenset[str]('\\/*?:[]')
    RESERVED: ClassVar[frozenset[str]] = frozenset[str]({'history'})
    FALLBACK: ClassVar[str] = 'Sheet'

    def __post_init__(self) -> None:
        """Reject a name that Excel would not accept.

        The check lives here so :meth:`create` does not duplicate the message.
        ``SheetName(raw)`` and ``SheetName.create(raw)`` then fail the same way.
        """
        found: list[str] = self.problems(self.value)
        if found:
            raise InvalidSheetNameError(
                f'Invalid sheet name {self.value!r}: '
                + '; '.join(found)
                + '. Pass on_invalid="fix" to repair it automatically.'
            )

    @classmethod
    def create(cls, raw: str, *, on_invalid: OnInvalid = OnInvalid.RAISE) -> Self:
        """Build a sheet name, applying the given policy to invalid input.

        Args:
            raw: Desired name.
            on_invalid: What to do if Excel would reject the name.

        Returns:
            A valid sheet name.

        Raises:
            InvalidSheetNameError: If the name is invalid and the policy is :attr:`OnInvalid.RAISE`.
        """
        found: list[str] = cls.problems(raw)
        if not found:
            return cls(raw)
        if on_invalid is OnInvalid.RAISE:
            # Не бросаем ошибку здесь. Конструктор вызовет __post_init__, и сообщение об ошибке останется в одном месте.
            return cls(raw)
        fixed: str = cls._sanitize(raw)
        if on_invalid is OnInvalid.WARN:
            warnings.warn(
                f'Sheet name {raw!r} was rewritten to {fixed!r}: ' + '; '.join(found),
                UserWarning,
                stacklevel=3,
            )
        return cls(fixed)

    @staticmethod
    def problems(raw: str) -> list[str]:
        """List every reason ``raw`` is unacceptable.

        Collecting every reason at once makes the message useful: the caller
        sees both "too long" and "contains :" in one go.

        Args:
            raw: Candidate name, not yet validated.

        Returns:
            Human-readable problems. An empty list means the name is valid.
        """
        found: list[str] = []
        if not raw:
            found.append('name is empty')
            return found
        if len(raw) > SheetName.MAX_LENGTH:
            found.append(f'length is {len(raw)}, Excel allows at most {SheetName.MAX_LENGTH}')
        bad: list[str] = sorted(set(raw) & SheetName.FORBIDDEN)
        if bad:
            found.append('contains forbidden characters ' + ' '.join(bad))
        if raw.startswith("'") or raw.endswith("'"):
            found.append('starts or ends with an apostrophe')
        if raw != raw.strip():
            found.append('has leading or trailing whitespace')
        if raw.lower() in SheetName.RESERVED:
            found.append(f'{raw!r} is reserved by Excel')
        return found

    @staticmethod
    def _sanitize(raw: str) -> str:
        """Repair a name so that Excel will accept it.

        Reachable only through an explicit policy. The default path never calls this.

        Args:
            raw: Invalid name.

        Returns:
            A name that passes :meth:`problems`.
        """
        cleaned: str = _FORBIDDEN_RE.sub('_', raw).strip().strip("'").strip()
        if cleaned.lower() in SheetName.RESERVED:
            cleaned = f'{cleaned}_'
        if len(cleaned) > SheetName.MAX_LENGTH:
            cleaned = cleaned[: SheetName.MAX_LENGTH - 3] + '...'
        return cleaned or SheetName.FALLBACK

    def __str__(self) -> str:
        """Return the validated name.

        Returns:
            The sheet name string.
        """
        return self.value

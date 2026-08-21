"""Normalize formulas to the en-US form that xlsx actually stores.

Excel files never store ``СУММ`` or a comma as the decimal mark. A Russian
UI shows those; the file still holds ``SUM`` and ``1.5``. This module is
that translation. It does not evaluate anything.
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import Final

from flowxl.primitives.exceptions import UsageError
from flowxl.primitives.formula_names import BARE_CONSTANTS, RU_TO_EN, fold_name

__all__ = ['to_excel_formula']

_CYRILLIC_START: Final[int] = 0x0400
_CYRILLIC_END: Final[int] = 0x04FF

_STRING: Final[str] = 'string'
_SHEET: Final[str] = 'sheet'
_IDENT: Final[str] = 'ident'
_NUMBER: Final[str] = 'number'
_WS: Final[str] = 'ws'
_PUNCT: Final[str] = 'punct'


def to_excel_formula(expression: str) -> str:
    """Return the en-US formula Excel stores in the file.

    Russian function names become English. ``;`` becomes the argument
    comma. A decimal comma becomes a period. Text in quotes is left alone.
    An English formula is returned unchanged.

    Args:
        expression: Formula text, starting with ``=``.

    Returns:
        A formula Excel can store, still starting with ``=``.

    Raises:
        UsageError: If the text does not start with ``=``, or a Cyrillic
            name looks like a function but is not in the translation table.
    """
    if not expression.startswith('='):
        raise UsageError(
            f'A formula must start with "=", got {expression!r}. Write it as "={expression}" if that was the intent.'
        )
    tokens: list[tuple[str, str]] = _tokenize(expression)
    russian: bool = _uses_russian_locale(tokens)
    return _emit(tokens, russian=russian)


def _tokenize(expression: str) -> list[tuple[str, str]]:
    """Split a formula into strings, names, numbers and punctuation.

    Args:
        expression: Full formula including the leading ``=``.

    Returns:
        Tokens as ``(kind, text)`` pairs. ``text`` is the original slice.
    """
    tokens: list[tuple[str, str]] = []
    index: int = 0
    length: int = len(expression)
    while index < length:
        char: str = expression[index]
        if char.isspace():
            start = index
            while index < length and expression[index].isspace():
                index += 1
            tokens.append((_WS, expression[start:index]))
            continue
        if char == '"':
            text, index = _read_quoted(expression, index, '"')
            tokens.append((_STRING, text))
            continue
        if char == "'":
            text, index = _read_quoted(expression, index, "'")
            tokens.append((_SHEET, text))
            continue
        if _is_ident_start(char):
            start = index
            index += 1
            while index < length and _is_ident_continue(expression[index]):
                index += 1
            tokens.append((_IDENT, expression[start:index]))
            continue
        if char.isdigit():
            text, index = _read_number(expression, index)
            tokens.append((_NUMBER, text))
            continue
        tokens.append((_PUNCT, char))
        index += 1
    return tokens


def _read_quoted(expression: str, start: int, quote: str) -> tuple[str, int]:
    """Read a quoted run, doubling the quote to escape it as Excel does.

    Args:
        expression: Full formula.
        start: Index of the opening quote.
        quote: ``"`` for a string or ``'`` for a sheet name.

    Returns:
        The quoted slice, including quotes, and the index after it.
    """
    index: int = start + 1
    length: int = len(expression)
    while index < length:
        if expression[index] != quote:
            index += 1
            continue
        if index + 1 < length and expression[index + 1] == quote:
            index += 2
            continue
        index += 1
        return expression[start:index], index
    return expression[start:], length


def _read_number(expression: str, start: int) -> tuple[str, int]:
    """Read an English-style number: digits, optional period, optional exponent.

    A decimal comma is not consumed here. The emitter joins ``1``, ``,``,
    ``5`` only after it knows the formula is Russian.

    Args:
        expression: Full formula.
        start: Index of the first digit.

    Returns:
        The number slice and the index after it.
    """
    index: int = start
    length: int = len(expression)
    while index < length and expression[index].isdigit():
        index += 1
    if index < length and expression[index] == '.':
        index += 1
        while index < length and expression[index].isdigit():
            index += 1
    if index < length and expression[index] in 'eE':
        next_index: int = index + 1
        if next_index < length and expression[next_index] in '+-':
            next_index += 1
        if next_index < length and expression[next_index].isdigit():
            index = next_index
            while index < length and expression[index].isdigit():
                index += 1
    return expression[start:index], index


def _is_ident_start(char: str) -> bool:
    """Return whether ``char`` can start a function name or cell reference."""
    return char.isalpha() or char == '_'


def _is_ident_continue(char: str) -> bool:
    """Return whether ``char`` can continue a function name or cell reference."""
    return char.isalnum() or char in '._'


def _has_cyrillic(text: str) -> bool:
    """Return whether ``text`` contains a Cyrillic letter."""
    return any(_CYRILLIC_START <= ord(char) <= _CYRILLIC_END for char in text)


def _uses_russian_locale(tokens: Sequence[tuple[str, str]]) -> bool:
    """Guess whether separators should be read the Russian way.

    A semicolon outside quotes is decisive: Russian Excel uses it as the
    argument comma. A translated or Cyrillic function name is too.
    ``1,5`` without either of those is a decimal only when no bare comma
    sits between non-numeric tokens the English way.

    Args:
        tokens: Output of :func:`_tokenize`.

    Returns:
        True when ``;`` and decimal commas should be converted.
    """
    saw_semicolon: bool = False
    saw_russian_name: bool = False
    saw_call: bool = False
    for index, (kind, text) in enumerate(tokens):
        if kind == _PUNCT and text == ';':
            saw_semicolon = True
        elif kind == _IDENT and (fold_name(text) in RU_TO_EN or _has_cyrillic(text)):
            saw_russian_name = True
        if kind == _IDENT and _is_call(tokens, index):
            saw_call = True
    if saw_semicolon or saw_russian_name:
        return True
    return _has_unspaced_decimal_comma(tokens) and not saw_call


def _has_unspaced_decimal_comma(tokens: Sequence[tuple[str, str]]) -> bool:
    """Return whether ``1,5`` appears with no space around the comma."""
    for index, (kind, text) in enumerate(tokens):
        if kind != _PUNCT or text != ',':
            continue
        if index == 0 or index + 1 >= len(tokens):
            continue
        prev_kind, _prev = tokens[index - 1]
        next_kind, _nxt = tokens[index + 1]
        if prev_kind == _NUMBER and next_kind == _NUMBER:
            return True
    return False


def _nearest_code(tokens: Sequence[tuple[str, str]], index: int, *, step: int) -> tuple[str, str] | None:
    """Return the nearest non-whitespace token left or right of ``index``.

    Args:
        tokens: Full token list.
        index: Starting position. The token at this index is skipped.
        step: ``-1`` looks left, ``1`` looks right.

    Returns:
        The token, or ``None`` at the edge.
    """
    cursor: int = index + step
    while 0 <= cursor < len(tokens):
        kind, text = tokens[cursor]
        if kind != _WS:
            return kind, text
        cursor += step
    return None


def _emit(tokens: Sequence[tuple[str, str]], *, russian: bool) -> str:
    """Rebuild the formula, translating names and separators when needed.

    Args:
        tokens: Output of :func:`_tokenize`.
        russian: Whether to convert ``;`` and decimal commas.

    Returns:
        The en-US formula text.

    Raises:
        UsageError: If a Cyrillic identifier is used as a function and is
            missing from the translation table.
    """
    parts: list[str] = []
    array_depth: int = 0
    index: int = 0
    while index < len(tokens):
        kind, text = tokens[index]
        if kind == _IDENT:
            parts.append(_translate_ident(text, tokens, index))
            index += 1
            continue
        if kind == _PUNCT:
            if text == '{':
                array_depth += 1
            elif text == '}' and array_depth:
                array_depth -= 1
            if russian and text == ';' and array_depth == 0:
                parts.append(',')
                index += 1
                continue
            if russian and text == '\\' and array_depth > 0:
                parts.append(',')
                index += 1
                continue
            if russian and text == ',' and _is_decimal_comma(tokens, index):
                parts.append('.')
                index += 1
                continue
            parts.append(text)
            index += 1
            continue
        parts.append(text)
        index += 1
    return ''.join(parts)


def _translate_ident(text: str, tokens: Sequence[tuple[str, str]], index: int) -> str:
    """Translate one identifier if it is a Russian function or TRUE/FALSE.

    Args:
        text: Identifier as written.
        tokens: Full token list, to look at the following ``(``.
        index: Position of this identifier.

    Returns:
        English function name, or ``text`` unchanged.

    Raises:
        UsageError: If a Cyrillic name is called as a function but unknown.
    """
    folded: str = fold_name(text)
    english: str | None = RU_TO_EN.get(folded)
    call: bool = _is_call(tokens, index)
    if english is not None and (call or folded in BARE_CONSTANTS):
        return english
    if call and _has_cyrillic(text) and english is None:
        raise UsageError(
            f'Unknown Russian Excel function {text!r}. Write the English name, or check the spelling (yo vs ye).'
        )
    return text


def _is_call(tokens: Sequence[tuple[str, str]], index: int) -> bool:
    """Return whether the token after ``index`` is ``(``, ignoring spaces."""
    nxt = _nearest_code(tokens, index, step=1)
    return nxt is not None and nxt[0] == _PUNCT and nxt[1] == '('


def _is_decimal_comma(tokens: Sequence[tuple[str, str]], index: int) -> bool:
    """Return whether the comma at ``index`` sits between two number tokens.

    No whitespace around the comma: ``1,5`` is a decimal, ``1, 5`` is not.
    """
    if index == 0 or index + 1 >= len(tokens):
        return False
    prev_kind, _prev = tokens[index - 1]
    next_kind, _nxt = tokens[index + 1]
    return prev_kind == _NUMBER and next_kind == _NUMBER

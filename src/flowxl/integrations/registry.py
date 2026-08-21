"""Dispatcher from user objects to TabularSource. Imports nothing up front."""

from __future__ import annotations

from collections.abc import Callable
from typing import Final

from flowxl.integrations.protocol import TabularSource
from flowxl.primitives.exceptions import UnsupportedTypeError

__all__ = ['SourceRegistry', 'adapt', 'registry']

Factory = Callable[[object], TabularSource]


class SourceRegistry:
    """Map a library's root module to an adapter factory.

    The dispatch key is the object's type root module, so polars is not
    imported until a polars object appears.
    """

    def __init__(self) -> None:
        """Create an empty registry."""
        self._factories: dict[str, Factory] = {}

    def register(self, module: str, factory: Factory) -> None:
        """Bind a root module name to a factory. Imports nothing.

        Args:
            module: Root module name, for example ``'polars'`` or ``'builtins'``.
            factory: Callable that wraps data in a :class:`TabularSource`.
        """
        self._factories[module] = factory

    def adapt(self, obj: object) -> TabularSource:
        """Wrap ``obj`` in a :class:`TabularSource`.

        Args:
            obj: A dataframe, a list of dicts, or anything that already
                implements the protocol.

        Returns:
            A tabular source.

        Raises:
            UnsupportedTypeError: If nobody knows how to adapt ``obj``.
        """
        if isinstance(obj, TabularSource):
            return obj
        # Модуль polars.dataframe.frame.DataFrame даёт ключ polars.
        # Так мы не импортируем библиотеку, пока объект этого типа не появился.
        root: str = type(obj).__module__.split('.', 1)[0]
        factory: Factory | None = self._factories.get(root)
        if factory is None:
            supported: str = ', '.join(sorted(self._factories)) or 'nothing'
            raise UnsupportedTypeError(
                f'Cannot write objects of type {type(obj).__name__!r} (from {root!r}). '
                f'Supported sources: {supported}. '
                f'Install the matching extra, e.g. pip install flowxl[polars].'
            )
        return factory(obj)


def _sequence_factory(obj: object) -> TabularSource:
    """Wrap a built-in list or tuple.

    Args:
        obj: A sequence of records.

    Returns:
        A :class:`~flowxl.integrations.sequences.SequenceSource`.
    """
    from flowxl.integrations.sequences import SequenceSource

    return SequenceSource.from_records(obj)


def _polars_factory(obj: object) -> TabularSource:
    """Wrap a polars DataFrame. Imports polars only at call time.

    Args:
        obj: A polars DataFrame.

    Returns:
        A :class:`~flowxl.integrations.polars.PolarsSource`.
    """
    from flowxl.integrations.polars import PolarsSource

    return PolarsSource(obj)


def _pandas_factory(obj: object) -> TabularSource:
    """Wrap a pandas DataFrame. Imports pandas only at call time.

    Args:
        obj: A pandas DataFrame.

    Returns:
        A :class:`~flowxl.integrations.pandas.PandasSource`.
    """
    from flowxl.integrations.pandas import PandasSource

    return PandasSource(obj)


registry: Final[SourceRegistry] = SourceRegistry()
registry.register('builtins', _sequence_factory)
registry.register('polars', _polars_factory)
registry.register('pandas', _pandas_factory)


def adapt(obj: object) -> TabularSource:
    """Module-level shortcut for ``registry.adapt``.

    Args:
        obj: Object to wrap.

    Returns:
        A tabular source.
    """
    return registry.adapt(obj)

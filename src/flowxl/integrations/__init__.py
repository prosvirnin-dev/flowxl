"""Optional adapters that turn foreign tables into one shared source.

The API layer does not import polars or pandas. This package turns their
objects and plain Python sequences into :class:`TabularSource`.
"""

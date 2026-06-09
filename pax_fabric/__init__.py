"""PAX Purview Audit Log Processor — Fabric notebook package.

Phase A0 (CSV-to-Files bridge): a near-verbatim port of the legacy ``pax``
package whose only behavioural change is that all output (CSV bundles,
logs, checkpoints) is written under ``/lakehouse/default/Files/pax/...``
instead of a local ``OutputPath``.

Public surface:
    run(params: dict) -> dict       — execute the full pipeline.
    files_io                         — lakehouse Files/ path resolver.
"""

from __future__ import annotations

from . import files_io  # noqa: F401  (re-exported for notebook callers)


def run(params: dict | None = None) -> dict:
    """Lazy proxy to :func:`pax_fabric.pipeline.run`.

    Importing the pipeline module is deferred so that ``import pax_fabric``
    in a notebook stays cheap and never raises if optional dependencies
    are missing — the error surfaces only when ``run()`` is actually
    invoked.
    """
    from .pipeline import run as _run
    return _run(params or {})


__all__ = ["run", "files_io"]

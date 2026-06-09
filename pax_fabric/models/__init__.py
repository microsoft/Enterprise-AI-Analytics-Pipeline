"""
PAX Models Package
==================
Central dataclass definitions shared across all PAX modules.

- PAXConfig:      All configuration parameters (CLI + computed).
- PAXRunContext:  Runtime state container passed to all functions.
- PAXMetrics:     Run-level counters, timings, and diagnostics.
"""

from .config import PAXConfig
from .context import PAXRunContext
from .metrics import PAXMetrics

__all__ = ["PAXConfig", "PAXRunContext", "PAXMetrics"]

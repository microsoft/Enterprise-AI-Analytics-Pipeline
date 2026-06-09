"""
PAXConfig — Central configuration dataclass.
=============================================
Replaces the PowerShell param() block (72+ parameters) from the original
PAX_Purview_Audit_Log_Processor.ps1 (L1206–2600).

This is a re-export of the canonical PAXConfig already defined in
mod1_pax_config.py. The models/ package provides a clean import path
(`from models import PAXConfig`) for downstream consumers without
requiring knowledge of the module numbering scheme.

PS Source: L1206–1483 (param block) + L1485–2600 (validation)
"""

from __future__ import annotations

# Re-export PAXConfig from its canonical definition in mod1_pax_config.
# This avoids duplicating the dataclass and keeps a single source of truth.
from ..mod1_pax_config import PAXConfig  # noqa: E402, F401

__all__ = ["PAXConfig"]

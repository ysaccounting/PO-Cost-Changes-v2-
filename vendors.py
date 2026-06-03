"""TC-vendors lookup from data/Vendors_Open.xlsx (third tab).

The expense pair's offset Category is decided by the *final* (renamed) vendor:
  - If the vendor appears on the **third tab** ("Consolidated (without
    company)", the single "Account" column) → "<Vendor> (TC)".
  - Otherwise → "Due from Vendors - Open".

In other words, the third tab is the consolidated TC list; everything not on it
defaults to "Due from Vendors - Open". Membership is case-insensitive — both the
lookup set and the input vendor are lowercased when checking.
"""
from __future__ import annotations

import logging
import os
from functools import lru_cache
from pathlib import Path

import pandas as pd

log = logging.getLogger(__name__)

DEFAULT_PATH = Path("data") / "Vendors_Open.xlsx"
# Third tab ("Consolidated (without company)") holds the TC vendor list.
TC_SHEET_INDEX = 2


def load_tc_vendors(path: Path | str | None = None) -> set[str]:
    """Load the set of TC vendor names from the third tab, lowercased for
    case-insensitive matching at lookup time."""
    path = Path(path) if path else Path(os.getenv("OPEN_VENDORS_PATH", DEFAULT_PATH))
    if not path.exists():
        log.warning("Vendors file not found at %s; using empty TC set.", path)
        return set()

    try:
        df = pd.read_excel(path, sheet_name=TC_SHEET_INDEX)
    except (ValueError, IndexError):
        log.warning("Vendors file %s has no third tab; using empty TC set.", path)
        return set()

    if df.shape[1] == 0:
        return set()
    col = df.columns[0]  # single "Account" column
    vendors = {
        str(v).strip().lower()
        for v in df[col].dropna().tolist()
        if str(v).strip()
    }
    log.info("Loaded %d TC vendors from %s (third tab)", len(vendors), path)
    return vendors


@lru_cache(maxsize=1)
def get_tc_vendors() -> set[str]:
    """Cached accessor used by the processor."""
    return load_tc_vendors()


def reset_cache() -> None:
    """Clear the cache (used in tests)."""
    get_tc_vendors.cache_clear()


def offset_category(vendor: str, tc_vendors: set[str] | None = None) -> str:
    """Return the offset Category for an expense pair's second line.

    - If `vendor` is on the TC list (third tab) → '<vendor> (TC)'
    - Otherwise → 'Due from Vendors - Open'
    """
    if tc_vendors is None:
        tc_vendors = get_tc_vendors()
    v = str(vendor).strip()
    if v.lower() in tc_vendors:
        return f"{v} (TC)"
    return "Due from Vendors - Open"

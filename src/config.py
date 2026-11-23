"""Shared settings for BranchPulse. Edit here (or on the Settings sheet in Excel)."""
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
RAW_DIR = ROOT / "data" / "raw"
CLEAN_DIR = ROOT / "data" / "clean"
OUT_DIR = ROOT / "output"

FY_LABEL = "FY 2025-26"
FY_START = "2025-04-01"
FY_END = "2026-03-31"

# Branch Health Score weights (must add up to 1.0)
WEIGHTS = {"achievement": 0.40, "settlement": 0.25, "renewal": 0.20, "tat": 0.15}

# Thresholds used for RAG status
GREEN_MIN = 90      # score >= 90  -> Green
AMBER_MIN = 82      # score >= 82  -> Amber, below -> Red

# Red-flag triggers
FLAG_ACHIEVEMENT_BELOW = 0.85   # premium achievement below 85% of target
FLAG_SETTLEMENT_BELOW = 0.85    # claim settlement ratio below 85%
FLAG_TAT_ABOVE = 10             # average claim turnaround above 10 days
FLAG_RENEWAL_BELOW = 0.75       # renewal rate below 75%

# TAT scoring: full marks up to TAT_GOOD days, then lose TAT_PENALTY points per extra day
TAT_GOOD = 7
TAT_PENALTY = 8

"""Run with:  python -m pytest -q   (or just: python tests/test_kpis.py)"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
from src import kpi


def test_health_score_perfect_branch():
    assert round(kpi.health_score(1.2, 1.0, 1.0, 5), 6) == 100.0


def test_achievement_is_capped_at_100_percent():
    assert kpi.health_score(2.0, 0.9, 0.9, 7) == kpi.health_score(1.0, 0.9, 0.9, 7)


def test_tat_score_floor_and_ceiling():
    assert kpi.tat_score(3) == 100 and kpi.tat_score(7) == 100
    assert kpi.tat_score(10) == 76 and kpi.tat_score(40) == 0


def test_status_thresholds():
    assert kpi.status_for(95) == "Green" and kpi.status_for(85) == "Amber" and kpi.status_for(70) == "Red"


def test_flags():
    assert kpi.flags_for(1.0, 0.95, 6, 0.9) == "None"
    assert kpi.flags_for(0.5, 0.5, 20, 0.5) == "Below target; Low claim settlement; Slow claim TAT; Low renewal rate"


def test_branch_kpis_small_example():
    master = pd.DataFrame({"Branch_ID": ["A"], "Branch_Name": ["Alpha"], "Region": ["R"],
                           "Branch_Manager": ["M"], "Annual_Target": [1000]})
    pol = pd.DataFrame({"Policy_ID": ["P1", "P2"], "Branch_ID": ["A", "A"], "Premium": [400, 700]})
    claims = pd.DataFrame({"Policy_ID": ["P1", "P2"], "Status": ["Settled", "Rejected"], "TAT_Days": [6, None]})
    ren = pd.DataFrame({"Branch_ID": ["A"] * 4, "Renewed": ["Yes", "Yes", "Yes", "No"]})
    r = kpi.branch_kpis(master, pol, claims, ren).iloc[0]
    assert r.Premium == 1100 and round(r.Achievement, 2) == 1.10
    assert r.Settlement == 0.5 and r.Avg_TAT == 6 and r.Renewal == 0.75


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn(); print("PASS", name)

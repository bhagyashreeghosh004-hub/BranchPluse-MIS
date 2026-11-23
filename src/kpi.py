"""KPI logic in pandas. Mirrors the Excel formulas exactly (used for the email digest and for tests)."""
import numpy as np
import pandas as pd
from src import config as C


def tat_score(tat):
    return max(0.0, min(100.0, 100 - (tat - C.TAT_GOOD) * C.TAT_PENALTY))


def health_score(ach, settle, renew, tat):
    w = C.WEIGHTS
    return (w["achievement"] * min(ach, 1) * 100 + w["settlement"] * settle * 100
            + w["renewal"] * renew * 100 + w["tat"] * tat_score(tat))


def status_for(score):
    return "Green" if score >= C.GREEN_MIN else ("Amber" if score >= C.AMBER_MIN else "Red")


def flags_for(ach, settle, tat, renew):
    f = []
    if ach < C.FLAG_ACHIEVEMENT_BELOW: f.append("Below target")
    if settle < C.FLAG_SETTLEMENT_BELOW: f.append("Low claim settlement")
    if tat > C.FLAG_TAT_ABOVE: f.append("Slow claim TAT")
    if renew < C.FLAG_RENEWAL_BELOW: f.append("Low renewal rate")
    return "; ".join(f) if f else "None"


def load_clean():
    d = C.CLEAN_DIR
    return (pd.read_csv(d / "branch_master.csv"),
            pd.read_csv(d / "policies.csv", parse_dates=["Policy_Date"]),
            pd.read_csv(d / "claims.csv", parse_dates=["Claim_Date"]),
            pd.read_csv(d / "renewals.csv", parse_dates=["Due_Month"]))


def branch_kpis(master, pol, claims, ren):
    claims = claims.merge(pol[["Policy_ID", "Branch_ID"]], on="Policy_ID", how="left")
    rows = []
    for m in master.itertuples():
        p = pol[pol.Branch_ID == m.Branch_ID]
        c = claims[claims.Branch_ID == m.Branch_ID]
        r = ren[ren.Branch_ID == m.Branch_ID]
        premium = p.Premium.sum()
        ach = premium / m.Annual_Target
        settled = (c.Status == "Settled").sum()
        settle = settled / len(c) if len(c) else 0
        tat = c.loc[c.Status == "Settled", "TAT_Days"].mean()
        renew = (r.Renewed == "Yes").sum() / len(r) if len(r) else 0
        score = health_score(ach, settle, renew, tat)
        rows.append(dict(Branch_ID=m.Branch_ID, Branch=m.Branch_Name, Region=m.Region, Manager=m.Branch_Manager,
                         Target=m.Annual_Target, Premium=premium, Achievement=ach, Policies=len(p),
                         Claims=len(c), Settled=settled, Settlement=settle, Avg_TAT=tat,
                         Renewals_Due=len(r), Renewed=(r.Renewed == "Yes").sum(), Renewal=renew,
                         Score=score, Status=status_for(score), Flags=flags_for(ach, settle, tat, renew)))
    df = pd.DataFrame(rows)
    df["Rank"] = df.Score.rank(ascending=False, method="min").astype(int)
    return df.sort_values("Rank").reset_index(drop=True)


def totals(df, claims_all=None):
    t = dict(Premium=df.Premium.sum(), Target=df.Target.sum(), Policies=int(df.Policies.sum()),
             Claims=int(df.Claims.sum()), Settled=int(df.Settled.sum()),
             Renewals_Due=int(df.Renewals_Due.sum()), Renewed=int(df.Renewed.sum()))
    t["Achievement"] = t["Premium"] / t["Target"]
    t["Settlement"] = t["Settled"] / t["Claims"]
    t["Renewal"] = t["Renewed"] / t["Renewals_Due"]
    return t

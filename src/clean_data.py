"""Data Quality Gate: validates raw files, fixes what is safe to fix, drops what is not,
and writes an audit log of every action (nothing is changed silently)."""
import pandas as pd
from src.config import RAW_DIR, CLEAN_DIR, OUT_DIR


def clean():
    log = []

    def note(table, issue, n, action):
        log.append({"Table": table, "Issue Found": issue, "Rows Affected": int(n), "Action Taken": action})

    master = pd.read_csv(RAW_DIR / "branch_master.csv")
    pol = pd.read_csv(RAW_DIR / "policies_raw.csv", parse_dates=["Policy_Date"])
    claims = pd.read_csv(RAW_DIR / "claims_raw.csv", parse_dates=["Claim_Date"])
    ren = pd.read_csv(RAW_DIR / "renewals_raw.csv", parse_dates=["Due_Month"])
    n0 = len(pol)

    # ---- Policies ----
    fixed = pol["Branch_ID"].astype(str).str.strip().str.upper()
    note("Policies", "Branch code had spaces / lowercase", (fixed != pol["Branch_ID"]).sum(), "Standardised to upper-case, trimmed")
    pol["Branch_ID"] = fixed

    d = pol.duplicated(subset="Policy_ID", keep="first")
    note("Policies", "Duplicate Policy_ID", d.sum(), "Removed duplicates (kept first)")
    pol = pol[~d]

    bad = (pol["Premium"].isna()) | (pol["Premium"] <= 0)
    note("Policies", "Zero / negative / missing premium", bad.sum(), "Removed (reversal entries are not new business)")
    pol = pol[~bad]

    bad = pol["LOB"].isna() | (pol["LOB"].astype(str).str.strip() == "")
    note("Policies", "Missing line of business (LOB)", bad.sum(), "Removed (cannot be classified in reports)")
    pol = pol[~bad]

    bad = ~pol["Branch_ID"].isin(master["Branch_ID"])
    note("Policies", "Branch code not in Branch_Master", bad.sum(), "Removed")
    pol = pol[~bad]

    # ---- Claims ----
    orphan = ~claims["Policy_ID"].isin(pol["Policy_ID"])
    note("Claims", "Claim linked to a policy that does not exist (VLOOKUP would return #N/A)", orphan.sum(), "Removed")
    claims = claims[~orphan].copy()

    badtat = (claims["Status"] == "Settled") & (claims["TAT_Days"] <= 0)
    note("Claims", "Settled claim with impossible TAT (<= 0 days)", badtat.sum(), "TAT set to blank so averages are not distorted")
    claims.loc[badtat, "TAT_Days"] = pd.NA

    # ---- Renewals ----
    bad = ren["Renewed"].isna()
    note("Renewals", "Missing renewal outcome", bad.sum(), "Removed")
    ren = ren[~bad]

    note("Policies", "Rows in raw file / rows after cleaning", n0, f"{len(pol)} rows kept ({n0 - len(pol)} removed in total)")

    CLEAN_DIR.mkdir(parents=True, exist_ok=True)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    master.to_csv(CLEAN_DIR / "branch_master.csv", index=False)
    pol.to_csv(CLEAN_DIR / "policies.csv", index=False)
    claims.to_csv(CLEAN_DIR / "claims.csv", index=False)
    ren.to_csv(CLEAN_DIR / "renewals.csv", index=False)
    pd.DataFrame(log).to_csv(OUT_DIR / "data_quality_log.csv", index=False)
    print(f"[clean] policies={len(pol)} claims={len(claims)} renewals={len(ren)}  "
          f"({sum(r['Rows Affected'] for r in log[:-1])} issues logged)")
    return log


if __name__ == "__main__":
    clean()

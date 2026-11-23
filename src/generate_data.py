"""Generate a realistic, SYNTHETIC insurance-branch dataset (no real customer or company data).

Creates deliberately 'messy' raw files (duplicates, blanks, negatives, orphan claims)
so the Data Quality Gate in clean_data.py has real work to do.
"""
import numpy as np
import pandas as pd
from src.config import RAW_DIR, FY_START, FY_END

rng = np.random.default_rng(42)

# id, name, region, manager (fictional), annual target (INR), performance factor,
# claim settle rate, avg claim TAT (days), renewal rate
BRANCHES = [
    ("PN01", "Kothrud",       "West Pune",  "Neha Kulkarni",  6200000, 1.09, 0.95, 6.0, 0.89),
    ("PN02", "Hinjewadi",     "West Pune",  "Rohan Deshmukh", 7400000, 1.12, 0.93, 6.5, 0.86),
    ("PN03", "Baner",         "West Pune",  "Isha Patwardhan",5800000, 0.97, 0.91, 7.5, 0.82),
    ("PN04", "Aundh",         "West Pune",  "Sameer Joshi",   5200000, 1.02, 0.92, 7.0, 0.84),
    ("PN05", "Viman Nagar",   "East Pune",  "Pooja Bhosale",  6800000, 1.04, 0.90, 8.0, 0.83),
    ("PN06", "Kharadi",       "East Pune",  "Aditya More",    7000000, 0.88, 0.87, 9.0, 0.79),
    ("PN07", "Hadapsar",      "East Pune",  "Kavita Shinde",  6000000, 0.94, 0.89, 8.5, 0.80),
    ("PN08", "Wakad",         "North Pune", "Vikram Pawar",   5600000, 0.99, 0.90, 7.8, 0.81),
    ("PN09", "Pimpri",        "North Pune", "Anjali Gaikwad", 5000000, 0.81, 0.86, 9.5, 0.66),
    ("PN10", "Chinchwad",     "North Pune", "Mihir Kale",     5400000, 0.91, 0.88, 8.8, 0.77),
    ("PN11", "Katraj",        "South Pune", "Sneha Jadhav",   4600000, 0.78, 0.80, 13.5, 0.72),
    ("PN12", "Shivajinagar",  "South Pune", "Tejas Wagh",     6400000, 1.01, 0.93, 7.2, 0.85),
]

LOBS = ["Motor", "Health", "Travel", "Home", "Personal Accident"]
LOB_AVG = {"Motor": 11000, "Health": 19000, "Travel": 3200, "Home": 7500, "Personal Accident": 3800}
LOB_MIX = np.array([0.38, 0.30, 0.12, 0.08, 0.12])
LOB_CLAIM_P = {"Motor": 0.12, "Health": 0.09, "Travel": 0.04, "Home": 0.02, "Personal Accident": 0.03}
LOB_CLAIM_AMT = {"Motor": 22000, "Health": 45000, "Travel": 12000, "Home": 60000, "Personal Accident": 80000}
CHANNELS = ["Agent", "Online", "Bancassurance", "Direct"]
CHANNEL_P = [0.45, 0.25, 0.18, 0.12]

MONTHS = pd.date_range(FY_START, FY_END, freq="MS")
MONTH_W = np.array([.07, .07, .07, .075, .075, .08, .10, .08, .08, .08, .08, .12])
MONTH_W = MONTH_W / MONTH_W.sum()


def make_master():
    return pd.DataFrame(
        [(b[0], b[1], b[2], b[3], b[4]) for b in BRANCHES],
        columns=["Branch_ID", "Branch_Name", "Region", "Branch_Manager", "Annual_Target"],
    )


def make_policies():
    rows = []
    for (bid, _n, _r, _m, target, perf, *_rest) in BRANCHES:
        mix = LOB_MIX * rng.uniform(0.8, 1.2, len(LOBS))
        mix = mix / mix.sum()
        exp_avg = float((mix * np.array([LOB_AVG[l] for l in LOBS])).sum())
        n = int(target * perf / exp_avg)
        lobs = rng.choice(LOBS, size=n, p=mix)
        months = rng.choice(len(MONTHS), size=n, p=MONTH_W)
        for lob, mi in zip(lobs, months):
            m0 = MONTHS[mi]
            day = int(rng.integers(0, m0.days_in_month))
            prem = float(rng.lognormal(np.log(LOB_AVG[lob]), 0.35))
            rows.append((m0 + pd.Timedelta(days=day), bid, lob, round(prem, 0),
                         "Renewal" if rng.random() < 0.38 else "New",
                         rng.choice(CHANNELS, p=CHANNEL_P)))
    df = pd.DataFrame(rows, columns=["Policy_Date", "Branch_ID", "LOB", "Premium", "Policy_Type", "Channel"])
    df = df.sort_values("Policy_Date").reset_index(drop=True)
    df.insert(0, "Policy_ID", [f"P-{i+1:06d}" for i in range(len(df))])
    return df


def make_claims(pol):
    settle = {b[0]: b[6] for b in BRANCHES}
    tat = {b[0]: b[7] for b in BRANCHES}
    end = pd.Timestamp(FY_END)
    rows = []
    for r in pol.itertuples():
        if rng.random() < LOB_CLAIM_P[r.LOB]:
            cdate = min(r.Policy_Date + pd.Timedelta(days=int(rng.integers(15, 240))), end)
            amt = round(float(rng.lognormal(np.log(LOB_CLAIM_AMT[r.LOB]), 0.5)), 0)
            if rng.random() < settle[r.Branch_ID]:
                status, t = "Settled", max(1, int(round(rng.normal(tat[r.Branch_ID], 2.5))))
            else:
                status, t = ("Rejected" if rng.random() < 0.4 else "Pending"), np.nan
            rows.append((r.Policy_ID, cdate, amt, status, t))
    df = pd.DataFrame(rows, columns=["Policy_ID", "Claim_Date", "Claim_Amount", "Status", "TAT_Days"])
    df = df.sort_values("Claim_Date").reset_index(drop=True)
    df.insert(0, "Claim_ID", [f"C-{i+1:05d}" for i in range(len(df))])
    return df


def make_renewals(pol):
    renew = {b[0]: b[8] for b in BRANCHES}
    pol = pol.assign(Month=pol["Policy_Date"].dt.to_period("M").dt.to_timestamp())
    rows = []
    for (bid, m), g in pol.groupby(["Branch_ID", "Month"]):
        due = int(round(len(g) * 0.55 * rng.uniform(0.8, 1.2)))
        for _ in range(due):
            rows.append((bid, m, "Yes" if rng.random() < renew[bid] else "No"))
    df = pd.DataFrame(rows, columns=["Branch_ID", "Due_Month", "Renewed"])
    df = df.sort_values(["Due_Month", "Branch_ID"]).reset_index(drop=True)
    df.insert(0, "Due_ID", [f"R-{i+1:06d}" for i in range(len(df))])
    return df


def inject_mess(pol, claims, ren):
    """Plant realistic data-entry problems so the cleaning step is meaningful."""
    pol, claims, ren = pol.copy(), claims.copy(), ren.copy()
    pol["Branch_ID"] = pol["Branch_ID"].astype(object)
    pol["LOB"] = pol["LOB"].astype(object)
    # messy branch codes (lowercase / stray spaces)
    for i in rng.choice(len(pol), 12, replace=False):
        pol.at[i, "Branch_ID"] = " " + pol.at[i, "Branch_ID"].lower() + " "
    # negative premiums (reversal entries)
    for i in rng.choice(len(pol), 10, replace=False):
        pol.at[i, "Premium"] = -abs(pol.at[i, "Premium"])
    # blank LOB
    for i in rng.choice(len(pol), 9, replace=False):
        pol.at[i, "LOB"] = np.nan
    # duplicate rows
    pol = pd.concat([pol, pol.sample(25, random_state=7)], ignore_index=True)
    # orphan claims (policy id does not exist)
    orphans = claims.sample(7, random_state=3).copy()
    orphans["Claim_ID"] = [f"C-9{i:04d}" for i in range(7)]
    orphans["Policy_ID"] = [f"P-99{i:04d}" for i in range(7)]
    claims = pd.concat([claims, orphans], ignore_index=True)
    # impossible TAT on settled claims
    idx = claims.index[claims["Status"] == "Settled"]
    for i in rng.choice(idx, 5, replace=False):
        claims.at[i, "TAT_Days"] = -3
    # blank renewal outcome
    ren["Renewed"] = ren["Renewed"].astype(object)
    for i in rng.choice(len(ren), 8, replace=False):
        ren.at[i, "Renewed"] = np.nan
    return pol, claims, ren


def main():
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    master = make_master()
    pol = make_policies()
    claims = make_claims(pol)
    ren = make_renewals(pol)
    pol_raw, claims_raw, ren_raw = inject_mess(pol, claims, ren)
    master.to_csv(RAW_DIR / "branch_master.csv", index=False)
    pol_raw.to_csv(RAW_DIR / "policies_raw.csv", index=False)
    claims_raw.to_csv(RAW_DIR / "claims_raw.csv", index=False)
    ren_raw.to_csv(RAW_DIR / "renewals_raw.csv", index=False)
    print(f"[generate] branches={len(master)} policies={len(pol_raw)} "
          f"claims={len(claims_raw)} renewals={len(ren_raw)}  ->  {RAW_DIR}")


if __name__ == "__main__":
    main()

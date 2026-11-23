"""Loads the clean CSVs into SQLite and runs sql/branch_kpis.sql so you can see SQL and Excel agree."""
import sqlite3
import pandas as pd
from src import config as C

TABLES = {"branch_master": "branch_master.csv", "policies": "policies.csv",
          "claims": "claims.csv", "renewals": "renewals.csv"}


def main():
    db = C.OUT_DIR / "mis.db"
    if db.exists():
        db.unlink()
    con = sqlite3.connect(db)
    for name, f in TABLES.items():
        pd.read_csv(C.CLEAN_DIR / f).to_sql(name, con, index=False)
    statements = [s.strip() for s in (C.ROOT / "sql" / "branch_kpis.sql").read_text().split(";") if "SELECT" in s]
    for s in statements:
        print(pd.read_sql_query(s, con).to_string(index=False), "\n")
    con.close()


if __name__ == "__main__":
    main()

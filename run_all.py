"""One command to rebuild everything:  python run_all.py"""
from src import generate_data, clean_data, build_workbook, email_digest, config as C

if __name__ == "__main__":
    generate_data.main()
    clean_data.clean()
    build_workbook.build(C.OUT_DIR / "BranchPulse_MIS_Report.xlsx")
    email_digest.main()
    print("\nDone. Open output/BranchPulse_MIS_Report.xlsx and output/weekly_mis_email.html")

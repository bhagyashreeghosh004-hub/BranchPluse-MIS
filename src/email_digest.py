"""Auto-writes the weekly MIS email (HTML + plain text) from the same KPIs as the Excel pack."""
from src import config as C, kpi


def inr_cr(x):
    return f"₹{x/1e7:.2f} Cr"


def inr_lakh(x):
    return f"₹{x/1e5:.1f} lakh"


def build_digest():
    m, p, c, r = kpi.load_clean()
    df = kpi.branch_kpis(m, p, c, r)
    t = kpi.totals(df)
    top, bottom = df.head(3), df.tail(3).iloc[::-1]
    flagged = df[df.Flags != "None"]
    n = df.Status.value_counts()

    actions = []
    for x in flagged.itertuples():
        a = []
        if "Below target" in x.Flags: a.append(f"agree a recovery plan to close the {inr_lakh(x.Target - x.Premium)} gap to target")
        if "Slow claim TAT" in x.Flags: a.append(f"review pending claims (average TAT is {x.Avg_TAT:.1f} days)")
        if "Low claim settlement" in x.Flags: a.append(f"audit rejected and pending claims (settlement ratio {x.Settlement:.0%})")
        if "Low renewal rate" in x.Flags: a.append(f"start renewal follow-ups (renewal rate {x.Renewal:.0%})")
        actions.append(f"{x.Branch} ({x.Manager}): " + "; ".join(a) + ".")

    subject = f"Weekly Branch MIS | {C.FY_LABEL} | Network health {df.Score.mean():.1f}/100 | {n.get('Red', 0)} branch(es) need attention"
    lines = [
        "Dear Team,", "",
        f"Please find attached the branch MIS pack for {C.FY_LABEL}. Key points:", "",
        f"- Premium achieved: {inr_cr(t['Premium'])} against a target of {inr_cr(t['Target'])} ({t['Achievement']:.1%} achievement).",
        f"- Policies issued: {t['Policies']:,}.",
        f"- Claim settlement ratio: {t['Settlement']:.1%} ({t['Settled']} of {t['Claims']} claims).",
        f"- Renewal rate: {t['Renewal']:.1%}.",
        f"- Branch status: {n.get('Green', 0)} Green, {n.get('Amber', 0)} Amber, {n.get('Red', 0)} Red.", "",
        "Top performers: " + ", ".join(f"{x.Branch} ({x.Score:.1f})" for x in top.itertuples()) + ".",
        "Needs attention: " + ", ".join(f"{x.Branch} ({x.Score:.1f})" for x in bottom.itertuples()) + ".", "",
        "Suggested actions:",
    ] + [f"{i}. {a}" for i, a in enumerate(actions, 1)] + [
        "", "All figures were reconciled against the source data (see the DQ_Report sheet). "
        "Please reply with any corrections before this is circulated further.", "",
        "Regards,", "Bhagyashree Ghosh", "MIS Team",
    ]
    text = f"Subject: {subject}\n\n" + "\n".join(lines)

    rows = "".join(
        f"<tr><td>{x.Rank}</td><td>{x.Branch}</td><td>{x.Score:.1f}</td>"
        f"<td style='background:{ {'Green':'#C6EFCE','Amber':'#FFEB9C','Red':'#FFC7CE'}[x.Status] }'>{x.Status}</td>"
        f"<td>{x.Flags}</td></tr>" for x in df.itertuples())
    html = f"""<html><body style="font-family:Calibri,Arial,sans-serif;font-size:14px;color:#222">
<p><b>Subject:</b> {subject}</p><hr>
<p>Dear Team,</p>
<p>Please find attached the branch MIS pack for <b>{C.FY_LABEL}</b>.</p>
<ul>
<li>Premium achieved: <b>{inr_cr(t['Premium'])}</b> vs target {inr_cr(t['Target'])} (<b>{t['Achievement']:.1%}</b>)</li>
<li>Policies issued: <b>{t['Policies']:,}</b></li>
<li>Claim settlement ratio: <b>{t['Settlement']:.1%}</b></li>
<li>Renewal rate: <b>{t['Renewal']:.1%}</b></li>
</ul>
<table border="1" cellpadding="6" style="border-collapse:collapse"><tr style="background:#0B2545;color:#fff">
<th>Rank</th><th>Branch</th><th>Score</th><th>Status</th><th>Red flags</th></tr>{rows}</table>
<p><b>Suggested actions</b></p><ol>{''.join(f'<li>{a}</li>' for a in actions)}</ol>
<p>Regards,<br>Bhagyashree Ghosh<br>MIS Team</p></body></html>"""
    return text, html


def main():
    text, html = build_digest()
    C.OUT_DIR.mkdir(parents=True, exist_ok=True)
    (C.OUT_DIR / "weekly_mis_email.txt").write_text(text, encoding="utf-8")
    (C.OUT_DIR / "weekly_mis_email.html").write_text(html, encoding="utf-8")
    print("[email] wrote weekly_mis_email.txt / .html")


if __name__ == "__main__":
    main()

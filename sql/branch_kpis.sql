-- BranchPulse: the same branch KPIs as the Excel pack, written in SQL (SQLite syntax).
-- Run with:  python -m src.sql_demo

SELECT
    m.Branch_ID,
    m.Branch_Name,
    m.Region,
    SUM(p.Premium)                                   AS Premium,
    m.Annual_Target                                  AS Target,
    ROUND(100.0 * SUM(p.Premium) / m.Annual_Target, 1) AS Achievement_Pct,
    COUNT(p.Policy_ID)                               AS Policies
FROM branch_master m
JOIN policies p ON p.Branch_ID = m.Branch_ID
GROUP BY m.Branch_ID, m.Branch_Name, m.Region, m.Annual_Target
ORDER BY Achievement_Pct DESC;

-- Claim settlement ratio and average TAT per branch (JOIN replaces the Excel VLOOKUP)
SELECT
    p.Branch_ID,
    COUNT(*)                                                    AS Claims,
    SUM(CASE WHEN c.Status = 'Settled' THEN 1 ELSE 0 END)       AS Settled,
    ROUND(100.0 * SUM(CASE WHEN c.Status = 'Settled' THEN 1 ELSE 0 END) / COUNT(*), 1) AS Settlement_Pct,
    ROUND(AVG(CASE WHEN c.Status = 'Settled' THEN c.TAT_Days END), 1) AS Avg_TAT_Days
FROM claims c
JOIN policies p ON p.Policy_ID = c.Policy_ID
GROUP BY p.Branch_ID
ORDER BY Settlement_Pct ASC;

-- Month x line-of-business premium (the SQL version of a PivotTable)
SELECT
    strftime('%Y-%m', Policy_Date) AS Month,
    SUM(CASE WHEN LOB = 'Motor'  THEN Premium ELSE 0 END) AS Motor,
    SUM(CASE WHEN LOB = 'Health' THEN Premium ELSE 0 END) AS Health,
    SUM(CASE WHEN LOB = 'Travel' THEN Premium ELSE 0 END) AS Travel,
    SUM(CASE WHEN LOB = 'Home'   THEN Premium ELSE 0 END) AS Home,
    SUM(CASE WHEN LOB = 'Personal Accident' THEN Premium ELSE 0 END) AS Personal_Accident,
    SUM(Premium) AS Total
FROM policies
GROUP BY Month
ORDER BY Month;

-- Renewal rate per branch
SELECT Branch_ID,
       COUNT(*) AS Due,
       SUM(CASE WHEN Renewed = 'Yes' THEN 1 ELSE 0 END) AS Renewed,
       ROUND(100.0 * SUM(CASE WHEN Renewed = 'Yes' THEN 1 ELSE 0 END) / COUNT(*), 1) AS Renewal_Pct
FROM renewals
GROUP BY Branch_ID
ORDER BY Renewal_Pct ASC;

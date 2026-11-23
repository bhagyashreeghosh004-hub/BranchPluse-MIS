# Data dictionary (all data is synthetic)

| Table | Column | Meaning |
|---|---|---|
| Branch_Master | Branch_ID, Branch_Name, Region, Branch_Manager | 12 fictional Pune branches in 4 regions |
| Branch_Master | Annual_Target | Premium target for the year (INR) |
| Policies | Policy_ID, Policy_Date, Branch_ID | One row per policy sold |
| Policies | LOB | Motor, Health, Travel, Home, Personal Accident |
| Policies | Premium | Premium in INR |
| Policies | Policy_Type, Channel | New or Renewal; Agent, Online, Bancassurance, Direct |
| Claims | Claim_ID, Policy_ID, Claim_Date, Claim_Amount | One row per claim |
| Claims | Status, TAT_Days | Settled / Rejected / Pending; days to settle (settled claims only) |
| Renewals | Due_ID, Branch_ID, Due_Month, Renewed | One row per policy that came up for renewal; Yes or No |

## KPI definitions

| KPI | Formula |
|---|---|
| Achievement % | Premium achieved / Annual target |
| Claim settlement ratio | Settled claims / Claims received |
| Avg claim TAT | Average days to settle, settled claims only |
| Renewal rate | Renewed / Renewals due |
| **Branch Health Score** | 40% x min(Achievement, 100%) + 25% x Settlement + 20% x Renewal + 15% x TAT score |
| TAT score | 100 - (Avg TAT - 7) x 8, kept between 0 and 100 |
| Status | Green >= 90, Amber >= 82, otherwise Red |

Weights, thresholds and red-flag triggers live on the **Settings** sheet (yellow cells) and in `src/config.py`.

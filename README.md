# Company Expense Dashboard (Google Sheets)

A Google Sheets expense dashboard for the three projects under development:

1. **Aqua Bot** — funded (₹2,00,000 grant)
2. **IC Tester** — no budget yet
3. **CPS Training Kit** — no budget yet

Plus a **Company (General)** track for non-project expenses, funded by the
founders' contribution (₹5,000 × 2 for opening the bank account).

Everything is generated automatically by one Apps Script —
[`apps-script/dashboard_builder.gs`](apps-script/dashboard_builder.gs).
You never build the sheets by hand.

## What gets created

| Sheet | Purpose |
|---|---|
| **Dashboard** | KPI tiles: Funding Received, Total Spent, Account Balance, % of Funds Used. Per-project table (funding / spent / balance / usage bar), monthly spend table, and 3 charts (Funding vs Spent, Spend Share by Project, Monthly Spend). |
| **Funding** | Money received — pre-filled with the ₹2,00,000 Aqua Bot grant and the two ₹5,000 founder contributions (edit the founder names on the sheet). Add a row whenever new funds come in (date, source, project/track, amount). |
| **\<Project\> - Expenses** ×3 | The simple expense log you asked for: **Si No (auto-numbers itself), Date, Particulars of Expenditure, Amount, Remarks.** The Aqua Bot sheet is pre-loaded with the 16 expenditures already incurred (₹93,469.96). |
| **Company - Expenses** | Same format, for non-project company expenses (bank charges, registration, stationery, …). Shows up on the Dashboard as the **Company (General)** row. |
| **Aqua Bot - Utilization Certificate** | Print-ready UC in the standard format — Si No / Particulars of Expenditure / Amount (₹), with **Total Expenditure** and **Balance Amount (if any)** rows — generated from the expense log. A UC is created automatically for every project that has funding recorded. |

You only ever type into the **Funding** sheet and the **Expenses** sheets.
The Dashboard recalculates by itself; the Utilization Certificate is refreshed
from the menu.

## Setup (one time, ~2 minutes)

1. Create a new Google Sheet at [sheets.new](https://sheets.new) and name it
   (e.g. *Company Expenses FY 2026-27*).
2. In the sheet, open **Extensions → Apps Script**.
3. Delete the placeholder code, paste the full contents of
   [`apps-script/dashboard_builder.gs`](apps-script/dashboard_builder.gs), and save (💾).
4. In the function dropdown at the top, select **`buildDashboard`** and click **Run**.
5. Google asks for authorization the first time — choose your account →
   *Advanced → Go to (project) → Allow*. (The script only touches this one
   spreadsheet.)
6. Switch back to the spreadsheet tab — all sheets are generated.

After the first run a **⚙ Dashboard Tools** menu appears inside the spreadsheet:

- **Build / Rebuild all sheets** — regenerates everything. Rows already
  entered in Funding and the Expenses sheets are **kept**; the Dashboard and
  UC sheets are rebuilt from scratch.
- **Refresh Utilization Certificates** — regenerates the UC sheet(s) from the
  current expense log. Run this after adding expenses, before printing/sharing a UC.

## Using it day-to-day

- **Log an expense:** open the project's `- Expenses` sheet, enter Date,
  Particulars, and Amount on the next row. The Si No numbers itself.
- **Record funds received:** add a row on the **Funding** sheet. Projects with
  funding automatically get balance tracking, usage bars, and a UC sheet.
- **Produce a UC:** ⚙ Dashboard Tools → *Refresh Utilization Certificates*,
  then print or download that sheet as PDF (File → Download → PDF, or
  File → Print with "Current sheet").
- The pre-loaded Aqua Bot entries have no dates (the source record didn't
  include them). Totals and the UC are unaffected; add dates when known and
  the monthly-spend chart will pick them up.

## Customizing

All configuration is at the top of `dashboard_builder.gs`:

| Constant | What it controls |
|---|---|
| `PROJECTS` | Project names — add/remove and rebuild; sheets, dashboard rows, and charts adapt. |
| `INITIAL_FUNDING` | Funding rows seeded on first build (₹2 L Aqua Bot grant + 2 × ₹5,000 founder contributions). |
| `AQUABOT_EXPENSES` | The pre-loaded Aqua Bot expenditure rows. |
| `CURRENCY_FORMAT` | Default `₹#,##0.00`. |
| `FY_START_MONTH` | First month of the monthly table. Default `3` (April, Indian FY); `0` for January. |

After editing constants, run **⚙ Dashboard Tools → Build / Rebuild all sheets**.
(Rebuild wipes entered data, so customize before entering real expenses, or
copy your rows out first.)

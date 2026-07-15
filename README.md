# Company Expense Dashboard (Google Sheets)

A Google Sheets expense dashboard for the three projects under development:

1. **Aqua Bot**
2. **IC Tester**
3. **CPS Training Kit**

Everything is generated automatically by one Apps Script —
[`apps-script/dashboard_builder.gs`](apps-script/dashboard_builder.gs).
You never build the sheets by hand.

## What gets created

| Sheet | Purpose |
|---|---|
| **Dashboard** | Company-wide overview: KPI tiles (Total Budget / Utilized / Remaining / % Utilized), a per-project summary table with usage bars, monthly spend table, and 3 charts (Budget vs Utilized, Spend Share by Project, Monthly Spend by Project). |
| **\<Project\> - Expenses** ×3 | The expense log where you enter data. Columns: Date, Description, Category (dropdown), Vendor, Invoice No., Amount, Payment Mode (dropdown), Paid By, Notes. |
| **\<Project\> - Fund Utilization** ×3 | Budget vs actuals per category: you type the **Allocated Budget** per category (yellow cells); Utilized / Remaining / % Utilized and the usage bars calculate themselves from the expense log. Also has a monthly spend table, a category pie chart, and a monthly column chart. |
| **Lists** (hidden) | Sources for the Category / Payment Mode dropdowns. |

Everything is formula-driven: **you only ever type into the Expenses sheets and
the yellow "Allocated Budget" cells** — every dashboard, chart, and utilization
figure updates by itself.

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
6. Switch back to the spreadsheet tab — all 8 sheets are generated.

After the first run, a **⚙ Dashboard Tools** menu also appears inside the
spreadsheet with a *Build / Rebuild all sheets* option.
⚠️ Rebuilding erases and regenerates the sheets (it asks for confirmation first).

## Using it day-to-day

- **Log an expense:** open the project's `- Expenses` sheet and add a row.
  Use the dropdowns for Category and Payment Mode so the utilization math stays accurate.
- **Set budgets:** on each `- Fund Utilization` sheet, fill in the yellow
  *Allocated Budget* column for each category.
- **Delete the sample rows:** each Expenses sheet ships with 3 sample rows
  (marked "Sample row — delete") so the charts render immediately — remove them
  and the pre-filled sample budgets before real use.
- **Read the dashboard:** the % Utilized cells turn amber above 75% and red
  above 100% of budget, on both the Dashboard and each Fund Utilization sheet.

## Customizing

All configuration is at the top of `dashboard_builder.gs`:

| Constant | What it controls |
|---|---|
| `PROJECTS` | Project names — add/remove projects here and rebuild; expense + utilization sheets and all dashboard rows/charts adapt automatically. |
| `CATEGORIES` | Expense categories (dropdown + utilization rows). |
| `PAYMENT_MODES` | Payment mode dropdown options. |
| `CURRENCY_FORMAT` | Default `₹#,##0` — change to `$#,##0.00` etc. |
| `FY_START_MONTH` | First month of the monthly tables. Default `3` (April, Indian FY); use `0` for January. |

After editing constants, run **⚙ Dashboard Tools → Build / Rebuild all sheets**.
(Rebuild wipes entered data, so customize before you start entering real expenses,
or copy your rows out first.)

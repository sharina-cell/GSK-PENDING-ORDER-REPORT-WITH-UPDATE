# GSK Pending Order Report

Daily report generator + NOT PUSHED order updater for GSK/Haleon marketplace orders (Shopee, TikTok, Lazada).

---

## Repo contents

| File | Purpose |
|---|---|
| `app.py` | Streamlit app — all logic in one file |
| `requirements.txt` | Python dependencies |

---

## Deploy on Streamlit Community Cloud

1. Push this repo to GitHub (public or private)
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**
3. Select repo, branch `main`, main file `app.py`
4. Click **Deploy**

---

## Local setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

---

## How to use

### Tab 1 — Generate Report

Upload the four daily files:

| Upload | File |
|---|---|
| TC Order Report | `GSK_TC_ORDER_REPORT.csv` |
| Shopee Order File | `GSK_SHOPEE_ORDER.xlsx` |
| WMS File(s) | One or more `GSK_WMS_*.xlsx` (auto-combined) |
| Inventory Report | `InventoryWarehouse_Report_*.xlsx` |

Click **Generate Report** → download the Excel file.

### Tab 2 — Update NOT PUSHED Orders

After checking your NOT PUSHED sheet and confirming certain orders are now in WMS:

1. The report from Tab 1 is already loaded — or upload an existing report file
2. Paste the invoice numbers (one per line or comma-separated)
3. Select the WMS status (`PENDING_FULFILLMENT` by default)
4. Click **Update Report** → download the corrected file

The app will:
- Move the matching rows from **NOT PUSHED** → **FILTERED DATA** with the new WMS status
- Remove them from **CLOSE & NOT FOUND**
- Rebuild the **PIVOT** table
- Leave PARTIAL & NON ALLOCATED, ORIGINAL, and RMA sheets untouched

---

## Report sheet reference

| Sheet | Tab colour | Contents |
|---|---|---|
| FILTERED DATA | Blue | Active pending orders (excl. NOT FOUND, CANCEL, zero-stock partial) |
| PIVOT | Blue | Order count by MP SLA date × Order Item Status |
| NOT PUSHED | Red | Orders not yet found in WMS |
| PARTIAL & NON ALLOCATED | Orange | Partial/none allocated with stock levels and CANCEL/FULFILL remarks |
| CLOSE & NOT FOUND | Grey | CLOSE or NOT FOUND WMS status |
| ORIGINAL (All Data) | Black | Full TC report, unfiltered |
| RMA | Dark red | Cancelled orders already fulfilled in WMS |

---

## Business rules baked in

- **WMS CANCEL / CANCELLED** rows are excluded from FILTERED DATA and PIVOT (same as NOT FOUND)
- **Partial orders with zero warehouse stock** → Remarks = `CANCEL` (red); positive stock → `FULFILL` (green)
- **MP SLA logic**: Shopee = Estimated Ship Out Date from Seller Centre file; TikTok / Lazada = ordered date + 1 day
- **PIVOT date colours**: past dates = red, today = yellow, future = blue

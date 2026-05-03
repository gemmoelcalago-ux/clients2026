# St. Peter Payment Tracker — Google Apps Script Setup Guide

## What You're Getting

Two files that replace your localStorage-based HTML tracker with a live Google Sheets database:

| File       | Purpose                                           |
|------------|---------------------------------------------------|
| `Code.gs`  | Server-side logic: reads/writes Google Sheets     |
| `index.html` | The full UI — served directly from Apps Script  |

All payment data, notes, and payment methods are stored in Google Sheets in real time.
No more browser-only storage — any device, any browser, always in sync.

---

## Step 1 — Create a Google Spreadsheet

1. Go to [sheets.google.com](https://sheets.google.com) and create a **new blank spreadsheet**.
2. Name it: `St. Peter Payment Tracker`

---

## Step 2 — Open Apps Script

1. In your spreadsheet, click **Extensions → Apps Script**.
2. You'll see a default `Code.gs` file.

---

## Step 3 — Add the Code

### Code.gs
1. Delete everything in the default `Code.gs` editor.
2. Paste the entire contents of **`Code.gs`** from the files provided.

### index.html
1. Click the **+** button next to "Files" in the left sidebar.
2. Select **HTML**.
3. Name it exactly: `index` (Apps Script will add `.html` automatically).
4. Delete the default content and paste the entire contents of **`index.html`**.

---

## Step 4 — Run Setup (Seeds the Sheet with Client Data)

1. In the Apps Script editor, make sure `Code.gs` is selected.
2. In the toolbar dropdown (showing function names), select **`setupSheets`**.
3. Click the ▶ **Run** button.
4. When prompted, click **Review permissions → Allow**.
5. You should see an alert: ✅ *Setup complete! All sheets created and seeded.*

This creates 4 sheets in your spreadsheet:
- **Clients** — all 21 clients with DOB
- **Payments** — one row per payment recorded
- **Notes** — one row per client with notes
- **PaymentMethods** — one row per client with their payment method

---

## Step 5 — Deploy as Web App

1. In Apps Script, click **Deploy → New deployment**.
2. Click the gear icon ⚙ next to "Select type" → choose **Web app**.
3. Set:
   - **Description**: `St. Peter Tracker v1`
   - **Execute as**: `Me`
   - **Who has access**: `Anyone` *(so clients can log in without a Google account)*
4. Click **Deploy**.
5. Copy the **Web app URL** — this is your tracker link!

> 💡 Every time you edit the code, go to **Deploy → Manage deployments → Edit → New version → Deploy** to publish the update.

---

## Step 6 — Change the Admin Password

In `Code.gs`, find line 8:

```javascript
const ADMIN_PASSWORD = 'admin123'; // ← CHANGE THIS
```

Change `'admin123'` to something secure, then redeploy.

---

## How the Sheets Database Works

### Clients (read-only after setup)
| id | name | policy | plan | planType | batch | dob |
|----|------|--------|------|----------|-------|-----|

### Payments (written on every payment record/clear)
| clientId | monthKey | paymentDate | recordedAt |
|----------|----------|-------------|------------|

### Notes (written on every notes save)
| clientId | notes | updatedAt |
|----------|-------|-----------|

### PaymentMethods (written on every method change)
| clientId | method | updatedAt |
|----------|--------|-----------|

---

## Adding New Clients

Simply add a new row to the **Clients** sheet in your spreadsheet.
The app reads clients live from Sheets, so no code changes needed.

Columns: `id | name | policy | plan | planType | batch | dob`

---

## Adding New Months

In `Code.gs` and `index.html`, find the `MONTHS` array and add entries:

```javascript
{ key: 'jul27', label: 'Jul 2027' },
```

Then redeploy.

---

## Sharing with Clients

1. Log in as Admin.
2. Click the **🔗 Share** button next to any client.
3. Copy the link and send it to them.
4. The client visits the link, selects their name, and enters their Date of Birth.
5. They see only their own payment history — nothing else.

---

## Troubleshooting

| Issue | Fix |
|-------|-----|
| "Could not connect to database" | Make sure you deployed as Web App with "Anyone" access |
| Blank page / no clients | Re-run `setupSheets` from the Apps Script editor |
| Changes not showing | Redeploy: Deploy → Manage deployments → Edit → New version |
| "Incorrect password" | Check `ADMIN_PASSWORD` in Code.gs matches what you're typing |

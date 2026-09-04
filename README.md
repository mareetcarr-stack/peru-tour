# Yenrri's App (Peru Tour Manager)

A light-themed single-page PWA for managing Yenrri's Peru tours — flights, passenger check-in status, daily checklist, optional tours, tickets, traveller info/birthdays, and recommended places. Reads and writes live data from/to a Google Sheet through a small Google Apps Script backend. No PIN, no accounts, no OAuth — the deployed Apps Script URL is the only thing gating access.

## How it fits together

```
index.html (GitHub Pages, static)  →  Apps Script Web App  →  Google Sheet
```

GitHub Pages can only serve static files, so the Apps Script Web App is the tiny backend that actually reads/writes the Sheet. It runs under **your** Google account — no service account key is needed for this app (the existing `credentials.json` in `latam-cloud/` is untouched and keeps powering your separate Node automation scripts).

## 1. Deploy the Apps Script backend (one-time, ~5 minutes)

1. Open the Google Sheet → **Extensions → Apps Script**.
2. Delete whatever's in `Code.gs` and paste in the contents of [`apps-script/Code.gs`](apps-script/Code.gs) from this repo.
3. **Deploy → New deployment**:
   - Type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
4. Click **Deploy**, authorize the permissions prompt if shown (that's you granting your own script access to your own Sheet), then copy the **Web app URL** (ends in `/exec`).
5. Paste that URL into `CONFIG.APPS_SCRIPT_URL` in `index.html` if it ever changes (it's already set from the first deploy).

### One manual Sheet change needed

The **Valencia (V)** column and the **Tips** column in the `TOURS` tab used to be plain checkboxes (TRUE/FALSE). They now store a currency + amount instead (e.g. `USD` or `PEN:120.00`), so Sheets' checkbox validation on those columns will otherwise fight the app. Select those columns → **Data → Data validation → Remove rule**. (The app also calls `clearDataValidations()` on write as a safety net, but removing the rule up front keeps the Sheet looking right when you glance at it directly.)

### Re-deploying after editing Code.gs
Deploy → Manage deployments → pencil icon → New version → Deploy. This keeps the same `/exec` URL working (a brand new deployment would get a new URL, breaking the app until you update it).

## 2. Test locally before publishing

Open `index.html` directly in a browser (double-click it, or `open index.html`). Confirm each tab loads live data, tick a checklist item / tour / ticket / birthday, then check the actual Google Sheet to confirm the cell updated.

## 3. Publish to GitHub Pages

```bash
cd peru-tour
git add .
git commit -m "Update"
git push
```

Pages is already enabled (Settings → Pages → main branch, root). Live at:

```
https://mareetcarr-stack.github.io/peru-tour/
```

**Note:** this is a **public** repo (GitHub Pages on the free plan requires a public repo). That's fine here — no traveller data or secrets are stored in the code itself; everything is fetched live into the browser at runtime and never committed. Since there's no PIN anymore, anyone who finds the URL can read/write the Sheet's trip data — acceptable for this use case, but worth keeping in mind.

## 4. Install on Yenrri's iPhone/iPad

1. Open **https://mareetcarr-stack.github.io/peru-tour/** in **Safari** (must be Safari, not Chrome, for "Add to Home Screen" to create a standalone app).
2. Tap the **Share** button → **Add to Home Screen**.

## Notes / known limitations

- **Currency converter** is a static calculator with fixed exchange rates — not sheet-backed, matches the original design.
- **Ticket notes** (the textarea on the Tickets tab) save to that device's local storage only, not to the Sheet — there's no dedicated column for it. Ask if you'd like a `Notes` column added to the `TICKETS` tab so it syncs properly.
- **Dietary requirements** are directly editable in the app and sync straight to the `B-DAYS, AGES, DIET` sheet.
- **Arrived at airport** count on the Passengers tab reads the `CHECK-IN STATUS` column D `Arrived` flag and excludes Yenrri (row id `0`) and cancelled travellers from both the numerator and denominator. Ticking a passenger's box writes column D for that row; the **Select all / Deselect all** button next to the list header does the same for every traveller in one write (again skipping Yenrri and anyone cancelled), and flips to "Deselect all" once everyone is ticked.
- Sync is "live enough", not instant multi-device real-time: every tap writes to the Sheet immediately, and the app re-fetches the full data set every ~45 seconds and whenever the app regains focus.
- Tour/tip prices are editable per-tour in the Optional Tours table header — editing one writes straight back to the Sheet's price row for everyone.
- Optional Tours grand totals show both currencies with the other shown in brackets as a rough conversion (fixed rates, same as the currency converter) — not a live FX rate.
- `RECEIPTS`, `VALENCIA TRAVEL`, `Hoja 22`, `Hoja 25` tabs aren't used by any screen.

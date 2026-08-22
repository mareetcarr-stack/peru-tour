# Peru Tour Manager

A dark-themed single-page PWA for managing Yenrri's Peru tours — flights, passenger check-in status, daily checklist, optional tours, tickets, traveller info/birthdays, and recommended places. Reads and writes live data from/to a Google Sheet through a small Google Apps Script backend, gated by a 4-digit PIN. No accounts, no OAuth — Yenrri just enters a PIN once on his iPhone/iPad.

## How it fits together

```
index.html (GitHub Pages, static)  →  Apps Script Web App  →  Google Sheet
```

GitHub Pages can only serve static files, so the Apps Script Web App is the tiny backend that actually reads/writes the Sheet. It runs under **your** Google account — no service account key is needed for this app (the existing `credentials.json` in `latam-cloud/` is untouched and keeps powering your separate Node automation scripts).

## 1. Deploy the Apps Script backend (one-time, ~5 minutes)

1. Open the Google Sheet → **Extensions → Apps Script**.
2. Delete whatever's in `Code.gs` and paste in the contents of [`apps-script/Code.gs`](apps-script/Code.gs) from this repo.
3. In the function dropdown at the top of the editor, select **`setPin`** and click ▶ **Run**. This sets the PIN to `1234` in Script Properties (change the value in the code first if you want something else, or change it later — see "Changing the PIN" below).
   - The first run will prompt you to authorize the script — that's you granting your own script access to your own Sheet. Click through it (Advanced → Go to project (unsafe) is expected for a script you just pasted yourself).
4. **Deploy → New deployment**:
   - Type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Click **Deploy**, authorize again if prompted, then copy the **Web app URL** (ends in `/exec`).
6. Open `index.html` in this folder and paste that URL into:
   ```js
   const CONFIG = {
     APPS_SCRIPT_URL: 'https://script.google.com/macros/s/AKfycby0-SNX19SxU-KRk64b1ptP1NBjYghoIFri0GXTujx16vjGkh0wEh1WzGhXFNjG6JndeQ/exec',
   };
   ```

### Changing the PIN later
Extensions → Apps Script → Project Settings (gear icon) → Script Properties → edit `PIN`. No redeploy needed.

### Re-deploying after editing Code.gs
Deploy → Manage deployments → pencil icon → New version → Deploy. This keeps the same `/exec` URL working (a brand new deployment would get a new URL, breaking the app until you update it).

## 2. Test locally before publishing

Open `index.html` directly in a browser (double-click it, or `open index.html`). Enter the PIN, confirm each tab loads live data, tick a checklist item / tour / ticket / birthday, then check the actual Google Sheet to confirm the cell updated.

## 3. Publish to GitHub Pages

```bash
cd peru-tour
git init
git add .
git commit -m "Peru Tour Manager PWA"
gh repo create peru-tour --public --source=. --push
```

(Or without the `gh` CLI: create a repo called `peru-tour` on github.com, then `git remote add origin <url>` and `git push -u origin main`.)

Then: repo **Settings → Pages → Source: Deploy from a branch → main / (root)**. GitHub gives you a URL like:

```
https://<your-username>.github.io/peru-tour/
```

**Note:** this makes the repo public (GitHub Pages on the free plan requires a public repo). That's fine here — no traveller data or secrets are stored in the code itself; everything is fetched live into the browser at runtime and never committed.

## 4. Install on Yenrri's iPhone/iPad

1. Open the GitHub Pages URL in **Safari** (must be Safari, not Chrome, for "Add to Home Screen" to create a standalone app).
2. Tap the **Share** button → **Add to Home Screen**.
3. Enter the PIN once — it's remembered on that device (stored in Safari's local storage) until Safari data is cleared.

## Notes / known limitations

- **Currency converter** is a static calculator with fixed exchange rates — not sheet-backed, matches the original design.
- **Ticket notes** (the passport-discrepancy textarea) save to that device's local storage only, not to the Sheet — there's no dedicated column for it. Ask if you'd like a `Notes` column added to the `TICKETS` tab so it syncs properly.
- **Dietary requirements** are read-only in the app (display only), matching the original design — editing them means editing the `B-DAYS, AGES, DIET` sheet directly. The Apps Script backend already has a `setDiet` write action ready if you want an edit UI added later.
- Sync is "live enough", not instant multi-device real-time: every tap writes to the Sheet immediately, and the app re-fetches the full data set every ~45 seconds and whenever the app regains focus.
- `RECEIPTS`, `VALENCIA TRAVEL`, `Hoja 22`, `Hoja 25` tabs aren't used by any screen.

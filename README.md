# Email sender plus PDF creator

Google Apps Script tools for a Google Sheet: personalized email blasts, Gmail
bounce tracking, contact lookup, and personalized PDF bundles with mailing labels.

## Installing into a Sheet

1. Open the Google Sheet you want the tools on
2. **Extensions → Apps Script**
3. Delete whatever is in `Code.gs`
4. Paste all of [`dist/Code.gs`](dist/Code.gs)
5. Save, reload the Sheet

A **📧 Email Tools** menu appears. The `People`, `email details` and `Bounced`
sheets are created automatically on first open, and their header rows are
enforced every open — a renamed header gets reset, since the sheets are read by
column position.

That is the whole install. Two features quietly no-op until you enable their
Google service under **Services +** in the Apps Script editor:

| Service | Powers | Without it |
|---|---|---|
| `Gmail` | pulls your Gmail signature onto outgoing mail | no signature appended |
| `People` | "Fill Emails from Google Contacts" | falls back to legacy Contacts + Gmail history |

Everything else — sending, drafts, bounce checking, PDFs, VCF import — works
from the paste alone.

## Sheets

**`People`** — one row per recipient.

| A | B | C | D | E |
|---|---|---|---|---|
| Name | PAC Names | Email | Phone | Address |

**`email details`** — the templates.

| Cell | Holds |
|---|---|
| `A2` | body template — `[first name]` or `{{first name}}` placeholders |
| `B2` | subject template |
| `C2` | Drive URL or file ID of a PDF to attach (optional) |
| `D2`, `D3`, … | CC addresses, one per row (optional) |

**`Bounced`** — created automatically by the bounce checker.

| Email | Name | Bounced On | Type | Reason | Message |
|---|---|---|---|---|---|

## Bounce checking

**📧 Email Tools → Bounces**

- **Check for Bounces Now** — scans Gmail for delivery failures, matches them
  against `People`, appends anything new to `Bounced`
- **Enable Auto-Check on Open** — runs that scan every time the Sheet opens

Notes:

- A simple `onOpen(e)` trigger runs **without authorization** and cannot read
  Gmail, so auto-check installs a proper trigger via `ScriptApp.newTrigger()`.
  That is why it is an explicit opt-in rather than automatic.
- Auto-check adds a few seconds to Sheet load. The manual menu item is often
  the better trade.
- The first scan looks back 30 days; later scans resume from the last one.
- The failed address comes from the RFC 3464 `Final-Recipient` field in the raw
  message, falling back to scraping Gmail's wording.
- Results are deduped by email, so re-running is safe.

## Repo layout

```
src/         modules — edit these
dist/Code.gs generated bundle — paste this, never edit
build.mjs    concatenates src/ into dist/Code.gs
```

Apps Script shares one global scope across every `.gs` file in a project, so
bundling them into one file changes nothing about how the code runs.

### Making changes

```bash
node build.mjs           # regenerate dist/Code.gs
node build.mjs --check   # fail if the bundle is stale (CI uses this)
```

CI rebuilds `dist/Code.gs` on every push to `main` and fails any PR whose
bundle is out of date, so the file you paste is never behind `src/`.

Adding a module: drop it in `src/` and add it to `ORDER` in `build.mjs`.
Unlisted files still get bundled, alphabetically, after the listed ones.

## Sending limits

`GmailApp.sendEmail` is capped by Google at roughly **100 recipients/day** on a
consumer Gmail account and **1,500/day** on Workspace. The send loop has no
quota guard and no resume — if it hits the cap mid-list it throws, and the rows
already sent are not recorded anywhere.

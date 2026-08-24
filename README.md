# Distributor Invoice Generator

A small Flask web app that turns your distributor Excel sheet into one
Word (.docx) TAX INVOICE per distributor, automatically, with GST
calculated per your rules — plus an Email Config tab to send every
invoice straight to its distributor.

## Run locally

```bash
cd distributor_invoice_app
pip install -r requirements.txt
python app.py
```

Open http://127.0.0.1:5000 in your browser.

## Put it online for your team (free) — Render

[Render](https://render.com) gives you a free, always-public URL with
no credit card. The free tier "sleeps" after ~15 minutes of no
traffic, so the first request after a break takes 30-60 seconds to
wake up — fine for an internal tool.

1. **Push this folder to a GitHub repo** (public or private both work).
2. Go to [render.com](https://render.com) → sign up (GitHub login is
   fastest) → **New +** → **Web Service** → connect your repo.
3. Render should auto-detect Python. Set:
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app`
   - **Instance Type:** Free
4. Add environment variables under **Environment**:
   - `TEAM_PASSWORD` — a password your team will type in before using
     the app. Leave unset if you want it open to anyone with the link.
   - `SECRET_KEY` — any random string (signs the login session cookie).
   - `EMAIL_ADDRESS` — the sending address, e.g. `info@myalternates.com`
     (defaults to that if you leave it unset).
   - `EMAIL_APP_PASSWORD` — a Gmail **App Password** for that account
     (Google Account → Security → 2-Step Verification → App Passwords).
     Required only for the Email Config tab; invoice generation works
     without it. **Never commit this to the repo or paste it in chat/
     Slack/email — set it only as an environment variable.** If a
     password has ever been shared outside this variable, rotate/
     regenerate it in Google immediately.
5. Click **Create Web Service**. In a couple of minutes you'll get a
   URL like `https://distributor-invoices.onrender.com` — share that
   with your team.

That's it — no server to manage. Every push to your GitHub repo
redeploys automatically.

### Alternative: PythonAnywhere

Also free, no credit card, and doesn't sleep — good if you'd rather
not use Git. Trade-off: limited daily CPU seconds on the free plan, so
it's best for occasional/light use. Sign up at
[pythonanywhere.com](https://www.pythonanywhere.com), create a Flask
web app from the dashboard, upload this folder via their Files tab or
`git clone`, point the WSGI config at `app.py`'s `app` object, install
`requirements.txt` in a virtualenv, and set the same environment
variables from step 4 above in the "Web" tab's env var section.

## How it works

### 1. Generate Invoices tab

1. Upload the distributor Excel file with columns: Partner Code,
   Partner Name, Total, GST No, GST, Address, PAN, Bank Name,
   Bank Account Number, ifsc code, Email.
   - **Address, PAN and Email are optional** — if a distributor's cell
     is blank, that line/field is simply left off (Email is only
     needed for the Email Config step).
2. Pick the payout month and the invoice date (default to today) —
   these apply to every invoice in that batch, rendered as `Apr-2023`
   and `30-Apr-2023` respectively.
3. Click **Generate invoices (.zip)** — you get a ZIP with one
   `.docx` per distributor, named `<PartnerCode>_<Month>.docx`, plus a
   `summary.csv`. That same batch is also kept on the server so you
   can switch to the **Email Config** tab (top right) and send it out.

### 2. Email Config tab

After generating a batch, switch to this tab to review/edit each
distributor's email (pulled from the sheet's Email column, editable
here without touching the Excel file), write one subject + body using
`{name}` / `{code}` placeholders, optionally add director BCC
addresses, and click **Send emails to all distributors** — each
distributor gets their own email with only their own invoice attached.
A results table shows sent / failed / skipped per distributor.

## GST rule applied to every distributor row

The `Total` column in your sheet is the **base (pre-tax) commission**
amount, shown as-is on the invoice. The tax amount is read directly
from the sheet's **`GST`** column (a rupee value) — it is *not*
calculated from a percentage — and added on top:

| GST No. on file            | Tax charged                                    |
|-----------------------------|-------------------------------------------------|
| Blank / `0`                 | None — invoice = Total as-is (GST column ignored) |
| Present, doesn't start "33" | Full GST column value, as IGST                 |
| Present, starts with "33"   | GST column value split in half: SGST + CGST    |

```
base (Particulars amount) = Total
tax  (IGST, or SGST+CGST)  = GST column value
grand total                = Total + GST
```

Myalternates' own details (name, address, GSTIN, place of supply, SAC
code) are fixed on every invoice, same as the sample.

## Layout / design details

- **Myalternates logo** in the top-right corner of every invoice
  (`assets/myalternates_logo.png` — swap this file to change it).
- **Tight, consistent spacing** throughout — no stray gaps between
  rows in the From/To blocks.
- **Address** wraps at roughly half the page width and continues on
  extra lines below if long, instead of stretching edge to edge.
- **Bank Details** laid out in a clean, borderless aligned column
  (label, colon, value) rather than tab-separated text.
- **Particulars table** has generously padded, taller rows for a more
  professional look.
- **Amount Chargeable (in words)** and **Tax Amount (in words)** shown
  as two lines directly under the particulars table.
- Extra spacing between the **"For &lt;distributor&gt;"** line and
  **AUTHORISED SIGNATORY** so there's room to actually sign in between.
- Every invoice fits on **a single page**.

## Files

- `app.py` — Flask routes: optional team-password gate, upload →
  generate → zip, in-memory batch store, Email Config page, and the
  send-emails route (SMTP over SSL with an app password)
- `invoice_engine.py` — GST math, number-to-words, and the .docx builder
- `templates/index.html` — Generate Invoices tab (upload + pickers)
- `templates/email_config.html` — Email Config tab (distributor list,
  subject/body compose, BCC, send, results)
- `templates/login.html` — password gate page (only used if
  `TEAM_PASSWORD` is set)
- `assets/myalternates_logo.png` — logo placed top-right on every invoice
- `requirements.txt`

## Customizing

- To change the fixed Myalternates details, SAC code, or GST split
  logic, edit the constants/`compute_gst` in `invoice_engine.py`.
- To accept a different set of Excel column names, edit
  `COLUMN_ALIASES` in `app.py`.
- To change how wide the address wraps, adjust the `right_indent` set
  on the address paragraph in `build_invoice_docx`.
- To resize/reposition the logo, edit `_add_logo_header` in
  `invoice_engine.py`.
- To turn off the password gate, just don't set `TEAM_PASSWORD` on
  your host.
- To change the outgoing email account or SMTP provider, edit
  `EMAIL_ADDRESS` / `SMTP_HOST` / `SMTP_PORT` via environment
  variables in `app.py` — no code change needed for a same-provider
  address change.

## Notes / limitations

- The batch store (used to bridge Generate → Email Config) lives in
  server memory and is cleared after 6 hours or on restart. If you
  ever scale this to more than one gunicorn worker, that store needs
  to move to something shared (e.g. SQLite/Redis) — a single worker
  (Render's free tier default) is fine as-is.
- Sending relies on the receiving mail server accepting SMTP over SSL
  from Gmail with an App Password; if your account has stricter
  org policies (Google Workspace with 2FA enforcement, etc.), check
  with your admin if sends start failing.

"""
Distributor Commission-Payout Invoice Generator
================================================

Run:
    pip install -r requirements.txt
    python app.py

Then open http://127.0.0.1:5000 , upload the distributor Excel sheet,
fill in the payout month and invoice date, and download a ZIP containing
one Word (.docx) TAX INVOICE per distributor - built from the sample
CP Payout format, with GST taken per-distributor from the sheet:

  - No GST No on file                -> no tax lines, invoice = Total
  - GST No NOT starting with '33'    -> the sheet's GST value is charged
                                         in full as IGST
  - GST No starting with '33'        -> the sheet's GST value is split
                                         in half between SGST and CGST

The Excel 'Total' column is the BASE (pre-tax) commission amount. The
tax amount is read directly from the 'GST' column (a rupee value) and
added on top: grand total = Total + GST.

Expected Excel columns (case-insensitive, order doesn't matter):
    Partner Code | Partner Name | Total | GST No | GST | Address | PAN |
    Bank Name | Bank Account Number | ifsc code | Email
('Address', 'PAN' and 'Email' are optional - if missing, those lines /
fields are simply omitted. An 'RM' / RM-name column, if present, is
ignored. 'Email' is only needed for the email-sending step.)

After generating a batch, use the "Email Config" tab to review each
distributor's email address, write a subject/body (with optional
{name} / {code} placeholders), add director BCCs, and send every
invoice out as an individually attached email.
"""

import io
import os
import shutil
import smtplib
import tempfile
import time
import traceback
import uuid
import zipfile
from datetime import datetime
from email.message import EmailMessage
from functools import wraps

import pandas as pd
from flask import Flask, request, render_template, send_file, flash, redirect, url_for, session

from invoice_engine import build_invoice_docx, safe_filename, compute_gst, is_blank_gst

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "change-me")

# Optional team password gate. Set the TEAM_PASSWORD environment variable
# on your host (e.g. Render -> Environment) to require a shared password
# before anyone can use the app. Leave it unset to allow open access.
TEAM_PASSWORD = os.environ.get("TEAM_PASSWORD")

# Outgoing email account. Set these as environment variables on your host
# (same way as TEAM_PASSWORD / SECRET_KEY) - never hardcode a real
# password in this file. EMAIL_APP_PASSWORD must be a Gmail "App
# Password" (Google Account -> Security -> App Passwords), not the
# normal account password.
EMAIL_ADDRESS = os.environ.get("EMAIL_ADDRESS", "info@myalternates.com")
EMAIL_APP_PASSWORD = os.environ.get("EMAIL_APP_PASSWORD")
SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "465"))

# In-memory store of generated invoice batches: batch_id -> {...}.
# NOTE: this lives in process memory. It's fine for a single dev/gunicorn
# worker (the default for a small internal tool). If you ever scale to
# multiple gunicorn workers, this needs to move to a shared store
# (e.g. a small SQLite file or Redis) since each worker has its own copy.
BATCHES = {}
BATCH_MAX_AGE_SECONDS = 6 * 60 * 60  # 6 hours


def _cleanup_old_batches():
    now = time.time()
    stale_ids = [bid for bid, b in BATCHES.items() if now - b["created"] > BATCH_MAX_AGE_SECONDS]
    for bid in stale_ids:
        shutil.rmtree(BATCHES[bid]["dir"], ignore_errors=True)
        BATCHES.pop(bid, None)


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if TEAM_PASSWORD and not session.get("authed"):
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)
    return wrapped


@app.route("/login", methods=["GET", "POST"])
def login():
    if not TEAM_PASSWORD:
        return redirect(url_for("index"))
    if request.method == "POST":
        if request.form.get("password") == TEAM_PASSWORD:
            session["authed"] = True
            return redirect(request.args.get("next") or url_for("index"))
        flash("Incorrect password.")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.pop("authed", None)
    return redirect(url_for("login"))

# Map of accepted (lower-cased, stripped) header names -> canonical field
COLUMN_ALIASES = {
    "partner code": "partner_code",
    "partner name": "partner_name",
    "total": "total",
    "gst no": "gst_no",
    "gst number": "gst_no",
    "gst": "gst_amount",       # rupee tax value (NOT the GST registration no.)
    "email": "email",
    "address": "address",
    "pan": "pan_no",
    "pan no": "pan_no",
    "bank name": "bank_name",
    "bank account number": "account_number",
    "account number": "account_number",
    "ifsc code": "ifsc",
    "ifsc": "ifsc",
}

REQUIRED_FIELDS = ["partner_code", "partner_name", "total", "bank_name", "account_number", "ifsc"]


def format_month_period(raw_value):
    """HTML <input type=month> gives 'YYYY-MM' -> 'Apr-2023'."""
    dt = datetime.strptime(raw_value, "%Y-%m")
    return dt.strftime("%b-%Y")


def format_invoice_date(raw_value):
    """HTML <input type=date> gives 'YYYY-MM-DD' -> '30-Apr-2023'."""
    dt = datetime.strptime(raw_value, "%Y-%m-%d")
    return dt.strftime("%d-%b-%Y")


def _clean_text_cell(df, i, col):
    if col not in df.columns:
        return ""
    val = str(df.iloc[i][col]).strip()
    return "" if val.lower() == "nan" else val


def _numeric_cell(df_numeric, i, col, default=0.0):
    if col not in df_numeric.columns:
        return default
    val = df_numeric.iloc[i][col]
    try:
        f = float(val)
        return default if pd.isna(f) else f
    except (TypeError, ValueError):
        return default


def load_distributors(file_storage):
    """Read the uploaded Excel file into a list of normalized dicts."""
    df = pd.read_excel(file_storage, dtype=str)  # read everything as text first
    # also grab a numeric-safe version for Total / GST amount columns
    file_storage.seek(0)
    df_numeric = pd.read_excel(file_storage)

    normalized_cols = {}
    for col in df.columns:
        key = str(col).strip().lower()
        if key in COLUMN_ALIASES:
            normalized_cols[col] = COLUMN_ALIASES[key]
    df = df.rename(columns=normalized_cols)
    df_numeric = df_numeric.rename(columns=normalized_cols)

    missing = [f for f in REQUIRED_FIELDS if f not in df.columns]
    if missing:
        raise ValueError(
            "The uploaded sheet is missing required column(s): "
            + ", ".join(missing)
            + ". Expected headers like: Partner Code, Partner Name, Total, GST No, GST, "
              "Bank Name, Bank Account Number, ifsc code."
        )

    rows = []
    for i in range(len(df)):
        rows.append({
            "partner_code": _clean_text_cell(df, i, "partner_code"),
            "partner_name": _clean_text_cell(df, i, "partner_name"),
            "total": _numeric_cell(df_numeric, i, "total"),
            "gst_no": df.iloc[i]["gst_no"] if "gst_no" in df.columns else None,
            "gst_amount": _numeric_cell(df_numeric, i, "gst_amount"),
            "email": _clean_text_cell(df, i, "email"),
            "address": _clean_text_cell(df, i, "address"),
            "pan_no": _clean_text_cell(df, i, "pan_no"),
            "bank_name": _clean_text_cell(df, i, "bank_name"),
            "account_number": _clean_text_cell(df, i, "account_number"),
            "ifsc": _clean_text_cell(df, i, "ifsc"),
        })
    return rows


@app.route("/", methods=["GET"])
@login_required
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
@login_required
def generate():
    excel_file = request.files.get("excel_file")
    month_period_raw = (request.form.get("month_period") or "").strip()
    invoice_date_raw = (request.form.get("invoice_date") or "").strip()

    if not excel_file or excel_file.filename == "":
        flash("Please choose an Excel file to upload.")
        return redirect(url_for("index"))
    if not month_period_raw:
        flash("Please select the payout month/period.")
        return redirect(url_for("index"))
    if not invoice_date_raw:
        flash("Please select the invoice date.")
        return redirect(url_for("index"))

    try:
        month_period = format_month_period(month_period_raw)
    except ValueError:
        flash("Payout month/period is not a valid month.")
        return redirect(url_for("index"))

    try:
        invoice_date = format_invoice_date(invoice_date_raw)
    except ValueError:
        flash("Invoice date is not a valid date.")
        return redirect(url_for("index"))

    try:
        rows = load_distributors(excel_file)
    except Exception as exc:
        flash(f"Could not read the Excel file: {exc}")
        return redirect(url_for("index"))

    if not rows:
        flash("No distributor rows found in the uploaded sheet.")
        return redirect(url_for("index"))

    _cleanup_old_batches()

    batch_id = uuid.uuid4().hex
    batch_dir = tempfile.mkdtemp(prefix=f"invoices_{batch_id}_")
    distributors_meta = []

    zip_buffer = io.BytesIO()
    summary_lines = ["Partner Code,Partner Name,Base,IGST,SGST,CGST,Total"]

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for row in rows:
            fname = f"{safe_filename(row['partner_code'])}_{safe_filename(month_period)}.docx"
            out_path = os.path.join(batch_dir, fname)
            try:
                gst = build_invoice_docx(
                    row=row,
                    month_period=month_period,
                    invoice_date=invoice_date,
                    ref_prefix=row["partner_code"],
                    out_path=out_path,
                )
            except Exception:
                traceback.print_exc()
                continue
            zf.write(out_path, arcname=fname)
            distributors_meta.append({
                "code": row["partner_code"],
                "name": row["partner_name"],
                "email": row.get("email", ""),
                "filename": fname,
            })
            summary_lines.append(
                f"{row['partner_code']},{row['partner_name']},"
                f"{gst['base']},{gst['igst']},{gst['sgst']},{gst['cgst']},{gst['total']}"
            )
        zf.writestr("summary.csv", "\n".join(summary_lines))

    zip_buffer.seek(0)

    if not distributors_meta:
        shutil.rmtree(batch_dir, ignore_errors=True)
        flash("No invoices could be generated from that sheet - check the file and try again.")
        return redirect(url_for("index"))

    BATCHES[batch_id] = {
        "dir": batch_dir,
        "month_period": month_period,
        "invoice_date": invoice_date,
        "distributors": distributors_meta,
        "created": time.time(),
    }
    session["batch_id"] = batch_id

    zip_name = f"Distributor_Invoices_{safe_filename(month_period)}.zip"
    return send_file(
        zip_buffer,
        as_attachment=True,
        download_name=zip_name,
        mimetype="application/zip",
    )


@app.route("/email_config", methods=["GET"])
@login_required
def email_config():
    batch_id = session.get("batch_id")
    batch = BATCHES.get(batch_id)
    if not batch:
        flash("Generate a batch of invoices first, then come back to Email Config.")
        return redirect(url_for("index"))
    return render_template(
        "email_config.html",
        batch=batch,
        results=None,
        subject="",
        body="",
        bcc="",
        email_configured=bool(EMAIL_APP_PASSWORD),
    )


def _send_invoice_email(to_email, bcc_list, subject, body, attachment_path, attachment_name):
    msg = EmailMessage()
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = to_email
    if bcc_list:
        msg["Bcc"] = ", ".join(bcc_list)
    msg["Subject"] = subject
    msg.set_content(body)

    with open(attachment_path, "rb") as f:
        data = f.read()
    msg.add_attachment(
        data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=attachment_name,
    )

    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_APP_PASSWORD)
        smtp.send_message(msg)


@app.route("/send_emails", methods=["POST"])
@login_required
def send_emails():
    batch_id = session.get("batch_id")
    batch = BATCHES.get(batch_id)
    if not batch:
        flash("No invoice batch found. Please generate invoices first.")
        return redirect(url_for("index"))

    subject_template = (request.form.get("subject") or "").strip()
    body_template = request.form.get("body") or ""
    bcc_raw = (request.form.get("bcc") or "").strip()
    bcc_list = [e.strip() for e in bcc_raw.split(",") if e.strip()]

    if not EMAIL_APP_PASSWORD:
        flash(
            "EMAIL_APP_PASSWORD is not set on the server, so emails can't be sent yet. "
            "Add it as an environment variable (see README) and try again."
        )
        return render_template(
            "email_config.html", batch=batch, results=None,
            subject=subject_template, body=body_template, bcc=bcc_raw,
            email_configured=False,
        )

    if not subject_template or not body_template:
        flash("Please fill in both the subject and body before sending.")
        return render_template(
            "email_config.html", batch=batch, results=None,
            subject=subject_template, body=body_template, bcc=bcc_raw,
            email_configured=True,
        )

    results = []
    for d in batch["distributors"]:
        field_name = f"email_{d['code']}"
        to_email = (request.form.get(field_name) or d.get("email") or "").strip()

        if not to_email:
            results.append((d["code"], d["name"], "Skipped - no email address"))
            continue

        subject = subject_template.replace("{name}", d["name"]).replace("{code}", d["code"])
        body = body_template.replace("{name}", d["name"]).replace("{code}", d["code"])
        attachment_path = os.path.join(batch["dir"], d["filename"])

        if not os.path.isfile(attachment_path):
            results.append((d["code"], d["name"], "Skipped - invoice file missing (regenerate batch)"))
            continue

        try:
            _send_invoice_email(
                to_email=to_email,
                bcc_list=bcc_list,
                subject=subject,
                body=body,
                attachment_path=attachment_path,
                attachment_name=d["filename"],
            )
            results.append((d["code"], d["name"], f"Sent to {to_email}"))
        except Exception as exc:
            results.append((d["code"], d["name"], f"FAILED - {exc}"))

    sent_count = sum(1 for _, _, r in results if r.startswith("Sent"))
    flash(f"Done: {sent_count} of {len(results)} email(s) sent.")

    return render_template(
        "email_config.html",
        batch=batch,
        results=results,
        subject=subject_template,
        body=body_template,
        bcc=bcc_raw,
        email_configured=True,
    )


if __name__ == "__main__":
    app.run(debug=True)


"""Teacher portal for Individual Student-wise Statistical Result Analysis."""

import os
import re
import qrcode
import sqlite3
from functools import wraps
from io import BytesIO




import pandas as pd
from dotenv import load_dotenv
from flask import (
    Flask, flash, redirect, render_template, request,
    send_file, session, url_for
)
from werkzeug.security import check_password_hash, generate_password_hash

# --------------------------------------------------
# Environment configuration
# --------------------------------------------------

BASE = os.path.dirname(os.path.abspath(__file__))

# Load environment variables from the project-root .env file
load_dotenv(os.path.join(BASE, ".env"))

# Import your backend after loading environment variables
import student_result_analysis as sra


# --------------------------------------------------
# Project paths and settings
# --------------------------------------------------

DB = os.path.join(BASE, "data.db")
PDF_DIR = os.path.join(BASE, "generated_pdfs")
CHART_DIR = os.path.join(BASE, "static", "charts")

os.makedirs(PDF_DIR, exist_ok=True)
os.makedirs(CHART_DIR, exist_ok=True)

app = Flask(__name__)

app.secret_key = os.environ.get("SECRET_KEY", "change-this-secret")

INSTITUTION = os.environ.get("INSTITUTION_NAME", "").strip()
BRAND = INSTITUTION or "Result Analysis Portal"

PASS_PERCENT = float(os.environ.get("PASS_PERCENT", 40))

# Configure the backend output directories and chart settings
sra.OUT_DIR = PDF_DIR
sra.CHART_DIR = CHART_DIR
sra.MAX_MARKS = int(os.environ.get("MAX_MARKS", 100))
sra.BIN_WIDTH = int(os.environ.get("BIN_WIDTH", 5))

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

IDS = ("Roll", "Name", "Email")


# --------------------------------------------------
# Database functions
# --------------------------------------------------

def db():
    return sqlite3.connect(DB)


def init_db():
    with db() as c:
        c.execute("""
            CREATE TABLE IF NOT EXISTS teachers(
                username TEXT PRIMARY KEY,
                pw TEXT
            )
        """)

        c.execute("""
            CREATE TABLE IF NOT EXISTS email_log(
                id INTEGER PRIMARY KEY,
                roll TEXT,
                name TEXT,
                sent_to TEXT,
                status TEXT,
                at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Create a default admin account only if no teacher exists
        if not c.execute("SELECT 1 FROM teachers").fetchone():
            c.execute(
                "INSERT INTO teachers VALUES(?, ?)",
                (
                    os.environ.get("ADMIN_USER", "admin"),
                    generate_password_hash(
                        os.environ.get("ADMIN_PASS", "admin123")
                    )
                )
            )


# --------------------------------------------------
# Data loading and helper functions
# --------------------------------------------------

def load():
    try:
        with db() as c:
            df = pd.read_sql("SELECT * FROM results", c)

    except Exception:
        return None, []

    subjects = [x for x in df.columns if x not in IDS]

    return df, subjects


def safe(t):
    return re.sub(r"[^A-Za-z0-9_-]+", "_", str(t)).strip("_")


def login_required(f):
    @wraps(f)
    def w(*a, **k):
        if "user" not in session:
            return redirect(url_for("login"))

        return f(*a, **k)

    return w


# --------------------------------------------------
# Student result calculations
# --------------------------------------------------

def enrich(df, subs):
    df = df.copy()

    df["Total"] = df[subs].sum(axis=1)

    df["Percent"] = (
        df["Total"] / (len(subs) * sra.MAX_MARKS) * 100
    ).round(2)

    df["Rank"] = (
        df["Total"]
        .rank(method="min", ascending=False)
        .astype(int)
    )

    pass_mark = sra.MAX_MARKS * PASS_PERCENT / 100

    df["Pass"] = (df[subs] >= pass_mark).all(axis=1)

    return df.sort_values("Rank")


# --------------------------------------------------
# Excel validation
# --------------------------------------------------

def validate(file):
    errs = []

    try:
        raw = pd.read_excel(file)

    except Exception as e:
        return None, [f"Could not read the file: {e}"]

    raw = raw.dropna(how="all")
    raw.columns = [str(c).strip() for c in raw.columns]

    lower = {c.lower(): c for c in raw.columns}

    # Check required columns
    for k in IDS:
        if k.lower() not in lower:
            errs.append(f"Missing column: {k}")

    if errs:
        return None, errs

    # Normalize required column names
    raw = raw.rename(
        columns={lower[k.lower()]: k for k in IDS}
    )

    # Identify subject columns
    subs = [c for c in raw.columns if c not in IDS]

    if not subs:
        return None, [
            "No subject columns found after Roll, Name, Email."
        ]

    raw = raw[list(IDS) + subs]

    raw["Roll"] = raw["Roll"].astype(str).str.strip()
    raw["Name"] = raw["Name"].astype(str).str.strip()
    raw["Email"] = raw["Email"].astype(str).str.strip()

    # Convert subject marks to numbers
    for s in subs:
        raw[s] = pd.to_numeric(raw[s], errors="coerce")

    # Check duplicate roll numbers
    for r in raw[raw["Roll"].duplicated()]["Roll"]:
        errs.append(f"Duplicate roll number: {r}")

    # Check email addresses (optional)
    if "Email" in raw.columns:
        emails = raw["Email"].fillna("").astype(str).str.strip()

        invalid_emails = raw[
            (emails != "") & (~emails.str.match(EMAIL_RE))
        ]

        for _, r in invalid_emails.iterrows():
            errs.append(
                f"Invalid email for roll {r['Roll']} ({r['Name']})"
            )

    # Check empty or invalid marks
    n_empty = int(raw[subs].isna().sum().sum())

    if n_empty:
        errs.append(
            f"{n_empty} marks cell(s) are empty or not numbers."
        )

    # Check marks outside the allowed range
    n_over = int(
        (
            (raw[subs] > sra.MAX_MARKS)
            | (raw[subs] < 0)
        ).sum().sum()
    )

    if n_over:
        errs.append(
            f"{n_over} mark(s) are outside 0-{sra.MAX_MARKS}."
        )

    return raw, errs


# --------------------------------------------------
# Template context
# --------------------------------------------------

@app.context_processor
def inject():
    return {
        "brand": BRAND,
        "user": session.get("user")
    }


# --------------------------------------------------
# Login and logout
# --------------------------------------------------

@app.route("/login", methods=["GET", "POST"])
def login():

    if request.method == "POST":

        with db() as c:
            row = c.execute(
                "SELECT pw FROM teachers WHERE username=?",
                (request.form["username"],)
            ).fetchone()

        if row and check_password_hash(
            row[0], request.form["password"]
        ):
            session["user"] = request.form["username"]

            return redirect(url_for("dashboard"))

        flash("Wrong username or password.", "error")

    return render_template("login.html")


@app.route("/logout")
def logout():

    session.clear()

    return redirect(url_for("login"))


# --------------------------------------------------
# Dashboard
# --------------------------------------------------

@app.route("/")
@login_required
def dashboard():

    df, subs = load()

    if df is None:
        return render_template("dashboard.html", empty=True)

    stats = sra.subject_stats(df, subs)

    e = enrich(df, subs)

    info = {
        "students": len(df),
        "subjects": len(subs),
        "avg": round(e["Percent"].mean(), 2),
        "pass_pct": round(e["Pass"].mean() * 100, 1),
        "topper": e.iloc[0]["Name"],
        "top_pct": e.iloc[0]["Percent"]
    }

    return render_template(
        "dashboard.html",
        empty=False,
        info=info,
        stats=stats,
        max_marks=sra.MAX_MARKS,
        pass_percent=PASS_PERCENT
    )


# --------------------------------------------------
# Excel upload
# --------------------------------------------------

@app.route("/upload", methods=["GET", "POST"])
@login_required
def upload():

    errors = []

    if request.method == "POST":

        f = request.files.get("file")

        if not f or not f.filename:
            errors = ["Please choose an Excel file."]

        else:
            df, errors = validate(f)

            if not errors:

                with db() as c:
                    df.to_sql(
                        "results",
                        c,
                        if_exists="replace",
                        index=False
                    )

                    c.execute("DELETE FROM email_log")

                # Remove old generated PDFs
                for p in os.listdir(PDF_DIR):
                    file_path = os.path.join(PDF_DIR, p)

                    if os.path.isfile(file_path):
                        os.remove(file_path)

                flash(
                    f"Uploaded {len(df)} students.",
                    "ok"
                )

                return redirect(url_for("dashboard"))

    return render_template("upload.html", errors=errors)


# --------------------------------------------------
# Students list
# --------------------------------------------------

@app.route("/students")
@login_required
def students():

    df, subs = load()

    if df is None:
        return redirect(url_for("upload"))

    e = enrich(df, subs)

    q = request.args.get("q", "").strip().lower()

    if q:
        e = e[
            e["Name"].str.lower().str.contains(q)
            | e["Roll"].str.lower().str.contains(q)
        ]

    return render_template(
        "students.html",
        rows=e.to_dict("records"),
        q=q
    )


# --------------------------------------------------
# Individual student page
# --------------------------------------------------

@app.route("/student/<roll>")
@login_required
def student(roll):

    df, subs = load()

    if df is None or roll not in set(df["Roll"]):
        return redirect(url_for("students"))

    stats = sra.subject_stats(df, subs)

    row = df[df["Roll"] == roll].iloc[0]

    e = enrich(df, subs).set_index("Roll").loc[roll]

    charts = []

    for s in subs:

        fn = f"{safe(roll)}_{safe(s)}.png"

        sra.make_chart(
            df,
            s,
            row[s],
            row["Name"],
            stats[s],
            os.path.join(CHART_DIR, fn)
        )

        charts.append((s, fn))

    table = [
        (s, row[s], stats[s])
        for s in subs
    ]

    return render_template(
        "student.html",
        row=row,
        e=e,
        table=table,
        charts=charts
    )


# --------------------------------------------------
# PDF generation
# --------------------------------------------------

def pdf_for(df, subs, stats, row):

    p = os.path.join(
        PDF_DIR,
        f"{safe(row['Roll'])}.pdf"
    )

    sra.build_pdf(df, subs, stats, row, p)

    return p


@app.route("/pdf/<roll>")
@login_required
def pdf(roll):

    df, subs = load()

    if df is None or roll not in set(df["Roll"]):
        return redirect(url_for("students"))

    row = df[df["Roll"] == roll].iloc[0]

    report_path = pdf_for(
        df,
        subs,
        sra.subject_stats(df, subs),
        row
    )

    return send_file(
        report_path,
        as_attachment=True,
        download_name=f"{safe(row['Name'])}_report.pdf"
    )


# --------------------------------------------------
# Reports and email sending through Brevo SMTP
# --------------------------------------------------

@app.route("/reports", methods=["GET", "POST"])
@login_required
def reports():

    df, subs = load()

    if df is None:
        return redirect(url_for("upload"))

    if request.method == "POST":

        stats = sra.subject_stats(df, subs)

        action = request.form.get("action", "").strip()
        test_to = request.form.get("test_to", "").strip()

        # Brevo SMTP credentials from .env
        smtp_login = os.getenv("BREVO_SMTP_LOGIN")
        smtp_key = os.getenv("BREVO_SMTP_KEY")
        mail_from = os.getenv("MAIL_FROM")

        # Validate email configuration before sending
        if action == "send" and not (
            smtp_login and smtp_key and mail_from
        ):
            flash(
                "Brevo SMTP credentials are not configured. "
                "Please check your .env file.",
                "error"
            )
            return redirect(url_for("reports"))

        sent = 0
        failed = 0

        # Generate PDFs for all students.
        # If test_to is provided, send only the first student's
        # report to that test address.
        for index, (_, row) in enumerate(df.iterrows()):

            p = pdf_for(df, subs, stats, row)

            if action != "send":
                continue

            if test_to:
                # Test mode: send only one email
                if index > 0:
                    continue

                to = test_to

            else:
                # Normal mode: send to each student's email
                to = str(row["Email"]).strip()

            try:

                sra.send_mail(
                    to,
                    row["Name"],
                    p,
                    smtp_login,
                    smtp_key,
                    from_addr=mail_from,
                    body=(
                        f"Dear {row['Name']},\n\n"
                        "Please find attached your subject-wise "
                        "statistical result analysis.\n\n"
                        f"Regards,\n"
                        f"{INSTITUTION or 'Results Team'}"
                    )
                )

                status = "sent"
                sent += 1

            except Exception as ex:

                status = f"failed: {ex}"
                failed += 1

            # Save email status in database
            with db() as c:
                c.execute(
                    """
                    INSERT INTO email_log
                    (roll, name, sent_to, status)
                    VALUES (?, ?, ?, ?)
                    """,
                    (
                        row["Roll"],
                        row["Name"],
                        to,
                        status
                    )
                )

        if action != "send":

            flash(
                f"PDFs ready for {len(df)} students.",
                "ok"
            )

        else:

            if test_to:
                message = (
                    f"Test email sent: {sent}, "
                    f"failed: {failed}."
                )
            else:
                message = (
                    f"Emails sent: {sent}, "
                    f"failed: {failed}."
                )

            flash(
                message,
                "ok" if failed == 0 else "error"
            )

        return redirect(url_for("reports"))

    # Fetch recent email history
    with db() as c:

        log = c.execute(
            """
            SELECT name, sent_to, status, at
            FROM email_log
            ORDER BY id DESC
            LIMIT 100
            """
        ).fetchall()

    return render_template(
        "reports.html",
        log=log,
        n=len(df)
    )


# --------------------------------------------------
# QR code and sharing
# --------------------------------------------------

@app.route("/share")
def share():

    return render_template(
        "share.html",
        url=os.environ.get("APP_URL", request.host_url)
    )


@app.route("/qr.png")
def qr():

    import qrcode

    img = qrcode.make(
        os.environ.get("APP_URL", request.host_url),
        box_size=10,
        border=2
    )

    buf = BytesIO()

    img.save(buf, "PNG")
    buf.seek(0)

    return send_file(
        buf,
        mimetype="image/png"
    )


# --------------------------------------------------
# Application startup
# --------------------------------------------------

init_db()

if __name__ == "__main__":

    app.run(
        debug=True,
        host="0.0.0.0",
        port=5000
    )


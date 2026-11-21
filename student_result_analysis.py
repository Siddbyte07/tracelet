"""
Individual Student-wise Statistical Result Analysis
---------------------------------------------------
1. Read an Excel file (Roll, Name, Email, <subject columns...>)
2. For every subject compute Mean, Median, Variance (and Std Dev)
3. Draw a vertical bar chart (marks vs number of students) per subject
   and mark the individual student's marks as a point on that chart
4. Build one PDF per student
5. Email each PDF to the student's mail id

Usage:
    python student_result_analysis.py students.xlsx            # only makes PDFs
    python student_result_analysis.py students.xlsx --send     # makes PDFs + emails them

For email set these environment variables first (Gmail example):
    EMAIL_USER = yourname@gmail.com
    EMAIL_PASS = <16-char Gmail App Password>   (NOT your normal password)
"""

import argparse
import os
import smtplib
import ssl
import sys
from email.message import EmailMessage

import matplotlib
matplotlib.use("Agg")  # no GUI needed
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (Image, PageBreak, Paragraph, SimpleDocTemplate,
                                Spacer, Table, TableStyle)

# ----------------------------- CONFIG ---------------------------------
ID_COLS = {"roll": "Roll", "name": "Name", "email": "Email"}  # column names in Excel
BIN_WIDTH = 5            # width of each vertical bar (marks range). Use 1 for exact marks
MAX_MARKS = 100          # full marks of each subject
OUT_DIR = "output_pdfs"
CHART_DIR = os.path.join(OUT_DIR, "_charts")
# -----------------------------------------------------------------------


def load_data(path):
    df = pd.read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in ID_COLS.values() if c not in df.columns]
    if missing:
        sys.exit(f"Excel is missing column(s): {missing}. Found: {list(df.columns)}")
    subjects = [c for c in df.columns if c not in ID_COLS.values()]
    for s in subjects:
        df[s] = pd.to_numeric(df[s], errors="coerce")
    return df, subjects


def subject_stats(df, subjects, ddof=0):
    stats = {}
    for s in subjects:
        col = df[s].dropna()
        stats[s] = {
            "Mean": col.mean(),
            "Median": col.median(),
            "Variance": col.var(ddof=ddof),
            "Std Dev": col.std(ddof=ddof),
            "Highest": col.max(),
            "Lowest": col.min(),
        }
    return stats


def make_chart(df, subject, student_marks, student_name, st, path):
    """Vertical bar chart of class marks distribution + the student's point."""
    data = df[subject].dropna()
    edges = np.arange(0, MAX_MARKS + BIN_WIDTH, BIN_WIDTH)
    counts, _ = np.histogram(data, bins=edges)
    centers = edges[:-1] + BIN_WIDTH / 2

    fig, ax = plt.subplots(figsize=(7, 3.6))
    ax.bar(centers, counts, width=BIN_WIDTH * 0.9, color="#9ecae1",
           edgecolor="#3182bd", label="No. of students")

    # student's point -> sits on top of the bar where his/her marks fall
    if pd.notna(student_marks):
        idx = min(int(student_marks // BIN_WIDTH), len(counts) - 1)
        ax.plot(student_marks, counts[idx], "o", color="red", markersize=10,
                zorder=5, label=f"{student_name}: {student_marks:g}")
        ax.annotate(f"{student_marks:g}", (student_marks, counts[idx]),
                    textcoords="offset points", xytext=(0, 10),
                    ha="center", color="red", fontweight="bold")

    ax.axvline(st["Mean"], color="green", linestyle="--", label=f"Mean {st['Mean']:.1f}")
    ax.axvline(st["Median"], color="orange", linestyle=":", label=f"Median {st['Median']:.1f}")

    ax.set_title(f"{subject} - Class Marks Distribution")
    ax.set_xlabel("Marks")
    ax.set_ylabel("Number of students")
    ax.set_xlim(0, MAX_MARKS)
    ax.set_xticks(edges)
    ax.yaxis.get_major_locator().set_params(integer=True)
    ax.legend(fontsize=8)
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


def build_pdf(df, subjects, stats, row, pdf_path):
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(pdf_path, pagesize=A4, leftMargin=1.5 * cm,
                            rightMargin=1.5 * cm, topMargin=1.5 * cm, bottomMargin=1.5 * cm)
    story = []
    if os.environ.get("INSTITUTION_NAME"):
        story.append(Paragraph(os.environ["INSTITUTION_NAME"], styles["Heading3"]))
    story += [
        Paragraph("Individual Student Result Analysis", styles["Title"]),
        Paragraph(f"<b>Name:</b> {row[ID_COLS['name']]} &nbsp;&nbsp; "
                  f"<b>Roll:</b> {row[ID_COLS['roll']]}", styles["Normal"]),
        Spacer(1, 12),
    ]

    # summary table
    table_data = [["Subject", "Your Marks", "Class Mean", "Median", "Variance", "Std Dev", "vs Mean"]]
    for s in subjects:
        m, st = row[s], stats[s]
        diff = "-" if pd.isna(m) else f"{m - st['Mean']:+.1f}"
        table_data.append([s, "-" if pd.isna(m) else f"{m:g}", f"{st['Mean']:.2f}",
                           f"{st['Median']:.2f}", f"{st['Variance']:.2f}",
                           f"{st['Std Dev']:.2f}", diff])
    t = Table(table_data, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#3182bd")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.white]),
    ]))
    story += [t, Spacer(1, 14)]

    # one chart per subject (2 per page)
    roll = str(row[ID_COLS["roll"]]).replace("/", "_")
    for i, s in enumerate(subjects):
        chart_path = os.path.join(CHART_DIR, f"{roll}_{s}.png".replace(" ", "_"))
        make_chart(df, s, row[s], row[ID_COLS["name"]], stats[s], chart_path)
        story.append(Image(chart_path, width=16 * cm, height=8.2 * cm))
        story.append(Spacer(1, 8))
        if (i + 1) % 2 == 0 and i + 1 < len(subjects):
            story.append(PageBreak())
    doc.build(story)


def send_mail(to_addr, name, pdf_path, user, password,
              from_addr=None,
              subject="Your Individual Result Analysis",
              body=None):

    msg = EmailMessage()

    msg["Subject"] = subject

    # Actual verified sender email address
    msg["From"] = from_addr or user

    msg["To"] = to_addr

    msg.set_content(
        body or (
            f"Dear {name},\n\n"
            "Please find attached your subject-wise "
            "statistical result analysis.\n\n"
            "Regards,\n"
            "Results Team"
        )
    )

    # Attach individual student PDF
    with open(pdf_path, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="pdf",
            filename=os.path.basename(pdf_path)
        )

    # Brevo SMTP connection using STARTTLS
    with smtplib.SMTP(
        "smtp-relay.brevo.com", 587
    ) as s:

        s.starttls(context=ssl.create_default_context())

        s.login(user, password)

        s.send_message(msg)




def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("excel")
    ap.add_argument("--send", action="store_true", help="email the PDFs")
    args = ap.parse_args()

    os.makedirs(CHART_DIR, exist_ok=True)
    df, subjects = load_data(args.excel)
    stats = subject_stats(df, subjects)

    user = os.environ.get("EMAIL_USER")
    password = os.environ.get("EMAIL_PASS")
    if args.send and not (user and password):
        sys.exit("Set EMAIL_USER and EMAIL_PASS environment variables first.")

    for _, row in df.iterrows():
        roll = str(row[ID_COLS["roll"]]).replace("/", "_")
        pdf_path = os.path.join(OUT_DIR, f"{roll}_{row[ID_COLS['name']]}.pdf".replace(" ", "_"))
        build_pdf(df, subjects, stats, row, pdf_path)
        print(f"PDF created: {pdf_path}")
        if args.send:
            try:
                send_mail(row[ID_COLS["email"]], row[ID_COLS["name"]], pdf_path, user, password)
                print(f"  mailed -> {row[ID_COLS['email']]}")
            except Exception as e:
                print(f"  FAILED for {row[ID_COLS['email']]}: {e}")


if __name__ == "__main__":
    main()

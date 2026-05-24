"""
Automated Employee Report Generator
=====================================
Project  : Automated Report Generation
Author   : PAT - CIS Automation | Cognizant
Purpose  : Reads employee Excel data, cleans it, and generates a
           formatted summary report saved as report.xlsx
Run      : python generate_report.py
Schedule : Windows Task Scheduler (runs automatically on set time)
"""

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import os

# ─────────────────────────────────────────────
# CONFIGURATION  — change only these paths
# ─────────────────────────────────────────────
INPUT_FILE = "employees.xlsx"  # Source Excel file
OUTPUT_FILE = "Employee_Report.xlsx"  # Generated report file
REPORT_DATE = datetime.today().strftime("%d-%b-%Y")

# ─────────────────────────────────────────────
# STEP 1 — READ DATA
# ─────────────────────────────────────────────
print(f"[{REPORT_DATE}] Reading data from '{INPUT_FILE}' ...")

df = pd.read_excel(INPUT_FILE, dtype={"EMPLOYEE_ID": str})

# ─────────────────────────────────────────────
# STEP 2 — CLEAN DATA
# ─────────────────────────────────────────────
print("Cleaning data ...")

df.columns = df.columns.str.strip().str.upper()

df["FIRST_NAME"] = df["FIRST_NAME"].str.strip().str.title()
df["LAST_NAME"] = df["LAST_NAME"].str.strip().str.title()
df["FULL_NAME"] = df["FIRST_NAME"] + " " + df["LAST_NAME"]
df["EMAIL"] = df["EMAIL"].str.strip().str.upper()
df["JOB_ID"] = df["JOB_ID"].str.strip().str.upper()

df["SALARY"] = pd.to_numeric(df["SALARY"], errors="coerce").fillna(0)
df["COMMISSION_PCT"] = pd.to_numeric(df["COMMISSION_PCT"], errors="coerce").fillna(0)

df["HIRE_DATE"] = pd.to_datetime(df["HIRE_DATE"], errors="coerce")
df["YEARS_OF_SERVICE"] = ((datetime.today() - df["HIRE_DATE"]).dt.days / 365).round(1)

df["DEPARTMENT_NAME"] = df["DEPARTMENT_ID"].map({
    10: "Administration", 20: "Marketing", 30: "Purchasing",
    40: "Human Resources", 50: "Shipping", 60: "IT",
    70: "Public Relations", 80: "Sales", 90: "Executive",
    100: "Finance", 110: "Accounting"
}).fillna("Unknown")

df.drop_duplicates(subset=["EMPLOYEE_ID"], inplace=True)
df.dropna(subset=["EMPLOYEE_ID"], inplace=True)

print(f"  Total employees after cleaning: {len(df)}")

# ─────────────────────────────────────────────
# STEP 3 — SUMMARY CALCULATIONS
# ─────────────────────────────────────────────
total_employees = len(df)
total_salary = df["SALARY"].sum()
avg_salary = df["SALARY"].mean()
max_salary = df["SALARY"].max()
min_salary = df["SALARY"].min()

dept_summary = (
    df.groupby("DEPARTMENT_NAME")
    .agg(
        Employee_Count=("EMPLOYEE_ID", "count"),
        Total_Salary=("SALARY", "sum"),
        Avg_Salary=("SALARY", "mean"),
        Max_Salary=("SALARY", "max"),
        Min_Salary=("SALARY", "min"),
    )
    .reset_index()
    .sort_values("Employee_Count", ascending=False)
)

job_summary = (
    df.groupby("JOB_ID")
    .agg(
        Count=("EMPLOYEE_ID", "count"),
        Avg_Salary=("SALARY", "mean"),
    )
    .reset_index()
    .sort_values("Count", ascending=False)
)

top_earners = df.nlargest(10, "SALARY")[
    ["FULL_NAME", "JOB_ID", "DEPARTMENT_NAME", "SALARY", "YEARS_OF_SERVICE"]
].reset_index(drop=True)


# ─────────────────────────────────────────────
# HELPER — Style utilities
# ─────────────────────────────────────────────
def header_style(cell, bg="1F4E79", fg="FFFFFF", size=11, bold=True):
    cell.font = Font(bold=bold, color=fg, size=size, name="Arial")
    cell.fill = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def title_style(cell, size=14):
    cell.font = Font(bold=True, size=size, color="1F4E79", name="Arial")
    cell.alignment = Alignment(horizontal="left", vertical="center")


def sub_title_style(cell):
    cell.font = Font(size=10, color="595959", name="Arial")
    cell.alignment = Alignment(horizontal="left")


def data_style(cell, number_fmt=None, bold=False, center=False):
    cell.font = Font(size=10, name="Arial", bold=bold)
    cell.alignment = Alignment(
        horizontal="center" if center else "left",
        vertical="center"
    )
    if number_fmt:
        cell.number_format = number_fmt


def thin_border():
    s = Side(style="thin", color="D3D3D3")
    return Border(left=s, right=s, top=s, bottom=s)


def set_col_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def write_df_to_sheet(ws, df_in, start_row, headers, col_fmts=None):
    """Write a dataframe with styled headers and bordered data rows."""
    bg_alt = "EBF3FB"
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=start_row, column=ci, value=h)
        header_style(c, bg="2E75B6")
        c.border = thin_border()
    for ri, row in enumerate(df_in.itertuples(index=False), 1):
        bg = bg_alt if ri % 2 == 0 else "FFFFFF"
        for ci, val in enumerate(row, 1):
            c = ws.cell(row=start_row + ri, column=ci, value=val)
            c.fill = PatternFill("solid", start_color=bg)
            c.border = thin_border()
            fmt = (col_fmts or {}).get(ci)
            data_style(c, number_fmt=fmt, center=(fmt is not None))
    return start_row + len(df_in) + 1


# ─────────────────────────────────────────────
# STEP 4 — BUILD WORKBOOK
# ─────────────────────────────────────────────
print("Building report workbook ...")

wb = Workbook()

# ── SHEET 1: Dashboard ──────────────────────
ws1 = wb.active
ws1.title = "Dashboard"
ws1.sheet_view.showGridLines = False
ws1.row_dimensions[1].height = 45
ws1.row_dimensions[2].height = 20

# Title banner
ws1.merge_cells("A1:F1")
c = ws1["A1"]
c.value = "EMPLOYEE SUMMARY REPORT"
c.font = Font(bold=True, size=18, color="FFFFFF", name="Arial")
c.fill = PatternFill("solid", start_color="1F4E79")
c.alignment = Alignment(horizontal="center", vertical="center")

ws1.merge_cells("A2:F2")
c = ws1["A2"]
c.value = f"Generated on: {REPORT_DATE}  |  Source: {INPUT_FILE}"
c.font = Font(size=10, color="FFFFFF", name="Arial")
c.fill = PatternFill("solid", start_color="2E75B6")
c.alignment = Alignment(horizontal="center", vertical="center")

# KPI Cards
ws1.row_dimensions[4].height = 18
ws1["A4"].value = "KEY METRICS"
title_style(ws1["A4"], size=12)

kpis = [
    ("A", "Total Employees", total_employees, None),
    ("B", "Total Salary", total_salary, '"$"#,##0'),
    ("C", "Average Salary", avg_salary, '"$"#,##0'),
    ("D", "Highest Salary", max_salary, '"$"#,##0'),
    ("E", "Lowest Salary", min_salary, '"$"#,##0'),
]
ws1.row_dimensions[5].height = 18
ws1.row_dimensions[6].height = 30

for col, label, value, fmt in kpis:
    lc = ws1[f"{col}5"]
    lc.value = label
    header_style(lc, bg="2E75B6", size=10)
    lc.border = thin_border()

    vc = ws1[f"{col}6"]
    vc.value = value
    vc.font = Font(bold=True, size=13, color="1F4E79", name="Arial")
    vc.fill = PatternFill("solid", start_color="DEEAF1")
    vc.alignment = Alignment(horizontal="center", vertical="center")
    vc.border = thin_border()
    if fmt:
        vc.number_format = fmt

# Dept summary table on dashboard
ws1["A8"].value = "DEPARTMENT BREAKDOWN"
title_style(ws1["A8"], size=12)

dept_headers = ["Department", "Employees", "Total Salary ($)", "Avg Salary ($)", "Max Salary ($)", "Min Salary ($)"]
dept_display = dept_summary.copy()
dept_display.columns = dept_headers

write_df_to_sheet(ws1, dept_display, start_row=9, headers=dept_headers,
                  col_fmts={2: "#,##0", 3: "#,##0", 4: "#,##0", 5: "#,##0", 6: "#,##0"})

set_col_widths(ws1, {"A": 22, "B": 12, "C": 18, "D": 16, "E": 16, "F": 16})

# ── SHEET 2: Employee Details ────────────────
ws2 = wb.create_sheet("Employee Details")
ws2.sheet_view.showGridLines = False
ws2.row_dimensions[1].height = 35
ws2.row_dimensions[2].height = 18

ws2.merge_cells("A1:H1")
c = ws2["A1"]
c.value = "EMPLOYEE DETAILS"
c.font = Font(bold=True, size=14, color="FFFFFF", name="Arial")
c.fill = PatternFill("solid", start_color="1F4E79")
c.alignment = Alignment(horizontal="center", vertical="center")

ws2.merge_cells("A2:H2")
c = ws2["A2"]
c.value = f"Report Date: {REPORT_DATE}"
c.font = Font(size=10, color="595959", name="Arial")
c.alignment = Alignment(horizontal="right")

emp_cols = ["EMPLOYEE_ID", "FULL_NAME", "JOB_ID", "DEPARTMENT_NAME", "SALARY", "HIRE_DATE", "YEARS_OF_SERVICE",
            "COMMISSION_PCT"]
emp_heads = ["Emp ID", "Full Name", "Job Role", "Department", "Salary ($)", "Hire Date", "Years", "Commission"]
emp_display = df[emp_cols].copy()
emp_display["HIRE_DATE"] = emp_display["HIRE_DATE"].dt.strftime("%d-%b-%Y")

write_df_to_sheet(ws2, emp_display, start_row=4, headers=emp_heads,
                  col_fmts={5: "#,##0", 7: "0.0", 8: "0.00%"})

set_col_widths(ws2, {"A": 10, "B": 22, "C": 14, "D": 20, "E": 14, "F": 14, "G": 8, "H": 12})

# ── SHEET 3: Top Earners ─────────────────────
ws3 = wb.create_sheet("Top 10 Earners")
ws3.sheet_view.showGridLines = False
ws3.row_dimensions[1].height = 35

ws3.merge_cells("A1:E1")
c = ws3["A1"]
c.value = "TOP 10 HIGHEST PAID EMPLOYEES"
c.font = Font(bold=True, size=14, color="FFFFFF", name="Arial")
c.fill = PatternFill("solid", start_color="833C00")
c.alignment = Alignment(horizontal="center", vertical="center")

top_heads = ["Full Name", "Job Role", "Department", "Salary ($)", "Years of Service"]
top_display = top_earners.copy()
top_display.columns = top_heads

write_df_to_sheet(ws3, top_display, start_row=3, headers=top_heads,
                  col_fmts={4: "#,##0", 5: "0.0"})

set_col_widths(ws3, {"A": 24, "B": 16, "C": 20, "D": 14, "E": 18})

# ── SHEET 4: Job Summary ─────────────────────
ws4 = wb.create_sheet("Job Summary")
ws4.sheet_view.showGridLines = False
ws4.row_dimensions[1].height = 35

ws4.merge_cells("A1:C1")
c = ws4["A1"]
c.value = "JOB ROLE SUMMARY"
c.font = Font(bold=True, size=14, color="FFFFFF", name="Arial")
c.fill = PatternFill("solid", start_color="375623")
c.alignment = Alignment(horizontal="center", vertical="center")

job_heads = ["Job ID", "Employee Count", "Avg Salary ($)"]
job_display = job_summary.copy()
job_display["Avg_Salary"] = job_display["Avg_Salary"].round(0)
job_display.columns = job_heads

write_df_to_sheet(ws4, job_display, start_row=3, headers=job_heads,
                  col_fmts={2: "#,##0", 3: "#,##0"})

set_col_widths(ws4, {"A": 18, "B": 18, "C": 18})

# ─────────────────────────────────────────────
# STEP 5 — SAVE REPORT
# ─────────────────────────────────────────────
wb.save(OUTPUT_FILE)
print(f"\n✅ Report saved successfully → '{OUTPUT_FILE}'")
print(f"   Sheets  : Dashboard | Employee Details | Top 10 Earners | Job Summary")
print(f"   Records : {total_employees} employees")
print(f"   Date    : {REPORT_DATE}")
print(f"\n   Next step: Power Automate will pick up '{OUTPUT_FILE}' and email it automatically.")
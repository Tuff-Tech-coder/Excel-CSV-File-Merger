"""Creates 5 deliberately messy sample input files for the Excel Merger demo."""
import pandas as pd
from pathlib import Path

OUT = Path(__file__).parent / "sample_input"

# ── File 1: sales_q1.xlsx ─────────────────────────────────────────────────
# Standard sales data, some mixed-case headers
df1 = pd.DataFrame({
    "Name": ["Alice Johnson", "Bob Smith", "Carol White", "Dave Brown", "Eve Davis"],
    "Email": ["alice@corp.com", "bob@corp.com", "carol@corp.com", "dave@corp.com", "eve@corp.com"],
    "Revenue": [12500.00, 8750.50, 23400.00, 5600.75, 19800.00],
    "Region": ["North", "South", "North", "West", "East"],
    "Date": ["2024-01-15", "2024-01-20", "2024-02-03", "2024-02-14", "2024-03-01"],
    "Department": ["Sales", "Sales", "Sales", "Sales", "Sales"],
})
df1.to_excel(OUT / "sales_q1.xlsx", index=False)

# ── File 2: sales_q2.xlsx ─────────────────────────────────────────────────
# Same data type but different column ORDER and some EXTRA columns
df2 = pd.DataFrame({
    "email": ["frank@corp.com", "grace@corp.com", "henry@corp.com",    # lowercase headers
              "irene@corp.com", "jake@corp.com", "alice@corp.com"],     # alice is a duplicate
    "REVENUE": [7200.00, 15300.00, 9800.00, 22100.00, 6400.00, 12500.00],  # uppercase header
    "Region": ["South", "East", "North", "West", "South", "North"],
    "name": ["Frank Lee", "Grace Kim", "Henry Chen", "Irene Park", "Jake Wu", "Alice Johnson"],
    "Date": ["2024-04-10", "2024-04-22", "2024-05-08", "2024-05-19", "2024-06-01", "2024-01-15"],
    "Notes": ["Top performer", "", "Needs review", "Excellent", "", "Duplicate entry"],  # extra col
    "Department": ["Sales", "Sales", "Sales", "Sales", "Sales", "Sales"],
})
df2.to_excel(OUT / "sales_q2.xlsx", index=False)

# ── File 3: hr_employees.csv ──────────────────────────────────────────────
# CSV format, missing Revenue and Date columns, extra HR-specific columns
df3 = pd.DataFrame({
    "Name": ["Laura Moss", "Mike Stone", "Nancy Hill", "Oscar Reed", "Paula Marsh",
             "Quinn Ford", "Rachel Moore"],
    "Email": ["laura@corp.com", "mike@corp.com", "nancy@corp.com", "oscar@corp.com",
              "paula@corp.com", "quinn@corp.com", "rachel@corp.com"],
    "Department": ["HR", "HR", "Engineering", "Engineering", "Marketing", "Marketing", "HR"],
    "Region": ["East", "West", "North", "South", "East", "West", "North"],
    "Hire Date": ["2021-03-15", "2019-07-01", "2022-11-20", "2020-04-05",
                  "2023-01-10", "2018-09-30", "2024-02-15"],   # different date column name
    "Salary": [55000, 72000, 98000, 115000, 61000, 68000, 52000],  # different money column
})
df3.to_csv(OUT / "hr_employees.csv", index=False)

# ── File 4: marketing_leads.csv ───────────────────────────────────────────
# CSV with inconsistent formatting: extra whitespace, mixed number formats
import csv
rows = [
    ["  Name  ", "Email", "Revenue", " Region ", "Date", "Department"],
    ["Sam Turner ", "sam@leads.com", "$5,400.00", "  North  ", "2024/01/08", "Marketing"],
    [" Tina Brooks", "tina@leads.com", "8200", "East", "2024-02-14", "Marketing"],
    ["Uma Patel  ", "uma@leads.com", "$11,750.50", "West  ", "March 5, 2024", "Marketing"],
    ["Vince Hall", "vince@leads.com", "3100.0", "South", "2024-04-22", "Marketing"],
    ["  Wendy Fox", "wendy@leads.com", "$19,900", "North", "2024-05-30", "Marketing"],
    ["Xavier Long", "xavier@leads.com", "7650.25", " East", "2024-06-15", "Marketing"],
]
with open(OUT / "marketing_leads.csv", "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerows(rows)

# ── File 5: ops_data.xlsx ─────────────────────────────────────────────────
# Different column names, some blank rows, purely numeric revenue (no formatting)
df5 = pd.DataFrame({
    "Full Name": ["Yara Singh", "Zach Adams", None, "Amy Clarke", "Brian Duke"],  # blank row
    "Contact Email": ["yara@ops.com", "zach@ops.com", None, "amy@ops.com", "brian@ops.com"],
    "Sales Amount": [14200.00, 9900.50, None, 31000.00, 7800.25],  # renamed revenue col
    "Territory": ["West", "North", None, "South", "East"],          # renamed region col
    "Transaction Date": ["2024-03-18", "2024-04-02", None, "2024-05-11", "2024-06-29"],
    "Team": ["Operations", "Operations", None, "Operations", "Operations"],
})
df5.to_excel(OUT / "ops_data.xlsx", index=False)

print("✓ Created 5 sample input files in sample_input/")

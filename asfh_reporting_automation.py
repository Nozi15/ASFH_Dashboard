
import pandas as pd
from datetime import datetime

# === 1. Load Data ===
file_path = "ASFH_Dashboard_Template_2025.xlsx"  # Replace with the path to your dashboard file
df = pd.read_excel(file_path)

# === 2. Create Performance Summary ===
summary = []
for index, row in df.iterrows():
    indicator = row["Indicator"]
    target = row["Target (2025)"]
    current = row["Current Status"]
    variance = row["Variance"]
    rag = row["Status RAG"]
    
    # Generate a line of narrative
    if rag == "Green":
        status_text = "on track"
    elif rag == "Amber":
        status_text = "moderately behind"
    else:
        status_text = "significantly behind"
    
    summary.append(
        f"Indicator '{indicator}' is currently {status_text} with a value of {current} "
        f"against a target of {target} (variance: {variance})."
    )

# === 3. Save Narrative Summary to Text File ===
report_date = datetime.now().strftime("%Y-%m-%d")
narrative_path = f"ASFH_Report_Summary_{report_date}.txt"
with open(narrative_path, "w") as f:
    f.write("ASFH Performance Summary Report\n")
    f.write(f"Generated on: {report_date}\n\n")
    for line in summary:
        f.write(line + "\n")

# === 4. Export Updated Report to Excel (Optional) ===
output_excel_path = f"ASFH_KPI_Report_{report_date}.xlsx"
df["Report Date"] = report_date
df.to_excel(output_excel_path, index=False)

print(f"Report summary saved as: {narrative_path}")
print(f"Updated KPI Excel report saved as: {output_excel_path}")

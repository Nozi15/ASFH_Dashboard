
import streamlit as st
import pandas as pd
from datetime import datetime

st.title("ASFH KPI Reporting Automation")
st.write("Upload your KPI dashboard Excel file to generate a narrative summary and updated Excel report.")

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file:
    # Read the Excel file
    df = pd.read_excel(uploaded_file)

    # Display the data
    st.subheader("KPI Dashboard Preview")
    st.dataframe(df)

    # Generate the narrative summary
    st.subheader("Narrative Summary")
    summary = []
    for index, row in df.iterrows():
        indicator = row["Indicator"]
        target = row["Target (2025)"]
        current = row["Current Status"]
        variance = row["Variance"]
        rag = row["Status RAG"]

        if rag == "Green":
            status_text = "on track"
        elif rag == "Amber":
            status_text = "moderately behind"
        else:
            status_text = "significantly behind"

        summary_line = (
            f"Indicator '{indicator}' is currently {status_text} with a value of {current} "
            f"against a target of {target} (variance: {variance})."
        )
        summary.append(summary_line)
        st.write(summary_line)

    # Add report date
    report_date = datetime.now().strftime("%Y-%m-%d")
    df["Report Date"] = report_date

    # Save files
    output_excel = f"ASFH_KPI_Report_{report_date}.xlsx"
    output_txt = f"ASFH_Report_Summary_{report_date}.txt"

    df.to_excel(output_excel, index=False)
    with open(output_txt, "w") as f:
        f.write("ASFH Performance Summary Report\n")
        f.write(f"Generated on: {report_date}\n\n")
        for line in summary:
            f.write(line + "\n")

    # Provide download links
    with open(output_excel, "rb") as file:
        st.download_button("Download Updated KPI Excel Report", file, file_name=output_excel)

    with open(output_txt, "rb") as file:
        st.download_button("Download Narrative Summary (TXT)", file, file_name=output_txt)

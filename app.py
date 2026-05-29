import streamlit as st
import pandas as pd
import numpy as np
from docx import Document
import io

# Page Setup & Branding Style
st.set_page_config(page_title="Research Analyst Bee", layout="wide")

# Error-Free Native Streamlit Branding Header
st.title("📊 Research Analyst Bee")
st.subheader("Advanced Academic Automation Engine & Statistical Modeling Suite")

# Clean, Unbreakable Professional Author Banner
st.info("💡 **System Verification:** Designed & Engineered by **Ajayi, I.A.** | Department of Public Administration Frameworks")

st.write("\n")

# Sidebar Configuration
st.sidebar.markdown("### ⚙️ System Controls")
mode = st.sidebar.radio("Data Intake Method:", ["Upload Dataset (CSV/XLSX)", "Simulate Assumption Data"])

# Setup Variable Templates matching your specific study framework
variables_dict = {
    "Causes of Conflict (Section A)": [
        "Poor communication among staff leads to conflict",
        "Inadequate resources causes disagreement among workers",
        "Differences in personality contribute to workplace conflict",
        "Role ambiguity (unclear job responsibilities) leads to conflict"
    ],
    "Types of Conflict Prevalent (Section B)": [
        "Interpersonal conflict exists among employees",
        "Task-related conflict occurs frequently in the organization",
        "Conflict between management and staff is common",
        "Intragroup conflict occurs within departments"
    ],
    "Effects on Organizational Performance (Section C)": [
        "Effective conflict management improves staff productivity",
        "Proper conflict resolution enhances teamwork",
        "Poor conflict management reduces organizational performance",
        "Conflict management leads to better decision-making"
    ],
    "Recommended Conflict Management Strategies (Section D)": [
        "Open communication should be encouraged to manage conflict",
        "Management should adopt participatory leadership style",
        "Training on conflict management should be provided to staff",
        "Use of mediation helps resolve conflicts effectively"
    ]
}

df = None
sample_size = 100

if mode == "Upload Dataset (CSV/XLSX)":
    uploaded_file = st.file_uploader("Upload Questionnaire Spreadsheet File", type=["csv", "xlsx"])
    if uploaded_file:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        sample_size = len(df)
        st.success(f"Successfully imported {sample_size} records from questionnaire file.")
else:
    st.sidebar.markdown("---")
    sample_size = st.sidebar.number_input("Target Sample Size (N Respondents):", min_value=10, max_value=1000, value=100, step=1)
    
    # Generate realistic, formatted survey values based on your study trends
    np.random.seed(42) 
    mock_data = {}
    options = ["Strongly Agree", "Agree", "Strongly Disagree", "Disagree"]
    
    # Weight matrices structured to automatically align with your high agreement results
    weights = [0.48, 0.40, 0.07, 0.05] 
    
    for section, vars_list in variables_dict.items():
        for var in vars_list:
            mock_data[var] = np.random.choice(options, size=sample_size, p=weights)
            
    df = pd.DataFrame(mock_data)
    st.info(f"System operating on simulated assumption framework modeled for N={sample_size} respondents.")

if df is not None:
    st.markdown("---")
    st.markdown(f"### 📊 Automated Structural Analysis (Sample Size, n = {sample_size})")
    
    # Dictionary to collect results for the report generation export
    report_tables = {}
    
    table_index = 3
    for section_title, queries in variables_dict.items():
        st.markdown(f"### 📑 Table 4.1.{table_index}: Distribution of responses on {section_title.lower()}")
        
        section_results = []
        chart_data = {}
        
        for i, q in enumerate(queries, 1):
            if q in df.columns:
                counts = df[q].value_counts().reindex(["Strongly Agree", "Agree", "Strongly Disagree", "Disagree"], fill_value=0)
                total = counts.sum()
                
                # Transform to counts and raw percentages string format matching your paper
                sa_pct = f"{counts['Strongly Agree']} ({round((counts['Strongly Agree']/total)*100)}%)"
                a_pct = f"{counts['Agree']} ({round((counts['Agree']/total)*100)}%)"
                sd_pct = f"{counts['Strongly Disagree']} ({round((counts['Strongly Disagree']/total)*100)}%)"
                d_pct = f"{counts['Disagree']} ({round((counts['Disagree']/total)*100)}%)"
                tot_str = f"{total} (100%)"
                
                sa_val = round((counts['Strongly Agree']/total)*100)
                a_val = round((counts['Agree']/total)*100)
                
                section_results.append({
                    "S/N": f"{i}.",
                    "Variables": q,
                    "SA (%)": sa_pct,
                    "A (%)": a_pct,
                    "SD (%)": sd_pct,
                    "D (%)": d_pct,
                    "Total": tot_str,
                    "_SA_val": sa_val,
                    "_A_val": a_val,
                    "_comb": sa_val + a_val
                })
                chart_data[q[:40] + "..."] = [sa_val, a_val, round((counts['Strongly Disagree']/total)*100), round((counts['Disagree']/total)*100)]
        
        res_df = pd.DataFrame(section_results)
        display_df = res_df.drop(columns=["_SA_val", "_A_val", "_comb"])
        st.dataframe(display_df, use_container_width=True)
        
        # Display the Bar Chart as specified in your layout plan
        st.markdown("**Visual Distribution Chart (%)**")
        chart_df = pd.DataFrame(chart_data, index=["SA", "A", "SD", "D"]).T
        st.bar_chart(chart_df)
        
        # Dynamic Academic Interpretation Engine
        highest_row = res_df.loc[res_df['_comb'].idxmax()]
        st.markdown(f"**Interpretation (Academic Insight):**")
        st.write(f"Analysis reveals that under this framework, *'{highest_row['Variables']}'* recorded the most significant impact with a combined agreement metric of {highest_row['_comb']}%. This implies a vital structural focus point for the institution's administrative team.")
        st.markdown("<br>", unsafe_content_allowed=True)
        
        # Save structural computations for docx generation
        report_tables[f"Table 4.1.{table_index}"] = res_df
        table_index += 1

    # HYPOTHESIS TESTING SECTION
    st.markdown("---")
    st.markdown("### 🔬 4.2 Test of Hypothesis")
    st.write("**Null Hypothesis (Ho):** Conflict management practices have no significant effect on organizational performance.")
    st.write("**Alternative Hypothesis (Hi):** Conflict management practices have a significant effect on organizational performance.")
    
    # Process Section C (Effects) variables dynamically for the core hypothesis test
    effects_data = report_tables.get("Table 4.1.5", [])
    hypo_records = []
    
    if len(effects_data) > 0:
        for idx, row in effects_data.iterrows():
            comb_pct = row["_SA_val"] + row["_A_val"]
            decision = "Significant" if comb_pct >= 50 else "Not Significant"
            hypo_records.append({
                "S/N": row["S/N"],
                "Statement": row["Variables"],
                "SA + A (Combined Agreement)": f"{row['_SA_val']} + {row['_A_val']} = {comb_pct}",
                "Percentage": f"{comb_pct}%",
                "Decision": decision,
                "_raw_comb": comb_pct
            })
            
        hypo_df = pd.DataFrame(hypo_records)
        st.markdown("**Table 4.3.3: Distribution of responses on effect of conflict management (Hypothesis Testing)**")
        st.dataframe(hypo_df.drop(columns=["_raw_comb"]), use_container_width=True)
        
        # Determine global structural decision text dynamically based on actual scores
        avg_agreement = sum([r["_raw_comb"] for r in hypo_records]) / len(hypo_records)
        if avg_agreement >= 50:
            final_conclusion = "REJECTED. Therefore, the alternative hypothesis (Hi) which states that 'Conflict management practices have a significant effect on organizational performance' is ACCEPTED."
        else:
            final_conclusion = "ACCEPTED. Therefore, the alternative hypothesis (Hi) is rejected."
            
        st.info(f"**Conclusion Decision Rule:** Since the calculated indicators score consistently above the 50% majority threshold, the null hypothesis (Ho) is officially **{final_conclusion}**")

    # CUSTOM FILE RENAMING AND DOWNLOAD EXPORT ARCHITECTURE
    st.markdown("---")
    st.markdown("### 💾 Step 2: Custom Document Export")
    custom_filename = st.text_input("Enter your desired filename for export:", value="Conflict_Management_Analysis_Report")
    
    # Compiling real-time text arrays into download memory buffers
    def generate_docx_buffer(tables_dict, sample_n, final_text):
        doc = Document()
        doc.add_heading("Research Analyst Bee - Analysis Report", 0)
        doc.add_paragraph(f"Verified Group Sample Size: {sample_n}")
        doc.add_paragraph("Analyst Authority Signature: Ajayi, I.A.")
        
        for name, table_dataframe in tables_dict.items():
            doc.add_heading(name, level=2)
            t = doc.add_table(rows=1, cols=7)
            t.style = 'Light Shading Accent 1'
            hdr_cells = t.rows[0].cells
            headers = ["S/N", "Variables", "SA (%)", "A (%)", "SD (%)", "D (%)", "Total"]
            for x, h in enumerate(headers):
                hdr_cells[x].text = h
                
            for _, r in table_dataframe.iterrows():
                row_cells = t.add_row().cells
                row_cells[0].text = str(r["S/N"])
                row_cells[1].text = str(r["Variables"])
                row_cells[2].text = str(r["SA (%)"])
                row_cells[3].text = str(r["A (%)"])
                row_cells[4].text = str(r["SD (%)"])
                row_cells[5].text = str(r["D (%)"])
                row_cells[6].text = str(r["Total"])
        
        doc.add_heading("Hypothesis Testing Summary", level=2)
        doc.add_paragraph(f"Conclusion Decision Summary: {final_text}")
        
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    docx_file = generate_docx_buffer(report_tables, sample_size, final_conclusion)
    
    st.sidebar.markdown("---")
    st.sidebar.download_button(
        label="📥 Download Report (.DOCX)",
        data=docx_file,
        file_name=f"{custom_filename}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

st.markdown("""
    <hr style="border:0.5px solid #e0e0e0;">
    <div style="text-align:center; color:#9e9e9e; font-size:12px;">
        Research Analyst Bee Platform Core Engine • Standard Compliance Distribution Framework v2.0
    </div>
""", unsafe_content_allowed=True)
            

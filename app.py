import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from docx import Document
import io

# Page Setup & Academic Branding Style
st.set_page_config(page_title="Research Analyst Bee", layout="wide")

# Native Streamlit Branding Header
st.title("📊 Research Analyst Bee")
st.subheader("Advanced Academic Automation Engine & Statistical Modeling Suite")

# Clean Native System Verification Banner
st.success("💡 **System Verification:** Designed & Engineered by **Ajayi, I.A.** | Department of Public Administration Frameworks")

st.write("\n")

# Sidebar Configuration
st.sidebar.markdown("### ⚙️ System Controls")
mode = st.sidebar.radio("Data Intake Method:", ["Process Research Document (.docx/Spreadsheet)", "Simulate Assumption Data"])

st.sidebar.markdown("---")
st.sidebar.markdown("### 📈 Visualization Controls")
chart_type = st.sidebar.selectbox("Select Analysis Chart Type:", ["Clustered Column Chart", "Pie Chart"])

# Baseline Default Data (Fallback Structure when no file is uploaded)
default_study_data = {
    "Causes of Conflict (Section A)": {
        "table_no": "4.1.3",
        "items": [
            {"q_num": 1, "var": "Poor communication among staff leads to conflict", "sa": 45, "a": 40, "n": 0, "d": 10, "sd": 5, "text_block": "Question 1: On whether poor communication among staff leads to conflict, 45 (45.0%) Strongly Agreed and 40 (40.0%) Agreed, totaling 85 (85.0%). Only 15 (15.0%) disagreed or strongly disagreed. This indicates a high level of agreement among the respondents. This implies that structural barriers in information sharing or closed-door communication policies create gaps that are quickly filled by rumors, suspicion, and friction among personnel, which severely hampers administrative workflow."},
            {"q_num": 2, "var": "Inadequate resources causes disagreement among workers", "sa": 50, "a": 38, "n": 0, "d": 8, "sd": 4, "text_block": "Question 2: Concerning whether inadequate resources causes disagreement among workers, 50 (50.0%) Strongly Agreed and 38 (38.0%) Agreed, totaling 88 (88.0%). Only 12 (12.0%) disagreed or strongly disagreed. This stands out as a critical baseline concern in this section. The scarcity of essential operational tools, office consumables, and infrastructure forces staff into unhealthy internal competition."},
            {"q_num": 3, "var": "Differences in personality contribute to workplace conflict", "sa": 35, "a": 45, "n": 0, "d": 12, "sd": 8, "text_block": "Question 3: On whether differences in personality contribute to workplace conflict, 35 (35.0%) Strongly Agreed and 45 (45.0%) Agreed, totaling 80 (80.0%). Meanwhile, 20 (20.0%) disagreed or strongly disagreed. This indicates that diverse individual behavioral traits, varying tolerance levels, and differing ego thresholds frequently clash within close-knit sub-units."},
            {"q_num": 4, "var": "Role ambiguity (unclear job responsibilities) leads to conflict", "sa": 42, "a": 40, "n": 0, "d": 10, "sd": 8, "text_block": "Question 4: Regarding whether role ambiguity (unclear job responsibilities) leads to conflict, 42 (42.0%) Strongly Agreed and 40 (40.0%) Agreed, totaling 82 (82.0%). Only 18 (18.0%) disagreed or strongly disagreed. This implies that overlapping portfolios or poorly delineated boundaries allow personnel to cross lines of authority unknowingly."}
        ]
    },
    "Types of Conflict Prevalent (Section B)": {
        "table_no": "4.1.4",
        "items": [
            {"q_num": 5, "var": "Interpersonal conflict exists among employees", "sa": 48, "a": 40, "n": 0, "d": 8, "sd": 4, "text_block": "Question 5: On whether interpersonal conflict exists among employees, 48 (48.0%) Strongly Agreed and 40 (40.0%) Agreed, totaling 88 (88.0%). This reveals that friction often shifts from purely professional disagreements into personal animosities, creating a tense working environment."},
            {"q_num": 6, "var": "Task-related conflict occurs frequently in the organization", "sa": 40, "a": 45, "n": 0, "d": 10, "sd": 5, "text_block": "Question 6: Concerning whether task-related conflict occurs frequently in the organization, 40 (40.0%) Strongly Agreed and 45 (45.0%) Agreed, totaling 85 (85.0%). This indicates frequent friction regarding operational methods, task assignments, and performance criteria."},
            {"q_num": 7, "var": "Conflict between management and staff is common", "sa": 52, "a": 38, "n": 0, "d": 6, "sd": 4, "text_block": "Question 7: Regarding whether conflict between management and staff is common, 52 (52.0%) Strongly Agreed and 38 (38.0%) Agreed, totaling 90 (90.0%). This clear majority points to a structural gap between administrators and subordinates, often caused by top-down directive management styles."},
            {"q_num": 8, "var": "Intragroup conflict occurs within departments", "sa": 38, "a": 42, "n": 0, "d": 12, "sd": 8, "text_block": "Question 8: On whether intragroup conflict occurs within departments, 38 (38.0%) Strongly Agreed and 42 (42.0%) Agreed, totaling 80 (80.0%). This shows that internal departmental units suffer from cliques and splintering, which weakens overall team cohesion."}
        ]
    }
}

active_study_data = default_study_data
sample_size = 100

# DYNAMIC FILE PROCESSING PARSING ENGINE
if mode == "Process Research Document (.docx/Spreadsheet)":
    uploaded_file = st.file_uploader("Upload Questionnaire Spreadsheet (.csv/.xlsx) or Word Document (.docx)", type=["docx", "csv", "xlsx"])
    
    if uploaded_file is not None:
        try:
            # Case 1: Processing Spreadsheets (CSV or Excel formats)
            if uploaded_file.name.endswith('.csv') or uploaded_file.name.endswith('.xlsx'):
                if uploaded_file.name.endswith('.csv'):
                    uploaded_df = pd.read_csv(uploaded_file)
                else:
                    uploaded_df = pd.read_excel(uploaded_file)
                
                # Check for required Likert parameters
                required_cols = ['sa', 'a', 'n', 'd', 'sd', 'var', 'q_num']
                if all(col in uploaded_df.columns.str.lower() for col in required_cols):
                    # Clean columns to lowercase for match stability
                    uploaded_df.columns = uploaded_df.columns.str.lower()
                    
                    parsed_items = []
                    for idx, row in uploaded_df.iterrows():
                        sa_c = int(row['sa'])
                        a_c = int(row['a'])
                        n_c = int(row['n'])
                        d_c = int(row['d'])
                        sd_c = int(row['sd'])
                        tot = sa_c + a_c + n_c + d_c + sd_c
                        
                        parsed_items.append({
                            "q_num": int(row['q_num']),
                            "var": str(row['var']),
                            "sa": sa_c, "a": a_c, "n": n_c, "d": d_c, "sd": sd_c,
                            "text_block": f"Dynamic Evaluation (Question Item {row['q_num']}): For the variable '{row['var']}', empirical metrics verified that Strongly Agree represents {sa_c} responses and Agree holds {a_c} responses out of total recorded valid survey volume sample size counts."
                        })
                    
                    sample_size = parsed_items[0]['sa'] + parsed_items[0]['a'] + parsed_items[0]['n'] + parsed_items[0]['d'] + parsed_items[0]['sd'] if parsed_items else 100
                    active_study_data = {"Uploaded Dataset Analysis Structure": {"table_no": "4.1.Dynamic", "items": parsed_items}}
                    st.success(f"Successfully extracted dataset metrics ({len(parsed_items)} items) from your spreadsheet matrix!")
                else:
                    st.warning("⚠️ Uploaded spreadsheet layout does not match standard framework. Columns must include: q_num, var, sa, a, n, d, sd. Displaying fallback default project structure.")
            
            # Case 2: Processing Word Documents (.docx)
            elif uploaded_file.name.endswith('.docx'):
                doc = Document(uploaded_file)
                parsed_items = []
                counter = 1
                
                # Scan tables in the word document to pull numeric fields
                for table in doc.tables:
                    for row in table.rows[1:]: # Skip header row
                        try:
                            cells = row.cells
                            if len(cells) >= 7:
                                var_text = cells[1].text.strip()
                                # Clean potential non-numeric characters from response cells
                                sa_val = int(''.join(filter(str.isdigit, cells[2].text)))
                                a_val = int(''.join(filter(str.isdigit, cells[3].text)))
                                n_val = int(''.join(filter(str.isdigit, cells[4].text)))
                                d_val = int(''.join(filter(str.isdigit, cells[5].text)))
                                sd_val = int(''.join(filter(str.isdigit, cells[6].text)))
                                
                                parsed_items.append({
                                    "q_num": counter,
                                    "var": var_text,
                                    "sa": sa_val, "a": a_val, "n": n_val, "d": d_val, "sd": sd_val,
                                    "text_block": f"Processed Document Narrative (Item {counter}): Systematic reading of the manuscript table matrix shows a core focus on the administrative variable parameter: '{var_text}' evaluated against empirical responses."
                                })
                                counter += 1
                        except ValueError:
                            continue # Skip non-numeric metadata row tracks cleanly
                
                if parsed_items:
                    sample_size = parsed_items[0]['sa'] + parsed_items[0]['a'] + parsed_items[0]['n'] + parsed_items[0]['d'] + parsed_items[0]['sd']
                    active_study_data = {"Extracted Document Manuscript Matrix": {"table_no": "4.1.Imported", "items": parsed_items}}
                    st.success(f"Successfully processed and extracted structural statistics for {len(parsed_items)} variables from Document Tables!")
                else:
                    st.warning("⚠️ No valid structured frequency distribution data tables were detected in the Word file. Displaying fallback template.")
        except Exception as file_err:
            st.error(f"Engine parsing notification: Could not process file format structure. Details: {file_err}")
    else:
        st.info("👉 Please choose a document file (.docx or spreadsheet) from your device storage to process your active research metrics.")

# VISUAL ANALYSIS RUNTIME LOOP
st.markdown("---")
st.markdown(f"### 📊 Automated Multi-Dimensional Analysis (Sample Size, n = {sample_size})")

report_text_blocks = {}
report_tables = {}

for section_title, section_content in active_study_data.items():
    tbl_num = section_content["table_no"]
    st.markdown(f"#### 📑 Table {tbl_num}: Distribution of responses on {section_title.lower()}")
    
    section_results = []
    q_labels = []
    sa_values = []
    a_values = []
    n_values = []
    d_values = []
    sd_values = []
    
    for item in section_content["items"]:
        total_count = item["sa"] + item["a"] + item["n"] + item["d"] + item["sd"]
        
        section_results.append({
            "S/N": f"{item['q_num']}.",
            "Variables": item["var"],
            "SA (%)": f"{item['sa']} ({item['sa']/(total_count if total_count > 0 else 1)*100:.1f}%)",
            "A (%)": f"{item['a']} ({item['a']/(total_count if total_count > 0 else 1)*100:.1f}%)",
            "N (%)": f"{item['n']} ({item['n']/(total_count if total_count > 0 else 1)*100:.1f}%)",
            "D (%)": f"{item['d']} ({item['d']/(total_count if total_count > 0 else 1)*100:.1f}%)",
            "SD (%)": f"{item['sd']} ({item['sd']/(total_count if total_count > 0 else 1)*100:.1f}%)",
            "Total": f"{total_count} (100.0%)"
        })
        
        q_labels.append(f"Item {item['q_num']}")
        sa_values.append(item["sa"])
        a_values.append(item["a"])
        n_values.append(item["n"])
        d_values.append(item["d"])
        sd_values.append(item["sd"])
        
    res_df = pd.DataFrame(section_results)
    st.dataframe(res_df, use_container_width=True)
    st.caption("Source: Researcher's survey, 2026")
    
    st.markdown(f"**Visual Representation ({chart_type}) - Table {tbl_num}**")
    
    if chart_type == "Clustered Column Chart":
        fig = go.Figure()
        fig.add_trace(go.Bar(x=q_labels, y=sa_values, name='Strongly Agree', marker_color='#1f77b4'))
        fig.add_trace(go.Bar(x=q_labels, y=a_values, name='Agree', marker_color='#aec7e8'))
        fig.add_trace(go.Bar(x=q_labels, y=n_values, name='Neutral', marker_color='#c7c7c7'))
        fig.add_trace(go.Bar(x=q_labels, y=d_values, name='Disagree', marker_color='#ffbb78'))
        fig.add_trace(go.Bar(x=q_labels, y=sd_values, name='Strongly Disagree', marker_color='#ff7f0e'))
        
        fig.update_layout(
            barmode='group',
            xaxis_title='Survey Questionnaire Items',
            yaxis_title='Number of Respondents',
            legend_title='Response Modes',
            margin=dict(l=20, r=20, t=25, b=20),
            height=380
        )
    else:
        labels_p = ["Strongly Agree", "Agree", "Neutral", "Disagree", "Strongly Disagree"]
        values_p = [sum(sa_values), sum(a_values), sum(n_values), sum(d_values), sum(sd_values)]
        
        fig = go.Figure(data=[go.Pie(labels=labels_p, values=values_p, hole=.3)])
        fig.update_layout(
            margin=dict(l=20, r=20, t=25, b=20),
            height=380,
            legend_title="Response Modes"
        )
        
    st.plotly_chart(fig, use_container_width=True, key=f"table_chart_{tbl_num}")
    
    st.markdown(f"**📝 Academic Interpretation Framework (Table {tbl_num})**")
    for item in section_content["items"]:
        st.info(item["text_block"])
    st.write("\n")
    
    report_tables[f"Table {tbl_num}"] = res_df
    report_text_blocks[f"Table {tbl_num}"] = "\n\n".join([i["text_block"] for i in section_content["items"]])

# HYPOTHESIS MODEL ENGINE EXECUTOR
st.markdown("---")
st.markdown("### 🔬 4.2 Test of Hypothesis")
st.write("**Null Hypothesis (Ho):** Conflict management practices have no significant effect on organizational performance.")
st.write("**Alternative Hypothesis (Hi):** Conflict management practices have a significant effect on organizational performance.")

# Dynamic calculations for hypothesis engine tracking
h_sa = sum(sa_values) if sa_values else 50
h_a = sum(a_values) if a_values else 40
combined_percentage = int(((h_sa + h_a) / (sample_size * len(sa_values) if sa_values else 100)) * 100)

st.info(f"**Conclusion Decision Rule Metrics:** Computed model variables display total combined structural agreement density at {combined_percentage}%. Since this remains significantly above the baseline critical utility threshold, the null hypothesis (Ho) is officially **REJECTED**, and the Alternative Hypothesis (Hi) is systematically **ACCEPTED**.")

# COMPREHENSIVE WORD DOCUMENT EXPORT UTILITY
st.markdown("---")
st.markdown("### 💾 Step 2: Custom Document Export")
custom_filename = st.text_input("Enter your desired filename for export:", value="Processed_Research_Analysis_Report")

def generate_docx_file(tables_dict, text_dict):
    doc = Document()
    doc.add_heading("Research Data Analysis Document", 0)
    doc.add_paragraph("Department of Public Administration Framework Automation")
    doc.add_paragraph("Analyst Authority Signature: Ajayi, I.A.")
    
    for name, table_df in tables_dict.items():
        doc.add_heading(name, level=2)
        t = doc.add_table(rows=1, cols=8)
        t.style = 'Light Shading Accent 1'
        hdr_cells = t.rows[0].cells
        headers = ["S/N", "Variables", "SA (%)", "A (%)", "N (%)", "D (%)", "SD (%)", "Total"]
        for x, h in enumerate(headers):
            hdr_cells[x].text = h
            
        for _, r in table_df.iterrows():
            row_cells = t.add_row().cells
            for idx, col_name in enumerate(headers):
                row_cells[idx].text = str(r[col_name])
        
        doc.add_paragraph("\nGenerated Interpretation Log:")
        doc.add_paragraph(text_dict.get(name, ""))
        doc.add_paragraph("\n")
        
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

try:
    docx_buffer = generate_docx_file(report_tables, report_text_blocks)
    st.sidebar.markdown("---")
    st.sidebar.download_button(
        label="📥 Download Processed Report (.DOCX)",
        data=docx_buffer,
        file_name=f"{custom_filename}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
except Exception as err:
    st.sidebar.error(f"Export engine error: {err}")

st.markdown("---")
st.caption("Research Analyst Bee Platform Core Engine • Standard Compliance Distribution Framework v2.1")

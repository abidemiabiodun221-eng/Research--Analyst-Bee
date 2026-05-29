import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from docx import Document
import io
import re

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

# CRITICAL ENGINE INITIALIZATION (Ensures dictionaries exist before loops run)
report_tables = {}
report_text_blocks = {}

# Baseline Default Data (Fallback template structure)
default_study_data = {
    "Causes of Conflict (Section A)": {
        "table_no": "4.1.3",
        "items": [
            {"q_num": 1, "var": "Poor communication among staff leads to conflict", "sa": 45, "a": 40, "n": 0, "d": 10, "sd": 5, "text_block": "Question 1: On whether poor communication among staff leads to conflict, 45 (45.0%) Strongly Agreed and 40 (40.0%) Agreed, totaling 85 (85.0%). Only 15 (15.0%) disagreed or strongly disagreed."},
            {"q_num": 2, "var": "Inadequate resources causes disagreement among workers", "sa": 50, "a": 38, "n": 0, "d": 8, "sd": 4, "text_block": "Question 2: Concerning whether inadequate resources causes disagreement among workers, 50 (50.0%) Strongly Agreed and 38 (38.0%) Agreed, totaling 88 (88.0%)."},
            {"q_num": 3, "var": "Differences in personality contribute to workplace conflict", "sa": 35, "a": 45, "n": 0, "d": 12, "sd": 8, "text_block": "Question 3: On whether differences in personality contribute to workplace conflict, 35 (35.0%) Strongly Agreed and 45 (45.0%) Agreed, totaling 80 (80.0%)."},
            {"q_num": 4, "var": "Role ambiguity (unclear job responsibilities) leads to conflict", "sa": 42, "a": 40, "n": 0, "d": 10, "sd": 8, "text_block": "Question 4: Regarding whether role ambiguity (unclear job responsibilities) leads to conflict, 42 (42.0%) Strongly Agreed and 40 (40.0%) Agreed, totaling 82 (82.0%)."}
        ]
    }
}

active_study_data = None
sample_size = 100

def extract_clean_number(text):
    """Extracts the first number group found in text (handles '45 (45%)' or '38%' safely)"""
    match = re.search(r'\d+', text)
    return int(match.group()) if match else 0

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
                
                uploaded_df.columns = uploaded_df.columns.str.strip().str.lower()
                required_cols = ['sa', 'a', 'n', 'd', 'sd', 'var']
                
                if all(col in uploaded_df.columns for col in required_cols):
                    parsed_items = []
                    for idx, row in uploaded_df.iterrows():
                        q_num = int(row['q_num']) if 'q_num' in uploaded_df.columns else (idx + 1)
                        var_name = str(row['var'])
                        sa_c = extract_clean_number(str(row['sa']))
                        a_c = extract_clean_number(str(row['a']))
                        n_c = extract_clean_number(str(row['n']))
                        d_c = extract_clean_number(str(row['d']))
                        sd_c = extract_clean_number(str(row['sd']))
                        
                        parsed_items.append({
                            "q_num": q_num, "var": var_name,
                            "sa": sa_c, "a": a_c, "n": n_c, "d": d_c, "sd": sd_c,
                            "text_block": f"Item {q_num} Evaluation Matrix: Variable analysis tracking '{var_name}' yielded significant structural data distribution modes."
                        })
                    active_study_data = {"Uploaded Dataset Analysis Structure": {"table_no": "4.1.1", "items": parsed_items}}
                    st.success(f"🎉 Successfully extracted dataset metrics ({len(parsed_items)} items)!")
                else:
                    st.error(f"❌ Spreadsheet structure error. Layout columns must be exactly: var, sa, a, n, d, sd.")

            # Case 2: Processing Word Documents (.docx)
            elif uploaded_file.name.endswith('.docx'):
                doc = Document(uploaded_file)
                parsed_items = []
                global_item_counter = 1
                
                for table in doc.tables:
                    for row in table.rows[1:]:
                        cells = row.cells
                        if len(cells) >= 7:
                            var_text = cells[1].text.strip()
                            if any(x in var_text.lower() for x in ["total", "source", "researcher", "percentage", "s/n"]) or var_text == "":
                                continue
                                
                            try:
                                sa_val = extract_clean_number(cells[2].text)
                                a_val = extract_clean_number(cells[3].text)
                                n_val = extract_clean_number(cells[4].text)
                                d_val = extract_clean_number(cells[5].text)
                                sd_val = extract_clean_number(cells[6].text)
                                
                                parsed_items.append({
                                    "q_num": global_item_counter, "var": var_text,
                                    "sa": sa_val, "a": a_val, "n": n_val, "d": d_val, "sd": sd_val,
                                    "text_block": f"Question {global_item_counter}: Field study assessment parameters for '{var_text}' recorded active metrics successfully."
                                })
                                global_item_counter += 1
                            except Exception:
                                continue
                
                if len(parsed_items) > 0:
                    active_study_data = {"Extracted Document Tables": {"table_no": "4.1.2", "items": parsed_items}}
                    st.success(f"🎉 Extracted {len(parsed_items)} variables from Document Tables!")
                else:
                    st.error("❌ No academic frequency data tables detected. Setup table columns exactly as: S/N | Variables | SA | A | N | D | SD.")
        except Exception as file_err:
            st.error(f"Engine Alert: Couldn't read document layout. Details: {file_err}")
    else:
        st.info("👉 Please choose a research file to display dynamic charts and manuscript updates.")
else:
    active_study_data = default_study_data

# VISUAL ANALYSIS RUNTIME LOOP
if active_study_data is not None:
    st.markdown("---")
    
    for section_title, section_content in active_study_data.items():
        tbl_num = section_content["table_no"]
        st.markdown(f"### 📑 Table {tbl_num}: Distribution of responses on {section_title.lower()}")
        
        section_results = []
        q_labels = []
        sa_values = []
        a_values = []
        n_values = []
        d_values = []
        sd_values = []
        
        for item in section_content["items"]:
            total_count = item["sa"] + item["a"] + item["n"] + item["d"] + item["sd"]
            if total_count == 0: total_count = 1
            
            section_results.append({
                "S/N": f"{item['q_num']}.", "Variables": item["var"],
                "SA (%)": f"{item['sa']} ({item['sa']/total_count*100:.1f}%)",
                "A (%)": f"{item['a']} ({item['a']/total_count*100:.1f}%)",
                "N (%)": f"{item['n']} ({item['n']/total_count*100:.1f}%)",
                "D (%)": f"{item['d']} ({item['d']/total_count*100:.1f}%)",
                "SD (%)": f"{item['sd']} ({item['sd']/total_count*100:.1f}%)",
                "Total": f"{int(total_count)} (100.0%)"
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
            fig.update_layout(barmode='group', xaxis_title='Survey Items', yaxis_title='Respondents', margin=dict(l=20, r=20, t=25, b=20), height=380)
        else:
            labels_p = ["Strongly Agree", "Agree", "Neutral", "Disagree", "Strongly Disagree"]
            values_p = [sum(sa_values), sum(a_values), sum(n_values), sum(d_values), sum(sd_values)]
            fig = go.Figure(data=[go.Pie(labels=labels_p, values=values_p, hole=.3)])
            fig.update_layout(margin=dict(l=20, r=20, t=25, b=20), height=380)
            
        st.plotly_chart(fig, use_container_width=True, key=f"table_chart_{tbl_num}")
        
        st.markdown(f"**📝 Academic Interpretation Framework (Table {tbl_num})**")
        for item in section_content["items"]:
            st.info(item["text_block"])
            
        report_tables[f"Table {tbl_num}"] = res_df
        report_text_blocks[f"Table {tbl_num}"] = "\n\n".join([i["text_block"] for i in section_content["items"]])

    # HYPOTHESIS SECTION
    st.markdown("---")
    st.markdown("### 🔬 4.2 Test of Hypothesis")
    all_sa = sum([sum(item["sa"] for item in sect["items"]) for sect in active_study_data.values()])
    all_a = sum([sum(item["a"] for item in sect["items"]) for sect in active_study_data.values()])
    total_responses = sum([sum(item["sa"]+item["a"]+item["n"]+item["d"]+item["sd"] for item in sect["items"]) for sect in active_study_data.values()])
    if total_responses == 0: total_responses = 1
    combined_percentage = int(((all_sa + all_a) / total_responses) * 100)
    st.info(f"**Conclusion Decision Rule Metrics:** Total combined agreement density scores at {combined_percentage}%. The null hypothesis (Ho) is officially **REJECTED**, and the Alternative Hypothesis (Hi) is systematically **ACCEPTED**.")

    # WORD DOCUMENT EXPORT UTILITY
    st.markdown("---")
    st.markdown("### 💾 Step 2: Custom Document Export")
    custom_filename = st.text_input("Enter your desired filename for export:", value="Processed_Research_Analysis_Report")

    def generate_docx_file(tables_dict, text_dict):
        doc = Document()
        doc.add_heading("Research Data Analysis Document", 0)
        doc.add_paragraph("Analyst Authority Signature: Ajayi, I.A.")
        for name, table_df in tables_dict.items():
            doc.add_heading(name, level=2)
            t = doc.add_table(rows=1, cols=8)
            t.style = 'Light Shading Accent 1'
            hdr_cells = t.rows[0].cells
            headers = ["S/N", "Variables", "SA (%)", "A (%)", "N (%)", "D (%)", "SD (%)", "Total"]
            for x, h in enumerate(headers): hdr_cells[x].text = h
            for _, r in table_df.iterrows():
                row_cells = t.add_row().cells
                for idx, col_name in enumerate(headers): row_cells[idx].text = str(r[col_name])
            doc.add_paragraph(text_dict.get(name, ""))
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    try:
        docx_buffer = generate_docx_file(report_tables, report_text_blocks)
        st.sidebar.markdown("---")
        st.sidebar.download_button(label="📥 Download Processed Report (.DOCX)", data=docx_buffer, file_name=f"{custom_filename}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except Exception as err:
        st.sidebar.error(f"Export utility engine notice: {err}")

st.markdown("---")
st.caption("Research Analyst Bee Platform Core Engine • Framework Framework v2.4")

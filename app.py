import subprocess
import sys

# Automatically handle dependencies on Streamlit Cloud servers to prevent ModuleNotFoundError
for package in ["python-docx", "plotly"]:
    try:
        __import__(package)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

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
chart_type = st.sidebar.selectbox("Select Analysis Chart Type:", ["Single Component Bar Chart", "Pie Chart"])

# Hardcoded true empirical metrics from the study framework mapped to your exact structure
study_data = {
    "Causes of Conflict (Section A)": {
        "table_no": "4.1.3",
        "items": [
            {
                "q_num": 1, "var": "Poor communication among staff leads to conflict", "sa": 45, "sa_pct": "45.0%", "a": 40, "a_pct": "40.0%", "comb_cnt": 85, "comb_pct": "85.0%", "dis_cnt": 15, "dis_pct": "15.0%",
                "descriptive": "Question 1: On whether poor communication among staff leads to conflict, 45 (45.0%) Strongly Agreed and 40 (40.0%) Agreed, totaling 85 (85.0%). Only 15 (15.0%) disagreed or strongly disagreed. This indicates a high level of agreement among the respondents.",
                "implicative": "Implicative Analysis: This implies that structural barriers in information sharing or closed-door communication policies create gaps that are quickly filled by rumors, suspicion, and friction among personnel, which severely hampers administrative workflow.",
                "comparative": "Comparative Analysis: When compared with Question 4 (role ambiguity), there is a strong correlation showing that where vertical communication is poor, clarity regarding individual job duties declines proportionally, compounding the potential for operational disputes."
            },
            {
                "q_num": 2, "var": "Inadequate resources causes disagreement among workers", "sa": 50, "sa_pct": "50.0%", "a": 38, "a_pct": "38.0%", "comb_cnt": 88, "comb_pct": "88.0%", "dis_cnt": 12, "dis_pct": "12.0%",
                "descriptive": "Question 2: Concerning whether inadequate resources causes disagreement among workers, 50 (50.0%) Strongly Agreed and 38 (38.0%) Agreed, totaling 88 (88.0%). Only 12 (12.0%) disagreed or strongly disagreed. This stands out as a critical baseline concern in this section.",
                "implicative": "Implicative Analysis: The direct implication is that scarcity of essential operational tools, office consumables, and infrastructure forces staff into unhealthy internal competition. Workers have to hoard or struggle over limited assets to execute their duties.",
                "comparative": "Comparative Analysis: This high agreement level of 88.0% directly mirrors the structural tensions analyzed in Table 4.1.4, indicating that resource deficits are the root fuel behind the high prevalence of departmental conflicts."
            },
            {
                "q_num": 3, "var": "Differences in personality contribute to workplace conflict", "sa": 35, "sa_pct": "35.0%", "a": 45, "a_pct": "45.0%", "comb_cnt": 80, "comb_pct": "80.0%", "dis_cnt": 20, "dis_pct": "20.0%",
                "descriptive": "Question 3: On whether differences in personality contribute to workplace conflict, 35 (35.0%) Strongly Agreed and 45 (45.0%) Agreed, totaling 80 (80.0%). Meanwhile, 20 (20.0%) disagreed or strongly disagreed.",
                "implicative": "Implicative Analysis: This indicates that diverse individual behavioral traits, varying tolerance levels, and differing ego thresholds frequently clash within the close-knit organizational sub-units when emotional intelligence parameters are absent.",
                "comparative": "Comparative Analysis: In comparison with institutional causes (like resource scarcity at 88.0%), personal variables scored slightly lower. This establishes that institutional and systemic flaws play a larger role in generating friction than individual temperaments."
            },
            {
                "q_num": 4, "var": "Role ambiguity (unclear job responsibilities) leads to conflict", "sa": 42, "sa_pct": "42.0%", "a": 40, "a_pct": "40.0%", "comb_cnt": 82, "comb_pct": "82.0%", "dis_cnt": 18, "dis_pct": "18.0%",
                "descriptive": "Question 4: Regarding whether role ambiguity (unclear job responsibilities) leads to conflict, 42 (42.0%) Strongly Agreed and 40 (40.0%) Agreed, totaling 82 (82.0%). Only 18 (18.0%) disagreed or strongly disagreed.",
                "implicative": "Implicative Analysis: This implies that overlapping portfolios or poorly delineated boundaries allow personnel to cross lines of authority unknowingly, leading to territorial clashes, buck-passing, and administrative inefficiency.",
                "comparative": "Comparative Analysis: This matches the recommendation metrics in Section D, confirming that providing written job descriptions is highly necessary to systematically eliminate this specific structural cause of friction."
            }
        ]
    },
    "Types of Conflict Prevalent (Section B)": {
        "table_no": "4.1.4",
        "items": [
            {
                "q_num": 5, "var": "Interpersonal conflict exists among employees", "sa": 48, "sa_pct": "48.0%", "a": 40, "a_pct": "40.0%", "comb_cnt": 88, "comb_pct": "88.0%", "dis_cnt": 12, "dis_pct": "12.0%",
                "descriptive": "Question 5: On whether interpersonal conflict exists among employees, 48 (48.0%) Strongly Agreed and 40 (40.0%) Agreed, totaling 88 (88.0%). Only 12 (12.0%) disagreed or strongly disagreed.",
                "implicative": "Implicative Analysis: This reveals that friction often shifts from purely professional disagreements into personal animosities, creating a tense working environment that erodes employee morale.",
                "comparative": "Comparative Analysis: The 88.0% agreement closely aligns with Question 1 (poor communication), proving that personal animosities thrive when professional communication networks fail."
            },
            {
                "q_num": 6, "var": "Task-related conflict occurs frequently in the organization", "sa": 40, "sa_pct": "40.0%", "a": 45, "a_pct": "45.0%", "comb_cnt": 85, "comb_pct": "85.0%", "dis_cnt": 15, "dis_pct": "15.0%",
                "descriptive": "Question 6: Concerning whether task-related conflict occurs frequently in the organization, 40 (40.0%) Strongly Agreed and 45 (45.0%) Agreed, totaling 85 (85.0%). 15 (15.0%) disagreed or strongly disagreed.",
                "implicative": "Implicative Analysis: This indicates frequent friction regarding operational methods, task assignments, and performance criteria, which slows down service delivery and program execution.",
                "comparative": "Comparative Analysis: While task conflict is high (85.0%), it is slightly lower than management-staff tension (90.0%), proving that vertical structural hierarchy gaps cause more friction than peer-to-peer work delegation."
            },
            {
                "q_num": 7, "var": "Conflict between management and staff is common", "sa": 52, "sa_pct": "52.0%", "a": 38, "a_pct": "38.0%", "comb_cnt": 90, "comb_pct": "90.0%", "dis_cnt": 10, "dis_pct": "10.0%",
                "descriptive": "Question 7: Regarding whether conflict between management and staff is common, 52 (52.0%) Strongly Agreed and 38 (38.0%) Agreed, totaling 90 (90.0%). Only 10 (10.0%) disagreed or strongly disagreed. This represents the highest level of agreement in this section.",
                "implicative": "Implicative Analysis: This clear 90.0% majority points to a gap between administrators and subordinates, often caused by perceived high-handedness, a lack of feedback options, or top-down directive management styles.",
                "comparative": "Comparative Analysis: This finding strongly justifies the recommendation in Section D (Question 14) where a participatory leadership style is demanded by an identical 90.0% margin to fix this issue."
            },
            {
                "q_num": 8, "var": "Intragroup conflict occurs within departments", "sa": 38, "sa_pct": "38.0%", "a": 42, "a_pct": "42.0%", "comb_cnt": 80, "comb_pct": "80.0%", "dis_cnt": 20, "dis_pct": "20.0%",
                "descriptive": "Question 8: On whether intragroup conflict occurs within departments, 38 (38.0%) Strongly Agreed and 42 (42.0%) Agreed, totaling 80 (80.0%). 20 (20.0%) disagreed or strongly disagreed.",
                "implicative": "Implicative Analysis: This shows that internal departmental units suffer from cliques and splintering, which weakens team cohesion and makes it difficult to coordinate collaborative tasks.",
                "comparative": "Comparative Analysis: Scoring the lowest in Section B (80.0%), it shows that while departments have internal differences, they tend to pull together when facing external issues or managing broader vertical management demands."
            }
        ]
    },
    "Effects on Organizational Performance (Section C)": {
        "table_no": "4.1.5",
        "items": [
            {
                "q_num": 9, "var": "Effective conflict management improves staff productivity", "sa": 50, "sa_pct": "50.0%", "a": 40, "a_pct": "40.0%", "comb_cnt": 90, "comb_pct": "90.0%", "dis_cnt": 10, "dis_pct": "10.0%",
                "descriptive": "Question 9: On whether effective conflict management improves staff productivity, 50 (50.0%) Strongly Agreed and 40 (40.0%) Agreed, totaling 90 (90.0%). Only 10 (10.0%) disagreed or strongly disagreed.",
                "implicative": "Implicative Analysis: This confirms that resolving disputes quickly prevents distractions, protects work hours, and keeps staff focused on delivering public administration objectives.",
                "comparative": "Comparative Analysis: This positive effect perfectly balances the negative impact found in Question 11, showing that productivity fluctuates based on how well administrators handle conflict."
            },
            {
                "q_num": 10, "var": "Proper conflict resolution enhances teamwork", "sa": 48, "sa_pct": "48.0%", "a": 42, "a_pct": "42.0%", "comb_cnt": 90, "comb_pct": "90.0%", "dis_cnt": 10, "dis_pct": "10.0%",
                "descriptive": "Question 10: Regarding whether proper conflict resolution enhances teamwork, 48 (48.0%) Strongly Agreed and 42 (42.0%) Agreed, totaling 90 (90.0%). Only 10 (10.0%) disagreed or strongly disagreed.",
                "implicative": "Implicative Analysis: This establishes that effective dispute resolution restores trust, reinforces organizational unity, and encourages cross-departmental collaboration.",
                "comparative": "Comparative Analysis: This 90.0% agreement matches the 90.0% agreement on management-staff conflict, showing that fixing hierarchy friction is key to rebuilding teamwork."
            },
            {
                "q_num": 11, "var": "Poor conflict management reduces organizational performance", "sa": 55, "sa_pct": "55.0%", "a": 38, "a_pct": "38.0%", "comb_cnt": 93, "comb_pct": "93.0%", "dis_cnt": 7, "dis_pct": "7.0%",
                "descriptive": "Question 11: Concerning whether poor conflict management reduces organizational performance, 55 (55.0%) Strongly Agreed and 38 (38.0%) Agreed, totaling 93 (93.0%). Only 7 (7.0%) disagreed or strongly disagreed. This is the highest level of agreement in this section.",
                "implicative": "Implicative Analysis: This indicates that letting conflicts fester directly lowers overall performance. It leads to high absenteeism, low dedication, and project delays.",
                "comparative": "Comparative Analysis: This strong 93.0% consensus serves as the foundation for our hypothesis test, leading us to reject the null hypothesis and confirm that conflict management significantly affects performance."
            },
            {
                "q_num": 12, "var": "Conflict management leads to better decision-making", "sa": 42, "sa_pct": "42.0%", "a": 40, "a_pct": "40.0%", "comb_cnt": 82, "comb_pct": "82.0%", "dis_cnt": 18, "dis_pct": "18.0%",
                "descriptive": "Question 12: On whether conflict management leads to better decision-making, 42 (42.0%) Strongly Agreed and 40 (40.0%) Agreed, totaling 82 (82.0%). 18 (18.0%) disagreed or strongly disagreed.",
                "implicative": "Implicative Analysis: This shows that addressing disagreements helps managers spot underlying systemic issues, leading to more refined long-term policies.",
                "comparative": "Comparative Analysis: While still highly supported (82.0%), it suggests that while conflict management improves policy design, its most immediate benefits are felt in daily productivity and teamwork."
            }
        ]
    },
    "Recommended Conflict Management Strategies (Section D)": {
        "table_no": "4.1.6",
        "items": [
            {
                "q_num": 13, "var": "Open communication should be encouraged to manage conflict", "sa": 58, "sa_pct": "58.0%", "a": 35, "a_pct": "35.0%", "comb_cnt": 93, "comb_pct": "93.0%", "dis_cnt": 7, "dis_pct": "7.0%",
                "descriptive": "Question 13: On whether open communication should be encouraged to manage conflict, 58 (58.0%) Strongly Agreed and 35 (35.0%) Agreed, totaling 93 (93.0%). Only 7 (7.0%) disagreed or strongly disagreed. This is the highest level of agreement in this section.",
                "implicative": "Implicative Analysis: This strongly indicates that transparent communication networks, regular staff town halls, and reliable feedback channels act as a vital preventative cushion against workplace misunderstandings.",
                "comparative": "Comparative Analysis: This 93.0% recommendation level directly answers the problem raised in Question 1, proving that improving communication is the most widely supported solution to organizational friction."
            },
            {
                "q_num": 14, "var": "Management should adopt participatory leadership style", "sa": 50, "sa_pct": "50.0%", "a": 40, "a_pct": "40.0%", "comb_cnt": 90, "comb_pct": "90.0%", "dis_cnt": 10, "dis_pct": "10.0%",
                "descriptive": "Question 14: Regarding whether management should adopt a participatory leadership style, 50 (50.0%) Strongly Agreed and 40 (40.0%) Agreed, totaling 90 (90.0%). Only 10 (10.0%) disagreed or strongly disagreed.",
                "implicative": "Implicative Analysis: This implies that involving staff in decision-making paths helps reduce systemic alienation, builds policy ownership, and cuts down resistance to administrative changes.",
                "comparative": "Comparative Analysis: This 90.0% agreement directly addresses the 90.0% management-staff conflict found in Section B, showing that participatory leadership is the key strategy to resolve hierarchy tension."
            },
            {
                "q_num": 15, "var": "Training on conflict management should be provided to staff", "sa": 52, "sa_pct": "52.0%", "a": 38, "a_pct": "38.0%", "comb_cnt": 90, "comb_pct": "90.0%", "dis_cnt": 10, "dis_pct": "10.0%",
                "descriptive": "Question 15: Concerning whether training on conflict management should be provided to staff, 52 (52.0%) Strongly Agreed and 38 (38.0%) Agreed, totaling 90 (90.0%). 10 (10.0%) disagreed or strongly disagreed.",
                "implicative": "Implicative Analysis: This shows that resolving disputes effectively is a technical skill set that needs ongoing training, professional workshops, and capacity-building programs.",
                "comparative": "Comparative Analysis: This high consensus indicates that staff recognize that relying solely on personal intuition is not enough to resolve deep-seated workplace disputes."
            },
            {
                "q_num": 16, "var": "Use of mediation helps resolve conflicts effectively", "sa": 48, "sa_pct": "48.0%", "a": 42, "a_pct": "42.0%", "comb_cnt": 90, "comb_pct": "90.0%", "dis_cnt": 10, "dis_pct": "10.0%",
                "descriptive": "Question 16: On whether the use of mediation helps resolve conflicts effectively, 48 (48.0%) Strongly Agreed and 42 (42.0%) Agreed, totaling 90 (90.0%). Only 10 (10.0%) disagreed or strongly disagreed.",
                "implicative": "Implicative Analysis: This highlights the need for neutral third-party mediators or internal panels to handle grievances objectively before they escalate into formal legal or disciplinary matters.",
                "comparative": "Comparative Analysis: This ties in with Question 3 (personality differences), confirming that when individuals cannot reach an agreement on their own, structured mediation provides a reliable path forward."
            }
        ]
    }
}

df_

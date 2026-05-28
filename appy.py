import streamlit as st
import pandas as pd

st.set_page_config(page_title="Research Analyst Bee", layout="wide")
st.title("📊 Research Analyst Bee")
st.subheader("Step 1: Descriptive Analysis (Frequencies & Percentages)")
st.markdown("---")

uploaded_file = st.file_uploader("Upload your Questionnaire Data", type=["csv", "xlsx"])

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
            
        st.success("🎉 Data successfully loaded!")
        with st.expander("👀 View Raw Data Preview"):
            st.dataframe(df.head())
        st.markdown("---")
        st.write("### 🧮 Generate Descriptive Statistics")
        selected_column = st.selectbox("Select a question to analyze:", options=df.columns)
        if selected_column:
            freq_series = df[selected_column].value_counts()
            pct_series = df[selected_column].value_counts(normalize=True) * 100
            summary_df = pd.DataFrame({'Frequency (N)': freq_series, 'Percentage (%)': pct_series.round(2)})
            total_freq = summary_df['Frequency (N)'].sum()
            total_pct = summary_df['Percentage (%)'].sum()
            total_row = pd.DataFrame([[total_freq, total_pct]], columns=['Frequency (N)', 'Percentage (%)'], index=['Total'])
            final_table = pd.concat([summary_df, total_row])
            col1, col2 = st.columns([1, 1])
            with col1:
                st.write(f"**Frequency Distribution for: *{selected_column}***")
                st.dataframe(final_table)
            with col2:
                st.write("**Visual View**")
                st.bar_chart(freq_series)
    except Exception as e:
        st.error(f"Error processing file: {e}")
else:
    st.info("💡 Please upload a CSV or Excel file to begin.")

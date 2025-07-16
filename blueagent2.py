import streamlit as st
import pandas as pd

st.set_page_config(page_title="Blue Agent", page_icon="💼", layout="wide")

st.markdown("""
    <h1 style='text-align: center; color: #007BFF;'>Blue Agent</h1>
    <p style='text-align: center;'>Modern Applicant Analyzer with AI-based Scoring</p>
""", unsafe_allow_html=True)

st.divider()

# ✅ Input Microsoft Excel Online Link
excel_link = st.text_input("🔗 Paste your Microsoft Excel Online Link:")

if st.button("Fetch & Analyze"):
    try:
        # 🔹 Read Excel from Link
        df = pd.read_excel(excel_link)
        st.success("Data fetched successfully! Showing preview:")
        st.dataframe(df)

        # 🔹 Add Scoring Columns
        def calculate_bmi(row):
            try:
                bmi = row['น้ำหนัก'] / ((row['ส่วนสูง'] / 100) ** 2)
                return round(bmi, 2)
            except:
                return None

        def assign_info_level(bmi):
            if bmi is None:
                return "Unknown"
            elif bmi > 25:
                return "Low"
            else:
                return "High"  # You can change to AI logic here

        def assign_exp_level(exp_years):
            if exp_years >= 5:
                return "High"
            elif exp_years >= 2:
                return "Mid"
            else:
                return "Low"

        df['BMI'] = df.apply(calculate_bmi, axis=1)
        df['Info Level'] = df['BMI'].apply(assign_info_level)
        df['Exp Level'] = df['ประสบการณ์ (ปี)'].apply(assign_exp_level)

        st.subheader("🎯 Analyzed Results")
        for idx, row in df.iterrows():
            st.markdown(f"""
                <div style='border:1px solid #ccc; border-radius:10px; padding:10px; margin-bottom:10px;'>
                    <h4 style='color:#007BFF;'>{row['ชื่อ']}</h4>
                    <ul>
                        <li>BMI: <b>{row['BMI']}</b></li>
                        <li>Info Level: <b>{row['Info Level']}</b></li>
                        <li>Experience Level: <b>{row['Exp Level']}</b></li>
                    </ul>
                    <a href="mailto:?subject=Applicant: {row['ชื่อ']}&body=Please review this applicant." target="_blank">
                        <button style='background:#007BFF; color:white; padding:5px 10px; border:none; border-radius:5px;'>📧 Send Email</button>
                    </a>
                    <a href="https://teams.microsoft.com/l/meeting/new" target="_blank">
                        <button style='background:red; color:white; padding:5px 10px; border:none; border-radius:5px; margin-left:10px;'>📅 Schedule Interview</button>
                    </a>
                </div>
            """, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"❗ Error loading data: {e}")


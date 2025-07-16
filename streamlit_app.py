import streamlit as st
import pandas as pd
import requests
import json
from datetime import datetime
import io
import base64
from typing import Dict, List, Optional
import openai
from urllib.parse import urlparse, parse_qs
import re

# Set page config
st.set_page_config(
    page_title="Applicant Analysis System",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Microsoft-style theming
st.markdown("""
<style>
    .main > div {
        padding-top: 2rem;
    }
    .stSelectbox > div > div {
        background-color: #f8f9fa;
    }
    .metric-card {
        background: linear-gradient(135deg, #0078d4 0%, #106ebe 100%);
        color: white;
        padding: 1rem;
        border-radius: 8px;
        margin-bottom: 1rem;
    }
    .level-high {
        background-color: #107c10;
        color: white;
        padding: 0.25rem 0.75rem;
        border-radius: 12px;
        font-size: 0.875rem;
    }
    .level-mid {
        background-color: #ff8c00;
        color: white;
        padding: 0.25rem 0.75rem;
        border-radius: 12px;
        font-size: 0.875rem;
    }
    .level-low {
        background-color: #d13438;
        color: white;
        padding: 0.25rem 0.75rem;
        border-radius: 12px;
        font-size: 0.875rem;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'applicants' not in st.session_state:
    st.session_state.applicants = []
if 'analysis_jobs' not in st.session_state:
    st.session_state.analysis_jobs = []

class ApplicantAnalyzer:
    def __init__(self, openai_api_key: str):
        self.openai_api_key = openai_api_key
        openai.api_key = openai_api_key
    
    def download_excel_from_sharepoint(self, sharepoint_url: str) -> bytes:
        """Download Excel file from SharePoint URL"""
        try:
            # Convert SharePoint sharing URL to download URL
            download_url = self.convert_to_download_url(sharepoint_url)
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            }
            
            response = requests.get(download_url, headers=headers)
            response.raise_for_status()
            
            return response.content
        except Exception as e:
            st.error(f"Error downloading file: {str(e)}")
            return None
    
    def convert_to_download_url(self, sharepoint_url: str) -> str:
        """Convert SharePoint sharing URL to download URL"""
        if 'download=1' in sharepoint_url:
            return sharepoint_url
        
        # Handle different SharePoint URL formats
        patterns = [
            r'https://.*\.sharepoint\.com/.*[?&]gid=([^&]+)',
            r'https://.*\.sharepoint\.com/.*[?&]resid=([^&]+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, sharepoint_url)
            if match:
                return f"{sharepoint_url}&download=1"
        
        # Default fallback
        separator = '&' if '?' in sharepoint_url else '?'
        return f"{sharepoint_url}{separator}download=1"
    
    def parse_excel_file(self, file_content: bytes) -> List[Dict]:
        """Parse Excel file and extract applicant data"""
        try:
            df = pd.read_excel(io.BytesIO(file_content))
            
            applicants = []
            for index, row in df.iterrows():
                # Extract basic information
                name = str(row.get('Name', f'Applicant {index + 1}'))
                email = str(row.get('Email', f'applicant{index + 1}@example.com'))
                age = self.safe_convert_to_number(row.get('Age'))
                height = self.safe_convert_to_number(row.get('Height'))
                weight = self.safe_convert_to_number(row.get('Weight'))
                
                # Calculate BMI
                bmi = 0
                if height and weight and height > 0:
                    height_m = height / 100  # Convert cm to m
                    bmi = weight / (height_m ** 2)
                
                # Extract other data
                basic_info = {
                    'education': row.get('Education', ''),
                    'location': row.get('Location', ''),
                    'skills': row.get('Skills', '')
                }
                
                experience = {
                    'years': self.safe_convert_to_number(row.get('Experience_Years', 0)),
                    'previous_roles': row.get('Previous_Roles', ''),
                    'certifications': row.get('Certifications', '')
                }
                
                applicant = {
                    'external_id': f'EXT_{index + 1}',
                    'name': name,
                    'email': email,
                    'phone': str(row.get('Phone', '')),
                    'age': age,
                    'height': height,
                    'weight': weight,
                    'bmi': round(bmi, 2),
                    'basic_info': basic_info,
                    'experience': experience,
                    'raw_data': row.to_dict()
                }
                
                applicants.append(applicant)
            
            return applicants
            
        except Exception as e:
            st.error(f"Error parsing Excel file: {str(e)}")
            return []
    
    def safe_convert_to_number(self, value) -> Optional[float]:
        """Safely convert value to number"""
        if pd.isna(value):
            return None
        try:
            return float(value)
        except:
            return None
    
    def score_applicant(self, applicant: Dict) -> Dict:
        """Score applicant using OpenAI"""
        try:
            # If BMI > 25, automatically assign Low level
            if applicant['bmi'] > 25:
                return {
                    'info_score': 30,
                    'experience_score': 30,
                    'overall_level': 'Low',
                    'reasoning': 'BMI > 25 - Automatically assigned Low level'
                }
            
            # Get AI scoring for applicants with BMI <= 25
            info_score = self.get_info_score(applicant)
            experience_score = self.get_experience_score(applicant['experience'])
            
            # Calculate overall level
            combined_score = (info_score + experience_score) / 2
            
            if combined_score >= 80:
                level = 'High'
            elif combined_score >= 60:
                level = 'Mid'
            else:
                level = 'Low'
            
            return {
                'info_score': info_score,
                'experience_score': experience_score,
                'overall_level': level,
                'reasoning': f'Combined score: {combined_score:.1f}%'
            }
            
        except Exception as e:
            st.error(f"Error scoring applicant: {str(e)}")
            return {
                'info_score': 50,
                'experience_score': 50,
                'overall_level': 'Mid',
                'reasoning': 'Error in scoring - default values assigned'
            }
    
    def get_info_score(self, applicant: Dict) -> float:
        """Get information score using OpenAI"""
        try:
            prompt = f"""
            Rate the following applicant's basic information on a scale of 0-100:
            
            Name: {applicant['name']}
            Age: {applicant['age']}
            BMI: {applicant['bmi']}
            Education: {applicant['basic_info'].get('education', 'N/A')}
            Location: {applicant['basic_info'].get('location', 'N/A')}
            Skills: {applicant['basic_info'].get('skills', 'N/A')}
            
            Consider factors like:
            - Age appropriateness for the role
            - Educational background
            - Relevant skills
            - Overall profile completeness
            
            Respond with only a number between 0-100.
            """
            
            response = openai.ChatCompletion.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=10
            )
            
            score = float(response.choices[0].message.content.strip())
            return max(0, min(100, score))
            
        except Exception:
            return 70  # Default score if API fails
    
    def get_experience_score(self, experience: Dict) -> float:
        """Get experience score using OpenAI"""
        try:
            prompt = f"""
            Rate the following applicant's experience on a scale of 0-100:
            
            Years of Experience: {experience.get('years', 0)}
            Previous Roles: {experience.get('previous_roles', 'N/A')}
            Certifications: {experience.get('certifications', 'N/A')}
            
            Consider factors like:
            - Years of relevant experience
            - Quality of previous roles
            - Relevant certifications
            - Career progression
            
            Respond with only a number between 0-100.
            """
            
            response = openai.ChatCompletion.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=10
            )
            
            score = float(response.choices[0].message.content.strip())
            return max(0, min(100, score))
            
        except Exception:
            return 60  # Default score if API fails

def main():
    st.title("📊 Applicant Analysis System")
    st.markdown("วิเคราะห์ข้อมูลผู้สมัครจากไฟล์ Excel บน SharePoint พร้อม AI-based scoring")
    
    # Sidebar for settings
    st.sidebar.header("🔧 Settings")
    
    # OpenAI API Key input
    openai_api_key = st.sidebar.text_input(
        "OpenAI API Key",
        type="password",
        help="ใส่ OpenAI API key เพื่อใช้งานการวิเคราะห์ด้วย AI"
    )
    
    if not openai_api_key:
        st.warning("⚠️ กรุณาใส่ OpenAI API Key ในแถบด้านข้างเพื่อใช้งานระบบ")
        st.stop()
    
    # Initialize analyzer
    analyzer = ApplicantAnalyzer(openai_api_key)
    
    # Main interface
    tab1, tab2, tab3 = st.tabs(["📥 Data Input", "📊 Analysis Results", "📈 Statistics"])
    
    with tab1:
        st.header("📥 Data Input")
        st.markdown("อัปโหลดข้อมูลผู้สมัครจาก SharePoint หรือไฟล์ Excel")
        
        input_method = st.radio(
            "เลือกวิธีการนำเข้าข้อมูล:",
            ["SharePoint URL", "Upload Excel File"]
        )
        
        if input_method == "SharePoint URL":
            sharepoint_url = st.text_input(
                "SharePoint URL",
                placeholder="https://company.sharepoint.com/sites/hr/Documents/applicants.xlsx",
                help="URL ของไฟล์ Excel บน SharePoint"
            )
            
            if st.button("🔄 Fetch Data from SharePoint", type="primary"):
                if sharepoint_url:
                    with st.spinner("กำลังดาวน์โหลดและวิเคราะห์ข้อมูล..."):
                        # Download and parse Excel file
                        file_content = analyzer.download_excel_from_sharepoint(sharepoint_url)
                        
                        if file_content:
                            applicants = analyzer.parse_excel_file(file_content)
                            
                            if applicants:
                                # Score each applicant
                                scored_applicants = []
                                progress_bar = st.progress(0)
                                
                                for i, applicant in enumerate(applicants):
                                    scoring_result = analyzer.score_applicant(applicant)
                                    
                                    scored_applicant = {
                                        **applicant,
                                        'info_score': scoring_result['info_score'],
                                        'experience_score': scoring_result['experience_score'],
                                        'overall_level': scoring_result['overall_level'],
                                        'reasoning': scoring_result['reasoning'],
                                        'created_at': datetime.now().isoformat()
                                    }
                                    
                                    scored_applicants.append(scored_applicant)
                                    progress_bar.progress((i + 1) / len(applicants))
                                
                                # Save to session state
                                st.session_state.applicants = scored_applicants
                                
                                st.success(f"✅ วิเคราะห์ข้อมูลผู้สมัครเรียบร้อยแล้ว! จำนวน {len(scored_applicants)} คน")
                                
                else:
                    st.error("⚠️ กรุณาใส่ SharePoint URL")
        
        else:  # Upload Excel File
            uploaded_file = st.file_uploader(
                "อัปโหลดไฟล์ Excel",
                type=['xlsx', 'xls'],
                help="อัปโหลดไฟล์ Excel ที่มีข้อมูลผู้สมัคร"
            )
            
            if uploaded_file is not None:
                if st.button("🔄 Analyze Uploaded File", type="primary"):
                    with st.spinner("กำลังวิเคราะห์ข้อมูล..."):
                        # Parse uploaded file
                        applicants = analyzer.parse_excel_file(uploaded_file.read())
                        
                        if applicants:
                            # Score each applicant
                            scored_applicants = []
                            progress_bar = st.progress(0)
                            
                            for i, applicant in enumerate(applicants):
                                scoring_result = analyzer.score_applicant(applicant)
                                
                                scored_applicant = {
                                    **applicant,
                                    'info_score': scoring_result['info_score'],
                                    'experience_score': scoring_result['experience_score'],
                                    'overall_level': scoring_result['overall_level'],
                                    'reasoning': scoring_result['reasoning'],
                                    'created_at': datetime.now().isoformat()
                                }
                                
                                scored_applicants.append(scored_applicant)
                                progress_bar.progress((i + 1) / len(applicants))
                            
                            # Save to session state
                            st.session_state.applicants = scored_applicants
                            
                            st.success(f"✅ วิเคราะห์ข้อมูลผู้สมัครเรียบร้อยแล้ว! จำนวน {len(scored_applicants)} คน")
    
    with tab2:
        st.header("📊 Analysis Results")
        
        if not st.session_state.applicants:
            st.info("ไม่มีข้อมูลผู้สมัคร กรุณานำเข้าข้อมูลในแท็บ Data Input ก่อน")
        else:
            # Filters
            col1, col2 = st.columns(2)
            
            with col1:
                level_filter = st.selectbox(
                    "กรองตามระดับ:",
                    ["All Levels", "High", "Mid", "Low"]
                )
            
            with col2:
                search_term = st.text_input(
                    "ค้นหาผู้สมัคร:",
                    placeholder="ชื่อหรืออีเมล"
                )
            
            # Filter applicants
            filtered_applicants = st.session_state.applicants
            
            if level_filter != "All Levels":
                filtered_applicants = [
                    a for a in filtered_applicants 
                    if a['overall_level'] == level_filter
                ]
            
            if search_term:
                filtered_applicants = [
                    a for a in filtered_applicants 
                    if search_term.lower() in a['name'].lower() or 
                       search_term.lower() in a['email'].lower()
                ]
            
            # Display results
            st.subheader(f"ผลการวิเคราะห์ ({len(filtered_applicants)} คน)")
            
            for applicant in filtered_applicants:
                with st.expander(f"👤 {applicant['name']} - {applicant['overall_level']} Level"):
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.metric("Info Score", f"{applicant['info_score']:.1f}%")
                        st.write(f"**Email:** {applicant['email']}")
                        st.write(f"**Age:** {applicant.get('age', 'N/A')}")
                        st.write(f"**BMI:** {applicant['bmi']}")
                    
                    with col2:
                        st.metric("Experience Score", f"{applicant['experience_score']:.1f}%")
                        st.write(f"**Phone:** {applicant.get('phone', 'N/A')}")
                        st.write(f"**Height:** {applicant.get('height', 'N/A')} cm")
                        st.write(f"**Weight:** {applicant.get('weight', 'N/A')} kg")
                    
                    with col3:
                        level_color = {
                            'High': '🟢',
                            'Mid': '🟡',
                            'Low': '🔴'
                        }
                        st.metric("Overall Level", f"{level_color.get(applicant['overall_level'], '')} {applicant['overall_level']}")
                        st.write(f"**Reasoning:** {applicant['reasoning']}")
                    
                    # Action buttons
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button(f"📅 Schedule Teams Meeting", key=f"teams_{applicant['external_id']}"):
                            st.info("Teams meeting scheduling feature - integrate with Microsoft Graph API")
                    
                    with col2:
                        if st.button(f"✉️ Generate Email", key=f"email_{applicant['external_id']}"):
                            st.info("Email generation feature - integrate with email service")
    
    with tab3:
        st.header("📈 Statistics")
        
        if not st.session_state.applicants:
            st.info("ไม่มีข้อมูลสำหรับแสดงสถิติ")
        else:
            applicants = st.session_state.applicants
            
            # Calculate statistics
            total_applicants = len(applicants)
            high_level = len([a for a in applicants if a['overall_level'] == 'High'])
            mid_level = len([a for a in applicants if a['overall_level'] == 'Mid'])
            low_level = len([a for a in applicants if a['overall_level'] == 'Low'])
            
            # Display metrics
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Total Applicants", total_applicants)
            
            with col2:
                st.metric("High Level", high_level, f"{high_level/total_applicants*100:.1f}%")
            
            with col3:
                st.metric("Mid Level", mid_level, f"{mid_level/total_applicants*100:.1f}%")
            
            with col4:
                st.metric("Low Level", low_level, f"{low_level/total_applicants*100:.1f}%")
            
            # Charts
            col1, col2 = st.columns(2)
            
            with col1:
                # Level distribution pie chart
                level_data = pd.DataFrame({
                    'Level': ['High', 'Mid', 'Low'],
                    'Count': [high_level, mid_level, low_level]
                })
                
                st.subheader("Level Distribution")
                st.bar_chart(level_data.set_index('Level'))
            
            with col2:
                # BMI distribution
                bmi_data = [a['bmi'] for a in applicants if a['bmi'] > 0]
                
                st.subheader("BMI Distribution")
                st.histogram(bmi_data, bins=20)
            
            # Export functionality
            st.subheader("📥 Export Data")
            
            if st.button("📊 Export to Excel"):
                # Create DataFrame
                df = pd.DataFrame(applicants)
                
                # Convert to Excel
                buffer = io.BytesIO()
                df.to_excel(buffer, index=False)
                buffer.seek(0)
                
                st.download_button(
                    label="📥 Download Excel File",
                    data=buffer,
                    file_name=f"applicant_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.ms-excel"
                )

if __name__ == "__main__":
    main()
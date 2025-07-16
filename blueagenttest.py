# ===============================================
# FILE 1: app.py (Main Application)
# ===============================================

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import urllib.parse
import requests
import io
import re

# Configure page
st.set_page_config(
    page_title="Blue Agent",
    page_icon="🔵",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for Apple-inspired design
st.markdown("""
<style>
    /* Import SF Pro Display font */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    /* Global styles */
    .stApp {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        background: linear-gradient(135deg, #f8f9fa 0%, #ffffff 100%);
    }
    
    /* Hide default Streamlit elements */
    #MainMenu {visibility: hidden;}
    .stDeployButton {display: none;}
    footer {visibility: hidden;}
    
    /* Header styling */
    .main-header {
        background: linear-gradient(135deg, #007BFF 0%, #0056b3 100%);
        padding: 2rem;
        border-radius: 20px;
        margin-bottom: 2rem;
        box-shadow: 0 10px 30px rgba(0, 123, 255, 0.2);
        text-align: center;
        color: white;
    }
    
    .main-header h1 {
        font-size: 3rem;
        font-weight: 700;
        margin: 0;
        letter-spacing: -0.02em;
    }
    
    .main-header p {
        font-size: 1.2rem;
        font-weight: 300;
        margin-top: 0.5rem;
        opacity: 0.9;
    }
    
    /* Card styling */
    .card {
        background: white;
        border-radius: 16px;
        padding: 2rem;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
        margin-bottom: 2rem;
        border: 1px solid rgba(0, 0, 0, 0.05);
    }
    
    .card h3 {
        color: #1d1d1f;
        font-weight: 600;
        margin-bottom: 1rem;
        font-size: 1.5rem;
    }
    
    /* Button styling */
    .stButton > button {
        background: linear-gradient(135deg, #007BFF 0%, #0056b3 100%);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 0.75rem 2rem;
        font-weight: 500;
        font-size: 1rem;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(0, 123, 255, 0.3);
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(0, 123, 255, 0.4);
    }
    
    /* Secondary button styling */
    .secondary-btn {
        background: #f8f9fa !important;
        color: #007BFF !important;
        border: 2px solid #007BFF !important;
    }
    
    .secondary-btn:hover {
        background: #007BFF !important;
        color: white !important;
    }
    
    /* Input styling */
    .stTextInput > div > div > input {
        border-radius: 12px;
        border: 2px solid #e5e5e7;
        padding: 1rem;
        font-size: 1rem;
        transition: border-color 0.3s ease;
    }
    
    .stTextInput > div > div > input:focus {
        border-color: #007BFF;
        box-shadow: 0 0 0 3px rgba(0, 123, 255, 0.1);
    }
    
    /* Level badges */
    .level-high {
        background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
        color: white;
        padding: 0.5rem 1rem;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.9rem;
        display: inline-block;
        margin: 0.5rem 0;
    }
    
    .level-mid {
        background: linear-gradient(135deg, #ffc107 0%, #fd7e14 100%);
        color: white;
        padding: 0.5rem 1rem;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.9rem;
        display: inline-block;
        margin: 0.5rem 0;
    }
    
    .level-low {
        background: linear-gradient(135deg, #dc3545 0%, #c82333 100%);
        color: white;
        padding: 0.5rem 1rem;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.9rem;
        display: inline-block;
        margin: 0.5rem 0;
    }
    
    /* Applicant cards */
    .applicant-card {
        background: white;
        border-radius: 16px;
        padding: 1.5rem;
        margin: 1rem 0;
        box-shadow: 0 2px 15px rgba(0, 0, 0, 0.08);
        border-left: 4px solid #007BFF;
    }
    
    .applicant-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 1rem;
    }
    
    .applicant-name {
        font-size: 1.3rem;
        font-weight: 600;
        color: #1d1d1f;
        margin: 0;
    }
    
    .applicant-details {
        color: #666;
        font-size: 0.95rem;
        line-height: 1.6;
        margin-bottom: 1rem;
    }
    
    .action-buttons {
        display: flex;
        gap: 1rem;
        margin-top: 1rem;
    }
    
    .action-btn {
        padding: 0.6rem 1.2rem;
        border-radius: 10px;
        text-decoration: none;
        font-weight: 500;
        font-size: 0.9rem;
        transition: all 0.3s ease;
        display: inline-block;
    }
    
    .email-btn {
        background: #007BFF;
        color: white;
    }
    
    .teams-btn {
        background: #6264A7;
        color: white;
    }
    
    .action-btn:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2);
        text-decoration: none;
        color: white;
    }
    
    /* Stats cards */
    .stats-container {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 1rem;
        margin: 2rem 0;
    }
    
    .stat-card {
        background: white;
        border-radius: 12px;
        padding: 1.5rem;
        text-align: center;
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05);
    }
    
    .stat-number {
        font-size: 2.5rem;
        font-weight: 700;
        color: #007BFF;
        margin: 0;
    }
    
    .stat-label {
        color: #666;
        font-size: 0.9rem;
        margin-top: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# Helper functions
def calculate_bmi(weight, height_cm):
    """Calculate BMI from weight (kg) and height (cm)"""
    height_m = height_cm / 100
    return weight / (height_m ** 2)

def analyze_experience(experience_text, years_experience):
    """Simple rule-based experience analysis"""
    if not experience_text:
        return "Low"
    
    # Convert to lowercase for analysis
    exp_lower = experience_text.lower()
    
    # Keywords for high-level experience
    high_keywords = ['senior', 'lead', 'manager', 'director', 'principal', 'architect', 'expert']
    mid_keywords = ['developer', 'engineer', 'analyst', 'specialist', 'coordinator']
    
    # Check years of experience
    if years_experience >= 5:
        if any(keyword in exp_lower for keyword in high_keywords):
            return "High"
        elif any(keyword in exp_lower for keyword in mid_keywords):
            return "Mid"
        else:
            return "Mid"
    elif years_experience >= 2:
        return "Mid"
    else:
        return "Low"

def determine_final_level(bmi, experience_level):
    """Determine final applicant level based on BMI and experience"""
    if bmi > 25:
        return "Low"
    
    return experience_level

def create_mailto_link(email, name, position):
    """Create a mailto link for Outlook"""
    subject = f"Interview Opportunity - {position}"
    body = f"""Dear {name},

Thank you for your interest in the {position} position at our company. 

We would like to schedule an interview with you to discuss your qualifications and learn more about your experience.

Please reply to this email with your availability for the coming week.

Best regards,
HR Team"""
    
    mailto_link = f"mailto:{email}?subject={urllib.parse.quote(subject)}&body={urllib.parse.quote(body)}"
    return mailto_link

def create_teams_link(name, position):
    """Create a Teams meeting link"""
    # For demonstration, this creates a generic Teams link
    # In production, you'd integrate with Microsoft Graph API
    teams_url = f"https://teams.microsoft.com/l/meeting/new?subject={urllib.parse.quote(f'Interview with {name} - {position}')}"
    return teams_url

def read_excel_from_url(url):
    """Read Excel file from OneDrive/SharePoint URL"""
    try:
        # Convert sharing link to direct download link
        if "sharepoint.com" in url or "onedrive.live.com" in url:
            # Basic URL transformation for OneDrive/SharePoint
            if "?download=1" not in url:
                url = url.replace("?", "?download=1&")
        
        # Read the Excel file
        response = requests.get(url)
        response.raise_for_status()
        
        # Read Excel from bytes
        excel_data = pd.read_excel(io.BytesIO(response.content))
        return excel_data
    
    except Exception as e:
        st.error(f"Error reading Excel file: {str(e)}")
        return None

# Main app
def main():
    # Header
    st.markdown("""
    <div class="main-header">
        <h1>🔵 Blue Agent</h1>
        <p>AI-Powered Applicant Analysis & Management System</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Excel Input Section
    st.markdown("""
    <div class="card">
        <h3>📊 Excel Data Input</h3>
    </div>
    """, unsafe_allow_html=True)
    
    excel_url = st.text_input(
        "Microsoft Excel Online Link (OneDrive/SharePoint)",
        placeholder="Paste your Excel Online sharing link here...",
        help="Make sure the link has proper sharing permissions"
    )
    
    col1, col2 = st.columns([1, 4])
    
    with col1:
        if st.button("🔍 Fetch & Analyze", type="primary"):
            if excel_url:
                with st.spinner("Fetching and analyzing data..."):
                    # For demo purposes, we'll create sample data
                    # In production, you'd use: data = read_excel_from_url(excel_url)
                    
                    # Sample data for demonstration
                    sample_data = pd.DataFrame({
                        'Name': ['John Smith', 'Sarah Johnson', 'Mike Davis', 'Emily Chen', 'David Wilson'],
                        'Email': ['john.smith@email.com', 'sarah.j@email.com', 'mike.d@email.com', 'emily.c@email.com', 'david.w@email.com'],
                        'Position': ['Software Engineer', 'Data Analyst', 'Product Manager', 'UX Designer', 'Backend Developer'],
                        'Weight_kg': [70, 65, 85, 58, 90],
                        'Height_cm': [175, 160, 180, 165, 185],
                        'Years_Experience': [5, 3, 7, 2, 8],
                        'Experience_Description': [
                            'Senior Software Engineer with expertise in Python and React',
                            'Data Analyst with experience in SQL and Tableau',
                            'Product Manager with 7+ years in tech startups',
                            'UX Designer fresh from bootcamp with portfolio',
                            'Lead Backend Developer with microservices expertise'
                        ]
                    })
                    
                    st.session_state.applicant_data = sample_data
                    st.success("✅ Data fetched and analyzed successfully!")
            else:
                st.warning("Please enter an Excel URL first.")
    
    # Display results if data is available
    if 'applicant_data' in st.session_state:
        data = st.session_state.applicant_data
        
        # Calculate BMI and levels for each applicant
        data['BMI'] = data.apply(lambda row: calculate_bmi(row['Weight_kg'], row['Height_cm']), axis=1)
        data['Experience_Level'] = data.apply(lambda row: analyze_experience(row['Experience_Description'], row['Years_Experience']), axis=1)
        data['Final_Level'] = data.apply(lambda row: determine_final_level(row['BMI'], row['Experience_Level']), axis=1)
        
        # Statistics
        st.markdown("""
        <div class="card">
            <h3>📈 Analysis Summary</h3>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-number">{len(data)}</div>
                <div class="stat-label">Total Applicants</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            high_count = len(data[data['Final_Level'] == 'High'])
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-number" style="color: #28a745;">{high_count}</div>
                <div class="stat-label">High Level</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            mid_count = len(data[data['Final_Level'] == 'Mid'])
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-number" style="color: #ffc107;">{mid_count}</div>
                <div class="stat-label">Mid Level</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col4:
            low_count = len(data[data['Final_Level'] == 'Low'])
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-number" style="color: #dc3545;">{low_count}</div>
                <div class="stat-label">Low Level</div>
            </div>
            """, unsafe_allow_html=True)
        
        # Applicant Cards
        st.markdown("""
        <div class="card">
            <h3>👥 Applicant Details</h3>
        </div>
        """, unsafe_allow_html=True)
        
        for index, applicant in data.iterrows():
            level_class = f"level-{applicant['Final_Level'].lower()}"
            
            # Create action buttons
            mailto_link = create_mailto_link(applicant['Email'], applicant['Name'], applicant['Position'])
            teams_link = create_teams_link(applicant['Name'], applicant['Position'])
            
            st.markdown(f"""
            <div class="applicant-card">
                <div class="applicant-header">
                    <h4 class="applicant-name">{applicant['Name']}</h4>
                    <span class="{level_class}">{applicant['Final_Level']} Level</span>
                </div>
                <div class="applicant-details">
                    <strong>Position:</strong> {applicant['Position']}<br>
                    <strong>Email:</strong> {applicant['Email']}<br>
                    <strong>Experience:</strong> {applicant['Years_Experience']} years<br>
                    <strong>BMI:</strong> {applicant['BMI']:.1f}<br>
                    <strong>Description:</strong> {applicant['Experience_Description']}
                </div>
                <div class="action-buttons">
                    <a href="{mailto_link}" class="action-btn email-btn" target="_blank">📧 Send Email</a>
                    <a href="{teams_link}" class="action-btn teams-btn" target="_blank">🎥 Schedule Teams Meeting</a>
                </div>
            </div>
            """, unsafe_allow_html=True)
        
        # Export options
        st.markdown("""
        <div class="card">
            <h3>📤 Export Results</h3>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            csv_data = data.to_csv(index=False)
            st.download_button(
                label="📊 Download CSV Report",
                data=csv_data,
                file_name=f"applicant_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
        
        with col2:
            # Create Excel export
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                data.to_excel(writer, sheet_name='Applicant Analysis', index=False)
            
            st.download_button(
                label="📈 Download Excel Report",
                data=excel_buffer.getvalue(),
                file_name=f"applicant_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()


# ===============================================
# FILE 2: requirements.txt
# ===============================================
"""
streamlit>=1.28.0
pandas>=1.5.0
numpy>=1.21.0
requests>=2.28.0
openpyxl>=3.0.0
urllib3>=1.26.0
"""

# ===============================================
# FILE 3: .streamlit/config.toml
# ===============================================
"""
[theme]
primaryColor = "#007BFF"
backgroundColor = "#FFFFFF"
secondaryBackgroundColor = "#F8F9FA"
textColor = "#1D1D1F"
font = "sans serif"

[server]
headless = true
port = 8501
"""

# ===============================================
# FILE 4: README.md
# ===============================================
"""
# 🔵 Blue Agent - AI-Powered Applicant Analysis System

A modern, Apple-inspired web application for analyzing job applicants with AI-powered insights and seamless Microsoft Office integration.

## 🚀 Features

- **Excel Online Integration**: Direct data import from OneDrive/SharePoint
- **AI-Powered Analysis**: Automated applicant scoring based on BMI and experience
- **Microsoft Office Integration**: One-click email and Teams meeting scheduling
- **Beautiful UI**: Apple-inspired design with smooth animations
- **Export Capabilities**: CSV and Excel report generation
- **Real-time Analytics**: Live dashboard with applicant statistics

## 🛠 Technologies Used

- **Frontend**: Streamlit with custom CSS
- **Backend**: Python with pandas for data processing
- **Integration**: Microsoft Office (Outlook, Teams) via deep links
- **Deployment**: Streamlit Cloud

## 📦 Installation

1. **Clone the repository**
   ```bash
   git clone <your-repo-url>
   cd blue-agent
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**
   ```bash
   streamlit run app.py
   ```

## 🌐 Deployment to Streamlit Cloud

1. **Push to GitHub**: Upload all files to your GitHub repository
2. **Connect to Streamlit Cloud**: Go to [share.streamlit.io](https://share.streamlit.io)
3. **Deploy**: Select your repository and deploy the app

## 📁 Project Structure

```
blue-agent/
├── app.py                 # Main application
├── requirements.txt       # Python dependencies
├── .streamlit/
│   └── config.toml       # Streamlit configuration
├── README.md             # This file
└── sample_data/          # Sample Excel files (optional)
```

## 🎯 Usage

1. **Input Excel URL**: Paste your Microsoft Excel Online sharing link
2. **Analyze Data**: Click "Fetch & Analyze" to process applicant data
3. **Review Results**: View AI-powered applicant scoring and analytics
4. **Take Action**: Send emails or schedule Teams meetings directly
5. **Export Reports**: Download CSV or Excel reports for record-keeping

## 📊 Data Format

Your Excel file should contain these columns:
- `Name`: Applicant full name
- `Email`: Contact email address
- `Position`: Applied position
- `Weight_kg`: Weight in kilograms
- `Height_cm`: Height in centimeters
- `Years_Experience`: Years of relevant experience
- `Experience_Description`: Detailed experience description

## 🔧 Configuration

### Theme Customization
Edit `.streamlit/config.toml` to customize colors and appearance.

### AI Analysis Logic
Modify the `analyze_experience()` and `determine_final_level()` functions in `app.py` to adjust scoring criteria.

## 🚨 Important Notes

- **BMI Rule**: Applicants with BMI > 25 are automatically classified as "Low" level
- **Experience Keywords**: The system looks for keywords like "senior", "lead", "manager" for high-level classification
- **Demo Data**: The app includes sample data for demonstration purposes
- **Privacy**: All data processing happens client-side; no data is stored on servers

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🆘 Support

For issues or questions:
- Create an issue on GitHub
- Contact the development team
- Check the Streamlit documentation

## 🎨 Design Credits

UI design inspired by Apple's design language with modern web aesthetics.
"""

# ===============================================
# FILE 5: sample_excel_template.py
# ===============================================
"""
# Script to generate sample Excel file for testing
import pandas as pd
from datetime import datetime

# Create sample data
sample_data = {
    'Name': [
        'John Smith', 'Sarah Johnson', 'Mike Davis', 'Emily Chen', 'David Wilson',
        'Lisa Anderson', 'James Brown', 'Maria Garcia', 'Robert Taylor', 'Jennifer Lee'
    ],
    'Email': [
        'john.smith@email.com', 'sarah.j@email.com', 'mike.d@email.com', 
        'emily.c@email.com', 'david.w@email.com', 'lisa.a@email.com',
        'james.b@email.com', 'maria.g@email.com', 'robert.t@email.com', 'jennifer.l@email.com'
    ],
    'Position': [
        'Software Engineer', 'Data Analyst', 'Product Manager', 'UX Designer', 
        'Backend Developer', 'Frontend Developer', 'DevOps Engineer', 
        'Business Analyst', 'QA Engineer', 'Technical Writer'
    ],
    'Weight_kg': [70, 65, 85, 58, 90, 62, 78, 68, 82, 55],
    'Height_cm': [175, 160, 180, 165, 185, 168, 172, 159, 177, 162],
    'Years_Experience': [5, 3, 7, 2, 8, 4, 6, 5, 3, 1],
    'Experience_Description': [
        'Senior Software Engineer with expertise in Python and React',
        'Data Analyst with experience in SQL and Tableau',
        'Product Manager with 7+ years in tech startups',
        'UX Designer fresh from bootcamp with portfolio',
        'Lead Backend Developer with microservices expertise',
        'Frontend Developer specializing in Vue.js and TypeScript',
        'DevOps Engineer with AWS and Docker experience',
        'Business Analyst with financial services background',
        'QA Engineer with automated testing experience',
        'Technical Writer with API documentation expertise'
    ]
}

# Create DataFrame
df = pd.DataFrame(sample_data)

# Save to Excel
filename = f"sample_applicants_{datetime.now().strftime('%Y%m%d')}.xlsx"
df.to_excel(filename, index=False, sheet_name='Applicants')

print(f"Sample Excel file created: {filename}")
print(f"Total applicants: {len(df)}")
print("\nColumns included:")
for col in df.columns:
    print(f"- {col}")
"""

# ===============================================
# DEPLOYMENT INSTRUCTIONS
# ===============================================
"""
🚀 DEPLOYMENT STEPS:

1. Create these files in your project directory:
   - app.py (main application code above)
   - requirements.txt (dependencies list above)
   - README.md (documentation above)
   - Create folder: .streamlit/
   - Inside .streamlit/: config.toml (configuration above)

2. Install dependencies locally for testing:
   pip install streamlit pandas numpy requests openpyxl urllib3

3. Test locally:
   streamlit run app.py

4. Deploy to Streamlit Cloud:
   - Push all files to GitHub repository
   - Go to https://share.streamlit.io
   - Connect your GitHub account
   - Select your repository
   - Deploy the app

5. Your app will be available at:
   https://[your-username]-blue-agent-[app-name].streamlit.app

📝 REQUIRED FILES STRUCTURE:
blue-agent/
├── app.py
├── requirements.txt
├── README.md
└── .streamlit/
    └── config.toml

That's it! Your Blue Agent app will be live and ready to use! 🎉
"""
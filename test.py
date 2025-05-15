import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# ----------------------------
# Core Excel Processing Functions
# ----------------------------

def remove_duplicates(source_df):
    duplicates = source_df.duplicated(subset=['Revised Copy'], keep='first')
    if duplicates.any():
        st.info("Duplicates found in Source Sheet. Removing duplicates...")
        source_df = source_df.drop_duplicates(subset=['Revised Copy'], keep='first')
    else:
        st.info("No duplicates found in 'Revised Copy'.")
    return source_df

def map_content(source_df, site_df):
    source_df = remove_duplicates(source_df)
    source_dict = {row['Design Copy']: (index, row['Revised Copy']) 
                   for index, row in source_df.iterrows()}

    revised_copy_col_index = source_df.columns.get_loc('Revised Copy')

    for col in ['Revised Copy', 'Mapped Cell']:
        if col not in site_df.columns:
            site_df[col] = ''

    for index, row in site_df.iterrows():
        design_copy = row['Design Copy']
        if design_copy in source_dict:
            source_index, revised_copy = source_dict[design_copy]
            source_cell = f"Source!{chr(65 + revised_copy_col_index)}{source_index + 2}"
            site_df.at[index, 'Mapped Cell'] = source_cell
            site_df.at[index, 'Revised Copy'] = str(revised_copy)
        else:
            site_df.at[index, 'Mapped Cell'] = 'Not Found'

    return site_df

def structure_and_format_data(raw_data, group_column='frame'):
    structured_data = pd.DataFrame(columns=raw_data.columns)
    
    for name, group in raw_data.groupby(group_column, sort=False):
        # Fixed syntax error in this line
        title_row = pd.DataFrame(
            [[name] + [''] * (len(raw_data.columns) - 1)],  # Added missing parenthesis
            columns=raw_data.columns
        )
        
        structured_data = pd.concat([
            structured_data,
            title_row,
            pd.DataFrame([[''] * len(raw_data.columns)], columns=raw_data.columns),  # <-- fix here
            group,
            pd.DataFrame([[''] * len(raw_data.columns)], columns=raw_data.columns)   # <-- and here
        ], ignore_index=True)

    return structured_data.iloc[:-1]

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        for row in worksheet.iter_rows():
            if row[0].value in df['frame'].dropna().unique():
                for cell in row:
                    cell.fill = fill
    return output.getvalue()

# ----------------------------
# Feedback System
# ----------------------------

def init_google_sheets():
    """Initialize Google Sheets connection with modern auth"""
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    return gspread.authorize(
        Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=scopes
        )
    )

def store_feedback(rating, comments):
    """Store feedback in Google Sheets"""
    try:
        client = init_google_sheets()
        sheet = client.open("App Feedback").sheet1
        
        user_email = st.user.email if st.user else 'Anonymous'
            
        sheet.append_row([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            str(rating),
            comments,
            user_email
        ])
        return True
    except Exception as e:
        st.error(f"Error storing feedback: {str(e)}")
        return False

# ----------------------------
# Streamlit UI
# ----------------------------

st.title("📄 Excel Content Mapper")

# File Processing Section
with st.expander("📤 Upload and Process Files", expanded=True):
    st.write("""
    1. **Source File** - Contains original and revised copies
    2. **Site File** - File to be updated with mapping information
    """)
    
    source_file = st.file_uploader("Source Excel File", type=['xlsx'])
    site_file = st.file_uploader("Site Excel File", type=['xlsx'])

    if source_file and site_file and st.button("🚀 Process Files"):
        try:
            source_df = pd.read_excel(source_file)
            site_df = pd.read_excel(site_file, dtype=str)
            
            with st.spinner('🔍 Mapping content...'):
                mapped_df = map_content(source_df, site_df)
                st.session_state.mapped_df = mapped_df
                st.success("✅ Mapping completed!")
                st.dataframe(mapped_df.head(10))
                
            with st.spinner('📊 Structuring data...'):
                structured_df = structure_and_format_data(mapped_df)
                st.session_state.structured_df = structured_df
                st.success("✅ Structuring completed!")
                st.dataframe(structured_df.head(10))
                
            st.download_button(
                "💾 Download Result",
                to_excel(structured_df),
                file_name=f"Mapped_{site_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")

# Feedback Section
with st.expander("💬 Provide Feedback", expanded=True):
    if 'feedback_submitted' not in st.session_state:
        st.session_state.feedback_submitted = False

    if not st.session_state.feedback_submitted:
        with st.form("feedback_form"):
            comments = st.text_area("Your comments/suggestions")
            
            if st.form_submit_button("📤 Submit Feedback"):
                    if store_feedback("", comments):
                        st.session_state.feedback_submitted = True
                        st.success("🎉 Thank you for your feedback!")
                    else:
                        st.error("❌ Failed to submit feedback")
    else:
        st.success("✅ Feedback submitted successfully!")
        if st.button("📝 Submit New Feedback"):
            st.session_state.feedback_submitted = False
            st.rerun()

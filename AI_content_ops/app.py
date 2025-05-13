import streamlit as st
import pandas as pd
from io import BytesIO

def remove_duplicates(source_df):
    duplicates = source_df.duplicated(subset=['Revised Copy'], keep='first')
    if duplicates.any():
        st.info("Duplicates found in 'Revised Copy'. Removing duplicates...")
        source_df = source_df.drop_duplicates(subset=['Revised Copy'], keep='first')
    else:
        st.info("No duplicates found in 'Revised Copy'.")
    return source_df

def map_content(source_df, site_df):
    source_df = remove_duplicates(source_df)
    source_dict = {row['Design Copy']: (index, row['Revised Copy']) for index, row in source_df.iterrows()}

    # Initialize columns if they don't exist
    if 'Mapped Cell' not in site_df.columns:
        site_df['Mapped Cell'] = ''
    if 'Revised Copy' not in site_df.columns:
        site_df['Revised Copy'] = ''

    for index, row in site_df.iterrows():
        design_copy = row['Design Copy']
        if design_copy in source_dict:
            source_index, revised_copy = source_dict[design_copy]
            site_df.at[index, 'Mapped Cell'] = f"Source!A{source_index + 2}"
            site_df.at[index, 'Revised Copy'] = str(revised_copy)
        else:
            site_df.at[index, 'Mapped Cell'] = 'Not Found'
    
    return site_df

def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False)
    writer.close()
    processed_data = output.getvalue()
    return processed_data

# Streamlit UI
st.title("Excel Content Mapper")

st.write("""
### Upload your files
1. **Source File** - Contains the original and revised copies
2. **Site File** - File to be updated with mapping information
""")

source_file = st.file_uploader("Upload Source Excel File", type=['xlsx'])
site_file = st.file_uploader("Upload Site Excel File", type=['xlsx'])

if source_file and site_file:
    try:
        # Read files
        source_df = pd.read_excel(source_file)
        site_df = pd.read_excel(site_file)

        # Process files
        with st.spinner('Processing files...'):
            updated_site_df = map_content(source_df, site_df)

        # Download button
        st.success("File processing completed!")
        
        st.write("### Preview of Updated Data")
        st.dataframe(updated_site_df.head())
        
        st.write("### Download Updated File")
        excel_data = to_excel(updated_site_df)
        st.download_button(
            label="Download Updated Site File",
            data=excel_data,
            file_name="updated_site_file.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    except KeyError as e:
        st.error(f"Missing required column: {e}")
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
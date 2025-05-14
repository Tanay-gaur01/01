# app.py
import streamlit as st
import pandas as pd
from utils import remove_duplicates, map_content, structure_and_format_data, to_excel

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
    # Read files
    source_df = pd.read_excel(source_file)
    site_df = pd.read_excel(site_file)

    # Button to structure the data
    if st.button("Structure Data"):
        try:
            with st.spinner('Structuring files...'):
                structured_site_df = structure_and_format_data(site_df, group_column='frame')

            st.success("File structuring completed!")
            st.write("### Preview of Structured Data")
            st.dataframe(structured_site_df.head())
            st.session_state['structured_site_df'] = structured_site_df

        except KeyError as e:
            st.error(f"Missing required column: {e}")
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")

    # Button to start mapping
    if st.button("Start Mapping") and 'structured_site_df' in st.session_state:
        try:
            with st.spinner('Mapping files...'):
                updated_site_df = map_content(source_df, st.session_state['structured_site_df'])

            st.success("File mapping completed!")
            st.write("### Preview of Mapped Data")
            st.dataframe(updated_site_df.head())

            original_filename = site_file.name.replace('.xlsx', '')
            output_filename = f"Mapped_{original_filename}.xlsx"

            st.write("### Download Final File")
            excel_data = to_excel(updated_site_df)
            st.download_button(
                label="Download Final File",
                data=excel_data,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except KeyError as e:
            st.error(f"Missing required column: {e}")
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")

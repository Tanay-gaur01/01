import streamlit as st
import pandas as pd
from utils import map_content, structure_and_format_data, to_excel

st.title("Excel Content Mapper")

st.write("""
### Upload your files
1. **Source File** - Contains original and revised copies
2. **Site File** - File to be updated with mapping information
""")

source_file = st.file_uploader("Upload Source Excel File", type=['xlsx'])
site_file = st.file_uploader("Upload Site Excel File", type=['xlsx'])

if source_file and site_file:
    source_df = pd.read_excel(source_file)
    site_df = pd.read_excel(site_file, dtype=str)

    if st.button("Map Data"):
        try:
            with st.spinner('Mapping content...'):
                mapped_df = map_content(source_df, site_df)
                st.session_state.mapped_df = mapped_df
            
            st.success("Mapping completed!")
            st.write("### Mapped Data Preview")
            st.dataframe(mapped_df.head(10))
            
        except KeyError as e:
            st.error(f"Missing required column: {e}")
        except Exception as e:
            st.error(f"Error during mapping: {str(e)}")

    if st.button("Structure Data") and 'mapped_df' in st.session_state:
        try:
            with st.spinner('Structuring data...'):
                structured_df = structure_and_format_data(st.session_state.mapped_df)
                st.session_state.structured_df = structured_df
            
            st.success("Structuring completed!")
            st.write("### Structured Data Preview")
            st.dataframe(structured_df.head(15))
            
        except KeyError as e:
            st.error(f"Missing required column: {e}")
        except Exception as e:
            st.error(f"Error during structuring: {str(e)}")

    if 'structured_df' in st.session_state:
        st.write("### Download Final File")
        output = to_excel(st.session_state.structured_df)
        
        original_name = site_file.name.split('.')[0]
        filename = f"{original_name}_Mapped_Structured.xlsx"
        
        st.download_button(
            label="Download Processed File",
            data=output,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

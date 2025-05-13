import pandas as pd

def remove_duplicates(source_df):
    # Check for duplicates in the 'Revised Copy' column
    duplicates = source_df.duplicated(subset=['Revised Copy'], keep='first')

    # Remove duplicates
    if duplicates.any():
        print("Duplicates found in 'Revised Copy'. Removing duplicates...")
        source_df = source_df.drop_duplicates(subset=['Revised Copy'], keep='first')
    else:
        print("No duplicates found in 'Revised Copy'.")

    return source_df

def map_content(source_path, site_path):
    # Load the source and site Excel files
    source_df = pd.read_excel(source_path)
    site_df = pd.read_excel(site_path)

    # Remove duplicates from the source DataFrame
    source_df = remove_duplicates(source_df)

    # Create dictionaries for quick lookup
    source_dict = {row['Design Copy']: (index, row['Revised Copy']) for index, row in source_df.iterrows()}

    # Initialize new columns in the site DataFrame if they don't exist
    if 'Mapped Cell' not in site_df.columns:
        site_df['Mapped Cell'] = site_df['Mapped Cell'].astype(str)
    if 'Revised Copy' not in site_df.columns:
        site_df['Revised Copy'] = site_df['Revised Copy'].astype(str)

    # Iterate over the site DataFrame and fill in the new columns
    for index, row in site_df.iterrows():
        design_copy = row['Design Copy']
        if design_copy in source_dict:
            source_index, revised_copy = source_dict[design_copy]
            site_df.at[index, 'Mapped Cell'] = f"Source!A{source_index + 2}"  # Assuming the data starts from row 2
            site_df.at[index, 'Revised Copy'] = str(revised_copy)
        else:
            site_df.at[index, 'Mapped Cell'] = 'Not Found'

    # Save the updated site DataFrame back to the same Excel file
    site_df.to_excel(site_path, index=False)

# Prompt the user for file paths
source_path = input("Enter the path to the source Excel file: ")
site_path = input("Enter the path to the site Excel file: ")

# Call the function with the provided file paths
map_content(source_path, site_path)

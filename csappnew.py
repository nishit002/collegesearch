import streamlit as st
import pandas as pd
from docx import Document
import os

# Title of the app
st.title("Comprehensive College Content Generation App")

# File upload
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

# Initialize empty data variable
data_sheets = {}

if uploaded_file is not None:
    # Read all sheets from the Excel file
    data_sheets = pd.read_excel(uploaded_file, sheet_name=None)

    # Assuming the primary college data is in the first sheet (Adjust as necessary)
    primary_sheet = list(data_sheets.keys())[0]
    df_colleges = data_sheets[primary_sheet]
    
    # Print column names for debugging
    st.write("Available columns:", df_colleges.columns.tolist())
    
    # Assuming the correct column name for college names
    college_name_column = 'College Name'  # Modify this based on the actual column name

    if college_name_column in df_colleges.columns:
        # Display the college names for selection
        college_names = df_colleges[college_name_column].unique().tolist()
        
        selected_colleges = st.multiselect("Select Colleges", college_names)
        
        # Button to generate the content
        if st.button("Generate Content"):
            if selected_colleges:
                # Iterate through selected colleges and generate articles
                for college in selected_colleges:
                    # Filter the data for the selected college
                    row = df_colleges[df_colleges[college_name_column] == college].iloc[0]
                    
                    # Create a Word document for the selected college
                    doc = Document()
                    
                    # Add college details based on the template from all sheets
                    doc.add_heading(row[college_name_column], 0)

                    if 'Establishment year' in row and not pd.isna(row['Establishment year']):
                        doc.add_paragraph(f"{row[college_name_column]} was established in {row['Establishment year']}.")

                    if 'College City' in row and 'College State' in row and not pd.isna(row['College City']) and not pd.isna(row['College State']):
                        doc.add_paragraph(f"The college is located in {row['College City']}, {row['College State']}.")

                    if 'College USP' in row and not pd.isna(row['College USP']):
                        doc.add_paragraph(f"The college is known for: {row['College USP']}.")

                    if 'NIRF RANK' in row and not pd.isna(row['NIRF RANK']):
                        doc.add_paragraph(f"Ranked {row['NIRF RANK']} in the NIRF rankings.")
                    
                    # Use data from other sheets (like placement, awards, faculty)
                    if 'Placements' in data_sheets:
                        df_placements = data_sheets['Placements']
                        placement_row = df_placements[df_placements[college_name_column] == college]
                        if not placement_row.empty:
                            placement_info = placement_row.iloc[0]
                            doc.add_heading(f"{college} Placements", level=1)
                            doc.add_paragraph(f"Highest package: INR {placement_info['Highest Package']} and Average package: INR {placement_info['Average Package']}.")

                    if 'Awards' in data_sheets:
                        df_awards = data_sheets['Awards']
                        award_row = df_awards[df_awards[college_name_column] == college]
                        if not award_row.empty:
                            award_info = award_row.iloc[0]
                            doc.add_heading(f"{college} Awards", level=1)
                            doc.add_paragraph(f"{award_info['Award']} awarded by {award_info['Awarding Authority']} in {award_info['Award Year']}.")

                    if 'Faculty' in data_sheets:
                        df_faculty = data_sheets['Faculty']
                        faculty_rows = df_faculty[df_faculty[college_name_column] == college]
                        if not faculty_rows.empty:
                            doc.add_heading(f"{college} Faculty", level=1)
                            for _, faculty in faculty_rows.iterrows():
                                doc.add_paragraph(f"{faculty['Faculty Name']}: {faculty['Specialty']} ({faculty['Education']})")
                    
                    if 'College Address' in row and not pd.isna(row['College Address']):
                        doc.add_heading(f"{college} Address", level=1)
                        doc.add_paragraph(f"{row['College Address']}")

                    if 'College Phone Number' in row or 'College Email' in row:
                        doc.add_paragraph(f"Contact the college at {row['College Phone Number']} or via email at {row['College Email']}.")

                    if 'College Website' in row and not pd.isna(row['College Website']):
                        doc.add_paragraph(f"Visit the official website at {row['College Website']}.")

                    # Save the document with the college name
                    doc_name = f"{college}_article.docx"
                    doc.save(doc_name)
                    
                    # Provide a download button for the Word document
                    with open(doc_name, "rb") as file:
                        st.download_button(label=f"Download {college} Article", data=file, file_name=doc_name)
                    
                    # Clean up by removing the generated file after download
                    os.remove(doc_name)
            else:
                st.warning("Please select at least one college.")
    else:
        st.error(f"Column '{college_name_column}' not found in the uploaded file.")

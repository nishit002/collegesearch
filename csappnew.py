import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt

# Function to create a table in Word document
def add_table_to_doc(doc, data, headers):
    table = doc.add_table(rows=1, cols=len(headers))
    hdr_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        hdr_cells[i].text = header

    for _, row in data.iterrows():
        row_cells = table.add_row().cells
        for i, header in enumerate(headers):
            row_cells[i].text = str(row[header])

st.title("College Information Portal")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Load the data from the Excel file
    excel_data = pd.ExcelFile(uploaded_file)

    # Load data into DataFrames
    college_df = pd.read_excel(excel_data, 'College')
    ranking_df = pd.read_excel(excel_data, 'Ranking')
    placement_df = pd.read_excel(excel_data, 'Placement')
    course_df = pd.read_excel(excel_data, 'Courses')
    faculty_df = pd.read_excel(excel_data, 'Faculty')
    recruiters_df = pd.read_excel(excel_data, 'Recruiters')
    awards_df = pd.read_excel(excel_data, 'Awards')
    admission_df = pd.read_excel(excel_data, 'Admission')
    contact_df = pd.read_excel(excel_data, 'Contact_Details')
    facilities_df = pd.read_excel(excel_data, 'Facilities')
    scholarships_df = pd.read_excel(excel_data, 'Scholarship')
    cutoff_df = pd.read_excel(excel_data, 'Cutoff')
    affiliation_df = pd.read_excel(excel_data, 'Affiliation')
    approval_df = pd.read_excel(excel_data, 'Approval')

    # Sidebar for College Selection
    st.sidebar.header('Select College')
    college_name = st.sidebar.selectbox('College Name', college_df['college_name'].unique())

    # Get selected college data
    college_info = college_df[college_df['college_name'] == college_name].iloc[0]

    # Generate the Word document and display content on the screen
    doc = create_word_document(college_info, ranking_df, placement_df, awards_df, 
                               faculty_df, recruiters_df, course_df, admission_df, 
                               contact_df, facilities_df, scholarships_df, cutoff_df,
                               affiliation_df, approval_df)

    # Prepare document for download
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.download_button(
        label="Download College Information in Word",
        data=buffer,
        file_name=f"{college_name}_info.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# Function to download Word file
def download_word_file(doc):
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# Streamlit app logic
st.title("College Information Portal")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Read the uploaded Excel file
    excel_data = pd.ExcelFile(uploaded_file)

    # Load data from the sheets into dataframes
    college_df = pd.read_excel(excel_data, 'College')
    ranking_df = pd.read_excel(excel_data, 'Ranking')
    placement_df = pd.read_excel(excel_data, 'Placement')
    course_df = pd.read_excel(excel_data, 'Courses')
    faculty_df = pd.read_excel(excel_data, 'Faculty')
    recruiters_df = pd.read_excel(excel_data, 'Recruiters')
    awards_df = pd.read_excel(excel_data, 'Awards')
    admission_df = pd.read_excel(excel_data, 'Admission')
    contact_df = pd.read_excel(excel_data, 'Contact_Details')
    facilities_df = pd.read_excel(excel_data, 'Facilities')
    scholarships_df = pd.read_excel(excel_data, 'Scholarship')
    cutoff_df = pd.read_excel(excel_data, 'Cutoff')
    affiliation_df = pd.read_excel(excel_data, 'Affiliation')
    approval_df = pd.read_excel(excel_data, 'Approval')

    # Sidebar for College Selection
    st.sidebar.header('Select College')
    college_name = st.sidebar.selectbox('College Name', college_df['college_name'].unique())

    # Display College Information
    st.header(f'{college_name} Information')

    college_info = college_df[college_df['college_name'] == college_name].iloc[0]
    st.write(f"**{college_name}** was established in {college_info['establishment_year']} and is located in {college_info['city']}, {college_info['state']}.")
    st.write(f"The college is known for its {college_info['usp']}. It is a {'Coed' if college_info['is_coed'] == 'Yes' else 'Non-Coed'} college.")
    st.write(f"The NIRF rank of the college is {college_info['nirf_rank']}.")

    # Generate and download Word document
    doc = create_word_document(college_info, ranking_df, placement_df, awards_df, 
                               faculty_df, recruiters_df, course_df, admission_df, 
                               contact_df, facilities_df, scholarships_df, cutoff_df,
                               affiliation_df, approval_df)
    buffer = download_word_file(doc)

    st.download_button(
        label="Download College Information in Word",
        data=buffer,
        file_name=f"{college_name}_info.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
else:
    st.write("Please upload the Excel file to get started.")


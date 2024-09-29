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

# Function to create Word document with the selected college data
def create_word_document(college_info, ranking_data, placement_data, awards_data, faculty_data, recruiters_data, course_data, admission_data, contact_data, facilities_data, scholarships_data, cutoff_data, affiliation_data, approval_data):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    # Title
    doc.add_heading(f"{college_info['college_name']} Information", level=1).style = style

    # General Info
    doc.add_paragraph(f"{college_info['college_name']} was established in the {college_info['establishment_year']} and is located in {college_info['city']}, {college_info['state']}.", style=style)
    doc.add_paragraph(f"The college is known for its {college_info['usp']}. It is a Coed College{'.' if college_info['is_coed'] == 'Yes' else ' and is not a Coed College.'}", style=style)
    nirf_rank = college_info['nirf_rank'] if pd.notna(college_info['nirf_rank']) else 'N/A'
    doc.add_paragraph(f"The NIRF rank of the college is {nirf_rank}.", style=style)

    # Approvals and Affiliations
    approved_by = ", ".join(approval_data['approval_body'].astype(str).tolist())
    affiliated_to = ", ".join(affiliation_data['affiliated_university'].astype(str).tolist())
    doc.add_paragraph(f"It has been approved by {approved_by}.", style=style)
    doc.add_paragraph(f"This college has a wide range of courses like {', '.join(course_data['course_name'].astype(str).tolist())}. It is affiliated with {affiliated_to}.", style=style)

    # NAAC Ranking
    doc.add_paragraph(f"NAAC has ranked the college at {college_info.get('naac_rank', 'N/A')}.", style=style)

    # Rankings Table
    doc.add_heading('Rankings', level=2).style = style
    if not ranking_data.empty:
        add_table_to_doc(doc, ranking_data, ['ranking_body', 'rank'])

    # Placements 
    doc.add_heading(f"{college_info['college_name']} Placements", level=2).style = style
    if not placement_data.empty:
        for index, row in placement_data.iterrows():
            doc.add_paragraph(f"The college placement records for {row['course_name']} are as follows, the Highest Package is INR {row['highest_package']} and the Average Package is INR {row['average_package']}.", style=style)

    # Top Recruiters
    doc.add_paragraph(f"The Top Recruiters of {college_info['college_name']} are {', '.join(recruiters_data['recruiter_name'].astype(str).tolist())}.", style=style)

    # Awards
    doc.add_heading(f"{college_info['college_name']} Awards", level=2).style = style
    if not awards_data.empty:
        for index, row in awards_data.iterrows():
            doc.add_paragraph(f"The college has been awarded with {row['award_name']} by {row['awarding_body']} in {row['year']}.", style=style)

    # Faculty Table
    doc.add_heading(f"{college_info['college_name']} Faculty", level=2).style = style
    if not faculty_data.empty:
        add_table_to_doc(doc, faculty_data, ['faculty_name', 'position', 'specialty', 'education'])

    # Contact Information
    doc.add_heading(f"{college_info['college_name']} Address", level=2).style = style
    contact_info = contact_data[contact_data['college_id'] == college_info['college_id']].iloc[0]
    doc.add_paragraph(f"Address: {contact_info['address']}", style=style)
    doc.add_paragraph(f"Phone: {contact_info['phone_number']}", style=style)
    doc.add_paragraph(f"Email: {contact_info['email']}", style=style)
    doc.add_paragraph(f"Website: {contact_info['website']}", style=style)

    # Courses and Fees
    doc.add_heading("Courses and Fees", level=2).style = style
    for index, row in course_data.iterrows():
        doc.add_paragraph(f"Course: {row['course_name']}, Duration: {row['duration']}, Fees: {row['fee']}.", style=style)

    return doc

# Function to download Word file
def download_word_file(doc):
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# Streamlit App
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

    # Display General Information about the selected college
    st.header(f'{college_name} Information')

    college_info = college_df[college_df['college_name'] == college_name].iloc[0]
    st.write(f"**{college_name}** was established in {college_info['establishment_year']} and is located in {college_info['city']}, {college_info['state']}.")
    st.write(f"The college is known for its {college_info['usp']}. It is a {'Coed' if college_info['is_coed'] == 'Yes' else 'Non-Coed'} college.")
    st.write(f"The NIRF rank of the college is {college_info['nirf_rank']}.")

    # Display rankings and other details
    st.subheader('Rankings')
    if not ranking_df.empty:
        st.dataframe(ranking_df)

    st.subheader('Placements')
    if not placement_df.empty:
        st.dataframe(placement_df)

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

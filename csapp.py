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
def create_word_document(college_info, ranking_data, placement_data, awards_data, faculty_data, recruiters_data, course_data):
    doc = Document()

    # Title
    doc.add_heading(f"{college_info['college_name']} Information", level=1)

    # General Info
    doc.add_paragraph(f"{college_info['college_name']} was established in {college_info['establishment_year']} and is located in {college_info['city']}, {college_info['state']}.")
    doc.add_paragraph(f"The college is known for its {college_info['usp']}. It is a {'Coed' if college_info['is_coed'] == 'Yes' else 'Non-Coed'} college.")
    doc.add_paragraph(f"The NIRF rank of the college is {college_info['nirf_rank']}.")

    # Rankings Table
    doc.add_heading('Rankings', level=2)
    if not ranking_data.empty:
        add_table_to_doc(doc, ranking_data, ['ranking_body', 'rank'])

    # Placements Table
    doc.add_heading(f"{college_info['college_name']} Placements", level=2)
    if not placement_data.empty:
        add_table_to_doc(doc, placement_data, ['course_name', 'highest_package', 'average_package'])

    # Awards
    doc.add_heading(f"{college_info['college_name']} Awards", level=2)
    for index, row in awards_data.iterrows():
        doc.add_paragraph(f"{row['award_name']} by {row['awarding_body']} in {row['year']}.")

    # Faculty Table
    doc.add_heading(f"{college_info['college_name']} Faculty", level=2)
    if not faculty_data.empty:
        add_table_to_doc(doc, faculty_data, ['faculty_name', 'position', 'specialty', 'education'])

    # Recruiters
    doc.add_heading(f"Top Recruiters at {college_info['college_name']}", level=2)
    doc.add_paragraph(", ".join(recruiters_data['recruiter_name'].tolist()))

    # Courses and Fees Table
    doc.add_heading(f"{college_info['college_name']} Courses and Fees", level=2)
    if not course_data.empty:
        add_table_to_doc(doc, course_data, ['course_name', 'level', 'duration', 'fee', 'seats'])

    return doc

# Function to download Word file
def download_word_file(doc):
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# Add a file uploader so the user can upload the Excel file
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

    # Sidebar for College Selection
    st.sidebar.header('Select College')
    college_name = st.sidebar.selectbox('College Name', college_df['college_name'].unique())

    # Display General Information about the selected college
    st.header(f'{college_name} Information')

    college_info = college_df[college_df['college_name'] == college_name].iloc[0]
    st.write(f"**{college_name}** was established in {college_info['establishment_year']} and is located in {college_info['city']}, {college_info['state']}.")
    st.write(f"The college is known for its {college_info['usp']}. It is a {'Coed' if college_info['is_coed'] == 'Yes' else 'Non-Coed'} college.")
    st.write(f"The NIRF rank of the college is {college_info['nirf_rank']}.")

    # Display Rankings Table
    st.subheader('Rankings')
    ranking_data = ranking_df[ranking_df['college_id'] == college_info['college_id']]
    if not ranking_data.empty:
        st.table(ranking_data[['ranking_body', 'rank']])
    else:
        st.write("No ranking data available.")

    # Display Placements Table
    st.subheader(f'{college_name} Placements')
    placement_data = placement_df[placement_df['college_id'] == college_info['college_id']]
    if not placement_data.empty:
        st.table(placement_data[['course_name', 'highest_package', 'average_package']])
    else:
        st.write("No placement data available.")

    # Display Awards
    st.subheader(f'{college_name} Awards')
    awards_data = awards_df[awards_df['college_id'] == college_info['college_id']]
    if not awards_data.empty:
        for index, row in awards_data.iterrows():
            st.write(f"{row['award_name']} by {row['awarding_body']} in {row['year']}.")
    else:
        st.write("No awards data available.")

    # Display Faculty Table
    st.subheader(f'{college_name} Faculty')
    faculty_data = faculty_df[faculty_df['college_id'] == college_info['college_id']]
    if not faculty_data.empty:
        st.table(faculty_data[['faculty_name', 'position', 'specialty', 'education']])
    else:
        st.write("No faculty data available.")

    # Display Recruiters
    st.subheader(f'Top Recruiters at {college_name}')
    recruiters_data = recruiters_df[recruiters_df['college_id'] == college_info['college_id']]
    st.write(", ".join(recruiters_data['recruiter_name'].tolist()) if not recruiters_data.empty else "No recruiters data available.")

    # Display Courses and Fees Table
    st.subheader(f'{college_name} Courses and Fees')
    course_data = course_df[course_df['college_id'] == college_info['college_id']]
    if not course_data.empty:
        st.table(course_data[['course_name', 'level', 'duration', 'fee', 'seats']])
    else:
        st.write("No course data available.")

    # Plot: Placements Graph
    st.subheader(f'Placements Overview for {college_name}')
    if not placement_data.empty:
        st.bar_chart(placement_data[['highest_package', 'average_package']])
    else:
        st.write("No placement data available for graph.")

    # Add Contact Information
    st.subheader(f'{college_name} Contact Information')
    st.write(f"Phone: {college_info.get('phone_number', 'Not Available')}")
    st.write(f"Email: {college_info.get('email', 'Not Available')}")
    st.write(f"Address: {college_info.get('address', 'Not Available')}")
    st.write(f"Website: [Visit Website]({college_info.get('website', '#')})")

    # Generate and download Word document
    doc = create_word_document(college_info, ranking_data, placement_data, awards_data, faculty_data, recruiters_data, course_data)
    buffer = download_word_file(doc)

    st.download_button(
        label="Download College Information in Word",
        data=buffer,
        file_name=f"{college_name}_info.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
else:
    st.write("Please upload the Excel file to get started.")

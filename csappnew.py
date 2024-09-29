import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt

# Helper function for adding sections
def add_section(doc, heading, content):
    doc.add_heading(heading, level=2).style = doc.styles['Normal']
    doc.add_paragraph(content, style=doc.styles['Normal'])

# Function to create Word document
def create_word_document(college_info, ranking_data, placement_data, awards_data, faculty_data, recruiters_data, course_data, admission_data, contact_data, facilities_data, scholarships_data, cutoff_data, affiliation_data, approval_data):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    # Title and General Info
    add_section(doc, f"{college_info['college_name']} Information",
                f"{college_info['college_name']} was established in {college_info['establishment_year']} and is located in {college_info['city']}, {college_info['state']}."
                f" The college is known for {college_info['usp']} and is {'Coed' if college_info['is_coed'] == 'Yes' else 'Non-Coed'}. NIRF rank: {college_info['nirf_rank'] or 'N/A'}")

    # Approvals and Affiliations
    doc.add_paragraph(f"Approved by: {', '.join(approval_data['approval_body'].astype(str).tolist())}", style=style)
    doc.add_paragraph(f"Affiliated with: {', '.join(affiliation_data['affiliated_university'].astype(str).tolist())}", style=style)

    # Add sections for Rankings, Placements, Recruiters, Awards, Scholarships, and Faculty
    if not ranking_data.empty:
        add_table_to_doc(doc, ranking_data, ['ranking_body', 'rank'])

    if not placement_data.empty:
        add_section(doc, f"{college_info['college_name']} Placements", 
                    "\n".join([f"{row['course_name']}: Highest INR {row['highest_package']}, Average INR {row['average_package']}" for _, row in placement_data.iterrows()]))

    doc.add_paragraph(f"Top Recruiters: {', '.join(recruiters_data['recruiter_name'].astype(str).tolist())}", style=style)

    if not awards_data.empty:
        add_section(doc, f"{college_info['college_name']} Awards", 
                    "\n".join([f"{row['award_name']} by {row['awarding_body']} ({row['year']})" for _, row in awards_data.iterrows()]))

    if not scholarships_data.empty:
        add_section(doc, "Scholarships", "\n".join([f"{row['scholarship_name']}: {row['description']}" for _, row in scholarships_data.iterrows()]))

    if not faculty_data.empty:
        add_table_to_doc(doc, faculty_data, ['faculty_name', 'position', 'specialty', 'education'])

    # Contact Info and Facilities
    contact_info = contact_data[contact_data['college_id'] == college_info['college_id']].iloc[0]
    doc.add_paragraph(f"Address: {contact_info['address']}. Contact: {contact_info['phone_number']}, {contact_info['email']}. Website: {contact_info['website']}", style=style)
    doc.add_paragraph(f"Facilities: {', '.join(facilities_data['facility_name'].astype(str).tolist())}", style=style)

    # Courses and Fees
    total_courses = len(course_data)
    doc.add_paragraph(f"{college_info['college_name']} offers {total_courses} courses across various levels.", style=style)

    for index, row in course_data.iterrows():
        doc.add_paragraph(f"Course: {row['course_name']}, Fee: {row['fee']}", style=style)
        admission_info = admission_data[admission_data['course_name'] == row['course_name']]
        if not admission_info.empty:
            doc.add_paragraph(f"Admissions: Start on {admission_info['start_date'].iloc[0].strftime('%Y-%m-%d')} and end on {admission_info['end_date'].iloc[0].strftime('%Y-%m-%d')}", style=style)

    # Admission Process
    add_section(doc, "Admission Process", "Follow these steps for admission: ...")
    
    # Cutoff Information
    if not cutoff_data.empty:
        add_table_to_doc(doc, cutoff_data, ['course_name', 'cutoff_score'])
    else:
        doc.add_paragraph("Cutoff information is not available.", style=style)

    return doc

# Streamlit functions to upload and process Excel file
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


import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt
import matplotlib.pyplot as plt

# Function to apply borders to a table (simplified to avoid errors)
def set_table_borders(table):
    for row in table.rows:
        for cell in row.cells:
            cell.paragraphs[0].runs[0].font.size = Pt(10)  # Adjust font size for clarity
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)

# Function to create a table in Word document with enhanced styling
def add_styled_table_to_doc(doc, data, headers):
    table = doc.add_table(rows=1, cols=len(headers))
    hdr_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        hdr_cells[i].text = header
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True  # Make header bold
        hdr_cells[i].paragraphs[0].alignment = 1  # Center alignment

    for _, row in data.iterrows():
        row_cells = table.add_row().cells
        for i, header in enumerate(headers):
            # Safely access columns and avoid KeyErrors
            row_cells[i].text = str(row.get(header.lower(), 'N/A'))  # Use lowercase column access
            row_cells[i].paragraphs[0].alignment = 1  # Center alignment
    set_table_borders(table)  # Apply borders to the table

# Function to create Word document with the selected college data
def create_word_document(college_info, ranking_data, placement_data, awards_data, faculty_data, recruiters_data, course_data, admission_data, contact_data, facilities_data, scholarships_data, cutoff_data, affiliation_data, approval_data):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    # Title
    doc.add_heading(f"{college_info['college_name']} Information", level=1)

    # About the College Section
    doc.add_heading('About the College', level=2)
    doc.add_paragraph(f"{college_info['college_name']} was established in the year {college_info['establishment_year']} and is located in {college_info['city']}, {college_info['state']}. It is a well-renowned institution known for its {college_info['usp']}.")
    doc.add_paragraph(f"The college is {'Coed' if college_info['is_coed'] == 'Yes' else 'Non-Coed'}. The NIRF rank of the college is {college_info['nirf_rank'] if pd.notna(college_info['nirf_rank']) else 'N/A'}.")

    # Rankings Table
    doc.add_heading('Rankings', level=2)

    # Ensure case insensitivity by converting all columns to lowercase
    ranking_data.columns = ranking_data.columns.str.lower()
    college_id = college_info['college_id']

    # Filter ranking data by college_id (with case-insensitive column handling)
    filtered_ranking_data = ranking_data[ranking_data['college_id'] == college_id]
    
    if not filtered_ranking_data.empty:
        doc.add_paragraph(f"{college_info['college_name']} has been ranked by various bodies for its academic and institutional performance. Below are the details of the rankings received:")

        # Check if the required columns exist and strip spaces if necessary
        if 'ranking_body' in filtered_ranking_data.columns and 'rank' in filtered_ranking_data.columns:
            add_styled_table_to_doc(doc, filtered_ranking_data[['ranking_body', 'rank']], ['Ranking Body', 'Rank'])
        else:
            doc.add_paragraph("Ranking data is not available due to missing columns.")
    else:
        doc.add_paragraph("Ranking data is not available.")

    # Placements Section
    doc.add_heading('Placements', level=2)
    placement_data.columns = placement_data.columns.str.lower()  # Ensure lowercase column names

    if not placement_data.empty and 'course_name' in placement_data.columns and 'highest_package' in placement_data.columns and 'average_package' in placement_data.columns:
        doc.add_paragraph(f"Placement records for various courses at {college_info['college_name']} are excellent. Below is a summary of the highest and average packages for each course:")
        add_styled_table_to_doc(doc, placement_data[['course_name', 'highest_package', 'average_package']], ['Course', 'Highest Package (INR)', 'Average Package (INR)'])
        doc.add_paragraph(f"The placement data indicates strong industry collaboration, with top recruiters participating in the placement process.")
    else:
        doc.add_paragraph("Placement data is not available.")

    # Recruiters Section
    recruiters_data.columns = recruiters_data.columns.str.lower()  # Ensure lowercase column names
    if not recruiters_data.empty:
        doc.add_paragraph(f"The Top Recruiters for {college_info['college_name']} include {', '.join(recruiters_data['recruiter_name'].astype(str).tolist())}. These recruiters represent leading companies in various sectors, ensuring bright career prospects for students.")
    
    # Facilities Section
    doc.add_heading('Facilities', level=2)
    facilities_data.columns = facilities_data.columns.str.lower()  # Ensure lowercase column names
    if not facilities_data.empty:
        doc.add_paragraph(f"{college_info['college_name']} offers a variety of modern facilities to its students, including:")
        doc.add_paragraph(f"{', '.join(facilities_data['facility_name'].astype(str).tolist())}")
        doc.add_paragraph(f"These facilities provide a conducive learning environment and support the holistic development of students.")

    # Courses Section
    doc.add_heading('Courses Offered', level=2)
    course_data.columns = course_data.columns.str.lower()  # Ensure lowercase column names
    if not course_data.empty:
        doc.add_paragraph(f"{college_info['college_name']} offers a wide range of undergraduate, postgraduate, and doctoral programs. Below are the details of some of the key courses offered:")
        add_styled_table_to_doc(doc, course_data[['course_name', 'duration', 'fee']], ['Course Name', 'Duration', 'Fee (INR)'])
        doc.add_paragraph(f"The courses offered cover a broad spectrum of specializations, ensuring that students can find programs suited to their career goals.")

    # Admission Section
    doc.add_heading('Admission Process', level=2)
    admission_data.columns = admission_data.columns.str.lower()  # Ensure lowercase column names
    if not admission_data.empty:
        doc.add_paragraph(f"Admissions at {college_info['college_name']} follow a structured process. Below are the important dates and criteria for the upcoming admissions:")
        add_styled_table_to_doc(doc, admission_data[['course_name', 'start_date', 'end_date']], ['Course Name', 'Admission Start Date', 'Admission End Date'])
        doc.add_paragraph(f"To apply, students must follow these steps:")
        doc.add_paragraph("1. Visit the official website and register.")
        doc.add_paragraph("2. Log in to the admission portal.")
        doc.add_paragraph("3. Submit the required documents.")
        doc.add_paragraph("4. Pay the application fee.")
        doc.add_paragraph("5. Keep a printout of the fee payment for future reference.")

    return doc

# Function to download Word file
def download_word_file(doc):
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# Function to generate a bar chart
def plot_placement_chart(placement_data):
    fig, ax = plt.subplots()
    ax.barh(placement_data['course_name'], placement_data['highest_package'], label='Highest Package', color='blue')
    ax.barh(placement_data['course_name'], placement_data['average_package'], label='Average Package', color='green')

    ax.set_xlabel('Package (INR)')
    ax.set_title('Placement Packages by Course')
    ax.legend()
    st.pyplot(fig)

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

    # Ensure all column names are case insensitive (lowercase)
    college_df.columns = college_df.columns.str.lower()
    ranking_df.columns = ranking_df.columns.str.lower()
    placement_df.columns = placement_df.columns.str.lower()
    course_df.columns = course_df.columns.str.lower()
    faculty_df.columns = faculty_df.columns.str.lower()
    recruiters_df.columns = recruiters_df.columns.str.lower()
    awards_df.columns = awards_df.columns.str.lower()
    admission_df.columns = admission_df.columns.str.lower()
    contact_df.columns = contact_df.columns.str.lower()
    facilities_df.columns = facilities_df.columns.str.lower()
    scholarships_df.columns = scholarships_df.columns.str.lower()
    cutoff_df.columns = cutoff_df.columns.str.lower()
    affiliation_df.columns = affiliation_df.columns.str.lower()
    approval_df.columns = approval_df.columns.str.lower()

    # Sidebar for College Selection
    st.sidebar.header('Select College')
    college_name = st.sidebar.selectbox('College Name', college_df['college_name'].unique())

    # Display General Information about the selected college
    st.header(f'{college_name} Information')

    college_info = college_df[college_df['college_name'] == college_name].iloc[0]
    st.write(f"**{college_name}** was established in {college_info['establishment_year']} and is located in {college_info['city']}, {college_info['state']}.")
    st.write(f"The college is known for its {college_info['usp']}. It is a {'Coed' if college_info['is_coed'] == 'Yes' else 'Non-Coed'} college.")
    st.write(f"The NIRF rank of the college is {college_info['nirf_rank']}.")

    # Display rankings
    st.subheader('Rankings')
    filtered_ranking_data = ranking_df[ranking_df['college_id'] == college_info['college_id']]
    if not filtered_ranking_data.empty:
        st.dataframe(filtered_ranking_data[['ranking_body', 'rank']])
    else:
        st.write("No ranking data available.")

    # Display placements
    st.subheader('Placements')
    if not placement_df.empty:
        st.dataframe(placement_df[['course_name', 'highest_package', 'average_package']])
        # Plot a bar chart for placement data
        plot_placement_chart(placement_df)

    # Display faculty
    st.subheader('Faculty')
    st.dataframe(faculty_df[['faculty_name', 'position', 'specialty', 'education']])

    # Display courses
    st.subheader('Courses')
    st.dataframe(course_df[['course_name', 'duration', 'fee']])

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

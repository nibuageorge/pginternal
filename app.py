import io
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.datavalidation import DataValidation


# ---------------- Grade Logic ----------------
GRADE_TO_POINT = {"A+": 5, "A": 4, "B": 3, "C": 2, "D": 1, "E": 0}
GRADES = ["A+", "A", "B", "C", "D", "E"]


def _grade_points_expr(cell_ref: str) -> str:
    return (
        f'IF({cell_ref}="A+",5,'
        f'IF({cell_ref}="A",4,'
        f'IF({cell_ref}="B",3,'
        f'IF({cell_ref}="C",2,'
        f'IF({cell_ref}="D",1,'
        f'IF({cell_ref}="E",0,""))))))'
    )


# ---------------- Excel Builder ----------------
THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
LEFT_NOWRAP = Alignment(horizontal="left", vertical="center", wrap_text=False)
CENTER_NOWRAP = Alignment(horizontal="center", vertical="center", wrap_text=False)
RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)


def _style_cell(cell, bold=False, align=CENTER, border=False, size=10):
    cell.font = Font(name="Calibri", size=size, bold=bold)
    cell.alignment = align
    if border:
        cell.border = BORDER


def build_excel(
    semester,
    programme,
    course_code,
    course_title,
    teacher_name,
    hod_name,
    total_students,
):
    wb = Workbook()
    ws = wb.active
    ws.title = "A1_Form"

    # ---------- Column widths (Base) ----------
    widths = {
        "A": 7, "B": 15, "C": 26, "D": 13,
        "E": 13, "F": 13, "G": 15, "H": 11,
        "I": 9, "J": 18,
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # ---------- Helpers ----------
    def write_merged(row, col_start, col_end, value, bold=False,
                     align=LEFT, size=10, border=False):
        ws.merge_cells(
            start_row=row, start_column=col_start,
            end_row=row, end_column=col_end,
        )
        c = ws.cell(row=row, column=col_start, value=value)
        _style_cell(c, bold=bold, align=align, size=size, border=border)
        for cc in range(col_start, col_end + 1):
            _style_cell(
                ws.cell(row=row, column=cc),
                bold=bold, align=align, size=size, border=border,
            )

    def write_cell(row, col, value, bold=False, align=CENTER,
                   size=10, border=False, number_format=None):
        c = ws.cell(row=row, column=col, value=value)
        _style_cell(c, bold=bold, align=align, size=size, border=border)
        if number_format:
            c.number_format = number_format

    # ---------- Header rows (NO borders) ----------
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    write_cell(1, 1, "FORM A : (To be kept in the college)", bold=True, align=LEFT)
    for cc in range(1, 6):
        _style_cell(ws.cell(row=1, column=cc), bold=True, align=LEFT)
        
    ws.merge_cells(start_row=1, start_column=6, end_row=1, end_column=10)
    c = ws.cell(row=1, column=6, value="Admission Batch:PG 2025-27")
    _style_cell(c, align=RIGHT)
    for cc in range(6, 11):
        _style_cell(ws.cell(row=1, column=cc), align=RIGHT)
        
    ws.row_dimensions[1].height = 20

    write_merged(
        2, 1, 10,
        "Mahatma Gandhi University, Kottayam",
        bold=True, align=CENTER, size=12,
    )
    ws.row_dimensions[2].height = 25

    write_merged(
        3, 1, 10,
        "INTERNAL ASSESSMENT (Theory) FOR PG PROGRAMME-CSS",
        bold=True, align=CENTER, size=11,
    )
    ws.row_dimensions[3].height = 22

    write_merged(4, 1, 3, "Center no.: E004", bold=True, align=LEFT)
    write_merged(4, 4, 6, "College: Baselius College", bold=True, align=LEFT)
    write_merged(4, 7, 10, f"Semester: {semester}", bold=True, align=LEFT)
    ws.row_dimensions[4].height = 22 

    write_merged(5, 1, 3, f"Programme: {programme}", bold=True, align=LEFT)
    write_merged(5, 4, 6, f"Course code: {course_code}", bold=True, align=LEFT)
    write_merged(5, 7, 10, f"Course title: {course_title}", bold=True, align=LEFT)
    ws.row_dimensions[5].height = 25 

    # ---------- Table header rows 6-7 (WITH borders) ----------
    ws.merge_cells(start_row=6, start_column=1, end_row=7, end_column=1)
    write_cell(6, 1, "Sl. No", bold=True, border=True)
    _style_cell(ws.cell(row=7, column=1), bold=True, border=True)

    ws.merge_cells(start_row=6, start_column=2, end_row=7, end_column=2)
    write_cell(6, 2, "Reg No.", bold=True, border=True)
    _style_cell(ws.cell(row=7, column=2), bold=True, border=True)

    ws.merge_cells(start_row=6, start_column=3, end_row=7, end_column=3)
    write_cell(6, 3, "Name of Candidate", bold=True, border=True)
    _style_cell(ws.cell(row=7, column=3), bold=True, border=True)

    ws.merge_cells(start_row=6, start_column=4, end_row=7, end_column=4)
    write_cell(6, 4, "Assignment\n(Wt 2)", bold=True, border=True)
    _style_cell(ws.cell(row=7, column=4), bold=True, border=True)

    ws.merge_cells(start_row=6, start_column=5, end_row=7, end_column=5)
    write_cell(6, 5, "Test Paper 1\n(Wt 2)", bold=True, border=True)
    _style_cell(ws.cell(row=7, column=5), bold=True, border=True)

    ws.merge_cells(start_row=6, start_column=6, end_row=7, end_column=6)
    write_cell(6, 6, "Test Paper 2\n(Wt 2)", bold=True, border=True)
    _style_cell(ws.cell(row=7, column=6), bold=True, border=True)

    ws.merge_cells(start_row=6, start_column=7, end_row=7, end_column=7)
    write_cell(6, 7, "Seminar\n(Wt 4)", bold=True, border=True)
    _style_cell(ws.cell(row=7, column=7), bold=True, border=True)

    ws.merge_cells(start_row=6, start_column=8, end_row=6, end_column=9)
    write_cell(6, 8, "Consolidated", bold=True, border=True)
    _style_cell(ws.cell(row=6, column=9), bold=True, border=True)
    write_cell(7, 8, "Grade Points", bold=True, border=True)
    write_cell(7, 9, "Grade", bold=True, border=True)

    ws.merge_cells(start_row=6, start_column=10, end_row=7, end_column=10)
    write_cell(6, 10, "Remarks", bold=True, border=True)
    _style_cell(ws.cell(row=7, column=10), bold=True, border=True)
    
    ws.row_dimensions[6].height = 20
    ws.row_dimensions[7].height = 20

    # ---------- Student rows (WITH borders) ----------
    data_start_row = 8
    n = total_students

    for i in range(n):
        r = data_start_row + i
        ws.row_dimensions[r].height = 20 

        write_cell(r, 1, i + 1, border=True)
        # Apply Number format "0" so it doesn't convert to scientific notation
        write_cell(r, 2, "", align=CENTER_NOWRAP, border=True, number_format="0")
        write_cell(r, 3, "", align=LEFT_NOWRAP, border=True)
        
        write_cell(r, 4, "", border=True)
        write_cell(r, 5, "", border=True)
        write_cell(r, 6, "", border=True)
        write_cell(r, 7, "", border=True)

        gp_d = _grade_points_expr(f"D{r}")
        gp_e = _grade_points_expr(f"E{r}")
        gp_f = _grade_points_expr(f"F{r}")
        gp_g = _grade_points_expr(f"G{r}")
        h_formula = (
            f'=IF(OR(D{r}="",E{r}="",F{r}="",G{r}=""),"",'
            f'ROUND(({gp_d}*2+{gp_e}*2+{gp_f}*2+{gp_g}*4)/10,2))'
        )
        write_cell(r, 8, h_formula, border=True)

        i_formula = (
            f'=IF(H{r}="","",'
            f'IF(H{r}>=4.5,"A+",'
            f'IF(H{r}>=4,"A",'
            f'IF(H{r}>=3,"B",'
            f'IF(H{r}>=2,"C",'
            f'IF(H{r}>=0.01,"D","E"))))))'
        )
        write_cell(r, 9, i_formula, border=True)

        write_cell(r, 10, "", border=True)

    # ---------- Data validation dropdowns ----------
    dv = DataValidation(
        type="list",
        formula1='"A+,A,B,C,D,E"',
        allow_blank=True,
        showErrorMessage=True,
        errorTitle="Invalid grade",
        error="Please select one of: A+, A, B, C, D, E",
    )
    ws.add_data_validation(dv)
    if n >= 1:
        dv.add(f"D{data_start_row}:G{data_start_row + n - 1}")

    # ---------- Footer (NO borders) ----------
    footer_row = data_start_row + n + 1

    ws.merge_cells(
        start_row=footer_row, start_column=1, end_row=footer_row, end_column=10
    )
    c = ws.cell(
        row=footer_row, column=1,
        value="Certified that  1: The entries are verified with the records and there is no error or omission.",
    )
    _style_cell(c, align=LEFT)
    for cc in range(1, 11):
        _style_cell(ws.cell(row=footer_row, column=cc), align=LEFT)

    ws.merge_cells(
        start_row=footer_row + 1, start_column=1,
        end_row=footer_row + 1, end_column=10,
    )
    c = ws.cell(
        row=footer_row + 1, column=1,
        value="2: These assessments were published and no grievances from the students are pending",
    )
    _style_cell(c, align=LEFT)
    for cc in range(1, 11):
        _style_cell(ws.cell(row=footer_row + 1, column=cc), align=LEFT)

    sig_row = footer_row + 3
    write_merged(sig_row, 1, 3, teacher_name or "xxxxxx", bold=True, align=CENTER)
    write_merged(sig_row, 4, 6, hod_name or "xxxxxxxx", bold=True, align=CENTER)
    ws.merge_cells(start_row=sig_row, start_column=7, end_row=sig_row, end_column=10)
    c = ws.cell(row=sig_row, column=7, value="Dr. Manoj Narayanan K.S.")
    _style_cell(c, bold=True, align=CENTER)
    for cc in range(7, 11):
        _style_cell(ws.cell(row=sig_row, column=cc), bold=True, align=CENTER)
    ws.row_dimensions[sig_row].height = 20

    desc_row = sig_row + 1
    write_merged(
        desc_row, 1, 3,
        "Name and Signature of \nthe Teacher in Charge",
        align=CENTER,
    )
    write_merged(
        desc_row, 4, 6,
        "Name and Signature of \nthe HOD",
        align=CENTER,
    )
    ws.merge_cells(
        start_row=desc_row, start_column=7, end_row=desc_row, end_column=10
    )
    c = ws.cell(
        row=desc_row, column=7,
        value="Name and Signature of \nthe Principal",
    )
    _style_cell(c, align=CENTER)
    for cc in range(7, 11):
        _style_cell(ws.cell(row=desc_row, column=cc), align=CENTER)
    ws.row_dimensions[desc_row].height = 35 

    seal_row = desc_row + 1
    ws.merge_cells(start_row=seal_row, start_column=7, end_row=seal_row, end_column=8)
    c = ws.cell(row=seal_row, column=7, value="College Seal")
    
    # Left-aligned seal
    _style_cell(c, align=LEFT) 
    for cc in range(7, 9):
        _style_cell(ws.cell(row=seal_row, column=cc), align=LEFT)
    ws.row_dimensions[seal_row].height = 30 

    # ---------- Page setup ----------
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins.left = 0.3
    ws.page_margins.right = 0.3
    ws.page_margins.top = 0.4
    ws.page_margins.bottom = 0.4
    ws.print_title_rows = "1:7"

    # ---------- Save ----------
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ---------------- Streamlit UI ----------------
st.set_page_config(page_title="Internal Mark Generator", layout="wide")
st.title("PG Internal Assessment – Form A Generator")
st.caption("Baselius COllege Kottayam · CSS PG Programme")

with st.sidebar:
    st.header("Course Details")
    semester = st.text_input("Semester", "xx")
    programme = st.text_input("Programme", "xxxxxxxx")
    course_code = st.text_input("Course Code", "xxxxxxxx")
    course_title = st.text_input("Course Title", "xxxxxxx")
    st.header("Signatures")
    teacher_name = st.text_input("Teacher in Charge (Name)", "xxxxxx")
    hod_name = st.text_input("HOD (Name)", "xxxxxxxx")
    st.header("Students")
    total_students = st.number_input(
        "Total number of students", min_value=1, max_value=200, value=10, step=1
    )

if st.button("📥 Generate Excel (Form A)", type="primary"):
    buf = build_excel(
        semester=semester,
        programme=programme,
        course_code=course_code,
        course_title=course_title,
        teacher_name=teacher_name,
        hod_name=hod_name,
        total_students=int(total_students),
    )

    file_name = f"FormA_{course_code or 'Course'}_Sem{semester or 'X'}.xlsx"
    st.success("Excel file generated successfully. Open it in Excel and fill the grades.")
    st.download_button(
        label="⬇️ Download Excel file",
        data=buf,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

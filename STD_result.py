import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import Workbook, load_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer
)
import os
import tempfile
import subprocess


# ==========================================================
# STUDENT RESULT MANAGEMENT SYSTEM
# Python Tkinter + Excel + PDF
# ==========================================================


FILE_NAME = "D:/STUDENT_RESULT.xlsx"


# ==========================================================
# COLOR COMBINATION
# ==========================================================

GHOST_WHITE = "#F8F8FF"
MIDNIGHT_BLUE = "#191970"
ROYAL_BLUE = "#4169E1"
WHITE = "#FFFFFF"


# ==========================================================
# ADMIN LOGIN DETAILS
# ONLY FOR SHOW ALL RESULTS
# ==========================================================

LOGIN_ID = "admin"
LOGIN_PASSWORD = "1234"


# ==========================================================
# EXCEL HEADERS
# ==========================================================

HEADERS = [
    "Name",
    "Roll No.",
    "Semester",
    "Section",
    "Probability",
    "OOP",
    "Data Structure",
    "CAO",
    "Communication",
    "Total Marks",
    "Percentage",
    "Result"
]


# ==========================================================
# CREATE EXCEL FILE
# ==========================================================

def create_excel_file():

    if not os.path.exists(FILE_NAME):

        workbook = Workbook()

        sheet = workbook.active

        sheet.title = "Student Results"

        sheet.append(HEADERS)

        workbook.save(FILE_NAME)

        workbook.close()


# ==========================================================
# CLEAR ADD STUDENT FIELDS
# ==========================================================

def clear_fields():

    name_entry.delete(0, tk.END)
    roll_entry.delete(0, tk.END)
    semester_entry.delete(0, tk.END)
    section_entry.delete(0, tk.END)

    for entry in subject_entries:
        entry.delete(0, tk.END)


# ==========================================================
# SAVE STUDENT
# ==========================================================

def save_student():

    name = name_entry.get().strip()
    roll_no = roll_entry.get().strip()
    semester = semester_entry.get().strip()
    section = section_entry.get().strip()

    # ------------------------------------------------------
    # CHECK BASIC FIELDS
    # ------------------------------------------------------

    if name == "" or roll_no == "" or semester == "" or section == "":

        messagebox.showerror(
            "Error",
            "Please enter Name, Roll No., Semester and Section."
        )

        return

    # ------------------------------------------------------
    # CHECK ROLL NO.
    # ------------------------------------------------------

    try:

        roll_no = int(roll_no)

    except ValueError:

        messagebox.showerror(
            "Error",
            "Roll No. must be a number."
        )

        return

    # ------------------------------------------------------
    # GET SUBJECT MARKS
    # ------------------------------------------------------

    marks = []

    for entry in subject_entries:

        value = entry.get().strip()

        if value == "":

            messagebox.showerror(
                "Error",
                "Please enter marks for all subjects."
            )

            return

        try:

            mark = int(value)

            if mark < 0 or mark > 100:

                messagebox.showerror(
                    "Error",
                    "Marks must be between 0 and 100."
                )

                return

            marks.append(mark)

        except ValueError:

            messagebox.showerror(
                "Error",
                "Marks must be numbers only."
            )

            return

    # ------------------------------------------------------
    # OPEN EXCEL
    # ------------------------------------------------------

    try:

        workbook = load_workbook(FILE_NAME)

        sheet = workbook.active

    except PermissionError:

        messagebox.showerror(
            "Excel Error",
            "Please close student_results.xlsx and try again."
        )

        return

    # ------------------------------------------------------
    # CHECK DUPLICATE ROLL NO.
    # ------------------------------------------------------

    for row in sheet.iter_rows(
        min_row=2,
        values_only=True
    ):

        if row[1] == roll_no:

            workbook.close()

            messagebox.showerror(
                "Duplicate Roll No.",
                "A student with this Roll No. already exists."
            )

            return

    # ------------------------------------------------------
    # CALCULATE TOTAL
    # ------------------------------------------------------

    total_marks = sum(marks)

    # ------------------------------------------------------
    # CALCULATE PERCENTAGE
    # ------------------------------------------------------

    percentage = total_marks / 5

    # ------------------------------------------------------
    # RESULT
    # Minimum 35 in every subject = PASS
    # ------------------------------------------------------

    if all(mark >= 35 for mark in marks):

        result = "Pass"

    else:

        result = "Fail"

    # ------------------------------------------------------
    # SAVE TO EXCEL
    # ------------------------------------------------------

    sheet.append([
        name,
        roll_no,
        semester,
        section,
        marks[0],
        marks[1],
        marks[2],
        marks[3],
        marks[4],
        total_marks,
        f"{percentage:.2f}%",
        result
    ])

    workbook.save(FILE_NAME)

    workbook.close()

    messagebox.showinfo(
        "Success",
        "Student record saved successfully!"
    )

    clear_fields()


# ==========================================================
# READ ALL STUDENTS FROM EXCEL
# ==========================================================

def read_all_students():

    workbook = load_workbook(FILE_NAME)

    sheet = workbook.active

    students = []

    for row in sheet.iter_rows(
        min_row=2,
        values_only=True
    ):

        if row[0] is not None:

            students.append(list(row))

    workbook.close()

    # ------------------------------------------------------
    # SORT BY ROLL NO. ASCENDING
    # ------------------------------------------------------

    def roll_sort(student):

        try:

            return int(student[1])

        except:

            return 999999999

    students.sort(key=roll_sort)

    return students


# ==========================================================
# FIND STUDENT BY ROLL NO.
# ==========================================================

def find_student(roll_no):

    students = read_all_students()

    for student in students:

        try:

            if int(student[1]) == int(roll_no):

                return student

        except:

            pass

    return None


# ==========================================================
# GET RESULT
# ==========================================================

def get_result():

    roll_no = get_roll_entry.get().strip()

    if roll_no == "":

        messagebox.showerror(
            "Error",
            "Please enter Roll No."
        )

        return

    try:

        roll_no = int(roll_no)

    except ValueError:

        messagebox.showerror(
            "Error",
            "Roll No. must be a number."
        )

        return

    student_data = find_student(roll_no)

    if student_data is None:

        messagebox.showerror(
            "Not Found",
            "Student record not found."
        )

        return

    open_student_result_window(student_data)


# ==========================================================
# CREATE PDF FOR ONE STUDENT
# ==========================================================

def create_student_pdf(student_data):

    file_path = os.path.join(
        tempfile.gettempdir(),
        "Student_Result_" + str(student_data[1]) + ".pdf"
    )

    document = SimpleDocTemplate(
        file_path,
        pagesize=A4,
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=40
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "TitleStyle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=20,
        textColor=colors.HexColor(MIDNIGHT_BLUE),
        alignment=TA_CENTER,
        spaceAfter=20
    )

    elements = []

    # ------------------------------------------------------
    # TITLE
    # ------------------------------------------------------

    elements.append(
        Paragraph(
            " STUDENT RESULT ",
            title_style
        )
    )

    # ------------------------------------------------------
    # STUDENT DATA
    # ------------------------------------------------------

    data = [
        ["Name", str(student_data[0])],
        ["Roll No.", str(student_data[1])],
        ["Semester", str(student_data[2])],
        ["Section", str(student_data[3])],
        ["Probability", str(student_data[4])],
        ["OOP", str(student_data[5])],
        ["Data Structure", str(student_data[6])],
        ["CAO", str(student_data[7])],
        ["Communication", str(student_data[8])],
        ["Total Marks", str(student_data[9])],
        ["Percentage", str(student_data[10])],
        ["Result", str(student_data[11])]
    ]

    table = Table(
        data,
        colWidths=[180, 250]
    )

    table.setStyle(
        TableStyle([
            (
                "BACKGROUND",
                (0, 0),
                (0, -1),
                colors.HexColor(GHOST_WHITE)
            ),
            (
                "TEXTCOLOR",
                (0, 0),
                (0, -1),
                colors.HexColor(MIDNIGHT_BLUE)
            ),
            (
                "FONTNAME",
                (0, 0),
                (0, -1),
                "Helvetica-Bold"
            ),
            (
                "FONTNAME",
                (1, 0),
                (1, -1),
                "Helvetica"
            ),
            (
                "GRID",
                (0, 0),
                (-1, -1),
                1,
                colors.HexColor(MIDNIGHT_BLUE)
            ),
            (
                "ALIGN",
                (0, 0),
                (-1, -1),
                "LEFT"
            ),
            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE"
            ),
            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                9
            ),
            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                9
            )
        ])
    )

    elements.append(table)

    document.build(elements)

    return file_path


# ==========================================================
# CREATE PDF FOR ALL STUDENTS
# ==========================================================

def create_all_students_pdf():

    students = read_all_students()

    file_path = os.path.join(
        tempfile.gettempdir(),
        "All_Student_Results.pdf"
    )

    document = SimpleDocTemplate(
        file_path,
        pagesize=landscape(A4),
        rightMargin=25,
        leftMargin=25,
        topMargin=30,
        bottomMargin=30
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "AllTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=20,
        textColor=colors.HexColor(MIDNIGHT_BLUE),
        alignment=TA_CENTER,
        spaceAfter=20
    )

    elements = []

    # ------------------------------------------------------
    # TITLE
    # ------------------------------------------------------

    elements.append(
        Paragraph(
            "STUDENTS RESULT",
            title_style
        )
    )

    # ------------------------------------------------------
    # TABLE HEADERS
    # ------------------------------------------------------

    data = [
        HEADERS
    ]

    # ------------------------------------------------------
    # ADD STUDENTS
    # Already sorted by Roll No.
    # ------------------------------------------------------

    for student in students:

        data.append([
            str(student[0]),
            str(student[1]),
            str(student[2]),
            str(student[3]),
            str(student[4]),
            str(student[5]),
            str(student[6]),
            str(student[7]),
            str(student[8]),
            str(student[9]),
            str(student[10]),
            str(student[11])
        ])

    # ------------------------------------------------------
    # TABLE
    # ------------------------------------------------------

    table = Table(
        data,
        repeatRows=1
    )

    table.setStyle(
        TableStyle([
            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.HexColor(ROYAL_BLUE)
            ),
            (
                "TEXTCOLOR",
                (0, 0),
                (-1, 0),
                colors.white
            ),
            (
                "FONTNAME",
                (0, 0),
                (-1, 0),
                "Helvetica-Bold"
            ),
            (
                "FONTSIZE",
                (0, 0),
                (-1, -1),
                7
            ),
            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.5,
                colors.HexColor(MIDNIGHT_BLUE)
            ),
            (
                "ALIGN",
                (0, 0),
                (-1, -1),
                "CENTER"
            ),
            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE"
            ),
            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                5
            ),
            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                5
            )
        ])
    )

    elements.append(table)

    document.build(elements)

    return file_path


# ==========================================================
# OPEN PDF
# ==========================================================

def open_pdf(file_path):

    try:

        if os.name == "nt":

            os.startfile(file_path)

        elif os.name == "posix":

            subprocess.Popen(
                ["xdg-open", file_path]
            )

    except Exception as e:

        messagebox.showerror(
            "PDF Error",
            str(e)
        )


# ==========================================================
# PRINT PDF
# ==========================================================

def print_pdf(file_path):

    try:

        if os.name == "nt":

            # Windows Print command
            os.startfile(
                file_path,
                "print"
            )

        else:

            messagebox.showinfo(
                "Print",
                "PDF has been created. Open it and select Print."
            )

            open_pdf(file_path)

    except Exception as e:

        messagebox.showerror(
            "Print Error",
            "Unable to print the PDF.\n\n" + str(e)
        )


# ==========================================================
# OPEN STUDENT RESULT WINDOW
# ==========================================================

def open_student_result_window(student_data):

    result_window = tk.Toplevel(get_window)

    result_window.title(
        "Student Result"
    )

    result_window.geometry(
        "700x780"
    )

    result_window.configure(
        bg=GHOST_WHITE
    )

    result_window.resizable(
        False,
        False
    )

    # ------------------------------------------------------
    # TOP BAR
    # ------------------------------------------------------

    top_frame = tk.Frame(
        result_window,
        bg=GHOST_WHITE
    )

    top_frame.pack(
        fill="x",
        padx=20,
        pady=15
    )

    # ------------------------------------------------------
    # TITLE
    # ------------------------------------------------------

    title = tk.Label(
        top_frame,
        text="STUDENT RESULT",
        font=("Arial", 22, "bold"),
        fg=MIDNIGHT_BLUE,
        bg=GHOST_WHITE
    )

    title.pack(
        side="left"
    )

    # ------------------------------------------------------
    # PDF BUTTON
    # ------------------------------------------------------

    pdf_button = tk.Button(
        top_frame,
        text="📄 PDF",
        font=("Arial", 11, "bold"),
        bg=ROYAL_BLUE,
        fg=WHITE,
        activebackground=ROYAL_BLUE,
        activeforeground=WHITE,
        width=10,
        command=lambda: open_pdf(
            create_student_pdf(student_data)
        )
    )

    pdf_button.pack(
        side="right",
        padx=5
    )

    # ------------------------------------------------------
    # PRINT BUTTON
    # ------------------------------------------------------

    print_button = tk.Button(
        top_frame,
        text="🖨 Print",
        font=("Arial", 11, "bold"),
        bg=ROYAL_BLUE,
        fg=WHITE,
        activebackground=ROYAL_BLUE,
        activeforeground=WHITE,
        width=10,
        command=lambda: print_pdf(
            create_student_pdf(student_data)
        )
    )

    print_button.pack(
        side="right",
        padx=5
    )

    # ------------------------------------------------------
    # RESULT FRAME
    # ------------------------------------------------------

    result_frame = tk.Frame(
        result_window,
        bg=GHOST_WHITE
    )

    result_frame.pack(
        padx=40,
        pady=10
    )

    # ------------------------------------------------------
    # ALL 12 FIELDS
    # ------------------------------------------------------

    fields = [
        ("Name", student_data[0]),
        ("Roll No.", student_data[1]),
        ("Semester", student_data[2]),
        ("Section", student_data[3]),
        ("Probability", student_data[4]),
        ("OOP", student_data[5]),
        ("Data Structure", student_data[6]),
        ("CAO", student_data[7]),
        ("Communication", student_data[8]),
        ("Total Marks", student_data[9]),
        ("Percentage", student_data[10]),
        ("Result", student_data[11])
    ]

    for i, (field, value) in enumerate(fields):

        field_label = tk.Label(
            result_frame,
            text=field,
            font=("Arial", 12, "bold"),
            fg=MIDNIGHT_BLUE,
            bg=GHOST_WHITE,
            width=20,
            anchor="w"
        )

        field_label.grid(
            row=i,
            column=0,
            padx=8,
            pady=5,
            sticky="w"
        )

        value_label = tk.Label(
            result_frame,
            text=str(value),
            font=("Arial", 12),
            bg=WHITE,
            fg="black",
            width=28,
            anchor="w",
            relief="solid",
            bd=1
        )

        value_label.grid(
            row=i,
            column=1,
            padx=8,
            pady=5,
            sticky="w"
        )

    # ------------------------------------------------------
    # EXIT
    # ------------------------------------------------------

    exit_button = tk.Button(
        result_window,
        text="Exit",
        font=("Arial", 12, "bold"),
        bg=ROYAL_BLUE,
        fg=WHITE,
        activebackground=ROYAL_BLUE,
        activeforeground=WHITE,
        width=15,
        command=result_window.destroy
    )

    exit_button.pack(
        pady=15
    )


# ==========================================================
# SHOW ALL RESULTS IN TREEVIEW
# ==========================================================

def show_all_results():

    # ------------------------------------------------------
    # CLEAR OLD DATA
    # ------------------------------------------------------

    for item in all_tree.get_children():

        all_tree.delete(item)

    # ------------------------------------------------------
    # READ SORTED STUDENTS
    # ------------------------------------------------------

    students = read_all_students()

    # ------------------------------------------------------
    # INSERT INTO TREEVIEW
    # ------------------------------------------------------

    for student in students:

        all_tree.insert(
            "",
            tk.END,
            values=(
                student[0],
                student[1],
                student[2],
                student[3],
                student[4],
                student[5],
                student[6],
                student[7],
                student[8],
                student[9],
                student[10],
                student[11]
            )
        )


# ==========================================================
# ADMIN LOGIN FOR SHOW ALL RESULTS
# ==========================================================

def open_all_results():

    login = tk.Toplevel(root)

    login.title(
        "Admin Login"
    )

    login.geometry(
        "500x350"
    )

    login.configure(
        bg=GHOST_WHITE
    )

    login.resizable(
        False,
        False
    )

    # ------------------------------------------------------
    # TITLE
    # ------------------------------------------------------

    title = tk.Label(
        login,
        text="ADMIN LOGIN",
        font=("Arial", 22, "bold"),
        fg=MIDNIGHT_BLUE,
        bg=GHOST_WHITE
    )

    title.pack(
        pady=30
    )

    # ------------------------------------------------------
    # FRAME
    # ------------------------------------------------------

    login_frame = tk.Frame(
        login,
        bg=GHOST_WHITE
    )

    login_frame.pack(
        pady=10
    )

    # ------------------------------------------------------
    # ID
    # ------------------------------------------------------

    tk.Label(
        login_frame,
        text="ID:",
        font=("Arial", 13, "bold"),
        fg=MIDNIGHT_BLUE,
        bg=GHOST_WHITE
    ).grid(
        row=0,
        column=0,
        padx=10,
        pady=15
    )

    admin_id_entry = tk.Entry(
        login_frame,
        font=("Arial", 13),
        width=25
    )

    admin_id_entry.grid(
        row=0,
        column=1,
        padx=10,
        pady=15
    )

    # ------------------------------------------------------
    # PASSWORD
    # ------------------------------------------------------

    tk.Label(
        login_frame,
        text="Password:",
        font=("Arial", 13, "bold"),
        fg=MIDNIGHT_BLUE,
        bg=GHOST_WHITE
    ).grid(
        row=1,
        column=0,
        padx=10,
        pady=15
    )

    admin_password_entry = tk.Entry(
        login_frame,
        font=("Arial", 13),
        width=25,
        show="*"
    )

    admin_password_entry.grid(
        row=1,
        column=1,
        padx=10,
        pady=15
    )

    # ------------------------------------------------------
    # CHECK LOGIN
    # ------------------------------------------------------

    def check_admin_login():

        entered_id = admin_id_entry.get().strip()

        entered_password = admin_password_entry.get().strip()

        if (
            entered_id == LOGIN_ID
            and entered_password == LOGIN_PASSWORD
        ):

            login.destroy()

            show_all_results_window()

        else:

            messagebox.showerror(
                "Access Denied",
                "Invalid ID or Password."
            )

    # ------------------------------------------------------
    # LOGIN BUTTON
    # ------------------------------------------------------

    login_button = tk.Button(
        login,
        text="🔐 Login",
        font=("Arial", 13, "bold"),
        bg=ROYAL_BLUE,
        fg=WHITE,
        activebackground=ROYAL_BLUE,
        activeforeground=WHITE,
        width=15,
        command=check_admin_login
    )

    login_button.pack(
        pady=15
    )

    # ------------------------------------------------------
    # EXIT
    # ------------------------------------------------------

    exit_button = tk.Button(
        login,
        text="Exit",
        font=("Arial", 12, "bold"),
        bg=ROYAL_BLUE,
        fg=WHITE,
        activebackground=ROYAL_BLUE,
        activeforeground=WHITE,
        width=15,
        command=login.destroy
    )

    exit_button.pack(
        pady=5
    )


# ==========================================================
# SHOW ALL RESULTS WINDOW
# ==========================================================

def show_all_results_window():

    global all_tree

    all_window = tk.Toplevel(root)

    all_window.title(
        "All Student Results"
    )

    all_window.geometry(
        "1450x700"
    )

    all_window.configure(
        bg=GHOST_WHITE
    )

    all_window.resizable(
        True,
        True
    )

    # ------------------------------------------------------
    # TOP BAR
    # ------------------------------------------------------

    top_frame = tk.Frame(
        all_window,
        bg=GHOST_WHITE
    )

    top_frame.pack(
        fill="x",
        padx=20,
        pady=15
    )

    # ------------------------------------------------------
    # TITLE
    # ------------------------------------------------------

    title = tk.Label(
        top_frame,
        text=" STUDENTS RESULT ",
        font=("Arial", 22, "bold"),
        fg=MIDNIGHT_BLUE,
        bg=GHOST_WHITE
    )

    title.pack(
        side="left"
    )

    # ------------------------------------------------------
    # PDF BUTTON
    # ------------------------------------------------------

    pdf_button = tk.Button(
        top_frame,
        text="📄 PDF",
        font=("Arial", 11, "bold"),
        bg=ROYAL_BLUE,
        fg=WHITE,
        activebackground=ROYAL_BLUE,
        activeforeground=WHITE,
        width=10,
        command=lambda: open_pdf(
            create_all_students_pdf()
        )
    )

    pdf_button.pack(
        side="right",
        padx=5
    )

    # ------------------------------------------------------
    # PRINT BUTTON
    # ------------------------------------------------------

    print_button = tk.Button(
        top_frame,
        text="🖨 Print",
        font=("Arial", 11, "bold"),
        bg=ROYAL_BLUE,
        fg=WHITE,
        activebackground=ROYAL_BLUE,
        activeforeground=WHITE,
        width=10,
        command=lambda: print_pdf(
            create_all_students_pdf()
        )
    )

    print_button.pack(
        side="right",
        padx=5
    )

    # ------------------------------------------------------
    # TREEVIEW FRAME
    # ------------------------------------------------------

    tree_frame = tk.Frame(
        all_window,
        bg=GHOST_WHITE
    )

    tree_frame.pack(
        fill="both",
        expand=True,
        padx=15,
        pady=10
    )

    # ------------------------------------------------------
    # 12 COLUMNS
    # ------------------------------------------------------

    columns = (
        "Name",
        "Roll No.",
        "Semester",
        "Section",
        "Probability",
        "OOP",
        "Data Structure",
        "CAO",
        "Communication",
        "Total Marks",
        "Percentage",
        "Result"
    )

    all_tree = ttk.Treeview(
        tree_frame,
        columns=columns,
        show="headings"
    )

    # ------------------------------------------------------
    # TREEVIEW HEADINGS
    # ------------------------------------------------------

    for column in columns:

        all_tree.heading(
            column,
            text=column
        )

        all_tree.column(
            column,
            width=120,
            minwidth=100,
            anchor="center"
        )

    # ------------------------------------------------------
    # VERTICAL SCROLLBAR
    # ------------------------------------------------------

    vertical_scrollbar = ttk.Scrollbar(
        tree_frame,
        orient="vertical",
        command=all_tree.yview
    )

    all_tree.configure(
        yscrollcommand=vertical_scrollbar.set
    )

    # ------------------------------------------------------
    # HORIZONTAL SCROLLBAR
    # ------------------------------------------------------

    horizontal_scrollbar = ttk.Scrollbar(
        tree_frame,
        orient="horizontal",
        command=all_tree.xview
    )

    all_tree.configure(
        xscrollcommand=horizontal_scrollbar.set
    )

    # ------------------------------------------------------
    # DISPLAY
    # ------------------------------------------------------

    all_tree.pack(
        side="top",
        fill="both",
        expand=True
    )

    vertical_scrollbar.pack(
        side="right",
        fill="y"
    )

    horizontal_scrollbar.pack(
        side="bottom",
        fill="x"
    )

    # ------------------------------------------------------
    # EXIT BUTTON
    # ------------------------------------------------------

    exit_button = tk.Button(
        all_window,
        text="Exit",
        font=("Arial", 12, "bold"),
        bg=ROYAL_BLUE,
        fg=WHITE,
        activebackground=ROYAL_BLUE,
        activeforeground=WHITE,
        width=15,
        command=all_window.destroy
    )

    exit_button.pack(
        pady=10
    )

    # ------------------------------------------------------
    # LOAD SORTED DATA
    # ------------------------------------------------------

    show_all_results()


# ==========================================================
# OPEN ADD STUDENT
# ==========================================================

def open_add_student():

    add_window = tk.Toplevel(root)

    add_window.title(
        "Add Student"
    )

    add_window.geometry(
        "600x760"
    )

    add_window.configure(
        bg=GHOST_WHITE
    )

    add_window.resizable(
        False,
        False
    )

    # ------------------------------------------------------
    # TITLE
    # ------------------------------------------------------

    title = tk.Label(
        add_window,
        text="ADD STUDENT",
        font=("Arial", 22, "bold"),
        fg=MIDNIGHT_BLUE,
        bg=GHOST_WHITE
    )

    title.pack(
        pady=20
    )

    # ------------------------------------------------------
    # FORM
    # ------------------------------------------------------

    form_frame = tk.Frame(
        add_window,
        bg=GHOST_WHITE
    )

    form_frame.pack(
        pady=5
    )

    # ------------------------------------------------------
    # NAME
    # ------------------------------------------------------

    tk.Label(
        form_frame,
        text="Name:",
        font=("Arial", 12, "bold"),
        fg=MIDNIGHT_BLUE,
        bg=GHOST_WHITE
    ).grid(
        row=0,
        column=0,
        padx=10,
        pady=7,
        sticky="w"
    )

    global name_entry

    name_entry = tk.Entry(
        form_frame,
        font=("Arial", 12),
        width=27
    )

    name_entry.grid(
        row=0,
        column=1,
        padx=10,
        pady=7
    )

    # ------------------------------------------------------
    # ROLL NO.
    # ------------------------------------------------------

    tk.Label(
        form_frame,
        text="Roll No.:",
        font=("Arial", 12, "bold"),
        fg=MIDNIGHT_BLUE,
        bg=GHOST_WHITE
    ).grid(
        row=1,
        column=0,
        padx=10,
        pady=7,
        sticky="w"
    )

    global roll_entry

    roll_entry = tk.Entry(
        form_frame,
        font=("Arial", 12),
        width=27
    )

    roll_entry.grid(
        row=1,
        column=1,
        padx=10,
        pady=7
    )

    # ------------------------------------------------------
    # SEMESTER
    # ------------------------------------------------------

    tk.Label(
        form_frame,
        text="Semester:",
        font=("Arial", 12, "bold"),
        fg=MIDNIGHT_BLUE,
        bg=GHOST_WHITE
    ).grid(
        row=2,
        column=0,
        padx=10,
        pady=7,
        sticky="w"
    )

    global semester_entry

    semester_entry = tk.Entry(
        form_frame,
        font=("Arial", 12),
        width=27
    )

    semester_entry.grid(
        row=2,
        column=1,
        padx=10,
        pady=7
    )

    # ------------------------------------------------------
    # SECTION
    # ------------------------------------------------------

    tk.Label(
        form_frame,
        text="Section:",
        font=("Arial", 12, "bold"),
        fg=MIDNIGHT_BLUE,
        bg=GHOST_WHITE
    ).grid(
        row=3,
        column=0,
        padx=10,
        pady=7,
        sticky="w"
    )

    global section_entry

    section_entry = tk.Entry(
        form_frame,
        font=("Arial", 12),
        width=27
    )

    section_entry.grid(
        row=3,
        column=1,
        padx=10,
        pady=7
    )

    # ------------------------------------------------------
    # SUBJECTS
    # ------------------------------------------------------

    global subject_entries

    subject_entries = []

    subjects = [
        "Probability",
        "OOP",
        "Data Structure",
        "CAO",
        "Communication"
    ]

    for i, subject in enumerate(subjects):

        tk.Label(
            form_frame,
            text=subject + " Marks:",
            font=("Arial", 12, "bold"),
            fg=MIDNIGHT_BLUE,
            bg=GHOST_WHITE
        ).grid(
            row=i + 4,
            column=0,
            padx=10,
            pady=7,
            sticky="w"
        )

        entry = tk.Entry(
            form_frame,
            font=("Arial", 12),
            width=27
        )

        entry.grid(
            row=i + 4,
            column=1,
            padx=10,
            pady=7
        )

        subject_entries.append(entry)

    # ------------------------------------------------------
    # SAVE BUTTON
    # ------------------------------------------------------

    save_button = tk.Button(
        add_window,
        text="💾 Save",
        font=("Arial", 12, "bold"),
        bg=ROYAL_BLUE,
        fg=WHITE,
        activebackground=ROYAL_BLUE,
        activeforeground=WHITE,
        width=15,
        command=save_student
    )

    save_button.pack(
        pady=8
    )

    # ------------------------------------------------------
    # CLEAR BUTTON
    # ------------------------------------------------------

    clear_button = tk.Button(
        add_window,
        text="Clear",
        font=("Arial", 12, "bold"),
        bg=ROYAL_BLUE,
        fg=WHITE,
        activebackground=ROYAL_BLUE,
        activeforeground=WHITE,
        width=15,
        command=clear_fields
    )

    clear_button.pack(
        pady=5
    )

    # ------------------------------------------------------
    # EXIT BUTTON
    # ------------------------------------------------------

    exit_button = tk.Button(
        add_window,
        text="Exit",
        font=("Arial", 12, "bold"),
        bg=ROYAL_BLUE,
        fg=WHITE,
        activebackground=ROYAL_BLUE,
        activeforeground=WHITE,
        width=15,
        command=add_window.destroy
    )

    exit_button.pack(
        pady=8
    )


# ==========================================================
# OPEN GET RESULT
# ==========================================================

def open_get_result():

    global get_window
    global get_roll_entry

    get_window = tk.Toplevel(root)

    get_window.title(
        "Get Result"
    )

    get_window.geometry(
        "700x350"
    )

    get_window.configure(
        bg=GHOST_WHITE
    )

    get_window.resizable(
        False,
        False
    )

    # ------------------------------------------------------
    # TITLE
    # ------------------------------------------------------

    title = tk.Label(
        get_window,
        text="GET STUDENT RESULT",
        font=("Arial", 22, "bold"),
        fg=MIDNIGHT_BLUE,
        bg=GHOST_WHITE
    )

    title.pack(
        pady=50
    )

    # ------------------------------------------------------
    # SEARCH FRAME
    # ------------------------------------------------------

    search_frame = tk.Frame(
        get_window,
        bg=GHOST_WHITE
    )

    search_frame.pack(
        pady=15
    )

    tk.Label(
        search_frame,
        text="Enter Roll No.:",
        font=("Arial", 13, "bold"),
        fg=MIDNIGHT_BLUE,
        bg=GHOST_WHITE
    ).grid(
        row=0,
        column=0,
        padx=10
    )

    get_roll_entry = tk.Entry(
        search_frame,
        font=("Arial", 13),
        width=20
    )

    get_roll_entry.grid(
        row=0,
        column=1,
        padx=10
    )

    # ------------------------------------------------------
    # GET RESULT BUTTON
    # ------------------------------------------------------

    search_button = tk.Button(
        search_frame,
        text="🔍 Get Result",
        font=("Arial", 11, "bold"),
        bg=ROYAL_BLUE,
        fg=WHITE,
        activebackground=ROYAL_BLUE,
        activeforeground=WHITE,
        command=get_result
    )

    search_button.grid(
        row=0,
        column=2,
        padx=10
    )

    # ------------------------------------------------------
    # EXIT
    # ------------------------------------------------------

    exit_button = tk.Button(
        get_window,
        text="Exit",
        font=("Arial", 12, "bold"),
        bg=ROYAL_BLUE,
        fg=WHITE,
        activebackground=ROYAL_BLUE,
        activeforeground=WHITE,
        width=15,
        command=get_window.destroy
    )

    exit_button.pack(
        pady=30
    )


# ==========================================================
# EXIT PROGRAM
# ==========================================================

def exit_program():

    answer = messagebox.askyesno(
        "Exit",
        "Are you sure you want to exit?"
    )

    if answer:

        root.destroy()


# ==========================================================
# MAIN DASHBOARD
# ==========================================================

def open_main_window():

    global root

    root = tk.Tk()

    root.title(
        "Student Result Management"
    )

    root.geometry(
        "800x600"
    )

    root.configure(
        bg=GHOST_WHITE
    )

    root.resizable(
        False,
        False
    )

    # ======================================================
    # DASHBOARD TITLE
    # ======================================================

    heading = tk.Label(
        root,
        text="STUDENT RESULT MANAGEMENT",
        font=("Courier", 28, "bold"),
        fg=MIDNIGHT_BLUE,
        bg=GHOST_WHITE
    )

    heading.pack(
        pady=55
    )

    # ======================================================
    # BUTTON FRAME
    # ======================================================

    button_frame = tk.Frame(
        root,
        bg=GHOST_WHITE
    )

    button_frame.pack(
        pady=10
    )

    # ------------------------------------------------------
    # COMMON BUTTON SETTINGS
    # ------------------------------------------------------

    button_width = 32
    button_height = 2

    # ------------------------------------------------------
    # 1. ADD STUDENT
    # ------------------------------------------------------

    add_button = tk.Button(
        button_frame,
        text="Add Student",
        font=("Times New Roman", 15, "bold"),
        bg=ROYAL_BLUE,
        fg=WHITE,
        activebackground=ROYAL_BLUE,
        activeforeground=WHITE,
        width=button_width,
        height=button_height,
        command=open_add_student
    )

    add_button.pack(
        pady=8
    )

    # ------------------------------------------------------
    # 2. GET RESULT
    # ------------------------------------------------------

    get_button = tk.Button(
        button_frame,
        text="Get Result",
        font=("Times New Roman", 15, "bold"),
        bg=ROYAL_BLUE,
        fg=WHITE,
        activebackground=ROYAL_BLUE,
        activeforeground=WHITE,
        width=button_width,
        height=button_height,
        
        command=open_get_result
    )

    get_button.pack(
        pady=8
    )

    # ------------------------------------------------------
    # 3. SHOW ALL RESULTS
    # ------------------------------------------------------

    all_button = tk.Button(
        button_frame,
        text="Show All Results",
        font=("Times New Roman", 15, "bold"),
        bg=ROYAL_BLUE,
        fg=WHITE,
        activebackground=ROYAL_BLUE,
        activeforeground=WHITE,
        width=button_width,
        height=button_height,
        
        command=open_all_results
    )

    all_button.pack(
        pady=8
    )

    # ------------------------------------------------------
    # 4. EXIT
    # ------------------------------------------------------

    exit_button = tk.Button(
        root,
        text="Exit",
        font=("Times New Roman", 13, "bold"),
        bg=ROYAL_BLUE,
        fg=WHITE,
        activebackground=ROYAL_BLUE,
        activeforeground=WHITE,
       width=button_width,
        height=button_height,
        
        command=exit_program
    )

    exit_button.pack(
        pady=8
    )

    root.mainloop()


# ==========================================================
# START PROGRAM
# ==========================================================

create_excel_file()

open_main_window()

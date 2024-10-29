import pandas as pd
import sqlite3
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import xlsxwriter
import os

# Connect to SQLite database (or create it if it doesn't exist)
conn = sqlite3.connect('attendance.db')
cursor = conn.cursor()

# Set up database tables
cursor.execute('''
    CREATE TABLE IF NOT EXISTS students (
        id INTEGER PRIMARY KEY,
        name TEXT,
        admission_number TEXT UNIQUE
    )
''')
cursor.execute('''
    CREATE TABLE IF NOT EXISTS attendance (
        id INTEGER PRIMARY KEY,
        student_id INTEGER,
        unit_name TEXT,
        date TEXT,
        is_present BOOLEAN,
        FOREIGN KEY(student_id) REFERENCES students(id)
    )
''')
conn.commit()

# 1. Function to upload students from Excel
def upload_students(file_path):
    try:
        # Read Excel file
        df = pd.read_excel(file_path)
        
        # Validate the presence of required columns
        if not set(['Name', 'Admission Number']).issubset(df.columns):
            print("Error: Excel file must contain 'Name' and 'Admission Number' columns.")
            return
        
        # Insert each student into the database
        for _, row in df.iterrows():
            try:
                cursor.execute("INSERT INTO students (name, admission_number) VALUES (?, ?)",
                               (row['Name'], row['Admission Number']))
            except sqlite3.IntegrityError:
                print(f"Skipping duplicate entry for admission number {row['Admission Number']}.")
        conn.commit()
        print("Students uploaded successfully.")
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")
    except Exception as e:
        print(f"Error processing file: {e}")

# 2. Function to take attendance for a specified unit
def take_attendance(unit_name):
    today = datetime.today().strftime('%Y-%m-%d')
    cursor.execute("SELECT * FROM students")
    students = cursor.fetchall()

    for student in students:
        while True:
            print(f"Mark attendance for {student[1]} (Admission Number: {student[2]})")
            status = input("Enter 'P' for present or 'A' for absent: ").strip().upper()
            if status in ('P', 'A'):
                is_present = status == 'P'
                cursor.execute("INSERT INTO attendance (student_id, unit_name, date, is_present) VALUES (?, ?, ?, ?)",
                               (student[0], unit_name, today, is_present))
                break
            else:
                print("Invalid input. Please enter 'P' for present or 'A' for absent.")
    conn.commit()
    print("Attendance recorded successfully.")

# 3. Function to generate attendance report with dynamic date headers
def generate_report(unit_name, output_format="excel"):
    cursor.execute("""
        SELECT s.name, s.admission_number, a.date, a.is_present
        FROM students s
        LEFT JOIN attendance a ON s.id = a.student_id
        WHERE a.unit_name = ?
        ORDER BY s.name, a.date
    """, (unit_name,))
    records = cursor.fetchall()

    # Extract unique dates for report columns
    dates = sorted(set(record[2] for record in records if record[2]))

    # Choose format for report generation
    if output_format == "excel":
        generate_excel_report(records, unit_name, dates)
    elif output_format == "pdf":
        generate_pdf_report(records, unit_name, dates)
    else:
        print("Unsupported format. Choose either 'excel' or 'pdf'.")

# 4a. Function to generate Excel report
def generate_excel_report(records, unit_name, dates):
    with xlsxwriter.Workbook(f"{unit_name}_Attendance_Report.xlsx") as workbook:
        worksheet = workbook.add_worksheet()

        # Write headers
        worksheet.write('A1', 'Name')
        worksheet.write('B1', 'Admission Number')
        for col_num, date in enumerate(dates, start=2):
            worksheet.write(0, col_num, date)

        # Populate data
        row = 1
        current_name = ""
        for record in records:
            if record[0] != current_name:
                # Start a new row for each student
                current_name = record[0]
                row += 1
                worksheet.write(row, 0, record[0])  # Name
                worksheet.write(row, 1, record[1])  # Admission Number

            # Write 'X' for present or 'O' for absent based on date
            col = dates.index(record[2]) + 2 if record[2] in dates else None
            if col is not None:
                worksheet.write(row, col, 'X' if record[3] else 'O')

    print(f"Excel report generated as {unit_name}_Attendance_Report.xlsx")

# 4b. Function to generate PDF report
def generate_pdf_report(records, unit_name, dates):
    file_name = f"{unit_name}_Attendance_Report.pdf"
    c = canvas.Canvas(file_name, pagesize=letter)
    width, height = letter

    # Set up headers
    c.drawString(30, height - 40, f"Attendance Report for {unit_name}")
    c.drawString(30, height - 60, "Name")
    c.drawString(150, height - 60, "Admission Number")

    # Draw date headers dynamically across the top
    x_pos = 300
    for date in dates:
        c.drawString(x_pos, height - 60, date)
        x_pos += 50  # Adjust spacing between date columns as needed

    # Populate data
    y = height - 80
    current_name = ""
    for record in records:
        if record[0] != current_name:
            # Start a new line for each student
            current_name = record[0]
            y -= 20
            c.drawString(30, y, record[0])  # Name
            c.drawString(150, y, record[1])  # Admission Number

        # Mark attendance for each date dynamically
        x_pos = 300 + dates.index(record[2]) * 50 if record[2] in dates else None
        if x_pos:
            c.drawString(x_pos, y, 'X' if record[3] else 'O')

        # Start a new page if needed
        if y < 50:
            c.showPage()
            y = height - 80

    c.save()
    print(f"PDF report generated as {file_name}")

# Main execution block for testing purposes
if __name__ == "__main__":
    # Example usage
    #upload_students(r'C:\Users\Pi\Desktop\knp\November 2024\DICTL\attend test.xlsx')
    #take_attendance('Unit 1')
    generate_report('Unit 1', output_format='excel')
    pass

import pdfplumber
import tabula
from openpyxl.utils import get_column_letter
from werkzeug.utils import secure_filename
import os
from flask import Flask, render_template, request, redirect, url_for, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
import fitz
import pandas as pd

app = Flask(__name__, template_folder='templates')

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_tables_from_pdf(pdf_path):
    """
    Extract tables from all pages of a PDF file using tabula-py.
    """
    all_tables = []
    with fitz.open(pdf_path) as doc:
        for page_num in range(len(doc)):
            tables = tabula.read_pdf(pdf_path, pages=page_num + 1, multiple_tables=True)
            if tables:
                all_tables.extend(tables)
    return all_tables
def extract_pdf_content(pdf_file):
    text_content = ""
    table_data = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            # Extract text content
            text_content += page.extract_text()

            # Extract table data
            tables = page.extract_tables()
            if tables:
                table_data.extend(tables)

    return text_content, table_data

def create_excel(text_content, table_data):
    wb = Workbook()
    ws = wb.active
    ws.title = "PDF Content"

    # Write text content to Excel
    ws['A1'] = ""
    text_lines = text_content.split('\n')
    current_row = 2
    max_row_position = 2
    table_data_used = set()
    for line in text_lines:
        line_in_table = any(line.strip() in " ".join([" ".join(map(str, row)) for row in table]) for table in table_data)
        if not line_in_table:
            ws.cell(row=current_row, column=1, value=line)
            current_row += 1

        for table_index, table in enumerate(table_data):
            table_row = " ".join([" ".join(map(str, row)) for row in table])
            if line.strip() in table_row and table_index not in table_data_used:
                table_cols = len(table[0])
                for row_index, row_data in enumerate(table, start=current_row):
                    for col_index, value in enumerate(row_data, start=2):  # Start from column B
                        ws.cell(row=row_index, column=col_index, value=value)
                # Update max_row_position after writing the table
                max_row_position = max(max_row_position, current_row + len(table) + 1)
                # Add the current table index to the table_data_used set
                table_data_used.add(table_index)
                # Increment the current row position after writing the table
                current_row += len(table) + 1  # Add some space between the table and next line of text

    # Update current_row to max_row_position for writing text after tables
    current_row = max_row_position

    # Apply font to maintain formatting
    for row in ws.iter_rows():
        for cell in row:
            cell.font = Font(name='Times New Roman', size=11)

    return wb
def write_tables_to_excel(tables, excel_path):
    """
    Write tables data to an Excel file.
    Each table is written to a separate sheet.
    """
    workbook = Workbook()

    for table_num, table in enumerate(tables, start=1):
        sheet = workbook.create_sheet(title=f'Table_{table_num}')

        # Set font styles for Excel
        title_font = Font(name='Times New Roma', size=11, bold=True)
        info_font = Font(name='Times New Roma', size=10)

        # Apply formatting to headers
        bold_font = Font(bold=True)
        alignment = Alignment(wrap_text=True, vertical='center')
        header_fill = PatternFill(start_color='ffffff', end_color='ffffff', fill_type='solid')

        # Convert table data to pandas DataFrame
        df = pd.DataFrame(table.values, columns=[col.title() if col else "" for col in table.columns])

        # Write headers to the sheet
        for col_num, column_header in enumerate(df.columns, start=1):
            if column_header:  # Check if column_header is not None
                cell = sheet.cell(row=1, column=col_num, value=column_header)
                cell.font = title_font
                cell.alignment = alignment
                cell.fill = header_fill
                sheet.column_dimensions[get_column_letter(col_num)].width = max(len(column_header) + 2, 10)

        # Write data rows
        for row_num, (_, row) in enumerate(df.iterrows(), start=2):  # Start from row 2 to skip headers
            for col_num, value in enumerate(row, start=1):
                cell = sheet.cell(row=row_num, column=col_num, value=value)
                cell.font = info_font
                cell.alignment = alignment
                sheet.column_dimensions[get_column_letter(col_num)].width = max(sheet.column_dimensions[get_column_letter(col_num)].width,
                                                                                len(str(value)) + 2)  # Adjust as needed

        # Adjust row height based on the maximum text length in each row
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
            max_text_length = max(len(str(cell.value)) for cell in row)
            sheet.row_dimensions[row[0].row].height = 35 + (max_text_length // 50) * 5  # Adjust as needed

    workbook.remove(workbook.active)
    workbook.save(excel_path)
def send_file_and_delete(file_path):
    try:
        response = send_file(file_path, as_attachment=True)
        return response
    finally:
        delete_files(file_path)

def delete_files(*file_paths):
    for path in file_paths:
        if os.path.exists(path):
            try:
                os.remove(path)
            except PermissionError:
                pass
@app.route('/')
def index():
    return render_template('pdftoexcel.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'pdfFile' not in request.files:
        return redirect(url_for('index'))

    file = request.files['pdfFile']

    if file.filename == '':
        return redirect(url_for('index'))

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        processing_option = request.form.get('processingOption')

        if processing_option == 'allText':
            text_content, table_data = extract_pdf_content(file)
            excel_file = create_excel(text_content, table_data)
            excel_file_path = 'output.xlsx'
            excel_file.save(excel_file_path)
            return send_file(excel_file_path, as_attachment=True)
        elif processing_option == 'tablesOnly':
            tables = extract_tables_from_pdf(filepath)
            if tables:
                excel_path = os.path.join(app.config['UPLOAD_FOLDER'], 'TablesOnly.xlsx')
                write_tables_to_excel(tables, excel_path)
                delete_files(filepath)
                return send_file_and_delete(excel_path)
            else:
                delete_files(filepath)
                return 'No tables found in the PDF.'
        else:
            delete_files(filepath)
            return 'Invalid processing option'

    return 'Invalid file. Please upload a PDF file.'


if __name__ == "__main__":
    app.run(debug=True)
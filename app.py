from flask import Flask, request, render_template, send_file
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import os
import shutil
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def compare_sheets(file1, file2, result_file):
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)
    df1['Row_Number_file1'] = df1.index + 1
    df2['Row_Number_file2'] = df2.index + 1
    merged_df = pd.merge(df1, df2, on='Product', how='outer', suffixes=('_file1', '_file2'))
    result_wb = openpyxl.Workbook()
    result_ws = result_wb.active
    result_ws.title = "Comparison Result"
    change_log_ws = result_wb.create_sheet(title="Change Log")
    diff_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    missing_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    result_ws.append(["Product", "Color", "Row_Number"])
    change_log_ws.append(["Product", "Old Value", "New Value"])

    for index, row in merged_df.iterrows():
        product = row['Product']
        color_file1 = row['Color_file1'] if pd.notna(row['Color_file1']) else ''
        color_file2 = row['Color_file2'] if pd.notna(row['Color_file2']) else ''
        row_num_file1 = int(row['Row_Number_file1']) if pd.notna(row['Row_Number_file1']) else 'MISSING'
        row_num_file2 = int(row['Row_Number_file2']) if pd.notna(row['Row_Number_file2']) else 'MISSING'

        if color_file1 != color_file2:
            color_result = f"{color_file1} | {color_file2}"
            change_log_ws.append([product, color_file1, color_file2])
        else:
            color_result = color_file1

        if row_num_file1 != row_num_file2:
            row_num_result = f"{row_num_file1} | {row_num_file2}"
            change_log_ws.append([product, row_num_file1, row_num_file2])
        else:
            row_num_result = row_num_file1

        result_ws.append([product, color_result, row_num_result])
        result_color_cell = result_ws.cell(row=index + 2, column=2)
        if color_file1 != color_file2:
            result_color_cell.fill = diff_fill

        result_row_num_cell = result_ws.cell(row=index + 2, column=3)
        if row_num_file1 == 'MISSING' or row_num_file2 == 'MISSING':
            result_row_num_cell.fill = missing_fill
        elif row_num_file1 != row_num_file2:
            result_row_num_cell.fill = diff_fill

    result_wb.save(result_file)
    print(f"Comparison complete. Results saved to {result_file}")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/compare', methods=['POST'])
def compare():
    if 'file1' not in request.files or 'file2' not in request.files:
        return "No file part"

    file1 = request.files['file1']
    file2 = request.files['file2']

    if file1.filename == '' or file2.filename == '':
        return "No selected file"

    if file1 and file2 and allowed_file(file1.filename) and allowed_file(file2.filename):
        filename1 = secure_filename(file1.filename)
        filename2 = secure_filename(file2.filename)
        file1_path = os.path.join(app.config['UPLOAD_FOLDER'], filename1)
        file2_path = os.path.join(app.config['UPLOAD_FOLDER'], filename2)
        result_file_path = os.path.join(app.config['UPLOAD_FOLDER'], "comparison_result.xlsx")
        
        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
        file1.save(file1_path)
        file2.save(file2_path)

        compare_sheets(file1_path, file2_path, result_file_path)

        response = send_file(result_file_path, as_attachment=True)

        # Cleanup uploaded files and result file
        try:
            os.remove(file1_path)
            os.remove(file2_path)
            os.remove(result_file_path)
        except OSError as e:
            print(f"Error: {e.strerror}")

        return response

    return "Invalid file type"

if __name__ == '__main__':
    app.run(debug=True)

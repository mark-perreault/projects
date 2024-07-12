from flask import Flask, request, render_template, send_file
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def compare_sheets(file1, file2, result_file, key_column):
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    # Trim column names to remove leading/trailing spaces
    df1.columns = df1.columns.str.strip()
    df2.columns = df2.columns.str.strip()

    # Ensure the key column exists
    if key_column not in df1.columns or key_column not in df2.columns:
        raise KeyError(f"Both files must contain the key column '{key_column}'.")

    # Check for missing columns
    missing_columns_file1 = set(df2.columns) - set(df1.columns)
    missing_columns_file2 = set(df1.columns) - set(df2.columns)
    if missing_columns_file1 or missing_columns_file2:
        missing_message = []
        if missing_columns_file1:
            missing_message.append(f"File 1 is missing columns: {', '.join(missing_columns_file1)}")
        if missing_columns_file2:
            missing_message.append(f"File 2 is missing columns: {', '.join(missing_columns_file2)}")
        raise KeyError(" | ".join(missing_message))

    # Add row numbers to each dataframe
    df1['Row_Number_file1'] = df1.index + 1
    df2['Row_Number_file2'] = df2.index + 1

    # Perform a full outer join on the key column
    merged_df = pd.merge(df1, df2, on=key_column, how='outer', suffixes=('_file1', '_file2'))

    # Create a new workbook for the result
    result_wb = openpyxl.Workbook()
    result_ws = result_wb.active
    result_ws.title = "Comparison Result"

    # Add a new sheet for the change log
    change_log_ws = result_wb.create_sheet(title="Change Log")

    # Define fill styles for differences and missing rows
    diff_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    missing_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    # Write the headers for the comparison result sheet
    result_ws.append([key_column, "Color", "Row_Number"])

    # Write the headers for the change log sheet
    change_log_ws.append([key_column, "Old Value", "New Value"])

    # Compare rows and handle missing rows
    for index, row in merged_df.iterrows():
        key_value = row[key_column]
        color_file1 = row['Color_file1'] if pd.notna(row['Color_file1']) else ''
        color_file2 = row['Color_file2'] if pd.notna(row['Color_file2']) else ''
        row_num_file1 = int(row['Row_Number_file1']) if pd.notna(row['Row_Number_file1']) else 'MISSING'
        row_num_file2 = int(row['Row_Number_file2']) if pd.notna(row['Row_Number_file2']) else 'MISSING'

        if color_file1 != color_file2:
            color_result = f"{color_file1} | {color_file2}"
            change_log_ws.append([key_value, color_file1, color_file2])
        else:
            color_result = color_file1

        if row_num_file1 != row_num_file2:
            row_num_result = f"{row_num_file1} | {row_num_file2}"
            change_log_ws.append([key_value, row_num_file1, row_num_file2])
        else:
            row_num_result = row_num_file1

        result_ws.append([key_value, color_result, row_num_result])
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
    key_column = request.form['key_column'].strip()

    if file1.filename == '' or file2.filename == '' or key_column == '':
        return "No selected file or key column"

    if file1 and file2 and allowed_file(file1.filename) and allowed_file(file2.filename):
        filename1 = secure_filename(file1.filename)
        filename2 = secure_filename(file2.filename)
        file1_path = os.path.join(app.config['UPLOAD_FOLDER'], filename1)
        file2_path = os.path.join(app.config['UPLOAD_FOLDER'], filename2)
        result_file_path = os.path.join(app.config['UPLOAD_FOLDER'], "comparison_result.xlsx")
        
        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
        file1.save(file1_path)
        file2.save(file2_path)

        try:
            compare_sheets(file1_path, file2_path, result_file_path, key_column)
        except KeyError as e:
            return str(e)

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

from flask import Flask, request, send_file, render_template
import pandas as pd
import io
import os
import openpyxl
import random

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    file1 = request.files.get('file1')
    file2 = request.files.get('file2')
    if file1 and file1.filename != '':
        return process_csv(file1, version=1)
    elif file2 and file2.filename != '':
        return process_csv(file2, version=2)
    else:
        return 'No selected file', 400

def process_csv(file, version):
    base_filename, file_ext = os.path.splitext(file.filename)
    output_filename = base_filename + " Processed" + file_ext

    if file_ext == ".csv":
        df = None
        if version == 1:
            df = process_dataframe(pd.read_csv(file))
        elif version == 2:
            df = process_dataframe_v2(pd.read_csv(file))
        
        output = io.StringIO()
        df.to_csv(output, index=False)
        output.seek(0)
        return send_file(output, mimetype='text/csv', as_attachment=True, attachment_filename=output_filename)
    
    elif file_ext == ".xlsx":
        if version == 1:
            process_xlsx(file, "temp.xlsx")
        elif version == 2:
            process_xlsx_v2(file, "temp_v2.xlsx")
        
        return send_file("temp" + ("_v2" if version == 2 else "") + ".xlsx", as_attachment=True, download_name=output_filename)
    
    else:
        return 'Unsupported file type', 400

def process_xlsx(file, output_path):
    file.save(output_path)
    wb = openpyxl.load_workbook(output_path)
    ws = wb.active
    for row in range(2, ws.max_row + 1):
        target_sum = ws[f'G{row}'].value
        if not isinstance(target_sum, int) or target_sum < 0 or target_sum > 5:
            continue
        combination = generate_combinations(target_sum)
        if combination:
            ws[f'C{row}'], ws[f'D{row}'], ws[f'E{row}'], ws[f'F{row}'] = combination
    wb.save(output_path)

def process_xlsx_v2(file, output_path):
    file.save(output_path)
    wb = openpyxl.load_workbook(output_path)
    ws = wb.active
    for row in range(2, ws.max_row + 1):
        target_sum = ws[f'I{row}'].value
        combination = generate_combinations_v2(target_sum)
        if combination:
            ws[f'C{row}'], ws[f'D{row}'], ws[f'E{row}'], ws[f'F{row}'], ws[f'G{row}'], ws[f'H{row}'] = combination
    wb.save(output_path)

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0')

from flask import Flask, request, send_file, render_template
import pandas as pd
import io
import os
import openpyxl
import random
import time  # For the delay

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')  # Make sure your HTML file is named 'index.html' and placed inside a folder named 'templates' in the same directory as your Flask app.

@app.route('/process', methods=['POST'])
def process():
    time.sleep(5)  # Simulate processing delay of 5 seconds
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

# Placeholder for processing dataframes
def process_dataframe(df):
    # Your processing logic for CSV version 1
    return df

def process_dataframe_v2(df):
    # Your processing logic for CSV version 2
    return df

def generate_combinations(target_sum):
    valid_combinations = []
    for c in [0, 1]:
        for d in [0, 1]:
            for e in [0, 1, 2]:
                for f in [0, 1]:
                    if c + d + e + f == target_sum:
                        valid_combinations.append((c, d, e, f))
    if valid_combinations:
        return random.choice(valid_combinations)
    return None

def generate_combinations_v2(target_sum):
    valid_combinations = []
    for c in [round(i*0.1, 2) for i in range(9)]:
        for d in [round(i*0.1, 2) for i in range(9)]:
            for e in [round(i*0.1, 2) for i in range(9)]:
                for f in [round(i*0.1, 2) for i in range(9)]:
                    for g in [round(i*0.1, 2) for i in range(33)]:
                        for h in [round(i*0.1, 2) for i in range(17)]:
                            if round(c + d + e + f + g + h, 2) == target_sum:
                                valid_combinations.append((c, d, e, f, g, h))
    if valid_combinations:
        return random.choice(valid_combinations)
    return None

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
    app.run(debug=True, host='0.0.0.0')

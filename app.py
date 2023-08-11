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
def process_csv():
    if 'file' not in request.files:
        return 'No file part', 400
    file = request.files['file']
    if file.filename == '':
        return 'No selected file', 400
    if file:
        file_ext = os.path.splitext(file.filename)[1]
        if file_ext == ".csv":
            df = pd.read_csv(file)
            df = process_dataframe(df)
            output = io.StringIO()
            df.to_csv(output, index=False)
            output.seek(0)
            return send_file(output, mimetype='text/csv', as_attachment=True, attachment_filename='output.csv')
        elif file_ext == ".xlsx":
            # Save the uploaded file temporarily
            file.save("temp.xlsx")
            process_xlsx("temp.xlsx")
            return send_file("temp.xlsx", as_attachment=True, download_name="processed_file.xlsx")
        else:
            return 'Unsupported file type', 400

def process_dataframe(df):
    for index, row in df.iterrows():
        target_sum = row['G']
        if not isinstance(target_sum, int) or target_sum < 0 or target_sum > 5:
            continue
        combination = generate_combinations(target_sum)
        if combination:
            df.at[index, 'C'], df.at[index, 'D'], df.at[index, 'E'], df.at[index, 'F'] = combination
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

def process_xlsx(filename):
    wb = openpyxl.load_workbook(filename)
    ws = wb.active
    for row in range(2, ws.max_row + 1):
        target_sum = ws[f'G{row}'].value
        if not isinstance(target_sum, int) or target_sum < 0 or target_sum > 5:
            continue
        combination = generate_combinations(target_sum)
        if combination:
            ws[f'C{row}'], ws[f'D{row}'], ws[f'E{row}'], ws[f'F{row}'] = combination
    wb.save(filename)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')

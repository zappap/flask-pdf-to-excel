from flask import Flask, request, render_template, send_file
import pdfplumber
import pandas as pd
import urllib.request
import xlrd
import xlwt
import os
import datetime

app = Flask(__name__)
OUTPUT_FOLDER = outputs
TEMPLATE_URL = httpwww.mavi.web.trygmsArac_Giris_Cikis_Aktarim_Sablon.xls

def process_pdf_to_xls(pdf_path, output_folder, yon, belge_no)
    template_path = os.path.join(output_folder, 'template.xls')
    urllib.request.urlretrieve(TEMPLATE_URL, template_path)

    pdf = pdfplumber.open(pdf_path)
    all_pages = []
    for page in pdf.pages
        table = page.extract_table()
        if table
            df = pd.DataFrame(table[1], columns=table[0])
            all_pages.append(df)

    all_data = pd.concat(all_pages)
    date_time = all_data.iloc[, -1].str.split(expand=True)
    date_time.columns = ['date', 'time']
    date_time['date'] = date_time['date'].str.replace('.', '', regex=True)

    df = pd.DataFrame(all_data['Araç Plaka'])
    df.insert(0, 'YÖN', yon)
    df.insert(1, 'BELGE_TÜRÜ', 3 if yon == 'Ç' else '')
    df.insert(2, 'BELGE_NO', belge_no)
    df.insert(4, 'DORSE1', '')
    df.insert(5, 'DORSE2', '')
    df.insert(6, 'date', date_time['date'])
    df.insert(7, 'time', date_time['time'])
    df = df.set_index('YÖN')

    output_path = os.path.join(output_folder, fSami-{datetime.datetime.today().strftime('%Y-%m-%d-%H-%M-%S')}.xls)
    workbook = xlrd.open_workbook(template_path, formatting_info=True)
    sheet = workbook.sheet_by_index(0)

    new_workbook = xlwt.Workbook()
    new_sheet = new_workbook.add_sheet(sheet.name)

    for row_idx in range(sheet.nrows)
        for col_idx in range(sheet.ncols)
            new_sheet.write(row_idx, col_idx, sheet.cell_value(row_idx, col_idx))

    start_row = 1
    for row_idx, row_data in enumerate(df.itertuples(index=False), start=start_row)
        for col_idx, value in enumerate(row_data)
            new_sheet.write(row_idx, col_idx, value)

    new_workbook.save(output_path)
    return output_path

@app.route(, methods=[GET, POST])
def index()
    if request.method == POST
        pdf_file = request.files[pdf]
        yon = request.form[yon]
        belge_no = request.form[belge_no]

        if pdf_file
            pdf_path = os.path.join(OUTPUT_FOLDER, input.pdf)
            pdf_file.save(pdf_path)

            output_path = process_pdf_to_xls(pdf_path, OUTPUT_FOLDER, yon, belge_no)
            return send_file(output_path, as_attachment=True)

    return render_template(index.html)

if __name__ == __main__
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    app.run(debug=True)

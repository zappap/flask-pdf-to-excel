from flask import Flask, render_template, request, redirect, url_for
import os
import pdfplumber
import pandas as pd
import xlwings as xw
import datetime
from PyPDF2 import PdfMerger
import urllib

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = "uploads"

if not os.path.exists("uploads"):
    os.makedirs("uploads")

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        files = request.files.getlist("pdfs")
        yon = request.form.get("yon")
        belge_no = request.form.get("beyanname")

        filenames = []
        for file in files:
            filename = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
            file.save(filename)
            filenames.append(filename)

        # Merge PDFs
        path = os.path.dirname(filenames[0])
        merger = PdfMerger()
        for pdf in filenames:
            merger.append(pdf)
        output_pdf = os.path.join(path, "combined_sami.pdf")
        merger.write(output_pdf)
        merger.close()

        # Extract data from PDF
        pdf = pdfplumber.open(output_pdf)
        all_pages = [pd.DataFrame(page.extract_table()[1:], columns=page.extract_table()[0]) for page in pdf.pages]
        all_data = pd.concat(all_pages)

        date_time = all_data.iloc[:, -1].str.split(expand=True)
        date_time.columns = ["date", "time"]
        date_time["date"] = date_time["date"].apply(lambda x: x if "/" in x else x.replace(".", "/"))

        df = pd.DataFrame(all_data["Araç Plaka"])
        df.insert(0, "YÖN", yon)
        df.insert(1, "BELGE_TÜRÜ", 3 if yon == "Cikis" else "")
        df.insert(2, "BELGE_NO", belge_no)
        df.insert(4, "DORSE1", "")
        df.insert(5, "DORSE2", "")
        df.insert(6, "date", date_time["date"])
        df.insert(7, "time", date_time["time"])
        df = df.set_index("YÖN")

        # Download template & save Excel file
        url = "http://www.mavi.web.tr/ygms/Arac_Giris_Cikis_Aktarim_Sablon.xls"
        template_path = os.path.join(path, "sablon.xls")
        urllib.request.urlretrieve(url, template_path)

        book = xw.Book(template_path)
        sheet = book.sheets[0]
        sheet["A2"].options(header=False).value = df
        output_xls = os.path.join(path, f"Sami-{datetime.datetime.today().strftime('%Y-%m-%d-%H-%M-%S')}.xls")
        book.save(output_xls)
        book.close()

        return f"Process completed! Download file: <a href='{output_xls}'>Here</a>"

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)

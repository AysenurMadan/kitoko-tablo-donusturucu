import os
import time
from flask import Flask, request, render_template, send_file, redirect, url_for
from werkzeug.utils import secure_filename
from docx import Document
import pandas as pd
from flask import send_file, url_for, redirect, request, render_template
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo

# ---------------------------------------------------------
# DOĞRUDAN ATAMA (ENV VAR GEREKMEZ)
endpoint = "https://kitoko.cognitiveservices.azure.com/"
key      = "9YY6WizjwBFCK36mGfr4f5LHUgkFXUcAtF2sMx2vfofPoCdztH5eJQQJ99BFACYeBjFXJ3w3AAALACOGWsaK"
# ---------------------------------------------------------

from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential

# Azure istemcisi
client = DocumentAnalysisClient(
    endpoint=endpoint,
    credential=AzureKeyCredential(key)
)

app = Flask(__name__)
app.config['UPLOAD_FOLDER']      = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        f = request.files.get('file')
        if not f:
            return redirect(request.url)
        fn = secure_filename(f.filename)
        save_path = os.path.join(app.config['UPLOAD_FOLDER'], fn)
        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
        f.save(save_path)
        return redirect(url_for('process_file', filename=fn))
    return render_template('index.html')

def process_with_form_recognizer(path):
    with open(path, "rb") as fd:
        poller = client.begin_analyze_document("prebuilt-layout", document=fd)
    result = poller.result()
    tbl = result.tables[0]
    df = pd.DataFrame(index=range(tbl.row_count), columns=range(tbl.column_count))
    for cell in tbl.cells:
        df.iat[cell.row_index, cell.column_index] = cell.content
    return df

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo

@app.route('/process/<filename>')
def process_file(filename):
    # 1) Görselden tabloyu alıyoruz
    path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    df   = process_with_form_recognizer(path)

    # ── Burada ilk satırı başlık olarak kullan ──
    # DataFrame’in ilk satırı gerçek kolon adları:
    header = df.iloc[0].tolist()
    # O satırı atıp index’i sıfırla
    df = df[1:].reset_index(drop=True)
    # Yeni kolon adlarını ata
    df.columns = header

    # 2) Excel dosyası ve XlsxWriter ayarları
    out_name = os.path.splitext(filename)[0] + ".xlsx"
    out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_name)

    with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:
        # Header=True ile df.columns’ı en üste basar
        df.to_excel(writer, index=False, sheet_name='Tablo')
        workbook  = writer.book
        worksheet = writer.sheets['Tablo']

        # 3) Excel Table objesi ekle (stil, filtre, zebra)
        (max_row, max_col) = df.shape
        tbl = {
            'columns': [{'header': col} for col in df.columns],
            'style': 'Table Style Medium 9',
            'autofilter': True,
            'show_row_stripes': True
        }
        worksheet.add_table(0, 0, max_row, max_col-1, tbl)

        # 4) Kolon genişliklerini içeriğe göre otomatik ayarla
        for i, col in enumerate(df.columns):
            max_len = max(
                df[col].astype(str).map(len).max(),
                len(str(col))
            ) + 2
            worksheet.set_column(i, i, max_len)

    # 5) İndir
    return send_file(out_path, as_attachment=True)

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    app.run(debug=True, port=5050)
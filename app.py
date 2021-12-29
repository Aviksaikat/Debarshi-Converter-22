import os
from flask import Flask, render_template, request, send_file
from io import BytesIO
from pdf2docx import parse
from docx2pdf import convert


app = Flask(__name__)


@app.route("/")
def hello():
    return render_template("Dashboard.html")


@app.route("/upload", methods=['POST'])
def upload():
    option1 = str(request.form.get('op1'))
    option2 = str(request.form.get('op2'))
    file = request.files['file']
    if (option1 == 'PDF') and (option2 == 'DOCX'):
        filename = file.filename + "-converted.docx"
        temp = 'test.pdf'
        doc_file = 'test.docx'
        with open(temp, 'wb') as f:
            f.write(BytesIO(file.read()).read())
            f.close()
        parse(temp, doc_file, start=0, end=None)
        os.remove(temp)
        return send_file(doc_file, download_name=filename, as_attachment=True)
    elif (option1 == 'DOCX') and (option2 == 'PDF'):
        filename = file.filename + "-converted.pdf"
        temp = 'test1.docx'
        pdf_file = 'test1.pdf'
        with open(temp, 'wb') as f:
            f.write(BytesIO(file.read()).read())
            f.close()
        convert('test1.docx')
        os.remove(temp)
        return send_file(pdf_file, download_name=filename, as_attachment=True)
    else:
        return "<h1>Please select different file formats</h1>"



if __name__ == "__main__":
    app.run(debug=True)
#import library yang dibutuhkan
import os
import processing as pr
from werkzeug.utils import secure_filename
from flask import Flask, render_template, flash, request, redirect, url_for, send_from_directory

#deklarasi lokasi upload file dan format file yang diperbolehkan
cwd = os.getcwd()
UPLOAD_FOLDER = f"{cwd}/upload"
ALLOWED_EXTENSIONS = {'xlsx', 'csv'}

#deklarasi aplikasi web menggunakan flask
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

#fungsi untuk memberikan aturan file yang diperbolehkan
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

#fungsi utama (routing) menggunakan flask
@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        #cek jika POST memiliki part-part file yang telah di upload
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        #jika user tidak memilih file, maka browser akan submit file dengan nama kosong
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            return redirect(url_for('download_file', name=filename, sheet_numb=1))
    return render_template('home.html')

#routing flask setelah proses upload untuk dilakukan proses natural language processing
@app.route('/process/<name>/<sheet_numb>')
def download_file(name, sheet_numb):
    filepath = "upload/" + name
    xlsx_list = pr.main(filepath)
    sheet_number, sheet_name, sheet_question = pr.sheet_num(int(sheet_numb))
    pr.stopword()
    pr.wordcloud()
    most_occur = pr.freq()
    avg_word, total_word = pr.avg()
    pr.ngrams(ngram=2)
    pr.network(top=3)
    return render_template("process.html",
                            most_occur=most_occur,
                            avg_word=avg_word,
                            total_word=total_word,
                            xlsxname=name,
                            sheet_number=sheet_number,
                            sheet_name=sheet_name,
                            sheet_question=sheet_question,
                            xlsx_list=xlsx_list)

if __name__ == '__main__':
    app.run(debug = True)
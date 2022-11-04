from pyrebase import pyrebase
import os
import html
import pandas as pd #panda.excelfiles.parse
import firebase_admin 
from firebase_admin import storage as admin_storage, credentials, firestore
from openpyxl import load_workbook, Workbook
from flask import Flask, render_template, request, redirect, send_file, send_from_directory, url_for

# Your credentials after create a app web project.
config = {
  "apiKey": "AIzaSyAYULJooDufFubZ-CKBdM03aOgHOZj7og",
  "authDomain": "data-brokers-f0705.firebaseapp.com",
  "databaseURL": "https://data-brokers-f0705-default-rtdb.firebaseio.com",
  "projectId": "data-brokers-f0705",
  "storageBucket": "data-brokers-f0705.appspot.com",
  "serviceAccount": "service_account.json",
  "messagingSenderId": "144918866815",
  "appId": "1:144918866815:web:7b03a6265eb122e284884a",
  "measurementId": "G-TQYCD14PX4"
}

cred = credentials.Certificate("service_account.json")
admin = firebase_admin.initialize_app(cred, {"storageBucket": "data-brokers-f0705.appspot.com"})
bucket = admin_storage.bucket()
firebase_storage = pyrebase.initialize_app(config)
storage = firebase_storage.storage()

app = Flask(__name__)
@app.route('/', methods=['GET', 'POST'])
def home():
    return render_template('home.html')

@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        request.form.get('upload')
        upload = request.files['upload']
        storage.child(upload.filename).put(upload)
        return redirect(url_for('home'))
    return render_template('upload.html')

@app.route('/modifyTest', methods=['GET', 'POST'])
def modifyTest():
    if request.method == 'POST':
        #try:
        path = "https://www.gs://data-brokers-f0705.appspot.com"
        excelTest = request.form.get('excelTest')
        url = storage.child(excelTest).get_url(None)
        file = storage.child(excelTest).download(path, filename=excelTest)
        print(excelTest)
        excel = pd.read_excel(url)

        sheet_name = request.form['sheet_name']
        col_name = request.form['col_name']
        row_name =  int(request.form['row_name'])
        new_val = request.form['new_value']


        excel.loc[row_name, col_name] = new_val


        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(excelTest, engine='xlsxwriter')

        # Convert the dataframe to an XlsxWriter Excel object.
        excel.to_excel(writer, sheet_name=sheet_name, index=False)

        # Close the Pandas Excel writer and output the Excel file.
        writer.close(path)

        #writer = pd.ExcelWriter(url, engine='xlsxwriter')
        #excel.to_excel(url, sheet_name='Sheet1')
        print(type(excel))
        print("this is before error")

        storage.child(excelTest).put(excelTest)
        return 'Success'
        #except:
            #return redirect(url_for('home'))
    return render_template("modifyTest.html")

@app.route('/modify_Excel', methods=['GET', 'POST'])
def modify():
    if request.method == 'POST':
        request.form.get('excel_file')
        excel = request.files['excel_file']
        sheet_name = request.form['sheet_name']
        col_name = request.form['col_name']
        row_name = request.form['row_name']
        new_val = request.form['new_value']

        wb = load_workbook(excel.filename)
        ws = wb[sheet_name]
        cell = col_name + row_name
        ws[cell] = new_val
        wb.save(excel.filename)
        storage.child(excel.filename).put(excel.filename)
        return redirect(url_for('home'))
    return render_template('modify_excel.html')

@app.route('/retrieve', methods=['GET', 'POST'])
def retrieve():
    if request.method == 'POST':
        try:
            retrieve = request.form['retrieve']
            url = storage.child(retrieve).get_url(None)
            words = retrieve.split(".")
            ext = words[1]
            
            #add additional conidtions for other extensions
            if ext == ('png') or ext == ('jpg'):
                return render_template('imageDown.html', url = url)

            elif ext == ('mp4'):
                return render_template('videoDown.html', url = url)

            elif ext == ('mp3'):
                return render_template('soundDown.html', url = url)
            
            else:
                return render_template("download.html", url = url)
        except:
            return render_template('fileError.html')
    return render_template('retrieve.html')

@app.route('/delete', methods=['GET', 'POST'])
def delete():
    if request.method == 'POST':
        try: 
            delete = request.form['delete']
            blob = bucket.blob(delete)
            blob.delete()
        except:
            return render_template('fileError.html')
    return render_template('delete.html')

if __name__ == '__main__':
    app.run(debug=True)
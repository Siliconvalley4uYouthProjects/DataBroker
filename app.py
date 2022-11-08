from pyrebase import pyrebase
from firebase_admin import storage as admin_storage, credentials
from openpyxl import load_workbook
from flask import Flask, render_template, request, redirect, url_for, session, flash
import requests.exceptions
import pandas as pd #panda.excelfiles.parse
import firebase_admin 
from firebase_admin import auth

# Your credentials after create a app web project.
config = {
  "apiKey": "AIzaSyAYULJoAoDufFubZ-CKBdM03aOgHOZj7og",
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
auth = firebase_storage.auth()
storage = firebase_storage.storage()


app = Flask(__name__)
app.secret_key = 'secret'

#Start route, login
@app.route('/', methods=['GET', 'POST'])
def login():
    #if('user' in session):
    #    return 'Hi, {}' .format(session['user'])
    if request.method == 'POST':
        email = request.form['email']
        password = request.form['password']
        try:
            user = auth.sign_in_with_email_and_password(email, password)
            session['user'] = email
            
            if('user' in session):
                return redirect(url_for('home'))
            else:
                return "Failed to create session"
        except:
            return render_template('login.html')

    return render_template('login.html')

#Logout route
@app.route('/logout', methods=['GET', 'POST'])
def logout():
    return render_template("logout.html")

#Account creation route
@app.route('/createAccount', methods=['GET', 'POST'])
def createAccount():
    if request.method == 'POST':
        email = request.form['email']
        password = request.form['password']
        try:
            auth.create_user_with_email_and_password(email, password)
            flash("User Created")
        except:
            return 'User already exists'
        return redirect(url_for('login'))
    return render_template('createAccount.html')

#home route
@app.route('/home', methods=['GET', 'POST'])
def home():
    if request.method == 'POST':
        return redirect(url_for('logout'))
    return render_template('home.html')

#upload route
@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        request.form.get('upload')
        upload = request.files['upload']
        storage.child(upload.filename).put(upload)
        return redirect(url_for('home'))
    return render_template('upload.html')

#Modify excel route pandas
@app.route('/modifyTest', methods=['GET', 'POST'])
def modifyTest():
    if request.method == 'POST':
        #try:
        path = "https://www.gs://data-brokers-f0705.appspot.com"
        excelTest = request.form.get('excelTest')
        url = storage.child(excelTest).get_url(None)
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
        writer.close()

        #writer = pd.ExcelWriter(url, engine='xlsxwriter')
        #excel.to_excel(url, sheet_name='Sheet1')

        storage.child(excelTest).put(excelTest)
        return render_template('modifyTest.html')
        #except:
            #return redirect(url_for('home'))
    return render_template("modifyTest.html")

# Modify Excel openpyxl route
@app.route('/modify_Excel', methods=['GET', 'POST'])
def modify():
    if request.method == 'POST':
        request.form.get('excel_file')
        excel = request.files['excel_file']
        sheet_name = request.form['sheet_name']
        col_name = request.form['col_name']
        row_name = request.form['row_name']
        new_val = request.form['new_value']

        wb = load_workbook(excel)
        ws = wb[sheet_name]
        cell = col_name + row_name
        ws[cell] = new_val
        wb.save(excel.filename)
        storage.child(excel.filename).put(excel)
        return redirect(url_for('home'))
    return render_template('modify_excel.html')

#Retreive route
@app.route('/retrieve', methods=['GET', 'POST'])
def retrieve():
    if request.method == 'POST':
        try:
            retrieve = request.form['retrieve']
            url = storage.child(retrieve).get_url(None)
            words = retrieve.split(".")
            ext = words[1]
            ext = ext.lower()
            
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

    all_files = storage.list_files()
    name = []

    for file in all_files:
        name.append(file.name)
    return render_template('retrieve.html', name=name)

#Delete route
@app.route('/delete', methods=['GET', 'POST'])
def delete():
    if request.method == 'POST':
        try: 
            delete = request.form['delete']
            blob = bucket.blob(delete)
            blob.delete()
        except:
            return render_template('fileError.html')
            
    all_files = storage.list_files()
    name = []

    for file in all_files:
        name.append(file.name)
    return render_template('delete.html', name=name)

if __name__ == '__main__':
    app.run(debug=True)
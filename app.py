from pyrebase import pyrebase
from firebase_admin import storage as admin_storage, credentials
from openpyxl import load_workbook
from flask import Flask, render_template, request, redirect, url_for, session, flash
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
        session.pop('user')
        return redirect(url_for('login'))
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

#Modify excel route (add drop downs of column names)
@app.route('/modify', methods=['GET', 'POST'])
def modify():
    if request.method == 'POST':
        try:
            excelItem = request.form.get('excel_file')
            url = storage.child(excelItem).get_url(None)
            excel = pd.read_excel(url)

            sheet_name = request.form['sheet_name']
            col_name = request.form['col_name']
            row_num = int(request.form['row_name'])
            new_val = request.form['new_value']

            excel.loc[excel.index[row_num], col_name] = new_val

            # Create a Pandas Excel writer using XlsxWriter as the engine.
            writer = pd.ExcelWriter(excelItem, engine='xlsxwriter')

            # Convert the dataframe to an XlsxWriter Excel object.
            excel.to_excel(writer, sheet_name=sheet_name, index=False)

            # Close the Pandas Excel writer and output the Excel file.
            writer.close()

            storage.child(excelItem).put(excelItem)
            return render_template('modifyExcel.html')
        except:
            return render_template('modifyExcel.html')
    return render_template("modifyExcel.html")

#Modify rows and columns route
@app.route('/modifyRowCol', methods=['GET', 'POST'])
def modifyRowCol():
    if request.method == 'POST':
        excelItem = request.form.get('excel_file')
        url = storage.child(excelItem).get_url(None)
        excel = pd.read_excel(url)

        sheet_name = request.form['sheet_name']     
        opt = request.form['modify_option']

        if opt == "add_column":
            col_name = request.form['option_name']
            excel.insert(0,col_name, " ")

        elif opt == "delete_column":
            col_name = request.form['option_name']
            excel = excel.drop(labels=col_name, axis=1)
        
        #(Add loop to add all rows if the index is farther down)
        elif opt == "add_row":
            row_num = request.form['option_name']
            shape = excel.shape
            count = int(shape[1])
            i = 0
            list = []

            while i < count:
                list.append(" ")
                i += 1
            excel.loc[len(excel.index)] = list
            print(excel)

        #Delete row since last option
        else:
            row_num = request.form['option_name']
            excel = excel.drop(labels=row_num, axis=0)

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(excelItem, engine='xlsxwriter')

        # Convert the dataframe to an XlsxWriter Excel object.
        excel.to_excel(writer, sheet_name=sheet_name, index=False)

        # Close the Pandas Excel writer and output the Excel file.
        writer.close()

        storage.child(excelItem).put(excelItem)

        return render_template("modifyRowCol.html")
    return render_template("modifyRowCol.html")

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
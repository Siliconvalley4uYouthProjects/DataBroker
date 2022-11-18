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

@app.route('/chooseExcel', methods=['GET', 'POST'])
def chooseExcel():
    all_files = storage.list_files()
    files = []

    for file in all_files:
        name = file.name
        words = name.split(".")
        ext = words[1]

        if ext == "xlsx" or ext == "csv":
            files.append(file.name)

    if request.method == 'POST':
        global excelName
        global sheets

        excelName = request.form.get('files')
        url = storage.child(excelName).get_url(None)

        xl = pd.ExcelFile(url)
        sheets = xl.sheet_names

        return redirect(url_for('chooseSheet'))
    return render_template("chooseExcel.html", files = files)

@app.route('/chooseSheet', methods=['GET', 'POST'])
def chooseSheet():
    url = storage.child(excelName).get_url(None)

    xl = pd.ExcelFile(url)
    sheets = xl.sheet_names
        
    if request.method == 'POST':
        global sheet_name
        global columns
        global rows
        sheet_name = request.form.get("sheet_name")

        url = storage.child(excelName).get_url(None)
        excel = pd.read_excel(url, sheet_name = sheet_name)

        xl = pd.ExcelFile(url)
        sheets = xl.sheet_names

        columns = excel.columns.tolist()
        shape = excel.shape
        row_count = int(shape[0])

        count = 0
        rows = []

        while count < row_count:
            rows.append(count)
            count += 1

        return redirect(url_for('modifyExcel'))
    return render_template('chooseSheet.html', excelName = excelName, sheets=sheets)
#Modify excel route (add drop downs of column names)
@app.route('/modifyExcel', methods=['GET', 'POST'])
def modifyExcel():
    url = storage.child(excelName).get_url(None)
    excel = pd.read_excel(url, sheet_name = sheet_name)

    if request.method == 'POST':
        col_name = request.form['col_name']
        row_num = int(request.form['row_name'])
        new_val = request.form['new_value']

        excel.loc[excel.index[row_num], col_name] = new_val
        print(excel)

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(excelName, engine='xlsxwriter')

        # Convert the dataframe to an XlsxWriter Excel object.
        excel.to_excel(writer, sheet_name=sheet_name, index=False)

        # Close the Pandas Excel writer and output the Excel file.
        writer.close()

        storage.child(excelName).put(excelName)
        return render_template('home.html')
    return render_template("modifyExcel.html", excelName = excelName, columns=columns, sheets=sheets, rows=rows)

@app.route('/addRow', methods=['GET', 'POST'])
def addRow():
    url = storage.child(excelName).get_url(None)
    excel = pd.read_excel(url)   
    print(excel)
    row_count = len(excel.index)
    print(row_count)
    if request.method == 'POST':


        def Insert_row(row_number, df, row_value):
            start_upper = 0
            end_upper = row_number
            start_lower = row_number
            end_lower = df.shape[0]
            upper_half = [*range(start_upper, end_upper, 1)]
            lower_half = [*range(start_lower, end_lower, 1)]
            lower_half = [x.__add__(1) for x in lower_half]
            index_ = upper_half + lower_half
            df.index = index_
            df.loc[row_number] = row_value
            df = df.sort_index()
            return df

        row_insert = int(request.form['row_num'])
        row_count = len(excel.index)
        print(excel)
        print(row_insert)
        print(row_count)

        if row_count == 0:
            shape = excel.shape
            count = int(shape[1])
            print(count)
            i = 0
            list = []
            #fill row with empty list
            while i < count:
                list.append(" ")
                i += 1
            excel.loc[len(excel.index)] = list
            print(excel)

        elif row_insert >= row_count: 
            while row_insert >= row_count:
                shape = excel.shape
                count = int(shape[1])
                i = 0
                list = []
                row_count += 1
                #fill row with empty list
                while i < count:
                    list.append(" ")
                    i += 1
                excel.loc[len(excel.index)] = list
        #inserts row between existing rows
        else:
            shape = excel.shape
            count = int(shape[1])
            i = 0
            list = []
            row_count += 1
            #fill row with empty list
            while i < count:
                list.append("Test")
                i += 1
            excel = Insert_row(row_insert, excel, list)
            print(excel)

                    # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(excelName, engine='xlsxwriter')

        # Convert the dataframe to an XlsxWriter Excel object.
        excel.to_excel(writer, sheet_name=sheet_name, index=False)

        # Close the Pandas Excel writer and output the Excel file.
        writer.close()

        storage.child(excelName).put(excelName)

        return redirect(url_for("home"))
    return render_template("addRow.html", excelName = excelName, sheets=sheets, rows=rows)

@app.route('/deleteRow', methods=['GET', 'POST'])
def deleteRow():
    if request.method == 'POST':
        url = storage.child(excelName).get_url(None)
        excel = pd.read_excel(url)    
        
        row_num = int(request.form['row_name'])
        excel = excel.drop(row_num)
        print(excel)

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(excelName, engine='xlsxwriter')

        # Convert the dataframe to an XlsxWriter Excel object.
        excel.to_excel(writer, sheet_name=sheet_name, index=False)

        # Close the Pandas Excel writer and output the Excel file.
        writer.close()

        storage.child(excelName).put(excelName)

        return redirect(url_for("home"))
    return render_template("deleteRow.html", excelName = excelName, sheets=sheets, rows=rows)

@app.route('/addColumn', methods=['GET', 'POST'])
def addColumn():
    if request.method == 'POST':
        url = storage.child(excelName).get_url(None)
        excel = pd.read_excel(url)

        col_name = request.form['col_name']
        excel.insert(0,col_name, " ")
        print(excel.keys())
        print(excel)
        
        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(excelName, engine='xlsxwriter')

        # Convert the dataframe to an XlsxWriter Excel object.
        excel.to_excel(writer, sheet_name=sheet_name, index=False)

        # Close the Pandas Excel writer and output the Excel file.
        writer.close()

        storage.child(excelName).put(excelName)
        return redirect('home') 
    return render_template("addColumn.html", excelName = excelName, sheets=sheets)

@app.route('/deleteColumn', methods=['GET', 'POST'])
def deleteColumn():
    if request.method == 'POST':
        url = storage.child(excelName).get_url(None)
        excel = pd.read_excel(url)

        col_name = request.form['column_name']
        excel = excel.drop(labels=col_name, axis=1)
        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(excelName, engine='xlsxwriter')

        # Convert the dataframe to an XlsxWriter Excel object.
        excel.to_excel(writer, sheet_name=sheet_name, index=False)

        # Close the Pandas Excel writer and output the Excel file.
        writer.close()

        storage.child(excelName).put(excelName)
        return redirect(url_for("home"))
    return render_template("deleteColumn.html", excelName=excelName, sheet_name=sheet_name, columns=columns)
    
@app.route('/addRowCol', methods=['GET', 'POST'])
def addRowCol():
    if request.method == 'POST':
        url = storage.child(excelName).get_url(None)
        excel = pd.read_excel(url)    

        opt = request.form['modify_option']

        if opt == "add_column":
            col_name = request.form['option_name']
            excel.insert(0,col_name, " ")
            print(excel.keys())
            print(excel)

        elif opt == "delete_column":
            col_name = request.form['option_name']
            excel = excel.drop(labels=col_name, axis=1)
            print(excel)
        
        elif opt == "add_row":
            # Function to insert row in the dataframe
            def Insert_row(row_number, df, row_value):
                start_upper = 0
                end_upper = row_number
                start_lower = row_number
                end_lower = df.shape[0]
                upper_half = [*range(start_upper, end_upper, 1)]
                lower_half = [*range(start_lower, end_lower, 1)]
                lower_half = [x.__add__(1) for x in lower_half]
                index_ = upper_half + lower_half
                df.index = index_
                df.loc[row_number] = row_value
                df = df.sort_index()
                return df

            row_insert = int(request.form['option_name'])
            row_count = len(excel.index)
            print(row_insert)
            print(row_count)
            #Checks if the row position is after the last row
            if row_insert >= row_count: 
                while row_insert >= row_count:
                    shape = excel.shape
                    count = int(shape[0])
                    i = 0
                    list = []
                    row_count += 1
                    #fill row with empty list
                    while i < count:
                        list.append(" ")
                        i += 1
                    test = excel.loc[len(excel.index)]
                    print(test)
                    excel.loc[len(excel.index)] = list
            #inserts row between existing rows
            else:
                shape = excel.shape
                count = int(shape[1])
                i = 0
                list = []
                row_count += 1
                #fill row with empty list
                while i < count:
                    list.append("Test")
                    i += 1
                excel = Insert_row(row_insert, excel, list)
                print(excel)
                

        #Delete row since last option
        else:
            row_num = int(request.form['option_name'])
            excel = excel.drop(row_num)
            print(excel)


        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(excelName, engine='xlsxwriter')

        # Convert the dataframe to an XlsxWriter Excel object.
        excel.to_excel(writer, sheet_name=sheet_name, index=False)

        # Close the Pandas Excel writer and output the Excel file.
        writer.close()

        storage.child(excelName).put(excelName)

        return render_template("addRowCol.html")
    return render_template("addRowCol.html")

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
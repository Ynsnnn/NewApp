import os
import pandas as pd
from flask import render_template, url_for, flash, redirect, request
from flaskblog import app, db, bcrypt
from flaskblog.forms import RegistrationForm, LoginForm
from flaskblog.models import User
from flask_login import login_user, current_user, logout_user
from PIL import ImageGrab
import xlwings as xw
from werkzeug.utils import secure_filename

def excel_catch_screen(shot_excel, shot_sheetname):
    app = xw.App(visible=True, add_book=False)  # Use xlwings Of app start-up
    wb = app.books.open(shot_excel)  # Open file
    sheet = wb.sheets(shot_sheetname)  # Selected sheet
    all = sheet.used_range  # Get content range
    print(all.value)
    all.api.CopyPicture()  # Copy picture area
    sheet.api.Paste()  # Paste
    img_name = 'data'
    pic = sheet.pictures[0]  # Current picture
    pic.api.Copy()  # Copy the picture
    img = ImageGrab.grabclipboard()  # Get the picture data of the clipboard
    img.save("C:\\Users\\Ynsnnn\\Desktop\\NewApp\\flaskblog\\static\\uploads\\" + img_name + ".png")  # Save the picture
    pic.delete()  # Delete sheet Pictures on
    wb.close()  # Do not save , Direct closure
    app.quit()

def processExcel():
    data = pd.read_excel('C:\\Users\\Ynsnnn\\Desktop\\NewApp\\flaskblog\\static\\uploads\\test.xlsx')
    print(data)
    data.fillna(method="ffill", inplace=True)
    print(data)

@app.route("/", methods=['GET', 'POST'])
@app.route("/login", methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('home'))
    form = LoginForm()
    if form.validate_on_submit():
        user = User.query.filter_by(email=form.email.data).first()
        if user and bcrypt.check_password_hash(user.password, form.password.data):
            login_user(user, remember=form.remember.data)
            next_page = request.args.get('next')
            return redirect(next_page) if next_page else redirect(url_for('home'))
        else:
            flash('Login Unsuccessful. Please check email and password', 'danger')
    return render_template('login.html', title='Login', form=form)

@app.route("/home")
def home():
    return render_template('home.html')

@app.route("/forms")
def forms():
    return render_template('forms.html')

@app.route("/modifydb")
def modifydb():
    return render_template('modifydb.html')

@app.route("/checkdb")
def checkdb():
    return render_template('checkdb.html')

@app.route("/logout")
def logout():
    logout_user()
    return redirect(url_for('/'))

@app.route('/uploader', methods=['GET', 'POST'])
def uploader():
    if request.method == 'POST':
        file = request.files['file']
        if file.filename == "":
            print("File must have a filename")
            return request.url

        filename = secure_filename(file.filename)
        file.save(os.path.join(app.config["UPLOAD_FOLDER"], filename))
        print("File saved")

        processExcel()
        print("File processed")

        excel_catch_screen("C:\\Users\\Ynsnnn\\Desktop\\NewApp\\flaskblog\\static\\uploads\\test.xlsx", "Sheet1")
        print("Capture saved")

        return redirect("http://localhost:5000/modifydb")

    return render_template("forms.html")

####################################################################################### Pastram asta ca sa ne facem noi conturi.
@app.route("/register", methods=['GET', 'POST'])
def register():
    if current_user.is_authenticated:
        return redirect(url_for('home'))
    form = RegistrationForm()
    if form.validate_on_submit():
        hashed_password = bcrypt.generate_password_hash(form.password.data).decode('utf-8')
        user = User(username=form.username.data, email=form.email.data, password=hashed_password)
        db.session.add(user)
        db.session.commit()
        flash('Your account has been created! You are now able to log in', 'success')
        return redirect(url_for('login'))
    return render_template('register.html', title='Register', form=form)

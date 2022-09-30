import io
import os
import string
from typing import Union, Any

from flask import Blueprint, render_template, request, flash, redirect, url_for, send_file, flash
from flask_login import login_user, login_required, logout_user, current_user
from werkzeug import Response
from werkzeug.security import check_password_hash
from werkzeug.utils import secure_filename
from flask import Flask, Request
from werkzeug.wsgi import FileWrapper

# from formatText import formatText
from .formatText import *
from .models import User
from .grammar import *


class dataFromHtml:
    bookTitle: string
    bookSubTitle: string
    author: string
    HalfTitlePage: string
    HalfTitleFont: string
    HalfTitleFontSize: string
    TitlePage: string
    TitlePageFont: string
    TitlePageFontSize: string
    SubtitleFont: string
    SubtitleFontSize: string
    AuthorFont: string
    AuthorFontSize: string
    CopyrightPage: string
    CopyrightPageType: string
    CopyrightFont: string
    CopyrightFontSize: string
    ISBN: string
    DedicationPage: string
    DedicationTitleFont: string
    DedicationTitleFontSize: string
    DedicationTextFont: string
    DedicationTextFontSize: string
    CapitalizeChapterNo: string
    ChapterTitleFont: string
    ChapterTitleFontSize: string
    ChapterDescription: string
    ChapterDescriptionFont: string
    ChapterDescriptionFontSize: string
    ParagraphFont: string
    ParagraphFontSize: string
    ParagraphDropCapital: string
    Interlining: string
    Justification: string
    PageSize: string
    InsideMargin: string
    OutsideMargin: string
    PublishOrEdit: string


auth = Blueprint('auth', __name__)



@auth.route('/login', methods=['GET', 'POST'])
def home():
    return render_template("home.html")


@auth.route('/contact', methods=['GET', 'POST'])
def contact():
    print('contact')
    dataFromHtml = loadDataFromHtml(request)
    print("#############################33")
    print(dataFromHtml.bookTitle)
    #file = request.files['file']
    #analyzeGrammar(file, dataFromHtml)

    return render_template("contact.html")


@auth.route('/login', methods=['GET', 'POST'])
def login():
    print('login')
    if request.method == 'POST':
        email = request.form.get('email')
        password = request.form.get('password')

        user = User.query.filter_by(email=email).first()
        if user:
            if check_password_hash(user.password, password):
                flash('Logged in successfully!', category='success')
                login_user(user, remember=True)
                return redirect(url_for('views.home'))
            else:
                flash('Incorrect password, try again.', category='error')
        else:
            flash('Email does not exist.', category='error')

    return render_template("login.html", user=current_user)


@auth.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('auth.login'))


@auth.route('/callisto', methods=['GET', 'POST'])
def file2():


    if request.method == 'POST' and request.files:
        dataFromHtml = loadDataFromHtml(request)


        file = request.files['file']
        if 'file' not in request.files:
            flash('No file uploaded', category='error')
            return redirect(request.url)

        if len(dataFromHtml.bookTitle) < 1:
            flash('Book title must be longer than 1 character', category='error')
            return redirect(request.url)
        elif len(dataFromHtml.bookTitle) > 256:
            flash('Book title must be shorter than 256 characters', category='error')
            return redirect(request.url)
        #elif 'bookSubTitle' in request.files:
         #   if len(dataFromHtml.bookSubTitle) > 256:
          #      flash('Book subtitle must be shorter than 256 characters', category='error')
           #     return redirect(request.url)
        elif len(dataFromHtml.author) < 1:
            flash('Author must be longer than 1 character', category='error')
            return redirect(request.url)
        elif file.filename == '':
            flash('No selected file.', category='error')
        elif len(dataFromHtml.ISBN) > 32:
                flash('ISBN must be shorter than 32 characters', category='error')
                return redirect(request.url)



        bookTitleSecure = secure_filename(dataFromHtml.bookTitle)
        dataFromHtml.bookTitle = bookTitleSecure

        bookSubtitleSecure = secure_filename(dataFromHtml.bookSubTitle)
        dataFromHtml.bookSubTitle = bookSubtitleSecure

        authorSecure = secure_filename(dataFromHtml.author)
        dataFromHtml.author = authorSecure

        ISBNSecure = secure_filename(dataFromHtml.ISBN)
        dataFromHtml.ISBN = ISBNSecure

        if file and allowed_file(file.filename):


            filename = secure_filename(file.filename)
            file.filename = filename

            #file.save(os.path.join("", filename))

            formatText(file, dataFromHtml)

            return (render_template("download.html"))

    return render_template("callisto.html")


@auth.route('/download', methods=['GET', 'POST'])
def index():
    return send_file('temporal/' + dataFromHtml.bookTitle + ' - Formatted by Callisto' + '.docx', as_attachment=True)


@auth.route('/download', methods=['GET', 'POST'])
def download_file():
    send_file('temporal/' + dataFromHtml.bookTitle + ' - Formatted by Callisto1' + '.docx', as_attachment=True)
    os.remove('temporal/' + dataFromHtml.bookTitle + ' - Formatted by Callisto' + '.docx')


################################################################################
@auth.route('/ganymede', methods=['GET', 'POST'])
def fileGrammar():

    dataFromHtml = loadDataFromHtml(request)



    if request.method == 'POST' and request.files:
        dataFromHtml = loadDataFromHtml(request)


        file = request.files['file']
        if 'file' not in request.files:
            flash('No file uploaded', category='error')
            return redirect(request.url)

        if len(dataFromHtml.bookTitle) < 1:
            flash('Book title must be longer than 1 character', category='error')
            return redirect(request.url)
        elif len(dataFromHtml.bookTitle) > 256:
            flash('Book title must be shorter than 256 characters', category='error')
            return redirect(request.url)
        elif file.filename == '':
            flash('No selected file.', category='error')

        bookTitleSecure = secure_filename(dataFromHtml.bookTitle)
        dataFromHtml.bookTitle = bookTitleSecure

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.filename = filename

            analyzeGrammar(file, dataFromHtml)

            return (render_template("downloadGanymede.html"))

    return (render_template("ganymede.html"))


@auth.route('/downloadGanymede', methods=['GET', 'POST'])
def download_file_ganymede():


    file_path = 'website/temporal/' + dataFromHtml.bookTitle + ' - Analyzed by Ganymede' + '.docx'

    return_data = io.BytesIO()

    with open(file_path, 'rb') as fo:
        return_data.write(fo.read())
# (after writing, cursor will be at last byte, so move it to start)
    return_data.seek(0)

    os.remove('website/temporal/' + dataFromHtml.bookTitle + ' - Analyzed by Ganymede' + '.docx')

    #return send_file(return_data, 'cristian.docx', mimetype='application/msword')
##################################
    file_wrapper = FileWrapper(return_data)
    headers = {
        'Content-Disposition': 'attachment; filename="{}"'.format(dataFromHtml.bookTitle + ' - Analyzed by Ganymede' + '.docx')
    }
    response = Response(file_wrapper,
                        mimetype='application//msword',
                        direct_passthrough=True,
                        headers=headers)
    return response

##################################


    #os.remove('website/temporal/' + dataFromHtml.bookTitle + ' - Analyzed by Ganymede' + '.docx')
    #return(send_file('temporal/' + dataFromHtml.bookTitle + ' - Analyzed by Ganymede' + '.docx', as_attachment=True))
    #return (render_template("deleteGanymede.html"))



#################################################################################
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def loadDataFromHtml(request):
    dataFromHtml.bookTitle = request.form.get('bookTitle')
    dataFromHtml.bookSubTitle = request.form.get('bookSubTitle')
    dataFromHtml.author = request.form.get('author')
    dataFromHtml.HalfTitlePage = request.form.get('Half Title Page')
    dataFromHtml.HalfTitleFont = request.form.get('Half Title Font')
    dataFromHtml.HalfTitleFontSize = request.form.get('Half Title Font Size')
    dataFromHtml.TitlePage = request.form.get('Title Page')
    dataFromHtml.TitlePageFont = request.form.get('Title Page Font')
    dataFromHtml.TitlePageFontSize = request.form.get('Title Page Font Size')
    dataFromHtml.SubtitleFont = request.form.get('Subtitle Font')
    dataFromHtml.SubtitleFontSize = request.form.get('Subtitle Font Size')
    dataFromHtml.AuthorFont = request.form.get('Author Font')
    dataFromHtml.AuthorFontSize = request.form.get('Author Font Size')
    dataFromHtml.CopyrightPage = request.form.get('Copyright Page')
    dataFromHtml.CopyrightPageType = request.form.get('Copyright Page Type')
    dataFromHtml.CopyrightFont = request.form.get('Copyright Font')
    dataFromHtml.CopyrightFontSize = request.form.get('Copyright Font Size')
    dataFromHtml.ISBN = request.form.get('ISBN')
    dataFromHtml.DedicationPage = request.form.get('Dedication Page')
    dataFromHtml.DedicationTitleFont = request.form.get('Dedication Title Font')
    dataFromHtml.DedicationTitleFontSize = request.form.get('Dedication Title Font Size')
    dataFromHtml.DedicationTextFont = request.form.get('Dedication Text Font')
    dataFromHtml.DedicationTextFontSize = request.form.get('Dedication Text Font Size')
    dataFromHtml.CapitalizeChapterNo = request.form.get('Capitalize Chapter No')
    dataFromHtml.ChapterTitleFont = request.form.get('Chapter Title Font')
    dataFromHtml.ChapterTitleFontSize = request.form.get('Chapter Title Font Size')
    dataFromHtml.ChapterDescription = request.form.get('Chapter Description')
    dataFromHtml.ChapterDescriptionFont = request.form.get('Chapter Description Font')
    dataFromHtml.ChapterDescriptionFontSize = request.form.get('Chapter Description Font Size')
    dataFromHtml.ParagraphFont = request.form.get('Paragraph Font')
    dataFromHtml.ParagraphFontSize = request.form.get('Paragraph Font Size')
    dataFromHtml.ParagraphDropCapital = request.form.get('Paragraph Drop Capital')
    dataFromHtml.Interlining = request.form.get('Interlining')
    dataFromHtml.Justification = request.form.get('Justification')
    dataFromHtml.PageSize = request.form.get('Page Size')
    dataFromHtml.InsideMargin = request.form.get('Inside')
    dataFromHtml.OutsideMargin = request.form.get('Outside')
    dataFromHtml.PublishOrEdit = request.form.get('drone')


############################################################################
    if dataFromHtml.PublishOrEdit == 'toEdit':
        dataFromHtml.HalfTitlePage = 'No'
        dataFromHtml.CopyrightPage = 'No'
        dataFromHtml.Interlining = '1.5'
        dataFromHtml.PageSize = 'Letter 8.5" x 11" (21,59 x 27,94 cm)'
        dataFromHtml.InsideMargin = '0.75 in (19.1 mm)'
        dataFromHtml.OutsideMargin = '0.75 in (19.1 mm)'
        dataFromHtml.ParagraphFontSize = '12'
        dataFromHtml.ChapterDescription = 'Yes'
        dataFromHtml.Justification ='JUSTIFY'
    ############################################################################


    return dataFromHtml


ALLOWED_EXTENSIONS = {'txt', 'doc', 'docx'}

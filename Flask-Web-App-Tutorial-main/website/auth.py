import io
import os
import string

from flask import Blueprint, render_template, request, flash, redirect, url_for, send_file, flash
from werkzeug import Response
from werkzeug.utils import secure_filename
from werkzeug.wsgi import FileWrapper
from .callisto import *
from .ganymede import *

class dataFromHtml:
    bookTitle: string
    bookSubTitle: string
    author: string
    HalfTitlePage: string
    TitlePage: string
    CopyrightPage: string
    ChapterDescriptionFont: string
    PageSize: string
    InsideMargin: string
    OutsideMargin: string
    PublishOrEdit: string
    FileName: string

auth = Blueprint('auth', __name__)


@auth.route('/', methods=['GET', 'POST'])
def home():
    return render_template("home.html")


@auth.route('/blog', methods=['GET', 'POST'])
def blog():
    return (render_template("blog.html"))


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
        elif len(dataFromHtml.author) < 1:
            flash('Author must be longer than 1 character', category='error')
            return redirect(request.url)
        elif file.filename == '':
            flash('No selected file.', category='error')

        if not allowed_file(file.filename):
            flash('Only .docx extension allowed', category='error')

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.filename = filename
            dataFromHtml.FileName = filename

            formatText(file, dataFromHtml)

            return render_template("downloadCallisto.html", book=dataFromHtml.bookTitle)

    return render_template("callisto.html")

@auth.route('/downloadCallisto/<book>', methods=['GET', 'POST'])
def download_file(book):

    file_path = PATH_TEMPORAL + book + FORMATTED_TEXT + '.docx'

    return_data = io.BytesIO()
    with open(file_path, 'rb') as fo:
        return_data.write(fo.read())
    # (after writing, cursor will be at last byte, so move it to start)
    return_data.seek(0)

    os.remove(file_path)

    file_wrapper = FileWrapper(return_data)
    headers = {'Content-Disposition': 'attachment; filename="{}"'.format(book + FORMATTED_TEXT + '.docx')}
    response = Response(file_wrapper, mimetype='application//msword', direct_passthrough=True, headers=headers)
    return response

@auth.route('/io', methods=['GET', 'POST'])
def download_character():
    return render_template("io.html")


@auth.route('/ioCharacter', methods=['GET', 'POST'])
def ioCharacter():
    return (send_file('doc templates/Character_Template.docx', 'Character_Template.docx', as_attachment=True))


@auth.route('/ioCauseEffect', methods=['GET', 'POST'])
def ioCauseEffect():
    return (send_file('doc templates/Ishikawa_Diagram.pptx', as_attachment=True))


@auth.route('/ganymede', methods=['GET', 'POST'])
def fileGrammar():

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

        if (not allowed_file(file.filename)):
            flash('Only .docx extension allowed', category='error')

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.filename = filename

            analyzeGrammar(file, dataFromHtml)

            return (render_template("downloadGanymede.html", book=dataFromHtml.bookTitle))

    return (render_template("ganymede.html"))


@auth.route('/downloadGanymede/<book>', methods=['GET', 'POST'])
def download_file_ganymede(book):
    file_path = PATH_TEMPORAL + book + ANALYZED_TEXT + '.docx'

    return_data = io.BytesIO()

    with open(file_path, 'rb') as fo:
        return_data.write(fo.read())
    # (after writing, cursor will be at last byte, so move it to start)
    return_data.seek(0)

    os.remove(file_path)

    file_wrapper = FileWrapper(return_data)
    headers = {'Content-Disposition': 'attachment; filename="{}"'.format(book + ANALYZED_TEXT + '.docx')}
    response = Response(file_wrapper, mimetype='application//msword', direct_passthrough=True, headers=headers)
    return response


def allowed_file(filename):
    if filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS:
        return True
    else:
        return False
def loadDataFromHtml(request):
    dataFromHtml.bookTitle = request.form.get('bookTitle')
    dataFromHtml.bookSubTitle = request.form.get('bookSubTitle')
    dataFromHtml.author = request.form.get('author')
    dataFromHtml.HalfTitlePage = request.form.get('Half Title Page')
    dataFromHtml.TitlePage = request.form.get('Title Page')
    dataFromHtml.CopyrightPage = request.form.get('Copyright Page')
    dataFromHtml.ChapterDescription = request.form.get('Chapter Description')
    dataFromHtml.PageSize = request.form.get('Page Size')
    dataFromHtml.PublishOrEdit = request.form.get('drone')
    #if not (hasattr(dataFromHtml, 'RefreshPage')):
     #   dataFromHtml.RefreshPage = False

 ############################################################################
    if dataFromHtml.PublishOrEdit == 'toEdit':
        dataFromHtml.HalfTitlePage = 'No'
        dataFromHtml.CopyrightPage = 'No'
        dataFromHtml.Interlining = '1.7'
        dataFromHtml.PageSize = 'Letter 8.5" x 11" (21,59 x 27,94 cm)'
        dataFromHtml.InsideMargin = '1 in (25.4 mm)'
        dataFromHtml.OutsideMargin = '1 in (25.4 mm)'
        dataFromHtml.ParagraphFontSize = '12'
        dataFromHtml.ChapterDescription = 'No'
        dataFromHtml.Justification = 'JUSTIFY'
    ############################################################################

    return dataFromHtml

ALLOWED_EXTENSIONS = {'docx'}

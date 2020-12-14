import os
import csv
import time
import shutil
import zipfile
from pyPreservica import cvs_to_xml, csv_to_search_xml, cvs_to_cmis_xslt, cvs_to_xsd
from flask import Flask, redirect, url_for, session, request, send_file,flash
from werkzeug.utils import secure_filename
from flask import render_template
from flask_bootstrap import Bootstrap
from flask_wtf import FlaskForm
from flask_wtf.file import FileField, FileAllowed, FileRequired
from wtforms import TextField, SelectField, HiddenField, ValidationError, RadioField,\
    BooleanField, SubmitField, IntegerField, FormField, validators
from wtforms.validators import Required, Length


NS_HELP = """
        This will become the namespace of the XML documents, e.g. http://www.openarchives.org/OAI/2.0/oai_dc/

        XML Namespaces are any valid URI or Uniform Resource Identifier.
        """
ROOT_HELP = 'This is Root Element of the XML documents, e.g. dc, Metadata, Root etc '

class DownloadForm(FlaskForm):
    xml_button = SubmitField('Download XML Documents as ZIP File')
    xsd_button = SubmitField('Download XSD Schema')
    search_button = SubmitField('Download Custom Search Index')
    cmis_button = SubmitField('Download Custom CMIS Transform for UA')

class CSVUploadForm(FlaskForm):
    namespace = TextField('XML Namespace', description=NS_HELP, validators=[Required()])
    root_element = TextField('XML Root Element', description=ROOT_HELP, validators=[Required(), Length(max=25)])
    cvs_file = FileField("CSV File", validators=[FileRequired(), FileAllowed(['csv'], 'CSV Files Only')], description='Make sure the Excel Spreadsheet has been saved as UTF-8 CSV')
    submit_button = SubmitField('Upload Spreadsheet')


class ColumnSelect(FlaskForm):
    column = SelectField('Unique CSV Column', description='Select a column which will be used to name the XML files. The value of this column should be different for each row of the spreadsheet.', coerce=str)
    options = list()
    options.append((".xml", "XML Convention (.xml)"))
    options.append((".metadata", "Preservica Convention (.metadata)"))
    xml_extenstion = RadioField(label="Select XML Naming Convention", description="Select XML Naming Convention", validators=[Required()], choices=options, default='.xml')
    submit_button = SubmitField('Generate XML', render_kw={"onclick": "showcursor()"})

app = Flask('app')
bootstrap = Bootstrap(app)


SECRET_KEY = os.urandom(32)
app.config['SECRET_KEY'] = SECRET_KEY
app.secret_key = SECRET_KEY
app.config['UPLOAD_FOLDER'] = '/home/pyPreservica/mysite/static/'

@app.route('/download', methods=('GET', 'POST'))
def download():
    form = DownloadForm()
    if form.validate_on_submit():
        if form.xml_button.data:
            return send_file(session['XML_ZIP'], mimetype="application/zip", as_attachment=True)
        if form.xsd_button.data:
            return send_file(session['XSD_File'], mimetype="application/xml", as_attachment=True)
        if form.search_button.data:
            return send_file(session['SEARCH_FILE'], mimetype="application/xml", as_attachment=True)
        if form.cmis_button.data:
            return send_file(session['CMIS_FILE'], mimetype="application/xml", as_attachment=True)

    return render_template('download.html', form=form)

@app.route('/select', methods=('GET', 'POST'))
def select():

    form = ColumnSelect()
    headers=session['HEADER']
    options = list()
    for h in headers:
        options.append((h,h))
        form.column.choices=options

    if form.validate_on_submit():
        column = form.column.data

        namespace = session['NS']
        element = session['ROOT']
        path = session['CSV']

        xml_extension = form.xml_extenstion.data


        client = session['client']
        folder = os.path.join(app.config['UPLOAD_FOLDER'], client)

        zipFile = os.path.join(folder, "xml.zip")
        try:
          os.remove(zipFile)
        except OSError:
            pass


        with zipfile.ZipFile(zipFile, 'w') as myzip:
            for xml in cvs_to_xml(csv_file=path, root_element=element, xml_namespace=namespace, file_name_column=column, export_folder=folder):
                file_name = os.path.basename(xml)
                if xml_extension == ".metadata":
                    head, _sep, tail = file_name.rpartition(".")
                    file_name = head + xml_extension
                myzip.write(xml, arcname=file_name)
                try:
                  os.remove(xml)
                except OSError:
                  pass

        search_xml = csv_to_search_xml(csv_file=path,  root_element=element, xml_namespace=namespace, title="Metadata Title", export_folder=folder)
        xsd = cvs_to_xsd(csv_file=path,  root_element=element, xml_namespace=namespace, export_folder=folder)
        cmis = cvs_to_cmis_xslt(csv_file=path,  root_element=element, xml_namespace=namespace, title="Metadata Title", export_folder=folder)

        session['XML_ZIP'] = zipFile
        session['XSD_File'] = xsd
        session['CMIS_FILE'] = cmis
        session['SEARCH_FILE'] = search_xml


        try:
          os.remove(path)
        except OSError:
          pass

        return redirect(url_for('download'))


    return render_template('choice.html', form=form)

@app.route('/restart')
def restart():
    try:
        client = session['csrf_token']
        if client is not None:
            folder = os.path.join(app.config['UPLOAD_FOLDER'], client)
            shutil.rmtree(os.path.join(folder))
            return redirect(url_for('start'))
    except KeyError:
        return redirect(url_for('start'))


@app.route('/', methods=('GET', 'POST'))
def start():

    form = CSVUploadForm()

    if form.validate_on_submit():
        f = form.cvs_file.data
        namespace = form.namespace.data
        element = form.root_element.data

        namespace = namespace.strip()
        element = element.strip()

        client = session['csrf_token']
        folder = os.path.join(app.config['UPLOAD_FOLDER'], client)

        try:
            os.mkdir(folder)
        except OSError:
            print(f"Directory {folder} failed")

        now = time.time()
        data_folder = app.config['UPLOAD_FOLDER']
        for fi in os.listdir(data_folder):
            if os.stat(os.path.join(data_folder, fi)).st_mtime < now - 2 * 86400:
                if os.path.isdir(os.path.join(data_folder, fi)):
                    shutil.rmtree(os.path.join(data_folder, fi))


        filename = secure_filename(f.filename)
        path = os.path.join(folder, filename)
        f.save(path)
        headers = list()
        with open(path, encoding='utf-8-sig', newline='') as csvfile:
            reader = csv.reader(csvfile)
            for row in reader:
              for header in row:
                headers.append(header)
              break

        if len(headers) == 0:
            flash("Could Not Find CSV Headers.. Please Upload A Different CSV File")
            return render_template('index.html', form=form)

        if " " in element:
            flash("Root Element Names Should Not Contain a Space")
            return render_template('index.html', form=form)


        session['HEADER'] = headers
        session['NS'] = namespace
        session['ROOT'] = element
        session['CSV'] = path
        session['client'] = client

        return redirect(url_for('select'))

    return render_template('index.html', form=form)





import csv
import os
import shutil
import time
import xml
import xml.etree.ElementTree as ET
import zipfile
from xml.dom.minidom import parse

from flask import Flask, redirect, url_for, session, send_file, flash
from flask import render_template
from flask_bootstrap import Bootstrap
from flask_wtf import FlaskForm
from flask_wtf.file import FileField, FileAllowed, FileRequired
from pyPreservica import cvs_to_xml, csv_to_search_xml, cvs_to_cmis_xslt, cvs_to_xsd
from werkzeug.utils import secure_filename
from wtforms import SelectField, RadioField, SubmitField, FieldList, StringField
from wtforms.validators import Length, DataRequired

opex_ns = {"opex": "http://www.openpreservationexchange.org/opex/v1.2"}

xml.etree.ElementTree.register_namespace("oai_dc", "http://www.openarchives.org/OAI/2.0/oai_dc/")
xml.etree.ElementTree.register_namespace("ead", "urn:isbn:1-931666-22-9")
xml.etree.ElementTree.register_namespace("opex", opex_ns['opex'])

NS_HELP = """
        This will become the default namespace of the XML documents, e.g. http://www.openarchives.org/OAI/2.0/oai_dc/

        XML Namespaces are any valid URI or Uniform Resource Identifier.
        """
ROOT_HELP = 'This is Root Element of the XML documents, e.g. oai_dc, dc, Metadata, Root etc '


class DownloadForm(FlaskForm):
    xml_button = SubmitField('Download XML Documents as ZIP File')
    xsd_button = SubmitField('Download XSD Schema')
    search_button = SubmitField('Download Custom Search Index')
    cmis_button = SubmitField('Download Custom CMIS Transform for UA')


class CSVUploadForm(FlaskForm):
    root_element = StringField('The Root Element Name', description=ROOT_HELP, validators=[DataRequired(), Length(max=25)])
    namespace = StringField('Default Namespace For The Root Element', description=NS_HELP, validators=[DataRequired()])
    cvs_file = FileField("CSV File", validators=[FileRequired(), FileAllowed(['csv'], 'CSV Files Only')],
                         description='Make sure the Excel Spreadsheet has been saved as UTF-8 CSV',
                         render_kw={"accept": ".csv,.CSV"})
    submit_button = SubmitField('Upload Spreadsheet')


class ColumnSelect(FlaskForm):
    column = SelectField('Unique CSV Column',
                         description='Select a column which will be used to name the XML files. The value of this '
                                     'column should be different for each row of the spreadsheet. '
                                     'If the spreadsheet has a column containing the file name, use that column',
                         coerce=str)
    options = list()
    options.append((".xml", "XML Convention (.xml)"))
    options.append((".metadata", "Preservica Convention: Compatible with SIP Creator (.metadata)"))
    options.append((".opex", "Preservica Convention: Compatible with PUT Tool (.opex)"))
    xml_extension = RadioField(label="Select XML Naming Convention", description="Select XML Naming Convention",
                               validators=[DataRequired()], choices=options, default='.xml')

    options_format = list()
    options_format.append(("pretty", "Format The XML for Humans"))
    options_format.append(("basic", "Leave it Compact for Computers"))
    xml_formatting = RadioField(label="Select XML Formatting", description="XML Formatting",
                                validators=[DataRequired()], choices=options_format, default='pretty')

    options_exclude_name = list()
    options_exclude_name.append(("include", "Include the Unique CSV Column in the XML"))
    options_exclude_name.append(("exclude", "Exclude the Unique CSV Column from the XML"))
    xml_exclude_name = RadioField(label="Include Unique CSV Column", description="",
                                  validators=[DataRequired()], choices=options_exclude_name, default='include')

    submit_button = SubmitField('Generate XML', render_kw={"onclick": "showcursor()"})

    optional_additional_namespaces = FieldList(StringField(label="", description="", render_kw={"size": "75"}),
                                               min_entries=0, max_entries=25)


app = Flask('app')
bootstrap = Bootstrap(app)

SECRET_KEY = os.urandom(32)
app.config['SECRET_KEY'] = SECRET_KEY
app.secret_key = SECRET_KEY
app.config['UPLOAD_FOLDER'] = '/home/opextest/mysite/static/'


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
    options = list()
    prefixes = set()

    headers = session['HEADER']
    for h in headers:
        options.append((h, h))
        # check for prefixes
        if ":" in h:
            head, sep, tail = h.partition(":")
            prefixes.add(head)

    for p in sorted(prefixes):
        field = form.optional_additional_namespaces.append_entry()
        field.label = f"xmlns:{p}"

    form.column.choices = options

    if form.validate_on_submit():
        column = form.column.data

        extra_ns = {}
        for prefix, namespace in zip(sorted(prefixes), form.optional_additional_namespaces.entries):
            if namespace.data:
                extra_ns[prefix] = namespace.data

        namespace = session['NS']
        element = session['ROOT']
        path = session['CSV']

        xml_extension = form.xml_extension.data
        xml_format = form.xml_formatting.data
        xml_exclude = form.xml_exclude_name.data

        client = session['client']
        folder = os.path.join(app.config['UPLOAD_FOLDER'], client)

        zipFile = os.path.join(folder, "xml.zip")
        try:
            os.remove(zipFile)
        except OSError:
            pass

        with zipfile.ZipFile(zipFile, 'w') as myzip:
            for xml_file in cvs_to_xml(csv_file=path, root_element=element, xml_namespace=namespace,
                                       file_name_column=column,
                                       export_folder=folder, additional_namespaces=extra_ns):
                file_name = os.path.basename(xml_file)

                if xml_exclude == "exclude":
                    tree = ET.parse(xml_file)
                    root = tree.getroot()
                    tag = ""
                    for item in root:
                        _, _, tag = item.tag.rpartition('}')
                        if tag == column.replace(" ", ""):
                            print(tag)
                            root.remove(item)
                    fd = open(xml_file, "w", encoding="utf-8")
                    fd.write(ET.tostring(root, encoding="UTF-8").decode("utf-8"))
                    fd.close()

                if xml_extension == ".opex":
                    opex_doc = xml.etree.ElementTree.Element(ET.QName(opex_ns["opex"], 'OPEXMetadata'))
                    dm = xml.etree.ElementTree.SubElement(opex_doc, ET.QName(opex_ns["opex"], "DescriptiveMetadata"))
                    with open(xml_file, 'r', encoding="utf-8") as md:
                        tree = xml.etree.ElementTree.parse(md)
                        dm.append(tree.getroot())
                        fd = open(xml_file, "w", encoding="utf-8")
                        fd.write(ET.tostring(opex_doc, encoding="UTF-8").decode("utf-8"))
                        fd.close()
                    head, _sep, tail = file_name.rpartition(".")
                    file_name = head +".opex"

                if xml_extension == ".metadata":
                    head, _sep, tail = file_name.rpartition(".")
                    file_name = head + xml_extension

                if xml_format == "pretty":
                    dom = xml.dom.minidom.parse(xml_file)
                    myzip.writestr(zinfo_or_arcname=file_name, data=dom.toprettyxml(encoding="UTF-8"))
                else:
                    myzip.write(xml_file, arcname=file_name)
                try:
                    os.remove(xml_file)
                except OSError:
                    pass

        search_xml = csv_to_search_xml(csv_file=path, root_element=element, xml_namespace=namespace,
                                       title="Metadata Title", export_folder=folder, additional_namespaces=extra_ns)

        if xml_format == "pretty":
            dom = xml.dom.minidom.parse(search_xml)
            xml_string = dom.toprettyxml(encoding="UTF-8").decode("utf-8")
            f = open(search_xml, "wt", encoding="UTF-8")
            f.write(xml_string)
            f.close()

        xsd = cvs_to_xsd(csv_file=path, root_element=element, xml_namespace=namespace, export_folder=folder,
                         additional_namespaces=extra_ns)
        if xml_format == "pretty":
            dom = xml.dom.minidom.parse(xsd)
            xml_string = dom.toprettyxml(encoding="UTF-8").decode("utf-8")
            f = open(xsd, "wt", encoding="UTF-8")
            f.write(xml_string)
            f.close()

        cmis = cvs_to_cmis_xslt(csv_file=path, root_element=element, xml_namespace=namespace, title="Metadata Title",
                                export_folder=folder, additional_namespaces=extra_ns)
        if xml_format == "pretty":
            dom = xml.dom.minidom.parse(cmis)
            xml_string = dom.toprettyxml(encoding="UTF-8").decode("utf-8")
            f = open(cmis, "wt", encoding="UTF-8")
            f.write(xml_string)
            f.close()

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
    except FileNotFoundError:
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
            flash("Could Not Find CSV Headers. Please Upload A Different CSV File")
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

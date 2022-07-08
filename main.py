from flask import Flask, render_template, send_from_directory
from flask_wtf import FlaskForm
from waitress import serve
from werkzeug.utils import secure_filename
from wtforms import FileField, SubmitField

from wtforms.validators import InputRequired
import os

from excel_app.services import process_xl

app = Flask(__name__)
app.secret_key = 'super secret key'
app.config['UPLOAD_FOLDER'] = 'static/files'

class UploadFileForm(FlaskForm):
    file = FileField('File', validators=[InputRequired()])
    submit = SubmitField('Download File') 

@app.route('/', methods=['GET', 'POST'])
def home_api():

    form = UploadFileForm()
    if form.validate_on_submit():

        file = form.file.data
        file.save(os.path.join(os.path.abspath(os.path.dirname(__file__)),app.config['UPLOAD_FOLDER'],secure_filename(file.filename)))
        process_xl(os.path.join(os.path.abspath(os.path.dirname(__file__)),app.config['UPLOAD_FOLDER'],secure_filename(file.filename)))

        #return 'File has been uploaded'
        return send_from_directory('static/files', file.filename )
    return render_template('index.html', form=form)
    

if __name__ == '__main__':
    serve(app)

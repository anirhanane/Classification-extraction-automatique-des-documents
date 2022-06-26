import os
from flask import render_template, request
from werkzeug.utils import secure_filename
from app import app
import pandas as pd

from app.modeleML import PreparationDataset


@app.route('/admin/updateDataset')
def updateDataset():
    datadocs = pd.DataFrame(columns=['Classe', 'texte'])
    datadocs = PreparationDataset(datadocs)
    return datadocs


@app.route('/admin/trainModel', methods=['GET', 'POST'])
def trainModel():
    if request.method == 'POST':
        # if 'file' not in request.files:
        # flash('No file part')
        # return redirect(request.url)
        f = request.files['File']

        if f.filename != '':
            filename = secure_filename(f.filename)
            f.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            return "File saved successfully"

import os
import shutil
import urllib
import requests
from flask import render_template, request, jsonify, make_response
from werkzeug.utils import secure_filename
from app import app
from app.modeleML import ClassificationFF2, ClassificationFF4


@app.route('/')
def homepage():
    # files = os.listdir(app.config['UPLOAD_FOLDER'])
    # return render_template('home.html', files=files)
    return render_template('home.html')


@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        # if 'file' not in request.files:
        # flash('No file part')
        # return redirect(request.url)
        f = request.files['File']

        if f.filename != '':
            print("Entree")
            filename = secure_filename(f.filename)
            path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            print(path)
            f.save(path)
            print("Saving file")
            monRepertoireimg = os.path.join(app.config['UPLOAD_FOLDER'] + '/tmp')
            score, pred, proba = ClassificationFF2(os.path.join(app.config['UPLOAD_FOLDER']) + '/' + filename,
                                                   monRepertoireimg)
            if os.path.exists(os.path.join(app.config['UPLOAD_FOLDER'], filename)):
                os.remove(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            else:
                print("Impossible de supprimer le fichier car il n'existe pas")
            return render_template('results.html', score=score[0], probacv=round(proba[0][0][0], 2),
                                   probalm=round(proba[0][0][1], 2))


@app.route('/api/classification/<path>/', methods=['GET'])
def classification(path):
    if path != '':
        path = path.replace(os.sep, '/')
        filename = path.split('/')[-1]
        try:
            f = open(os.path.join('../' + app.config['UPLOAD_FOLDER'], filename), 'wb')
            f.write(requests.get(path).content)
            f.close()
        except IOError as e:
            print("I/O error({0})".format(e))
        print(filename)
        # filename = os.path.basename(path)
        # destination = os.path.normpath(os.path.join('../' + app.config['UPLOAD_FOLDER'], filename))
        # f.save(os.path.join('../' + app.config['UPLOAD_FOLDER'], filename))
        monRepertoireimg = os.path.join('../' + app.config['UPLOAD_FOLDER'] + '/tmp')
        # shutil.copyfile(path, destination)
        prediction = ClassificationFF4(os.path.join('../' + app.config['UPLOAD_FOLDER']) + '/' + filename,
                                       monRepertoireimg)
        response = make_response(
            jsonify(
                {"classe": str(prediction)}
            ),
            200,
        )
        if os.path.exists(os.path.join('../' + app.config['UPLOAD_FOLDER'], filename)):
            os.remove(os.path.join('../' + app.config['UPLOAD_FOLDER'], filename))
        else:
            print("Impossible de supprimer le fichier car il n'existe pas")
        response.headers["Content-Type"] = "application/json"
        return response


@app.route('/api/classification2/<path>/', methods=['GET'])
def classification2(path):
    if path != '':
        '''path = path.replace(os.sep, '/')
        filename = path.split('/')[-1]
        try:
            f = open(os.path.join('../' + app.config['UPLOAD_FOLDER'], filename), 'wb')
            f.write(requests.get(path).content)
            f.close()
        except IOError as e:
            print("I/O error({0})".format(e))'''
        filename = os.path.basename(path)
        destination = os.path.normpath(os.path.join('../' + app.config['UPLOAD_FOLDER'], filename))
        # f.save(os.path.join('../' + app.config['UPLOAD_FOLDER'], filename))
        monRepertoireimg = os.path.join('../' + app.config['UPLOAD_FOLDER'] + '/tmp')
        shutil.copyfile(path, destination)
        prediction = ClassificationFF4(os.path.join('../' + app.config['UPLOAD_FOLDER']) + '/' + filename,
                                       monRepertoireimg)
        response = make_response(
            jsonify(
                {"classe": str(prediction)}
            ),
            200,
        )
        if os.path.exists(os.path.join('../' + app.config['UPLOAD_FOLDER'], filename)):
            os.remove(os.path.join('../' + app.config['UPLOAD_FOLDER'], filename))
        else:
            print("Impossible de supprimer le fichier car il n'existe pas")
        response.headers["Content-Type"] = "application/json"
        return response


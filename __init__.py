import os
import pickle

import joblib
from flask import Flask, url_for, current_app

UPLOAD_FOLDER = 'uploads'


'''def create_app(test_config=None):
    # create and configure the app
    app = Flask(__name__, instance_relative_config=True)
    app.config['MAX_CONTENT-PATH'] = 20
    app.config['MAX_CONTENT_LENGTH'] = 1024 * 1024
    app.config['UPLOAD_EXTENSIONS'] = ['.jpg', '.png', '.gif']
    app.config['UPLOAD_PATH'] = UPLOAD_FOLDER

    if test_config is None:
        # load the instance config, if it exists, when not testing
        app.config.from_pyfile('config.py', silent=True)
    else:
        # load the test config if passed in
        app.config.from_mapping(test_config)

    # ensure the instance folder exists
    try:
        os.makedirs(app.instance_path)
    except OSError:
        pass

    return app


myapp = create_app()'''
# create and configure the app
app = Flask(__name__)
app.config['MAX_CONTENT-PATH'] = 20
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.config['UPLOAD_EXTENSIONS'] = ['.jpg', '.png', '.gif']
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
#app.config['SERVER_NAME'] = "127.0.0.1:8050"





# from app import initDataModel
from app import views
from app import admin_views



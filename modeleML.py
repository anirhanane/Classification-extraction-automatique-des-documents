import os
import os.path
import shutil

# import pythoncom
# import win32com.client as win32
from flask import url_for
# from win32com import client
# from win32com.client import constants

# Gestion des Dataset
import pandas as pd

# Gestion des caractères
import string
import re
import numpy as np

# conversion des documents pdf
import PIL.Image as pili
import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import cv2

# Gestion des documents xlsx,docx

import docx

# ML Classification
from sklearn.svm import SVC
from sklearn import metrics
from sklearn.metrics import confusion_matrix, f1_score, classification_report
from sklearn.feature_extraction.text import CountVectorizer, TfidfVectorizer
from sklearn.svm import SVC
from sklearn import feature_extraction
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LogisticRegression, SGDClassifier
from sklearn.base import TransformerMixin
from sklearn.naive_bayes import MultinomialNB
from sklearn.ensemble import RandomForestClassifier
from sklearn.tree import DecisionTreeClassifier
from sklearn.neighbors import KNeighborsClassifier
from sklearn.pipeline import Pipeline

# Nlp
from nltk.corpus import stopwords

import spacy
from spacy.lang.fr.stop_words import STOP_WORDS as fr_stop

# serialisation
import pickle
import joblib
from flask import g

from win32com.client import Dispatch

# from app.initDataModel import classifieur1, X_test, y_test

STATIC_FOLDER = r'C:\Users\surfaC\Desktop\Stage\API\APIDocs\app\static'
listClasses = ['CV', 'LM']
model = 'Random_Forest'


# Lister les fichiers d'un repertoire


def pickleDataset(datadocs):
    # Let's pickle it for later use
    datadocs.to_pickle("datasetvf2.pkl")


def depickleDatasetX(dataset):
    # Let's depickle it for later use
    # with app.app_context():
    if 'datadocsx' not in g:
        with open('.' + url_for('static', filename='dataset/' + dataset), 'rb') as f1:
            g.datadocsx = pickle.load(f1)
            return g.datadocsx


def depickleDatasetY(dataset):
    # Let's depickle it for later use
    # with app.app_context():
    if 'datadocsy' not in g:
        with open('.' + url_for('static', filename='dataset/' + dataset), 'rb') as f1:
            g.datadocsy = pickle.load(f1)
            return g.datadocsy


def depickle(model):
    # Let's depickle it for later use
    # with app.app_context():
    if 'model' not in g:
        with open(os.path.join(STATIC_FOLDER + '/classifieurs/' + model), 'rb') as f1:
            g.model = pickle.load(f1)
            return g.model


def loadjob(model):
    # Let's depickle it for later use
    # with app.app_context():
    if 'model' not in g:
        with open('.' + url_for('static', filename='classifieurs/' + model), 'rb') as f1:
            g.model = joblib.load(f1)
            joblib.dump(g.model,
                        r'C:\Users\surfaC\Desktop\Stage\API\APIDocs\app\static\classifieurs\Random_Forest_n.pkl')
            return g.model


def ListFichiers(monRepertoire):
    fichiers = [f for f in os.listdir(monRepertoire) if os.path.isfile(os.path.join(monRepertoire, f))]
    return fichiers


# Définition des répertoires  des documents et celui où les documents convertis en image seront stockés
def initialisation(monRepertoire):
    liste = ListFichiers(monRepertoire)
    liste
    return monRepertoire, liste


def initialisation(monRepertoireimg):
    nlp = spacy.load("fr_core_news_md")
    # Create our list of punctuation marks
    punctuations = string.punctuation
    # stopWords = set(stopwords.words('french'))
    stopWords = list(fr_stop)
    return punctuations, stopWords, nlp


def save_as_docx(path):
    # Opening MS Word
    # pythoncom.CoInitialize()
    # word = client.gencache.EnsureDispatch('Word.Application')
    # word = Dispatch('Word.Application')
    real_path = os.path.realpath(path)
    # doc = word.Documents.Open(real_path)
    # doc.Activate()

    # Rename path with .docx

    new_file_abs = re.sub(r'\.\w+$', '.docx', os.path.abspath(real_path))

    # Save and Close
    # word.ActiveDocument.SaveAs(
    # new_file_abs, FileFormat=constants.wdFormatXMLDocument
    # )
    # doc.Close(False)


def reorganisationDataset(source, destination):
    for dossier, sous_dossiers, fichiers in os.walk(source):
        for fichier in fichiers:
            print(os.path.basename(dossier))
            parname = os.path.basename(os.path.dirname(os.path.normpath(dossier)))
            if os.path.basename(dossier) == "cv":
                print(fichier)

                nomfichier, ext = os.path.splitext(fichier)
                if nomfichier == "cv":
                    shutil.copy(os.path.join(dossier, fichier), destination)
                    newname = '.'.join(fichier.split('.')[:-1]) + "_" + parname + ext
                    if not os.path.exists(os.path.join(destination, newname)):
                        os.rename(os.path.join(destination, fichier), os.path.join(destination, newname))
            if os.path.basename(dossier) == "lm":
                print(fichier)

                nomfichier, ext = os.path.splitext(fichier)
                if nomfichier.startswith("lm") or nomfichier.startswith(
                        'LETTRE_DE_MOTIVATION') or nomfichier.startswith('LettredeMotivation') or nomfichier.startswith(
                    'Lettre_de_motivation') or nomfichier.startswith(
                    'Lettre-de-motivation') or nomfichier.startswith('Lettre-motivation') or nomfichier.startswith(
                    'Lettre_motivation') or nomfichier.startswith('LettreMotivation') or nomfichier.startswith(
                    'lettremotivation') or nomfichier.startswith('lettreMotivation') or nomfichier.startswith(
                    'Lettre_Motivation') or nomfichier.startswith('Lettremotivation') or nomfichier.startswith(
                    'LM'):
                    shutil.copy(os.path.join(dossier, fichier), destination)
                    newname = "lm_" + parname + ext
                    if not os.path.exists(os.path.join(destination, newname)):
                        os.rename(os.path.join(destination, fichier), os.path.join(destination, newname))


def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)


# conversion d'une image en texte en utilissant pillow et pytesseract
def imageTotexte(filename):
    """
    This function will handle the core OCR processing of images.
    """
    custom_config = r'-lang fra+eng+ara --psm 6'
    text = pytesseract.image_to_string(pili.open(filename), config=custom_config)
    return text


def imageTotexte2(img):
    """
    This function will handle the core OCR processing of images.
    """
    custom_config = r'-lang fra+eng+ara --psm 6'
    text = pytesseract.image_to_string(img, config=custom_config)
    return text


# Conversion du fichier pdf à image
def filetoimg(file, monRepertoire, monRepertoireimg):
    pages = convert_from_path(os.path.join(monRepertoire, '') + file, 350,
                              poppler_path=r'C:\Program Files\poppler-21.03.0\Library\bin', first_page=1,
                              last_page=2)
    lstNewimg = []

    for idx, page in enumerate(pages):
        newimg = '.'.join(file.split('.')[:-1]) + str(idx) + '.jpg'
        page.save(os.path.join(monRepertoireimg, '') + newimg, 'JPEG')
        lstNewimg.append(newimg)

    return lstNewimg


# Amélioration de la qualité des images
def monochromemode(image, monRepertoire):
    img = Image.open(os.path.join(monRepertoire, image))
    img = img.convert('L')
    img.save(os.path.join(monRepertoire, image))


def enhacement(image, monRepertoire):
    # load the original image
    imagE_path = os.path.join(monRepertoire, image)
    print(imagE_path)
    im = cv2.imread(imagE_path)
    if (im is not None):
        gray = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY)
        blur = cv2.GaussianBlur(gray, (9, 9), 0)
        thresh = cv2.adaptiveThreshold(blur, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 11, 30)

    return thresh


# transformation d'un pdfs en images puis en texte et retourne le texte
def TransformationPdfImage(pathPdf, monRepertoireimg):
    pdfRepertoire = os.path.dirname(pathPdf)
    print(pdfRepertoire)
    f = os.path.basename(pathPdf)
    print(f)
    images = filetoimg(f, pdfRepertoire, monRepertoireimg)
    texte = ""
    # for tmp in images :
    #    texte=texte+"\n"+imageTotexte2(enhacement(tmp))
    for img in images:
        monochromemode(img, monRepertoireimg)
        texte = texte + "\n" + imageTotexte(os.path.join(monRepertoireimg, img))
    for img in images:
        if os.path.exists(os.path.join(monRepertoireimg, img)):
            os.remove(os.path.join(monRepertoireimg, img))
        else:
            print("Impossible de supprimer le fichier car il n'existe pas")
    print(texte)
    return texte


def getText_pdf(f, monRepertoire, monRepertoireimg):
    images = filetoimg(f, monRepertoire)
    print(images)
    texte = ""
    for img in images:
        monochromemode(img)
        texte = texte + "\n" + imageTotexte(os.path.join(monRepertoireimg, '') + img)
    return texte


# transformation docx et des pdfs en images puis en textes et stocker le texte dans un dataframe
def PreparationDataset(datadocs, liste, monRepertoire, monRepertoireimg):
    texte = ""
    for f in liste:
        try:
            print(f)
            if f.startswith('cv'):
                print("cv")
                nomfichier, ext = os.path.splitext(f)
                print(ext)
                if ext == ".pdf":
                    print("cv.pdf")
                    texte = getText_pdf(f, monRepertoire, monRepertoireimg)
                    datadocs = datadocs.append({'Classe': 1, 'texte': texte}, ignore_index=True)
                elif ext == '.docx':
                    print("cv.docx")
                    texte = getText(f)
                    datadocs = datadocs.append({'Classe': 1, 'texte': texte}, ignore_index=True)
                elif ext == '.doc':
                    print("cv.doc")
                    texte = getText(f)
                    datadocs = datadocs.append({'Classe': 1, 'texte': texte}, ignore_index=True)

            if f.startswith('lm'):
                print("lm")
                nomfichier, ext = os.path.splitext(f)
                if ext == ".pdf":
                    print("lm.pdf")
                    texte = getText_pdf(f, monRepertoire, monRepertoireimg)
                    datadocs = datadocs.append({'Classe': 2, 'texte': texte}, ignore_index=True)
                elif ext == '.docx':
                    print("lm.docx")
                    texte = getText(os.path.join(monRepertoire, f))
                    datadocs = datadocs.append({'Classe': 2, 'texte': texte}, ignore_index=True)
                elif ext == '.doc':
                    print("lm.doc")
                    texte = getText(os.path.join(monRepertoire, f))
                    datadocs = datadocs.append({'Classe': 2, 'texte': texte}, ignore_index=True)

        except ValueError as e:
            print(str(e))
            pass
    return datadocs


# Application des techniques de nettoyage
def cleaning_round1(text):
    '''Make text lowercase, remove text in square brackets, remove punctuation and remove words containing numbers.'''
    text = text.lower()
    text = re.sub('\[-.*?\]', ' ', text)
    text = re.sub('[%s]' % re.escape(string.punctuation), ' ', text)
    text = re.sub('\s+', ' ', text)
    text = re.sub('^\s+|\s+?$', ' ', text)
    text = re.sub('\w*\d\w*', ' ', text)
    text = re.sub('\s-+\s', ' ', text)
    text = re.sub('\s-{2}', ' ', text)
    text = re.sub('\s-{3}', ' ', text)
    text = re.sub('\s-{4}', ' ', text)
    text = re.sub('\s—+', ' ', text)
    return text


# Apply a second round of cleaning
def cleaning_round2(text):
    '''Get rid of some additional punctuation and non-sensical text that was missed the first time around.'''
    text = re.sub('[«»°‘’“”—…$§¢€%]', ' ', text)
    text = re.sub('\n', ' ', text)
    return text


round2 = lambda x: cleaning_round2(x)


def cleaning(text):
    '''Get rid of some additional punctuation and non-sensical text that was missed the first time around.'''
    text = re.sub('\n', ' ', text)
    return text


# Apply a third round of cleaning
def cleaning_round3(text):
    '''remove stop words'''
    # Create our list of stopwords
    stopWords = set(stopwords.words('french'))
    text = ' '.join(term for term in text.split() if term not in stopWords)
    return text


# Apply a four round of cleaning
def cleaning_round4(text, nlp):
    '''remove lemmatization'''
    lemmatizer = nlp.get_pipe("lemmatizer")
    doc = nlp(text)

    text = ' '.join(token.lemma_ for token in doc)
    return text


def return_token(sentence, nlp):
    # Tokeniser la phrase
    doc = nlp(sentence)
    # Retourner le texte de chaque token
    text = ' '.join(X.text for X in doc)
    return text


def PretraitementTexte(X, monRepertoireimg):
    filename, file_extension = os.path.splitext(X)
    texte = ''
    if file_extension == '.pdf':
        texte = TransformationPdfImage(X, monRepertoireimg)
    elif file_extension == '.doc' or file_extension == '.docx':
        if file_extension == '.doc':
            save_as_docx(X)
            X = os.path.abspath(X)
            X = re.sub(r'\.\w+$', '.docx', X)
        texte = getText(X)
    else:
        print("Impossible de traiter ce document")
    # tokenisation(cleaning_round4(
    return cleaning_round3(cleaning_round2(cleaning_round1(texte)))


# Creating our tokenizer function
def spacy_tokenizer(sentence, nlp):
    # Creating our token object, which is used to create documents with linguistic annotations.
    mytokens = nlp(sentence)

    # Lemmatizing each token and converting each token into lowercase
    mytokens = [word.lemma_.lower().strip() if word.lemma_ != "-PRON-" else word.lower_ for word in mytokens]

    # Removing stop words
    # mytokens = [ word for word in mytokens if word not in stopWords and word not in punctuations ]

    # return preprocessed list of tokens
    return " ".join(mytokens)


def vectorisation_tfidf(stopWords, datadocs):
    vect = TfidfVectorizer(stop_words=stopWords, analyzer='word', min_df=10)  # , ngram_range=(1,2)
    tfidf_mat = vect.fit_transform(datadocs.texte)
    print(len(vect.get_feature_names()))
    feature_names = vect.get_feature_names()  # les noms des tokens
    dense = tfidf_mat.todense()  # convert sparse matrix to numpy array
    denselist = dense.tolist()  # convert array to list
    df2 = pd.DataFrame(dense.tolist(), columns=feature_names)
    df2.head()
    return vect, df2


def vectorisation_tfidf2(stopWords, datadocs):
    # define vectorizer parameters
    tfidf_vectorizer = TfidfVectorizer(max_df=0.8, max_features=1000,
                                       min_df=0.2, stop_words=stopWords,
                                       use_idf=True, tokenizer=spacy_tokenizer, ngram_range=(1, 2))

    tfidf_matrix = tfidf_vectorizer.fit_transform(datadocs.texte)  # fit the vectorizer to texte docs
    print(tfidf_matrix)


def vectorisation_tf(stopWords, datadocs):
    bow_vector = CountVectorizer(ngram_range=(1, 1), stop_words=stopWords, min_df=10)
    bow_vector

    vectdocs = bow_vector.fit_transform(datadocs["texte"])
    # vectdocs = bow_vector.fit(datadocs["texte"])
    print(len(bow_vector.get_feature_names()))
    return bow_vector, vectdocs


# Custom transformer using spaCy
class predictors(TransformerMixin):
    def transform(self, X, **transform_params):
        # Cleaning Text
        return [clean_text(text) for text in X]

    def fit(self, X, y=None, **fit_params):
        return self

    def get_params(self, deep=True):
        return {}


# Basic function to clean the text
def clean_text(text):
    # Removing spaces and converting text into lowercase
    return text.strip().lower()


def ClassificationFF(cheminfile, datadocs, bow_vector, monRepertoireimg):
    # Define models to train
    names = ['Random_Forest.joblib', 'Logistic_regression']
    classifiers = [RandomForestClassifier(), LogisticRegression()]
    models = zip(names, classifiers)
    X = datadocs['texte']  # the features we want to analyze
    ylabels = datadocs['Classe']  # the classe
    ylabels = ylabels.astype('int')
    X_train, X_test, y_train, y_test = train_test_split(X, ylabels, test_size=0.3)
    X_validation = PretraitementTexte(cheminfile, monRepertoireimg)
    pred = []
    proba = []
    for name, model in models:
        # Create pipeline using Bag of Words
        pipe = Pipeline([("cleaner", predictors()),
                         ('vectorizer', bow_vector),
                         ('classifier', model)])

        # model generation
        pipe.fit(X_train, y_train)
        # Predicting with a test dataset
        print("Predicting with a test dataset")
        predicted = pipe.predict(X_test)
        # print(predicted)
        # print(y_test)
        print('{}:Accuracy: {}'.format(name, metrics.accuracy_score(y_test, predicted)))
        print('{}:Precision micro: {}'.format(name, metrics.precision_score(y_test, predicted, average='micro')))
        print('{}:Recall micro: {}'.format(name, metrics.recall_score(y_test, predicted, average='micro')))
        print('{}:Precision N: {}'.format(name, metrics.precision_score(y_test, predicted, average=None)))
        print('{}:Recall N: {}'.format(name, metrics.recall_score(y_test, predicted, average=None)))
        print(" Confusion matrix: {}".format(confusion_matrix(predicted, y_test)))
        # Predicting with a validation dataset
        print('Predicting with a validation dataset')
        predicted = pipe.predict(np.array([X_validation]))
        probapred = pipe.predict_proba(np.array([X_validation]))
        y_validation = np.array([1])
        print(" Value Predicted: {}".format(predicted))
        pred.append(predicted)
        proba.append(probapred)
    return pred, proba


def ClassificationFF3(cheminfile, datadocs, bow_vector, monRepertoireimg):
    # Define models to train
    names = ['Random_Forest.joblib', 'Logistic_regression']
    classifiers = [RandomForestClassifier(), LogisticRegression()]
    models = zip(names, classifiers)
    X = datadocs['texte']  # the features we want to analyze
    ylabels = datadocs['Classe']  # the classe
    ylabels = ylabels.astype('int')
    X_train, X_test, y_train, y_test = train_test_split(X, ylabels, test_size=0.3)
    X_validation = PretraitementTexte(cheminfile, monRepertoireimg)
    pred = []
    proba = []
    for name, model in models:
        # Create pipeline using Bag of Words
        pipe = Pipeline([("cleaner", predictors()),
                         ('vectorizer', bow_vector),
                         ('classifier', model)])

        # model generation
        pipe.fit(X_train, y_train)
        print('Predicting with a validation dataset')
        predicted = pipe.predict(np.array([X_validation]))
        probapred = pipe.predict_proba(np.array([X_validation]))
        print(" Value Predicted: {}".format(predicted))
        pred.append(predicted)
        proba.append(probapred)
    return pred, proba


def ClassificationFF2(cheminfile, monRepertoireimg):
    # Define models to train

    X_validation = PretraitementTexte(cheminfile, monRepertoireimg)
    pred = []
    proba = []
    score = []
    classifieur1 = loadjob(model + '1.pkl')
    # classifieur2 = loadjob(model + '2.pkl')
    X_test = depickleDatasetX('Xtest.pkl')
    y_test = depickleDatasetY('ytest.pkl')
    predicted1 = classifieur1.predict(np.array([X_validation]))
    probapred1 = classifieur1.predict_proba(np.array([X_validation]))
    # probapred2 = classifieur2.predict_proba(np.array([X_validation]))
    # predicted2 = classifieur2.predict(np.array([X_validation]))

    score.append(round(metrics.accuracy_score(y_test, classifieur1.predict(X_test)) * 100, 2))
    # score.append(metrics.accuracy_score(y_test, classifieur2.predict(X_test)))
    pred.append(predicted1)
    # pred.append(predicted2)
    proba.append(probapred1 * 100)
    # proba.append(probapred2)
    return score, pred, proba


def ClassificationFF4(cheminfile, monRepertoireimg):
    # Define models to train
    resultat = ''
    classifieur1 = loadjob(model + '1.pkl')
    # classifieur2 = loadjob(model + '2.pkl')
    X_validation = PretraitementTexte(cheminfile, monRepertoireimg)
    predicted1 = classifieur1.predict(np.array([X_validation]))
    prediction = listClasses[int(predicted1) - 1]
    print(predicted1)
    for c, pb in zip(classifieur1.classes_, classifieur1.predict_proba(np.array([X_validation]))[0]):
        print(int(predicted1[0]))
        print(int(c))
        if (int(predicted1[0]) == int(c)):
            if (pb > 0.7):
                resultat = prediction

    return resultat


def nerExtraction(cheminfile, monRepertoireimg):
    texte = PretraitementTexte(cheminfile, monRepertoireimg)
    nlp1 = spacy.load(r'C:\Users\surfaC\Desktop\Stage\API\APIDocs\app\static\nermodel')  # load the best model
    texte2 = cleaning(texte)
    print(texte2)
    doc = nlp1(texte2)
    # spacy.displacy.render(doc, style="ent", jupyter=True) # display in Jupyter
    for ent in doc.ents:
        print('Entities', [(ent.text, ent.label_)])
    return doc.ents

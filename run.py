from app import app
from app.modeleML import predictors
if __name__ == '__main__':

    app.run(port=5000,debug=True)
#port=8050,debug=True,
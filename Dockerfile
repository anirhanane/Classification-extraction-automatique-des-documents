FROM python:3.9-slim

WORKDIR /app

ADD . /app

RUN pip install -r requirements.txt

CMD["uwsgi","app.ini"]
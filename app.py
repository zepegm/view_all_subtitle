from powerpoint import pegarTextoSlideShow
from flask import Flask, render_template, request, redirect, url_for, jsonify, send_file, Response

app=Flask(__name__)
app.secret_key = "abc123"
app.config['SECRET_KEY'] = 'justasecretkeythatishouldputhere'

@app.route('/', methods=['GET', 'POST'])
def index():
    legenda = pegarTextoSlideShow()
    return render_template('subtitle.jinja', legenda=legenda)


if __name__ == '__main__':
    app.run('0.0.0.0',port=80)
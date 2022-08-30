import re
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for
from pymongo import MongoClient
# custom module
import utils

# database 연결
conn = MongoClient("127.0.0.1")
# collection setting
collection = conn.notice.information_notice

app = Flask(__name__)
# enable debugging mode
app.config["DEBUG"] = True


# Root URL
@app.route('/')
def index():
    return render_template('index.html')


# upload & convert excel2html
@app.route('/<string:template_type>', methods=['POST'])
def uploadFiles(template_type):

    # get the uploaded file
    uploaded_file = request.files['file']

    # database data 삽입
    if template_type == 'notice':
        utils.xlsx2db(uploaded_file, collection)

    if template_type == 'congress':
        print('hello')
        pass
    return redirect(url_for('index'))


# load notice html page
@app.route("/<template_type>/<string:notice_id>")
def show_template(template_type , notice_id):
    # db2 df & json
    df, json_file = utils.db2df(notice_id, collection)

    if template_type == 'notice':
        # QnA 형태를 template에 적합 하게 불러 오기
        qna_list = utils.get_notice_df(df)
        # make template
        template = render_template('notice.html', values=json_file, qna_list=qna_list)
        # html file save
        with open(f'./static/html/{notice_id}.html', 'w') as f:
            f.write(template)

    if template_type == 'congress':
        pass
    return template


if __name__ == "__main__":
    app.run(port=5000)

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
@app.route('/', methods=['POST'])
def uploadFiles():
    # get the uploaded file
    uploaded_file = request.files['file']

    # database data 삽입
    utils.xlsx2db(uploaded_file, collection)
    return redirect(url_for('index'))


# load notice html page
@app.route("/notice/<string:notice_id>")
def show_template(notice_id):
    # db2 df & json
    df, json_file = utils.db2df(notice_id, collection)

    # QnA 형태를 template에 적합 하게 불러 오기
    answer_index = df.loc[df.index.str.contains('답변')].index
    qna_list = {}
    for ans_name in answer_index:
        ans_num = re.findall('[0-9]', ans_name)
        qus_list = []
        for num in ans_num:
            df_qus = df.loc[df.index.str.contains(str(num)) & df.index.str.contains('질의')]
            qus_list.append({df_qus.index[0]: df_qus.values[0][0]})
        df_ans = df.loc[df.index.str.contains(str(num)) & df.index.str.contains('답변')]
        qna_list[f"질의 {', '.join(ans_num)} 답변"] = {'qus': qus_list, 'ans': {df_ans.index[0]: df_ans.values[0]}}

    # make template
    template = render_template('notice.html', values=json_file, qna_list=qna_list)

    # html file save
    with open(f'./static/html/{notice_id}.html', 'w') as f:
        f.write(template)

    return template


if __name__ == "__main__":
    app.run(port=5000)

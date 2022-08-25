import re
import pandas as pd
from pymongo import MongoClient
from datetime import datetime


def isnan(num):
    return num != num


def xlsx2db(uploaded_file, collection):
    '''
    :param uploaded_file:
    :param collection: DB collection
    :return: notice-id
    '''
    # get name to meta info
    notice_id = re.sub(r'[^0-9]', '', uploaded_file.filename)
    name = uploaded_file.filename.split('.')[0]

    # xlsx2df
    df = pd.read_excel(uploaded_file)
    df.columns = ['key', 'value']
    data = {}
    for i in range(len(df)):
        data[df.iloc[i]['key']] = df.iloc[i]['value']

    # json2DB
    json_file = {'notice-id': notice_id, 'name': name, 'data': data}

    # null 값일 때
    if isnan(json_file['data']['접수일']):
        json_file['data']['접수일'] = datetime.today().strftime('%y/%m/%d')
    if isnan(json_file['data']['접수번호']):
        json_file['data']['접수번호'] = '000000'
    if isnan(json_file['data']['담당자']):
        json_file['data']['담당자'] = 'OOO 주무관'
    if isnan(json_file['data']['전화번호']):
        json_file['data']['전화번호'] = '043-0000-0000'
    if isnan(json_file['data']['인사말']):
        json_file['data']['인사말'] = '안녕하십니까? 귀하께서 국민신문고를 통해 신청하신 민원에 대한 검토 결과를 다음과 같이 알려드립니다.'
    if isnan(json_file['data']['끝인사']):
        json_file['data']['끝인사'] = '본 회신에 대한 자세한 설명이 필요한 경우는 행정안전부 정보공개정책과 OOO 주무관에게 문의 주시면 친절하게 안내해드리겠습니다.'
    if isnan(json_file['data']['붙임']):
        json_file['data']['붙임'] = '없음'

    json_file['data']['붙임'] = json_file['data']['붙임'].split('\n')

    collection.insert_one(json_file)
    print(f'{notice_id} saved in DB')
    return notice_id


def db2df(notice_id, collection):
    # get id from DB document
    json_file = collection.find_one({'notice-id': notice_id})

    data = json_file['data']
    data['notice-id'] = json_file['notice-id']

    # json2df
    df = pd.DataFrame(data.values(), index=data.keys())
    df.columns = ['value']

    return df, json_file

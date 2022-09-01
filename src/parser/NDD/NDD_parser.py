import os
from bs4 import BeautifulSoup
import warnings

warnings.filterwarnings(action='ignore')


def get_NDD_data(file_path):

    # load data
    file = open(file_path, 'r')
    soup = BeautifulSoup(file, 'html.parser')

    NDD_dict = {}

    # get meta data
    meta_info = [x.text.strip() for x in soup.findAll('p' , class_ ='value')]
    NDD_dict['접수번호'] = meta_info[0]
    NDD_dict['접수일'] = meta_info[1]
    NDD_dict['공개여부'] = meta_info[2]
    NDD_dict['담당자'] = meta_info[3]
    NDD_dict['전화번호'] = meta_info[4]

    # get greetings
    NDD_dict['인사말'] = soup.find('p' , id = 'greetings').text.strip()

    # get qna
    qus_list = [x.text for x in soup.findAll('p' , class_ ="question")]
    ans_list = [x.text for x in soup.findAll('p' , class_ ="answer")]
    qna_dict = {}
    for i , (qus , ans) in enumerate(zip(qus_list , ans_list) , 1):
        qna_dict[f'질의{i}'] = qus
        qna_dict[f'답변{i}'] = ans

    NDD_dict = { **NDD_dict , **qna_dict}

    # get closing ,ment
    NDD_dict['끝인사'] = soup.find('p' , id ="closing").text.strip()

    # get attached
    NDD_dict['붙임'] = soup.find('p' , id='attached').text.strip()
    return NDD_dict

# if __name__== "__main__":
#     NDD_dict = get_NDD_data('/Users/mac/PycharmProjects/document-transformation/src/template_generator/NDD/static/html/0113.html')
#     print(NDD_dict)

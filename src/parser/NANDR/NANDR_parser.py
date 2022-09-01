from bs4 import BeautifulSoup
import warnings

warnings.filterwarnings("ignore")


def get_manager(soup):
    manager_list = []
    manager_info_list = soup.findAll('hp:tr', dtype="manager")
    for info in manager_info_list:
        manager_dict = {}
        manager_dict['부서'] = info.find('hp:tc', dtype="depart").text.strip()
        manager_dict['이름'] = info.find('hp:tc', dtype="name").text.strip()
        manager_dict['전화번호'] = info.find('hp:tc', dtype="number").text.strip()
        manager_list.append(manager_dict)

    manager_all_dict = {f"실무자{i}": manager for i, manager in enumerate(manager_list, 1)}
    return manager_all_dict


def get_request(soup):
    request_info_list = soup.findAll('hp:tr', dtype="request")
    request_dict = {}
    for request in request_info_list:
        request_list = [x.text for x in request.findAll('hp:t')]
        request_dict[request_list[0]] = request_list[1]

    return request_dict


def get_detail_request(soup):
    detail_request_info_list = soup.findAll('hp:tr', dtype="detail_request")
    detail_request_dict = {}
    for detail_request in detail_request_info_list:
        request_list = [x.text for x in detail_request.findAll('hp:t')]
        detail_request_dict[request_list[-2]] = request_list[-1]

    return detail_request_dict


def get_attached(soup):
    attach_info = soup.find('hp:tc', dtype="attach")
    attach_list = attach_info.text.strip().split('\n')
    for i, attach in enumerate(attach_list, 1):
        attached_dict = {f'붙임{i}': attach}
    return attached_dict

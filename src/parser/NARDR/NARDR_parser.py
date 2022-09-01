from bs4 import BeautifulSoup
import warnings

warnings.filterwarnings("ignore")


def _get_manager(soup):
    manager_list = []
    manager_info_list = soup.findAll('hp:tr', dtype="manager")
    manager_dict = {}
    for i , info in enumerate(manager_info_list , 1):

        manager_dict[f'실무자{i}부서'] = info.find('hp:tc', dtype="depart").text.strip()
        manager_dict[f'실무자{i}이름'] = info.find('hp:tc', dtype="name").text.strip()
        manager_dict[f'실무자{i}전화번호'] = info.find('hp:tc', dtype="number").text.strip()
        manager_list.append(manager_dict)

    return manager_dict


def _get_request(soup):
    request_info_list = soup.findAll('hp:tr', dtype="request")
    request_dict = {}
    for request in request_info_list:
        request_list = [x.text for x in request.findAll('hp:t')]
        request_dict[request_list[0]] = request_list[1]

    return request_dict


def _get_detail_request(soup):
    detail_request_info_list = soup.findAll('hp:tr', dtype="detail_request")
    detail_request_dict = {}
    for detail_request in detail_request_info_list:
        request_list = [x.text for x in detail_request.findAll('hp:t')]
        detail_request_dict[request_list[-2]] = request_list[-1]

    return detail_request_dict


def _get_attached(soup):
    attach_info = soup.find('hp:tc', dtype="attach")
    attach_list = attach_info.text.strip().split('\n')
    for i, attach in enumerate(attach_list, 1):
        attached_dict = {f'붙임{i}': attach}
    return attached_dict


def get_NANDR_data(file_path):
    file = open(file_path, 'r')
    soup = BeautifulSoup(file, 'html.parser')
    manager =_get_manager(soup)
    request = _get_request(soup)
    detail_request = _get_detail_request(soup)
    attached = _get_attached(soup)

    all_dict = {**manager , **request , **detail_request , **attached}

    return all_dict


# if __name__ == "__main__":
#     file_path =  '/Users/mac/PycharmProjects/document-transformation/src/template_generator/NARDR/output/Contents/section0.xml'
#     NANDR_data = get_NANDR_data(file_path)
#     print(NANDR_data)
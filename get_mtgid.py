# coding=utf-8
import requests
from bs4 import BeautifulSoup

import re


def get_mtgid_list(main_contents, latest_only=False):
    tag_list = main_contents.find(class_='tableCommon').find_all('a')

    mtgid_list = []
    for tmp in tag_list:
        if 'onclick' not in tmp.attrs:
            continue
        mtgid = re.findall(r'([0-9]+\.?[0-9]*)', tmp.get('onclick'))[0]
        mtgid_list.append(mtgid)
    if latest_only:
        return mtgid_list[0]
    else:
        return mtgid_list


def get_latest_mtgid(base_url):
    r = requests.get(base_url)
    bs_obj = BeautifulSoup(r.text, 'html.parser')
    main_contents = bs_obj.find(class_='contentDiv')
    mtgid = get_mtgid_list(main_contents, latest_only=True)
    return mtgid


if __name__ == '__main__':
    base_url = 'http://tmcsupport.php.xdomain.jp/urawan'
    tmp = get_latest_mtgid(base_url)
    print(tmp)
    base_url = 'http://tmcsupport.php.xdomain.jp/shimazuyamax/'
    tmp = get_latest_mtgid(base_url)
    print(tmp)



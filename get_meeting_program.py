# coding=utf-8
import re

import openpyxl
import requests
from bs4 import BeautifulSoup


def get_event_info(main_contents):
    event_tags = main_contents.find(class_='tableCommon')
    key_list = [t.text for t in event_tags.find_all(class_='titleList')]
    value_list = [t.text for t in event_tags.find_all(class_='nowrap')]
    event_info_dict = {k: v for k, v in zip(key_list, value_list)}
    return event_info_dict


def get_guest_info(main_contents):
    guest_tag = main_contents.find(id='span4guest')
    guest_detail_tags = guest_tag.find_all(class_='span4guestnm')
    guest_info_dict = {}
    for i, tag in enumerate(guest_detail_tags):
        guest_info_dict[str(i)] = tag.text.replace('\n', '').strip()
    return guest_info_dict


def get_assign_info(main_contents):
    assign_tag = main_contents.find(class_='tableCommon mainTbl')
    assign_tag_list = assign_tag.find_all(id=re.compile('^roleTr'))

    assign_dict = {}
    for tag in assign_tag_list:
        role = tag.find(id='rolenm').get('value')
        member = tag.find(id='asignnm').get('value').replace('　', ' ')
        if member == '':
            continue

        if re.match(r'スピーチ\d', role):
            role_detail_cnts = tag.find(class_='roleDetailStr1 nowrap').contents
            if len(role_detail_cnts) == 0:
                role_detail = ''
            elif len(role_detail_cnts) == 1:
                role_detail = str(role_detail_cnts[0].replace('\xa0', ' '))
            else:
                role_detail = '\n'.join([role_detail_cnts[0], role_detail_cnts[2], role_detail_cnts[4]])
            title = tag.find(id='roleDetail2Td').text
            time_range = tag.find(id='roleTmTd').text.replace('　', ' ')
            if time_range != '':
                time_min, time_max = re.findall(r'([0-9]+\.?[0-9]*)', time_range)
            else:
                time_min, time_max = '', ''
            assign_dict[role] = {
                'member': member,
                'role_detail': role_detail,
                'title': title,
                'time': {'min': time_min, 'max': time_max},
            }
        else:
            role_detail = tag.find(class_='roleDetailStr1 nowrap').text.replace('　', ' ')
            time_range = tag.find(id='roleTmTd').text.replace('　', ' ')
            assign_dict[role] = {
                'member': member,
                'role_detail': role_detail,
                'time': time_range
            }
    return assign_dict


def get_attendance_info(main_contents):
    attendance_info_dict = {}
    status_list = [sorted(tag.text.replace('　', ' ').replace('\xa0', '').replace('\n', ',').split(',')[1:3])
                   for tag in main_contents.find_all(bgcolor='white')]
    for data in status_list:
        status, member = data[0], data[1]
        if member in attendance_info_dict:
            continue
        attendance_info_dict[member] = status
    status_list = [sorted(tag.text.replace('　', ' ').replace('\xa0', '').replace('\n', ',').split(',')[1:3])
                   for tag in main_contents.find_all(bgcolor='aliceblue')]
    for data in status_list:
        status, member = data[0], data[1]
        if member in attendance_info_dict:
            continue
        attendance_info_dict[member] = status
    return attendance_info_dict


def set_event_info(sheet, meeting_info_dict):
    for i in range(3, 7 + 1):
        key = sheet['O{}'.format(i)].value
        sheet['P{}'.format(i)] = meeting_info_dict[key]


def set_assign_info(sheet, assign_info_dict):
    index = 10  # テンプレートを参照のこと
    ignore_time = 0
    while True:
        key = sheet['O{}'.format(index)].value
        if key in assign_info_dict:
            ignore_time = 0
            data = assign_info_dict[key]
            sheet['P{}'.format(index)] = '{}さん'.format(data['member'])
            if re.match(r'スピーチ\d', key):
                sheet['Q{}'.format(index)] = data['role_detail']
                sheet['R{}'.format(index)] = data['title']
                sheet['S{}'.format(index)] = '00:{:02d}:00'.format(int(data['time']['min'])) \
                    if data['time']['min'] != '' else ''
                sheet['T{}'.format(index)] = '00:{:02d}:00'.format(int(data['time']['max'])) \
                    if data['time']['max'] != '' else ''
            else:
                sheet['Q{}'.format(index)] = data['role_detail']
                sheet['S{}'.format(index)] = '00:{:02d}:00'.format(int(data['time'])) if data['time'] != '' else ''
        else:
            ignore_time += 1

        if ignore_time == 5:
            break

        index += 1


def set_attendance_info(sheet, attendance_info_dict):
    index = 3  # テンプレートを参照のこと
    ignore_time = 0
    while True:
        key = sheet['V{}'.format(index)].value
        if hasattr(key, 'replace') and key.replace('　', ' ').replace('さん', '') in attendance_info_dict:
            key = key.replace('　', ' ').replace('さん', '')
            sheet['W{}'.format(index)] = attendance_info_dict[key]
        else:
            ignore_time += 1

        if ignore_time == 5:
            break

        index += 1


def set_guest_info(sheet, guest_info_dict):
    num_guest = len(guest_info_dict)
    if num_guest == 0:
        return

    for i in range(10, 50):
        key = sheet['O{}'.format(i)].value
        if key != 'ゲスト':
            continue
        sheet['S{}'.format(i)] = '00:{:02d}:00'.format(int(num_guest))
        break


def get_all_info(base_url, mtgid):
    url = '{}/mtgDetail.php?mtgid={}'.format(base_url, mtgid)
    r = requests.get(url)
    bs_obj = BeautifulSoup(r.text, 'html.parser')

    main_contents = bs_obj.find(class_='contentDiv')
    meeting_info_dict = get_event_info(main_contents)
    guest_info_dict = get_guest_info(main_contents)
    assign_info_dict = get_assign_info(main_contents)

    url = '{}/mtgDetailGetMemberList.php?mtgid={}'.format(base_url, mtgid)
    r = requests.get(url, headers={'X-Requested-With': 'XMLHttpRequest'})
    main_contents = BeautifulSoup(r.text, 'html.parser')
    attendance_info_dict = get_attendance_info(main_contents)
    return meeting_info_dict, guest_info_dict, attendance_info_dict, assign_info_dict


def set_all_info(sheet, meeting_info_dict, guest_info_dict, attendance_info_dict, assign_info_dict):
    set_event_info(sheet, meeting_info_dict)
    set_assign_info(sheet, assign_info_dict)
    set_attendance_info(sheet, attendance_info_dict)
    set_guest_info(sheet, guest_info_dict)


def get_place_name(base_url):
    if 'urawa' in base_url:
        place_name = 'urawan'
    elif 'shimazuyamax' in base_url:
        place_name = 'shimazuyamax'
    else:
        place_name = None
    return place_name


def get_meeting_program(base_url, mtgid, output=True):
    place_name = get_place_name(base_url)
    if place_name is None:
        return None

    meeting_info_dict, guest_info_dict, attendance_info_dict, assign_info_dict = get_all_info(base_url, mtgid)

    wb = openpyxl.load_workbook('./templates/{}.xlsx'.format(place_name))
    sheet = wb['Sheet1']
    set_all_info(sheet, meeting_info_dict, guest_info_dict, attendance_info_dict, assign_info_dict)

    date = meeting_info_dict['開催日'].replace('/', '')
    if output:
        wb.save('./プログラム_{}TMC_{}.xlsx'.format(place_name, date))


if __name__ == '__main__':

    target_num = int(input('Please input number (0: 浦和，1: 島津山) << '))
    if target_num == 0:
        base_url = 'http://tmcsupport.php.xdomain.jp/urawan'
    else:
        base_url = 'http://tmcsupport.php.xdomain.jp/shimazuyamax/'

    print('Get number of meeting ... ', end='')
    mtgid = '40'
    # mtgid = '41'
    print('Success')
    print('Making timetable ... ', end='')
    get_meeting_program(base_url, mtgid)
    print('Success')
    print('\nComplete to make timetable!\n')


# coding=utf-8
import sys
from time import sleep

from get_meeting_program import get_meeting_program
from get_mtgid import get_latest_mtgid


def main(args=sys.argv):
    while True:
        target_num = int(input('番号を入力してください (浦和: 0 ，島津山: 1) << '))
        if target_num == 0:
            place_name = '浦和'
            base_url = 'http://tmcsupport.php.xdomain.jp/urawan'
            break
        elif target_num == 1:
            place_name = '島津山'
            base_url = 'http://tmcsupport.php.xdomain.jp/shimazuyamax/'
            break
        else:
            print('その番号は対応していません')

    try:
        print('\n{}TMCの例会のIDを取得中 ... '.format(place_name))
        mtgid = args[1] if len(args) >= 2 else get_latest_mtgid(base_url)
        print('成功')
        sleep(1.5)
        print('例会のプログラムを作成中 ... ')
        get_meeting_program(base_url, mtgid)
        print('成功')
        print('\nプログラムの作成が完了しました\n')
    except Exception as e:
        print('失敗')
        print('\nこのプログラムの製作者に連絡をしてください\n')
        print()
        raise e
    input('終了するには，何かキーを押してください ...')


if __name__ == "__main__":
    main()


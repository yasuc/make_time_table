import openpyxl as px
import sys
import pprint
import pickle
import os
import datetime
import argparse
from pathlib import Path

time_range = [
    ["8:50", "10:20"],
    ["10:30", "12:00"],
    ["13:10", "14:40"],
    ["14:50", "16:20"],
    ["13:10", "16:20"],
]

subjects = [
    [
        #    [ 曜日(1-5),
        #      科目名,
        #      時限(1-4 and 5(3+4)),
        #      場所,
        #      0:本科、1:専攻科],
        # 前期用
        [1, "情報セキュリティI", 2, "講1-2", 0],
        [1, "オブジェクト指向言語", 3, "講1-2", 0],
        [3, "卒業研究", 5, "OLab", 0],
        [4, "科学技術英語", 2, "OLab", 0],
        [4, "卒業研究", 5, "OLab", 0],
        [5, "コンピュータネットワークI", 1, "講2-2", 0],
        [5, "情報セキュリティ特論", 2, "NW演習室", 1],
        #    [4, "コンピュータネットワークII", 1, "講1-2", 0],
    ],
    [
        # 後期用
        [1, "情報セキュリティII", 2, "講1-2", 0],
        [1, "情報セキュリティII", 3, "講1-2", 0],
        [3, "卒業研究", 5, "OLab", 0],
        [3, "コンピュータネットワークI", 1, "講2-2", 0],
        [4, "OSとコンパイラI", 2, "講1-2", 0],
        [4, "卒業研究", 5, "OLab", 0],
    ],
]


def get_value_list(t_2d):
    return [[cell.value for cell in row] for row in t_2d]


def get_list_2d(sheet, start_row, end_row, start_col, end_col):
    return get_value_list(
        sheet.iter_rows(
            min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col
        )
    )


def excel_date(num):
    from datetime import datetime, timedelta

    return datetime(1899, 12, 30) + timedelta(days=num)


def old_pklfile_del(xlsx_file, pkl_file):
    if not os.path.isfile(pkl_file):
        return

    p_xlsx = Path(xlsx_file)
    p_pkl = Path(pkl_file)

    xlsx_update_time = datetime.datetime.fromtimestamp(p_xlsx.stat().st_mtime)
    pkl_update_time = datetime.datetime.fromtimestamp(p_pkl.stat().st_mtime)

    if xlsx_update_time > pkl_update_time:
        os.remove(pklfile)


all_2d = []

parser = argparse.ArgumentParser(description="時間割作成プログラム")

parser.add_argument("-x", "--xlsx", default="schedule.xlsx", help="スケジュールExcelファイル")
parser.add_argument("-s", "--start", default=1, type=int, help="開始回")
parser.add_argument("-e", "--end", default=15, type=int, help="終了回")
parser.add_argument("-t", "--term", default=1, type=int, help="前後期(1:前期, 2:後期)")

args = parser.parse_args()

xlsx_name = args.xlsx
start = args.start
end = args.end
term = args.term

if term == 1:
    # 前期
    c = 1
    mon = 5
elif term == 2:
    # 後期
    c = 86
    mon = 6

pklfile = "make_time_schedule_%i.pkl" % term

old_pklfile_del(xlsx_name, pklfile)

if not os.path.isfile(pklfile):
    print("Making pkl file.", file=sys.stderr)
    wb = px.load_workbook(xlsx_name, data_only=True)

    sheet = wb.active

    for i in range(mon):
        all_2d.append(get_list_2d(sheet, 5, 128, c, c + 16))
        c += 17
    with open(pklfile, "wb") as f:
        pickle.dump(all_2d, f)
else:
    with open(pklfile, "rb") as f:
        all_2d = pickle.load(f)

print("Subject,Start Date,Start Time,End Date,End Time")
for s in subjects[term - 1]:
    n = 1
    for l_2d in all_2d:
        for t in l_2d:
            if t[s[0] + 6 + s[-1] * 5] is not None:
                if n >= start and n <= end:
                    day = excel_date(t[0]).strftime("%Y/%m/%d")
                    csv = "講義:{:s}[{:s}]:{:d},{:s},{:s},{:s},{:s}".format(
                        s[1],
                        s[3],
                        n,
                        day,
                        time_range[s[2] - 1][0],
                        day,
                        time_range[s[2] - 1][1],
                    )
                    print(csv)
                n += 1

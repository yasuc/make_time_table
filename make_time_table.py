import openpyxl as px
import sys
import pprint
import pickle
import os
from pathlib import Path
import datetime

time_range = [
	["9:00", "10:30"],
	["10:40", "12:10"],
	["13:30", "15:00"],
	["15:10", "16:40"],
	["13:30", "16:40"],
]

subjects = [
#    [ 曜日,
#      科目名
#      時限(1-4 and 5(3+4)),
#      場所,
#      0:本科、1:専攻科],
# 前期用
    [1, "情報セキュリティ特論", 3, "NW演習室", 1],
    [1, "コンピュータネットワークII", 4, "講1-2", 0],
    [2, "オブジェクト指向言語", 3, "講義室1-2", 0],
    [3, "コンピュータネットワークI", 2, "講2-2", 0],
    [3, "卒業研究", 5, "OLab", 0],
    [4, "情報セキュリティI", 2, "NW演習室", 0],
    [4, "卒業研究", 5, "OLab", 0],
# 未定
#    [4, "科学技術英語", 2, "OLab", 0],

# 後期用
#    [1, "情報セキュリティII", "10:40", "12:10", "講義室1-1:2", 0],
#    [1, "情報セキュリティII", "14:00", "15:30", "講義室1-1:2", 0],
#    [2, "創造研究", "14:40", "16:10", "NW演習室"],
#    [4, "OSとコンパイラI", "10:40", "12:10", "講義室1-1:2", 0],
]

def get_value_list(t_2d):
    return([[cell.value for cell in row] for row in t_2d])


def get_list_2d(sheet, start_row, end_row, start_col, end_col):
    return get_value_list(sheet.iter_rows(min_row=start_row,
                                          max_row=end_row,
                                          min_col=start_col,
                                          max_col=end_col))

def excel_date(num):
    from datetime import datetime, timedelta
    return(datetime(1899, 12, 30) + timedelta(days=num))

def old_pklfile_del(xlsx_file, pkl_file):
    if not os.path.isfile(pkl_file):
        return

    p_xlsx = Path(xlsx_file)
    p_pkl = Path(pkl_file)

    xlsx_update_time = datetime.datetime.fromtimestamp(p_xlsx.stat().st_mtime)
    pkl_update_time = datetime.datetime.fromtimestamp(p_pkl.stat().st_mtime)

    if xlsx_update_time > pkl_update_time:
        os.remove(pklfile)


#前期
c = 1
mon = 5
#後期
#c = 86
#mon = 5

all_2d = []

xlsx_name = 'schedule.xlsx'
pklfile = "make_time_schedule.pkl"

args = sys.argv

if len(args) == 2:
    xlsx_name = args[1]

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
sotuken = False
for s in subjects:
    n = 1

#    if s[1] != "卒業研究":
#        n = 1
#    if s[1] == "卒業研究" and not sotuken:
#        n = 1
#        sotuken = True

    for l_2d in all_2d:
        for t in l_2d:
            if t[s[0] + 6 + s[-1]*5] is not None:
                day = excel_date(t[0]).strftime('%Y/%m/%d')
                csv = "講義:{:s}[{:s}]:{:d},{:s},{:s},{:s},{:s}".format(
                    s[1], s[3], n, day, time_range[s[2]-1][0], day, time_range[s[2]-1][1])
                print(csv)
                n += 1


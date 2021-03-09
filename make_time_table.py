import openpyxl as px
import sys
import pprint
import pickle
import os

subjects = [
#    [ 曜日,
#      科目名
#      開始時間,
#      終了時間,
#      場所,
#      0:本科、1:専攻科],
# 前期用
    [1, "情報セキュリティ特論", "13:00", "14:30", "NW演習室", 1],
    [1, "コンピュータネットワークII", "14:40", "16:10", "講1-2", 0],
    [2, "オブジェクト指向言語", "13:00", "14:30", "講義室1-2", 0],
    [3, "コンピュータネットワークI", "10:30", "12:00", "講2-2", 0],
    [3, "卒業研究", "13:00", "16:10", "OLab", 0],
    [4, "情報セキュリティI", "10:30", "12:00", "NW演習室", 0],
    [4, "卒業研究", "13:00", "16:10", "OLab", 0],
# 未定
#    [4, "科学技術英語", "8:50", "10:20", "OLab", 0],

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

#前期
c = 1
mon = 5
#後期
#c = 86
#mon = 5

all_2d = []
pklfile = "schedule.pkl"
xlsx_name = 'schedule.xlsx'

if not os.path.isfile(pklfile):
    args = sys.argv

    if len(args) == 2:
        xlsx_name = args[1]

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
                    s[1], s[4], n, day, s[2], day, s[3])
                print(csv)
                n += 1


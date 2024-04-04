import openpyxl as px
import sys
import re
import pickle
import os
from pathlib import Path
from datetime import datetime, timedelta

def get_value_list(t_2d):
    return([[cell.value for cell in row] for row in t_2d])


def get_list_2d(sheet, start_row, end_row, start_col, end_col):
    return get_value_list(sheet.iter_rows(min_row=start_row,
                                          max_row=end_row,
                                          min_col=start_col,
                                          max_col=end_col))

def excel_date(num):
    #return(datetime(1899, 12, 30) + timedelta(days=num))
    return(datetime.strptime(num, '%Y/%m/%d'))

def old_pklfile_del(xlsx_file, pkl_file):
    if not os.path.isfile(pkl_file):
        return

    p_xlsx = Path(xlsx_file)
    p_pkl = Path(pkl_file)

    xlsx_update_time = datetime.fromtimestamp(p_xlsx.stat().st_mtime)
    pkl_update_time = datetime.fromtimestamp(p_pkl.stat().st_mtime)

    if xlsx_update_time > pkl_update_time:
        os.remove(pklfile)

all_2d = []
c = 1
mon = 12

xlsx_name = 'schedule.xlsx'
pklfile = "schedule.pkl"

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


print("Subject,Start Date,All Day Event")
day = ""
for l_2d in all_2d:
    for t in l_2d:
#        print(f"t[0]={t[0]}")
        if t[0] is not None:
            #day = excel_date(t[0]).strftime('%Y/%m/%d')
            day = t[0].strftime('%Y/%m/%d')
        for j in range(2, 7):
            if t[j] is not None:
                subj = re.sub('※.*', '', t[j])
                subj = re.sub('[ 　]+', '', subj)
                if subj != "":
                    csv = "{:s},{:s},TRUE".format(subj, day)
                    print(csv)

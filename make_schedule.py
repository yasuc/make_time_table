import openpyxl as px
import sys
import re
import pickle
import os


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

pklfile = "schedule.pkl"
xlsx_name = 'schedule.xlsx'

all_2d = []
c = 1

if not os.path.isfile(pklfile):
    args = sys.argv

    if len(args) == 2:
        xlsx_name = args[1]

    wb = px.load_workbook(xlsx_name, data_only=True)

    sheet = wb.active

    for i in range(12):
        all_2d.append(get_list_2d(sheet, 5, 128, c, c + 16))
        c += 17
    with open(pklfile, "wb") as f:
        pickle.dump(all_2d, f)
else:
    with open(pklfile, "rb") as f:
        all_2d = pickle.load(f)

print("Subject,Start Date,All Day Event")
for l_2d in all_2d:
    for t in l_2d:
        if t[0] is not None:
            day = excel_date(t[0]).strftime('%Y/%m/%d')
            for j in range(2, 7):
                if t[j] is not None:
                    csv = "{:s},{:s},TRUE".format(re.sub('※.*', '', t[j]), day)
                    print(csv)
        else:
            for j in range(2, 7):
                if t[j] is not None:
                    csv = "{:s},{:s},TRUE".format(re.sub('※.*', '', t[j]), day)
                    print(csv)

import openpyxl as px
import sys
import re
import pickle
import os
from pathlib import Path
from datetime import datetime, timedelta

def get_value_list(t_2d):
    return [[cell.value for cell in row] for row in t_2d]

def get_list_2d(sheet, start_row, end_row, start_col, end_col):
    return get_value_list(sheet.iter_rows(min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col))

def excel_date(num):
    return datetime.strptime(num, '%Y/%m/%d')

def update_needed(xlsx_file, pkl_file):
    if not os.path.isfile(pkl_file):
        return True

    xlsx_update_time = datetime.fromtimestamp(Path(xlsx_file).stat().st_mtime)
    pkl_update_time = datetime.fromtimestamp(Path(pkl_file).stat().st_mtime)

    return xlsx_update_time > pkl_update_time

def remove_file_if_exists(file_path):
    if os.path.isfile(file_path):
        os.remove(file_path)

def process_schedule(xlsx_name, pklfile):
    if update_needed(xlsx_name, pklfile):
        remove_file_if_exists(pklfile)
        print("Making pkl file.", file=sys.stderr)
        wb = px.load_workbook(xlsx_name, data_only=True)
        sheet = wb.active
        all_2d = [get_list_2d(sheet, 5, 128, c, c + 16) for c in range(1, 193, 17)]
        with open(pklfile, "wb") as f:
            pickle.dump(all_2d, f)
    else:
        with open(pklfile, "rb") as f:
            all_2d = pickle.load(f)
    return all_2d

def print_schedule(all_2d):
    print("Subject,Start Date,All Day Event")
    for l_2d in all_2d:
        day = ""
        for t in l_2d:
            if t[0] is not None:
                day = t[0].strftime('%Y/%m/%d')
            for j in range(2, 7):
                if t[j] is not None:
                    subj = re.sub('※.*', '', t[j])
                    subj = re.sub('[ 　]+', '', subj)
                    if subj != "":
                        csv = f"{subj},{day},TRUE"
                        print(csv)

def main():
    xlsx_name = 'schedule.xlsx'
    pklfile = "schedule.pkl"
    args = sys.argv[1:]
    if args:
        xlsx_name = args[0]
        if len(args) > 1:
            pklfile = args[1]

    all_2d = process_schedule(xlsx_name, pklfile)
    print_schedule(all_2d)

if __name__ == "__main__":
    main()
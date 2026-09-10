import openpyxl as px
import sys
import pickle
import os
import datetime
import argparse
import json
from pathlib import Path

time_range = [
    ["8:50", "10:20"],
    ["10:30", "12:00"],
    ["13:00", "14:30"],
    ["14:40", "16:10"],
    ["13:00", "16:10"],
]

def get_subjects(term, json_file="subjects.json"):
    with open(json_file, "r", encoding="utf-8") as f:
        subjects_data = json.load(f)
    if term == 1:
        return subjects_data["term1"]
    else:
        return subjects_data["term2"]


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

    return num


def old_pklfile_del(xlsx_file, pkl_file):
    if not os.path.isfile(pkl_file):
        return

    p_xlsx = Path(xlsx_file)
    p_pkl = Path(pkl_file)

    xlsx_update_time = datetime.datetime.fromtimestamp(p_xlsx.stat().st_mtime)
    pkl_update_time = datetime.datetime.fromtimestamp(p_pkl.stat().st_mtime)

    if xlsx_update_time > pkl_update_time:
        os.remove(pklfile)


def main():
    all_2d = []

    parser = argparse.ArgumentParser(description="時間割作成プログラム")

    parser.add_argument(
        "-x", "--xlsx", default="schedule.xlsx", help="スケジュールExcelファイル"
    )
    parser.add_argument("-s", "--start", default=1, type=int, help="開始回")
    parser.add_argument("-e", "--end", default=15, type=int, help="終了回")
    parser.add_argument(
        "-t", "--term", default=1, type=int, help="前後期(1:前期, 2:後期)"
    )

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

    subjects = get_subjects(term)

    print("Subject,Start Date,Start Time,End Date,End Time")
    for s in subjects:
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


if __name__ == "__main__":
    main()

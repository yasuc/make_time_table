import sys
import datetime
import argparse
import json
import re
from pathlib import Path

try:
    import pymupdf
except ImportError as e:
    raise SystemExit(
        "PyMuPDF が必要です。pip install pymupdf でインストールしてください。"
    ) from e


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
    return subjects_data["term2"]


def _center_x(word):
    return (word[0] + word[2]) / 2.0


def _center_y(word):
    return (word[1] + word[3]) / 2.0


def _find_month_headers(page):
    """ページ内の「n月」ヘッダを左から順に (x中心, 月) で返す。"""
    headers = []
    for word in page.get_text("words"):
        text = word[4]
        if re.fullmatch(r"(?:[1-9]|1[0-2])月", text) and 115 <= word[1] <= 140:
            headers.append((_center_x(word), int(text[:-1])))

    # x座標で重複を除去
    unique = {}
    for x, month in headers:
        unique[round(x, 1)] = month
    return sorted(unique.items())


def _parse_month(page, header_x, month, year):
    """
    1か月分を、旧Excel版が扱っていた17列の2次元配列へ変換する。

    [0]      : 日付
    [1:6]    : 行事領域（未使用）
    [6:11]   : 本科 月～金
    [11:16]  : 専攻科 月～金
    [16]     : 末尾の未使用領域
    """
    words = page.get_text("words")

    # 月名は月ブロックの中央にある。PDFの帳票レイアウトから左端を求める。
    block_left = header_x - 54.6

    # 曜日ヘッダは「月 火 水 木 金 月 火 水 木 金」の10個。
    # 月ブロック内のx範囲だけを対象にする。
    weekday_words = []
    for word in words:
        if not (132 <= word[1] <= 142):
            continue
        if word[4] not in {"月", "火", "水", "木", "金"}:
            continue
        x = _center_x(word)
        if block_left + 105 <= x <= block_left + 175:
            weekday_words.append((x, word[4]))

    weekday_words.sort()
    if len(weekday_words) < 10:
        raise ValueError(
            f"{month}月の曜日ヘッダを10列取得できませんでした。"
            f"（取得数={len(weekday_words)}）"
        )
    weekday_x = [x for x, _ in weekday_words[:10]]

    # 日付は月ブロック左端付近にある。
    date_words = []
    for word in words:
        if not word[4].isdigit():
            continue
        x = _center_x(word)
        y = _center_y(word)
        if block_left - 2 <= x <= block_left + 10 and 140 <= y <= 760:
            day = int(word[4])
            if 1 <= day <= 31:
                date_words.append((day, y))

    # 同じ日付が重複抽出された場合に備える。
    dates = {}
    for day, y in sorted(date_words, key=lambda item: item[1]):
        dates.setdefault(day, y)

    rows = []
    for day, y in sorted(dates.items()):
        try:
            date = datetime.datetime(year, month, day)
        except ValueError:
            continue

        row = [None] * 17
        row[0] = date

        # PDFでは、授業回数が記載されているセルだけを授業日として扱う。
        # 数値そのものは旧プログラムでは判定に使わないが、そのまま保存する。
        for dept in range(2):
            for wd in range(5):
                x = weekday_x[wd + dept * 5]
                target_col = 6 + wd + dept * 5

                for word in words:
                    if not word[4].isdigit():
                        continue
                    wx = _center_x(word)
                    wy = _center_y(word)
                    if abs(wx - x) <= 2.2 and abs(wy - y) <= 4.0:
                        value = int(word[4])
                        if 0 <= value <= 99:
                            row[target_col] = value
                        break

        rows.append(row)

    return rows


def load_schedule_pdf(pdf_file):
    """年間行事予定表PDFを読み込み、旧Excel版と同じ形式へ変換する。"""
    pdf_path = Path(pdf_file)
    if not pdf_path.is_file():
        raise FileNotFoundError(f"PDFファイルがありません: {pdf_file}")

    with pymupdf.open(pdf_path) as doc:
        if len(doc) < 4:
            raise ValueError(
                f"年間行事予定表は4ページを想定していますが、{len(doc)}ページです。"
            )

        all_2d = []
        page_months = ([4, 5, 6], [7, 8, 9], [10, 11, 12], [1, 2, 3])

        for page_index, expected_months in enumerate(page_months):
            page = doc[page_index]
            headers = _find_month_headers(page)
            if len(headers) != 3:
                raise ValueError(
                    f"{page_index + 1}ページの月ヘッダを3個取得できませんでした: {headers}"
                )

            for (header_x, month), expected_month in zip(headers, expected_months):
                if month != expected_month:
                    raise ValueError(
                        f"{page_index + 1}ページの月順が想定と異なります: "
                        f"{month}月 / 期待値 {expected_month}月"
                    )
                year = 2026 if month >= 4 else 2027
                all_2d.append(_parse_month(page, header_x, month, year))

        return all_2d


def main():
    parser = argparse.ArgumentParser(
        description="年間行事予定表PDFから時間割用データを作成するプログラム"
    )
    parser.add_argument(
        "-p",
        "--pdf",
        default="r8schedule_20260624_1.pdf",
        help="年間行事予定表PDFファイル",
    )
    parser.add_argument("-s", "--start", default=1, type=int, help="開始回")
    parser.add_argument("-e", "--end", default=15, type=int, help="終了回")
    parser.add_argument(
        "-t",
        "--term",
        default=1,
        type=int,
        choices=(1, 2),
        help="前後期(1:前期, 2:後期)",
    )
    parser.add_argument(
        "-j",
        "--subjects",
        default="subjects.json",
        help="科目定義JSONファイル",
    )
    args = parser.parse_args()

    if args.start < 1 or args.end < args.start:
        parser.error("--start と --end の指定が不正です。")

    print("Reading schedule PDF.", file=sys.stderr)
    parsed = load_schedule_pdf(args.pdf)

    # 旧Excel版と同じ対象月：前期=4～8月、後期=9～2月。
    if args.term == 1:
        all_2d = parsed[:5]
    else:
        all_2d = parsed[5:11]

    subjects = get_subjects(args.term, args.subjects)

    print("Subject,Start Date,Start Time,End Date,End Time")
    for s in subjects:
        n = 1
        col = (s[0] - 1) + 6 + s[-1] * 5

        for month_rows in all_2d:
            for row in month_rows:
                if row[col] is not None:
                    if args.start <= n <= args.end:
                        day = row[0].strftime("%Y/%m/%d")
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

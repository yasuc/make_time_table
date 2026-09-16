#!/usr/bin/env python3

"""
【このプログラムの目的】
    年間行事予定表（PDF または Excel）から授業が行われた日を読み取り、
    「科目名・何回目・開始日時・終了日時」を CSV 形式で標準出力へ流すツールです。

    PDF 版: PyMuPDF を使って座標ベースで数字を抽出
    Excel 版: openpyxl で直接セル値を読む

【使い方】
    python3 make_time_table.py [オプション] > output.csv

【入力ソース】
    --url  : Web ページから PDF を取得（デフォルト: 沖縄高専）
    --pdf  : ローカルの PDF ファイル
    --xlsx : ローカルの Excel ファイル
    いずれも指定しない場合は沖縄高専の年間行事予定表 URL を使用
"""

import argparse
import sys

from schedule_common import (
    DEFAULT_PAGE_URL,
    TIME_SLOTS,
    Subject,
    current_academic_year,
    infer_academic_year,
    load_subjects,
    parse_schedule_excel,
    parse_schedule_pdf,
    resolve_pdf_source,
)


def generate_csv(all_months, subjects, start: int, end: int) -> None:
    """科目ごとに授業日をたどり、CSV 形式で標準出力へ書き出す。

    CSV の 1 行は
        講義:科目名[教室]:n回目, 日付, 開始時刻, 日付, 終了時刻
    という 5 列カンマ区切りです（2 列目と 4 列目は同じ日付）。

    各科目の列（target_column）には授業の回数順に数字が入っているので、
    セルが空でない日を「その科目の n 回目の授業日」として数えていき、
    開始回数（start）から終了回数（end）までの行だけを出力します。
    """
    print("Subject,Start Date,Start Time,End Date,End Time")
    for subject in subjects:
        session = 1  # これは何回目の授業かを表すカウンタ
        target_column = subject.target_column()

        for month_rows in all_months:
            for row in month_rows:
                if row[target_column] is None or row[0] is None:
                    continue  # この曜日に授業がなかった日は読み飛ばす

                if start <= session <= end:
                    day = row[0].strftime("%Y/%m/%d")
                    begin_time, end_time = TIME_SLOTS[subject.period - 1]
                    csv = (
                        f"講義:{subject.name}[{subject.room}]:{session},"
                        f"{day},{begin_time},{day},{end_time}"
                    )
                    print(csv)
                session += 1


def main() -> None:
    parser = argparse.ArgumentParser(
        description="年間行事予定表（PDF/Excel）から時間割用 CSV データを作成するプログラム"
    )
    source_group = parser.add_mutually_exclusive_group()
    source_group.add_argument(
        "-x",
        "--xlsx",
        help="ローカルの年間行事予定表 Excel ファイル (.xlsx)",
    )
    source_group.add_argument(
        "-p",
        "--pdf",
        help="ローカルの年間行事予定表 PDF ファイル",
    )
    source_group.add_argument(
        "-u",
        "--url",
        nargs="?",
        const=DEFAULT_PAGE_URL,
        metavar="URL",
        help=(
            "最初にリンクされた PDF を取得する Web ページの URL（PDF の直URLも可）。"
            f"URL 省略時: {DEFAULT_PAGE_URL}"
        ),
    )
    parser.add_argument("-s", "--start", default=1, type=int, help="開始回数")
    parser.add_argument("-e", "--end", default=15, type=int, help="終了回数")
    parser.add_argument(
        "-t",
        "--term",
        default=1,
        type=int,
        choices=(1, 2),
        help="前後期 (1: 前期, 2: 後期)",
    )
    parser.add_argument(
        "-j",
        "--subjects",
        default="subjects.json",
        help="科目定義 JSON ファイル",
    )
    parser.add_argument(
        "-y",
        "--year",
        default=None,
        type=int,
        help=(
            "学年度の開始年（例: 2025 年度の PDF なら 2025）。"
            "省略時は PDF ファイル名から推測します（r7 → 2025 など）。"
        ),
    )
    args = parser.parse_args()

    if args.start < 1 or args.end < args.start:
        parser.error("--start と --end の指定が不正です。")

    try:
        if args.xlsx:
            all_months = parse_schedule_excel(args.xlsx, args.term)
        elif args.pdf:
            pdf_source, source_name = resolve_pdf_source(args.pdf, None)
            print(f"Reading schedule PDF: {source_name}", file=sys.stderr)
            year = args.year or infer_academic_year(source_name) or current_academic_year()
            all_months = parse_schedule_pdf(pdf_source, year)
        else:
            # --url または引数なし → デフォルト URL
            url = args.url or DEFAULT_PAGE_URL
            pdf_source, source_name = resolve_pdf_source(None, url)
            print(f"Reading schedule PDF: {source_name}", file=sys.stderr)
            year = args.year or infer_academic_year(source_name) or current_academic_year()
            all_months = parse_schedule_pdf(pdf_source, year)
    except (FileNotFoundError, RuntimeError, ValueError) as e:
        parser.error(str(e))

    # Excel 版は既に前期/後期で絞り込み済み、PDF 版は全年を返すのでスライス
    if args.xlsx:
        term_months = all_months
    else:
        if args.term == 1:
            term_months = all_months[:5]  # 4 月 ～ 8 月
        else:
            term_months = all_months[5:11]  # 9 月 ～ 2 月

    subjects = load_subjects(args.subjects, args.term)
    generate_csv(term_months, subjects, args.start, args.end)


if __name__ == "__main__":
    main()
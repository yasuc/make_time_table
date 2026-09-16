#!/usr/bin/env python3

"""
【このプログラムの目的】
    年間行事予定表（PDF または Excel）から行事予定を読み取り、
    Google Calendar などへ取り込める形式
    「件名, 開始日, 終日イベント(TRUE)」
    で CSV 形式を標準出力へ流すツールです。

【使い方】
    python3 make_schedule.py [オプション] > events.csv

【入力ソース】
    --url  : Web ページから PDF を取得（デフォルト: 沖縄高専）
    --pdf  : ローカルの PDF ファイル
    --xlsx : ローカルの Excel ファイル
    いずれも指定しない場合は沖縄高専の年間行事予定表 URL を使用
"""

import argparse
import re
import sys

from schedule_common import (
    DEFAULT_PAGE_URL,
    EVENT_COL_START,
    EVENT_COL_END,
    current_academic_year,
    infer_academic_year,
    parse_schedule_excel_all,
    parse_schedule_pdf,
    resolve_pdf_source,
)


def generate_event_csv(all_months) -> None:
    """月ごとの表から行事を読み取り、CSV で標準出力へ流す。

    出力形式は
        件名, 開始日, TRUE
    の 3 列カンマ区切りです（3 列目は「終日イベント」フラグ）。

    各行の [0] で日付が確定していない間は、直前の日付を引き継ぎます。
    行事のテキストからは「※」から始まる注釈と空白を取り除きます。
    """
    print("Subject,Start Date,All Day Event")
    for month_rows in all_months:
        day = ""
        for row in month_rows:
            if row[0] is not None:
                day = row[0].strftime("%Y/%m/%d")
            for col in range(EVENT_COL_START, EVENT_COL_END):
                if row[col] is None:
                    continue
                subj = re.sub("※.*", "", row[col])
                subj = re.sub("[ 　]+", "", subj)
                if subj != "":
                    print(f"{subj},{day},TRUE")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="年間行事予定表（PDF/Excel）から行事予定の CSV データを作成するプログラム"
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

    try:
        if args.xlsx:
            all_months = parse_schedule_excel_all(args.xlsx)
        else:
            if args.pdf:
                pdf_source, source_name = resolve_pdf_source(args.pdf, None)
            else:
                url = args.url or DEFAULT_PAGE_URL
                pdf_source, source_name = resolve_pdf_source(None, url)
            print(f"Reading schedule PDF: {source_name}", file=sys.stderr)
            year = args.year or infer_academic_year(source_name) or current_academic_year()
            all_months = parse_schedule_pdf(pdf_source, year)
    except (FileNotFoundError, RuntimeError, ValueError) as e:
        parser.error(str(e))

    generate_event_csv(all_months)


if __name__ == "__main__":
    main()
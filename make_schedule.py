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
import csv
import re
import sys
from collections.abc import Iterator, Sequence
from typing import TextIO

from schedule_common import (
    DEFAULT_PAGE_URL,
    EVENT_COL_END,
    EVENT_COL_START,
    current_academic_year,
    infer_academic_year,
    parse_schedule_excel_all,
    parse_schedule_pdf,
    resolve_pdf_source,
)


def iter_event_rows(all_months: list[list[list]]) -> Iterator[tuple[str, str, str]]:
    day = ""
    for month_rows in all_months:
        for row in month_rows:
            if row[0] is not None:
                day = row[0].strftime("%Y/%m/%d")
            for col in range(EVENT_COL_START, EVENT_COL_END):
                if row[col] is None:
                    continue
                subj = re.sub("※.*", "", row[col])
                subj = re.sub("[ 　]+", "", subj)
                if subj != "":
                    yield subj, day, "TRUE"


def generate_event_csv(
    all_months: list[list[list]], output: TextIO | None = None
) -> None:
    """月ごとの表から行事を読み取り、CSV で標準出力へ流す。

    出力形式は
        件名, 開始日, TRUE
    の 3 列カンマ区切りです（3 列目は「終日イベント」フラグ）。

    各行の [0] で日付が確定していない間は、直前の日付を引き継ぎます。
    行事のテキストからは「※」から始まる注釈と空白を取り除きます。
    """
    writer = csv.writer(
        output if output is not None else sys.stdout, lineterminator="\n"
    )
    writer.writerow(("Subject", "Start Date", "All Day Event"))
    writer.writerows(iter_event_rows(all_months))


def build_parser() -> argparse.ArgumentParser:
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
    return parser


def load_all_months(
    xlsx: str | None, pdf: str | None, url: str | None, year: int | None
) -> list[list[list]]:
    if xlsx:
        return parse_schedule_excel_all(xlsx)

    page_url = None if pdf else url or DEFAULT_PAGE_URL
    pdf_source, source_name = resolve_pdf_source(pdf, page_url)
    print(f"Reading schedule PDF: {source_name}", file=sys.stderr)
    academic_year = year or infer_academic_year(source_name) or current_academic_year()
    return parse_schedule_pdf(pdf_source, academic_year)


def main(argv: Sequence[str] | None = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        all_months = load_all_months(args.xlsx, args.pdf, args.url, args.year)
    except (OSError, RuntimeError, ValueError) as e:
        parser.error(str(e))

    generate_event_csv(all_months)


if __name__ == "__main__":
    main()

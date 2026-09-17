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
import csv
import sys
from collections.abc import Iterator, Sequence
from itertools import islice
from typing import TextIO

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


def iter_time_table_rows(
    all_months: list[list[list]], subjects: Sequence[Subject], start: int, end: int
) -> Iterator[tuple[str, str, str, str, str]]:
    for subject in subjects:
        target_column = subject.target_column()
        dates = (
            row[0]
            for month_rows in all_months
            for row in month_rows
            if row[0] is not None and row[target_column] is not None
        )
        begin_time, end_time = TIME_SLOTS[subject.period - 1]
        for session, date in islice(enumerate(dates, start=1), start - 1, end):
            day = date.strftime("%Y/%m/%d")
            yield (
                f"講義:{subject.name}[{subject.room}]:{session}",
                day,
                begin_time,
                day,
                end_time,
            )


def generate_csv(
    all_months: list[list[list]],
    subjects: Sequence[Subject],
    start: int,
    end: int,
    output: TextIO | None = None,
) -> None:
    """科目ごとに授業日をたどり、CSV 形式で標準出力へ書き出す。

    CSV の 1 行は
        講義:科目名[教室]:n回目, 日付, 開始時刻, 日付, 終了時刻
    という 5 列カンマ区切りです（2 列目と 4 列目は同じ日付）。

    各科目の列（target_column）には授業の回数順に数字が入っているので、
    セルが空でない日を「その科目の n 回目の授業日」として数えていき、
    開始回数（start）から終了回数（end）までの行だけを出力します。
    """
    writer = csv.writer(
        output if output is not None else sys.stdout, lineterminator="\n"
    )
    writer.writerow(("Subject", "Start Date", "Start Time", "End Date", "End Time"))
    writer.writerows(iter_time_table_rows(all_months, subjects, start, end))


def load_term_months(
    xlsx: str | None, pdf: str | None, url: str | None, term: int, year: int | None
) -> list[list[list]]:
    if xlsx:
        return parse_schedule_excel(xlsx, term)

    page_url = None if pdf else url or DEFAULT_PAGE_URL
    pdf_source, source_name = resolve_pdf_source(pdf, page_url)
    print(f"Reading schedule PDF: {source_name}", file=sys.stderr)
    academic_year = year or infer_academic_year(source_name) or current_academic_year()
    all_months = parse_schedule_pdf(pdf_source, academic_year)
    return all_months[:5] if term == 1 else all_months[5:11]


def build_parser() -> argparse.ArgumentParser:
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
    return parser


def main(argv: Sequence[str] | None = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.start < 1 or args.end < args.start:
        parser.error("--start と --end の指定が不正です。")

    try:
        term_months = load_term_months(
            args.xlsx, args.pdf, args.url, args.term, args.year
        )
        subjects = load_subjects(args.subjects, args.term)
    except (OSError, RuntimeError, ValueError) as e:
        parser.error(str(e))

    generate_csv(term_months, subjects, args.start, args.end)


if __name__ == "__main__":
    main()

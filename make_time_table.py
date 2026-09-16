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
import datetime
import json
import os
import pickle
import re
import sys
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path
from typing import Optional
from urllib.error import HTTPError, URLError
from urllib.parse import urljoin, urlparse
from urllib.request import Request, urlopen

try:
    import pymupdf
except ImportError as e:
    pymupdf = None  # PDF 未対応時は後でエラー

try:
    import openpyxl
except ImportError as e:
    openpyxl = None  # Excel 未対応時は後でエラー

# ==========================================================================
# 設定値
# ==========================================================================

TIME_SLOTS = [
    ("8:50", "10:20"),   # 1 限
    ("10:30", "12:00"),  # 2 限
    ("13:00", "14:30"),  # 3 限
    ("14:40", "16:10"),  # 4 限
    ("13:00", "16:10"),  # 特殊な時限（3 限+4 限分の長さ）
]

CLASS_SESSION_MAX = 99

# PDF レイアウトに関する定数
MONTH_HEADER_Y_RANGE = (115, 140)
WEEKDAY_HEADER_Y_RANGE = (132, 142)
WEEKDAY_LABEL_X_MIN = 105.0
WEEKDAY_LABEL_X_MAX = 175.0
DATE_LABEL_X_MIN = -2.0
DATE_LABEL_X_MAX = 10.0
DATE_LABEL_Y_MIN = 140
DATE_LABEL_Y_MAX = 760
MONTH_TO_BLOCK_OFFSET = 54.6
CELL_X_TOLERANCE = 2.2
CELL_Y_TOLERANCE = 4.0
WEEKDAY_CHARS = "月火水木金"
PAGE_MONTHS = (
    (4, 5, 6),
    (7, 8, 9),
    (10, 11, 12),
    (1, 2, 3),
)
FIRST_YEAR_MONTH = 4
TOTAL_COLUMNS = 17
DATA_COL_OFFSET = 6
NUM_WEEKDAYS = 5

# URL 設定
DEFAULT_PAGE_URL = "https://www.okinawa-ct.ac.jp/campus_life/class/annualev/"
HTTP_TIMEOUT_SECONDS = 30
MAX_DOWNLOAD_BYTES = 50 * 1024 * 1024
USER_AGENT = "make_time_table/1.0"


# ==========================================================================
# 共通のデータクラス
# ==========================================================================

@dataclass
class Subject:
    group: int
    name: str
    period: int
    room: str
    course: int

    def target_column(self) -> int:
        return (self.group - 1) + DATA_COL_OFFSET + self.course * NUM_WEEKDAYS


def load_subjects(json_file: str, term: int) -> list[Subject]:
    with open(json_file, "r", encoding="utf-8") as f:
        data = json.load(f)
    key = "term1" if term == 1 else "term2"
    return [Subject(*row) for row in data[key]]


# ==========================================================================
# 共通の CSV 出力
# ==========================================================================

def generate_csv(all_months, subjects, start: int, end: int) -> None:
    print("Subject,Start Date,Start Time,End Date,End Time")
    for subject in subjects:
        session = 1
        target_column = subject.target_column()
        for month_rows in all_months:
            for row in month_rows:
                if row[target_column] is None or row[0] is None:
                    continue
                if start <= session <= end:
                    day = row[0].strftime("%Y/%m/%d")
                    begin_time, end_time = TIME_SLOTS[subject.period - 1]
                    csv = (
                        f"講義:{subject.name}[{subject.room}]:{session},"
                        f"{day},{begin_time},{day},{end_time}"
                    )
                    print(csv)
                session += 1


# ==========================================================================
# Excel 読み込み
# ==========================================================================

EXCEL_START_ROW = 5
EXCEL_END_ROW = 128
EXCEL_BLOCK_COLS = 17


def old_pklfile_del(xlsx_file: str, pkl_file: str) -> None:
    if not os.path.isfile(pkl_file):
        return
    p_xlsx = Path(xlsx_file)
    p_pkl = Path(pkl_file)
    if datetime.datetime.fromtimestamp(p_xlsx.stat().st_mtime) > \
       datetime.datetime.fromtimestamp(p_pkl.stat().st_mtime):
        os.remove(pkl_file)


def _normalize_excel_row(raw_row: list) -> list:
    """Excel の 17 列レイアウトを PDF 版と同じ 17 列レイアウトに揃える。

    Excel: [日付, 曜日, 行事, 未使用×4, 本科 月〜金, 専攻科 月〜金]
    PDF  : [日付, 未使用×5, 本科 月〜金, 専攻科 月〜金, 未使用]
    """
    row: list = [None] * TOTAL_COLUMNS
    row[0] = raw_row[0]  # 日付
    row[6:11] = raw_row[7:12]  # 本科 月〜金
    row[11:16] = raw_row[12:17]  # 専攻科 月〜金
    return row


def parse_schedule_excel(xlsx_file: str, term: int, pkl_cache: Optional[str] = None) -> list[list[list]]:
    """Excel ファイルから行事予定表を読み、全月の表を返す。

    返り値は PDF 版と同じ形式: [月ごとの表]。月の並びは 4 月 → 3 月です。
    """
    if openpyxl is None:
        raise SystemExit(
            "openpyxl が必要です。pip install openpyxl でインストールしてください。"
        )

    if pkl_cache is None:
        pkl_cache = f"make_time_schedule_{term}_pdflayout.pkl"

    old_pklfile_del(xlsx_file, pkl_cache)

    if os.path.isfile(pkl_cache):
        print(f"Reading cached pkl: {pkl_cache}", file=sys.stderr)
        with open(pkl_cache, "rb") as f:
            return pickle.load(f)

    print(f"Reading Excel: {xlsx_file}", file=sys.stderr)
    wb = openpyxl.load_workbook(xlsx_file, data_only=True)
    sheet = wb.active

    if term == 1:
        start_col = 1
        num_months = 5
    else:
        start_col = 86
        num_months = 6

    all_months = []
    c = start_col
    for _ in range(num_months):
        month_data = []
        for row in sheet.iter_rows(
            min_row=EXCEL_START_ROW, max_row=EXCEL_END_ROW,
            min_col=c, max_col=c + EXCEL_BLOCK_COLS - 1,
        ):
            month_data.append(_normalize_excel_row([cell.value for cell in row]))
        all_months.append(month_data)
        c += EXCEL_BLOCK_COLS

    with open(pkl_cache, "wb") as f:
        pickle.dump(all_months, f)

    return all_months


# ==========================================================================
# PDF 読み込み
# ==========================================================================

def word_center_x(word) -> float:
    return (word[0] + word[2]) / 2.0


def word_center_y(word) -> float:
    return (word[1] + word[3]) / 2.0


def find_month_headers(page) -> list[tuple[float, int]]:
    headers = []
    for word in page.get_text("words"):
        text = word[4]
        if (
            re.fullmatch(r"(?:[1-9]|1[0-2])月", text)
            and MONTH_HEADER_Y_RANGE[0] <= word[1] <= MONTH_HEADER_Y_RANGE[1]
        ):
            headers.append((word_center_x(word), int(text[:-1])))
    unique = {round(x, 1): month for x, month in headers}
    return sorted(unique.items())


def find_weekday_columns(page, block_left: float) -> list[float]:
    candidates = []
    for word in page.get_text("words"):
        if not (WEEKDAY_HEADER_Y_RANGE[0] <= word[1] <= WEEKDAY_HEADER_Y_RANGE[1]):
            continue
        if word[4] not in set(WEEKDAY_CHARS):
            continue
        x = word_center_x(word)
        if block_left + WEEKDAY_LABEL_X_MIN <= x <= block_left + WEEKDAY_LABEL_X_MAX:
            candidates.append((x, word[4]))
    candidates.sort()
    if len(candidates) < 10:
        raise ValueError(
            f"曜日ヘッダを 10 列取得できませんでした。（取得数={len(candidates)}）"
        )
    return [x for x, _ in candidates[:10]]


def find_date_positions(page, block_left: float) -> list[tuple[int, float]]:
    positions = []
    for word in page.get_text("words"):
        if not word[4].isdigit():
            continue
        x = word_center_x(word)
        y = word_center_y(word)
        if not (
            block_left + DATE_LABEL_X_MIN <= x <= block_left + DATE_LABEL_X_MAX
            and DATE_LABEL_Y_MIN <= y <= DATE_LABEL_Y_MAX
        ):
            continue
        day = int(word[4])
        if 1 <= day <= 31:
            positions.append((day, y))
    unique = {}
    for day, y in sorted(positions, key=lambda item: item[1]):
        unique.setdefault(day, y)
    return sorted(unique.items())


def get_cell_value(page, x: float, y: float, words) -> Optional[int]:
    for word in words:
        if not word[4].isdigit():
            continue
        wx = word_center_x(word)
        wy = word_center_y(word)
        if abs(wx - x) <= CELL_X_TOLERANCE and abs(wy - y) <= CELL_Y_TOLERANCE:
            value = int(word[4])
            return value if 0 <= value <= CLASS_SESSION_MAX else None
    return None


def parse_month(page, month_header_x: float, month: int, year: int) -> list[list]:
    words = page.get_text("words")
    block_left = month_header_x - MONTH_TO_BLOCK_OFFSET
    weekday_x = find_weekday_columns(page, block_left)
    date_positions = find_date_positions(page, block_left)

    rows = []
    for day, y in date_positions:
        try:
            date = datetime.datetime(year, month, day)
        except ValueError:
            continue
        row: list = [None] * TOTAL_COLUMNS
        row[0] = date
        for course in range(2):
            for weekday in range(NUM_WEEKDAYS):
                cell_x = weekday_x[weekday + course * NUM_WEEKDAYS]
                target_column = DATA_COL_OFFSET + weekday + course * NUM_WEEKDAYS
                row[target_column] = get_cell_value(page, cell_x, y, words)
        rows.append(row)
    return rows


def parse_schedule_pdf(pdf_source, academic_year: int) -> list[list[list]]:
    if pymupdf is None:
        raise SystemExit(
            "PyMuPDF が必要です。pip install pymupdf でインストールしてください。"
        )

    all_months = []
    if isinstance(pdf_source, bytes):
        doc_context = pymupdf.open(stream=pdf_source, filetype="pdf")
    else:
        doc_context = pymupdf.open(pdf_source)

    with doc_context as doc:
        if len(doc) < 4:
            raise ValueError(
                f"年間行事予定表は 4 ページを想定していますが、{len(doc)} ページです。"
            )
        for page_index, expected_months in enumerate(PAGE_MONTHS):
            page = doc[page_index]
            headers = find_month_headers(page)
            if len(headers) != 3:
                raise ValueError(
                    f"{page_index + 1} ページの月ヘッダを 3 個取得できませんでした: "
                    f"{headers}"
                )
            for (header_x, month), expected_month in zip(headers, expected_months):
                if month != expected_month:
                    raise ValueError(
                        f"{page_index + 1} ページの月の順序が想定と異なります: "
                        f"{month} 月 / 期待値 {expected_month} 月"
                    )
                year = academic_year if month >= FIRST_YEAR_MONTH else academic_year + 1
                all_months.append(parse_month(page, header_x, month, year))
    return all_months


# ==========================================================================
# URL / ファイル解決
# ==========================================================================

class PdfLinkParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.first_pdf_href: Optional[str] = None

    def handle_starttag(self, tag: str, attrs) -> None:
        if self.first_pdf_href is not None or tag.lower() != "a":
            return
        href = dict(attrs).get("href")
        if href and urlparse(href).path.lower().endswith(".pdf"):
            self.first_pdf_href = href


def _download(url: str) -> tuple[bytes, str, str]:
    request = Request(url, headers={"User-Agent": USER_AGENT})
    try:
        with urlopen(request, timeout=HTTP_TIMEOUT_SECONDS) as response:
            content_length = response.headers.get("Content-Length")
            if content_length and int(content_length) > MAX_DOWNLOAD_BYTES:
                raise ValueError(
                    f"ダウンロード対象が大きすぎます（上限 {MAX_DOWNLOAD_BYTES // 1024 // 1024} MB）。"
                )
            data = response.read(MAX_DOWNLOAD_BYTES + 1)
            if len(data) > MAX_DOWNLOAD_BYTES:
                raise ValueError(
                    f"ダウンロード対象が大きすぎます（上限 {MAX_DOWNLOAD_BYTES // 1024 // 1024} MB）。"
                )
            return data, response.geturl(), response.headers.get_content_type()
    except (HTTPError, URLError) as e:
        raise RuntimeError(f"URL を取得できませんでした: {url} ({e})") from e


def load_pdf_from_url(page_url: str) -> tuple[bytes, str]:
    data, final_url, content_type = _download(page_url)
    if content_type == "application/pdf" or data.startswith(b"%PDF-"):
        return data, final_url

    parser = PdfLinkParser()
    try:
        parser.feed(data.decode("utf-8", errors="replace"))
    except Exception as e:
        raise ValueError(f"Web ページの HTML を解析できませんでした: {final_url}") from e

    if parser.first_pdf_href is None:
        raise ValueError(f"Web ページに PDF へのリンクがありません: {final_url}")

    pdf_url = urljoin(final_url, parser.first_pdf_href)
    pdf_data, resolved_pdf_url, _ = _download(pdf_url)
    if not pdf_data.startswith(b"%PDF-"):
        raise ValueError(f"リンク先が PDF ではありません: {resolved_pdf_url}")
    return pdf_data, resolved_pdf_url


def resolve_pdf_source(pdf_file: Optional[str], page_url: Optional[str]):
    if page_url:
        pdf_data, pdf_url = load_pdf_from_url(page_url)
        return pdf_data, pdf_url
    path = Path(pdf_file or "r8schedule_20260624_1.pdf")
    if not path.is_file():
        raise FileNotFoundError(f"PDF ファイルがありません: {path}")
    return path, str(path)


# ==========================================================================
# エントリポイント
# ==========================================================================

def main() -> None:
    parser = argparse.ArgumentParser(
        description="年間行事予定表（PDF/Excel）から時間割用 CSV データを作成するプログラム"
    )
    source_group = parser.add_mutually_exclusive_group()
    source_group.add_argument(
        "-x", "--xlsx",
        help="ローカルの年間行事予定表 Excel ファイル (.xlsx)",
    )
    source_group.add_argument(
        "-p", "--pdf",
        help="ローカルの年間行事予定表 PDF ファイル",
    )
    source_group.add_argument(
        "-u", "--url",
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
        "-t", "--term", default=1, type=int, choices=(1, 2),
        help="前後期 (1: 前期, 2: 後期)",
    )
    parser.add_argument(
        "-j", "--subjects", default="subjects.json",
        help="科目定義 JSON ファイル",
    )
    parser.add_argument(
        "-y", "--year", default=2026, type=int,
        help="学年度の開始年（例: 2026 年度の PDF なら 2026）",
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
            all_months = parse_schedule_pdf(pdf_source, args.year)
        else:
            # --url または引数なし → デフォルト URL
            url = args.url or DEFAULT_PAGE_URL
            pdf_source, source_name = resolve_pdf_source(None, url)
            print(f"Reading schedule PDF: {source_name}", file=sys.stderr)
            all_months = parse_schedule_pdf(pdf_source, args.year)
    except (FileNotFoundError, RuntimeError, ValueError) as e:
        parser.error(str(e))

    # Excel 版は既に前期/後期で絞り込み済み、PDF 版は全年を返すのでスライス
    if args.xlsx:
        term_months = all_months
    else:
        if args.term == 1:
            term_months = all_months[:5]
        else:
            term_months = all_months[5:11]

    subjects = load_subjects(args.subjects, args.term)
    generate_csv(term_months, subjects, args.start, args.end)


if __name__ == "__main__":
    main()

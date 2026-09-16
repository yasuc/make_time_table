#!/usr/bin/env python3

"""
【このモジュールの目的】
    make_time_table.py / make_schedule.py の両方で使う共通機能をまとめたモジュールです。

    ・年間行事予定表 Excel / PDF の読み込み
    ・Web ページ URL からの PDF 取得
    ・科目データ (subjects.json) の読み込み
    ・レイアウト定数

    読み込んだデータは「月ごとの表」のリストで、各行は次の 17 列形式です。
        [0]      : 日付 (datetime)
        [1:6]    : 行事欄（行事最大 5 列分）
        [6:11]   : 本科 月〜金 の授業回数
        [11:16]  : 専攻科 月〜金 の授業回数
        [16]     : 未使用
    月の並びは 4 月 → 3 月（学年度順）です。
"""

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

# 表の列構成（PDF / Excel を正規化した共通の 17 列）
#   [0]        : 日付（datetime）
#   [1:6]      : 行事欄（行事最大 5 列分）
#   [6:11]     : 本科の 月～金 の 5 列
#   [11:16]    : 専攻科の 月～金 の 5 列
#   [16]       : 末尾の未使用領域
DATA_COL_OFFSET = 6  # 本科の開始列番号
NUM_WEEKDAYS = 5  # 月～金の 5 曜日
TOTAL_COLUMNS = 17  # 1 行の列数
EVENT_COL_START = 1  # 行事欄の開始列番号
EVENT_COL_END = 6  # 行事欄の終了列番号（この番号は含まない）

# 授業回数として認識する数値の上限（回数セルの数字チェック幅として使用）
CLASS_SESSION_MAX = 99

# 時限ごとの開始・終了時刻（1 限～ 5 限）
TIME_SLOTS = [
    ("8:50", "10:20"),  # 1 限
    ("10:30", "12:00"),  # 2 限
    ("13:00", "14:30"),  # 3 限
    ("14:40", "16:10"),  # 4 限
    ("13:00", "16:10"),  # 特殊な時限（3 限+4 限分の長さ）
]

# ---- Excel レイアウトに関する定数 ----
# Excel は 1 ブロック 17 列で、月ごとに並んでいます。
EXCEL_START_ROW = 5  # 表データの開始行（1 始まり）
EXCEL_END_ROW = 128  # 表データの終了行（1 始まり）
EXCEL_BLOCK_COLS = 17  # 1 か月分の列数
EXCEL_ALL_MONTHS = 12  # Excel に並んでいる月数（4 月〜 3 月）

# Excel の行構成（ブロック先頭からのオフセット）
#   [0]  : 日付
#   [1]  : 曜日番号（未使用）
#   [2:7]: 行事欄（行事最大 5 列分）
#   [7:12] : 本科 月〜金
#   [12:17]: 専攻科 月〜金
EXCEL_EVENT_COL_START = 2  # 行事欄の開始オフセット
EXCEL_EVENT_COL_END = 7  # 行事欄の終了オフセット（この番号は含まない）
EXCEL_HS_COL_START = 7  # 本科 月〜金 の開始オフセット
EXCEL_HS_COL_END = 12  # 本科 月〜金 の終了オフセット（この番号は含まない）
EXCEL_SS_COL_START = 12  # 専攻科 月〜金 の開始オフセット
EXCEL_SS_COL_END = 17  # 専攻科 月〜金 の終了オフセット（この番号は含まない）

# 前期（term=1）は 4 ～ 8 月の 5 か月、後期（term=2）は 9 ～ 2 月の 6 か月。
# Excel では前期が先頭ブロックから、後期はその後ろに続くため、開始列が異なります。
EXCEL_TERM1_START_COL = 1  # 前期の開始列（1 始まり）
EXCEL_TERM1_MONTHS = 5
EXCEL_TERM2_START_COL = 86  # 後期の開始列（1 始まり）
EXCEL_TERM2_MONTHS = 6

# ---- PDF レイアウトに関する定数（すべて座標・ドキュメント単位） ----
# 以下の「bb」は PyMuPDF の word 形式、つまり
#   (x0, y0, x1, y1, 文字列, ブロック番号, 行番号, 語番号)
# のタプルを指しています。

# 月ヘッダ（「4月」など）の文字として扱う y 範囲
MONTH_HEADER_Y_RANGE = (115, 140)

# 曜日ヘッダ（「月 火 水 木 金」...）の文字として扱う y 範囲
WEEKDAY_HEADER_Y_RANGE = (132, 142)

# 曜日ヘッダ文字を探す x 範囲の、ブロック左端からのオフセット
# （曜日ヘッダは月名の少し右側、同じ行に並んでいるため）
WEEKDAY_LABEL_X_MIN = 105.0
WEEKDAY_LABEL_X_MAX = 175.0

# 日付数字を探す x 範囲の、ブロック左端からのオフセット
DATE_LABEL_X_MIN = -2.0
DATE_LABEL_X_MAX = 10.0

# 日付数字を探す y 範囲（表の上端から下端まで）
DATE_LABEL_Y_MIN = 140
DATE_LABEL_Y_MAX = 760

# 行の境界線（罫線）を探す y 範囲
ROW_BOUNDARY_Y_MIN = 140.0
ROW_BOUNDARY_Y_MAX = 780.0

# 行の境界線とみなす線分の最小幅（縦罫線などを除外するため）
ROW_BOUNDARY_MIN_WIDTH = 20.0

# 行事欄のテキストを探す x 範囲の、ブロック左端からのオフセット
# （日付列の右端から最初の曜日列の前まで）
EVENT_LABEL_X_MIN = 10.0
EVENT_LABEL_X_MAX = 105.0

# 行事欄で「1 つのテキスト」とみなして結合する語間の最大ギャップ
# （途中で書体が変わるなどして語が分断される PDF 形式に対応する）
MAX_EVENT_WORD_GAP = 2.5

# 月名の文字中心とブロック左端の x 差（この PDF の帳票レイアウト由来）
MONTH_TO_BLOCK_OFFSET = 54.6

# ある日付の行に「回数」数字があるかどうかを判定する許容距離
# （セル内の数字の中心が、行の日付数字の中心とほぼ同じ位置にあることを確認する）
CELL_X_TOLERANCE = 2.2
CELL_Y_TOLERANCE = 4.0

# 曜日文字（本科・専攻科で同じ順に並ぶ）
WEEKDAY_CHARS = "月火水木金"

# PDF のページ構成：各ページにどの月が載っているか（順は左→右）
PAGE_MONTHS = (
    (4, 5, 6),  # 1 ページ目
    (7, 8, 9),  # 2 ページ目
    (10, 11, 12),  # 3 ページ目
    (1, 2, 3),  # 4 ページ目
)

# 学年度の考え方：4 ～ 12 月はその年度、1 ～ 3 月は翌年度
# （例）2026 年度 = 2026 年 4 月～ 2027 年 3 月
FIRST_YEAR_MONTH = 4

# 令和 n 年の西暦換算（令和1年 = 2019 年、令和7年 = 2025 年）
REIWA_TO_SEIREKI_OFFSET = 2018

# URL 取得時の設定
DEFAULT_PAGE_URL = "https://www.okinawa-ct.ac.jp/campus_life/class/annualev/"
HTTP_TIMEOUT_SECONDS = 30
MAX_DOWNLOAD_BYTES = 50 * 1024 * 1024
USER_AGENT = "make_schedule/1.0"


# ==========================================================================
# Excel 読み込み
# ==========================================================================


def old_pklfile_del(xlsx_file: str, pkl_file: str) -> None:
    """Excel ファイルの更新日時が pkl キャッシュより新しい場合は pkl を削除する。"""
    if not os.path.isfile(pkl_file):
        return
    p_xlsx = Path(xlsx_file)
    p_pkl = Path(pkl_file)
    if datetime.datetime.fromtimestamp(p_xlsx.stat().st_mtime) > \
       datetime.datetime.fromtimestamp(p_pkl.stat().st_mtime):
        os.remove(pkl_file)


def normalize_excel_row(raw_row: list) -> list:
    """Excel の 17 列レイアウトを共通の 17 列レイアウトに揃える。

    Excel: [日付, 曜日, 行事×5, 本科 月〜金, 専攻科 月〜金]
    共通 : [日付, 行事×5, 本科 月〜金, 専攻科 月〜金, 未使用]
    """
    row: list = [None] * TOTAL_COLUMNS
    row[0] = raw_row[0]  # 日付
    row[EVENT_COL_START:EVENT_COL_END] = raw_row[EXCEL_EVENT_COL_START:EXCEL_EVENT_COL_END]  # 行事欄
    row[DATA_COL_OFFSET:DATA_COL_OFFSET + NUM_WEEKDAYS] = raw_row[EXCEL_HS_COL_START:EXCEL_HS_COL_END]  # 本科
    row[DATA_COL_OFFSET + NUM_WEEKDAYS:DATA_COL_OFFSET + NUM_WEEKDAYS * 2] = raw_row[EXCEL_SS_COL_START:EXCEL_SS_COL_END]  # 専攻科
    return row


def _load_excel_months(xlsx_file: str, start_col: int, num_months: int) -> list[list[list]]:
    """Excel の指定した列範囲を行事予定表の共通形式で読み、月ごとの表を返す。"""
    wb = openpyxl.load_workbook(xlsx_file, data_only=True)
    sheet = wb.active

    all_months = []
    c = start_col
    for _ in range(num_months):
        month_data = []
        for row in sheet.iter_rows(
            min_row=EXCEL_START_ROW, max_row=EXCEL_END_ROW,
            min_col=c, max_col=c + EXCEL_BLOCK_COLS - 1,
        ):
            month_data.append(normalize_excel_row([cell.value for cell in row]))
        all_months.append(month_data)
        c += EXCEL_BLOCK_COLS
    return all_months


def parse_schedule_excel(xlsx_file: str, term: int, pkl_cache: Optional[str] = None) -> list[list[list]]:
    """Excel ファイルから指定した期の行事予定表を読み、月ごとの表を返す。

    term=1 は前期（4 ～ 8 月）、term=2 は後期（9 ～ 2 月）です。
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

    if term == 1:
        start_col = EXCEL_TERM1_START_COL
        num_months = EXCEL_TERM1_MONTHS
    else:
        start_col = EXCEL_TERM2_START_COL
        num_months = EXCEL_TERM2_MONTHS

    all_months = _load_excel_months(xlsx_file, start_col, num_months)

    with open(pkl_cache, "wb") as f:
        pickle.dump(all_months, f)

    return all_months


def parse_schedule_excel_all(xlsx_file: str, pkl_cache: Optional[str] = None) -> list[list[list]]:
    """Excel ファイルから年間行事予定表（全 12 か月）を読み、月ごとの表を返す。"""
    if openpyxl is None:
        raise SystemExit(
            "openpyxl が必要です。pip install openpyxl でインストールしてください。"
        )

    if pkl_cache is None:
        pkl_cache = "schedule_pdflayout.pkl"

    old_pklfile_del(xlsx_file, pkl_cache)

    if os.path.isfile(pkl_cache):
        print(f"Reading cached pkl: {pkl_cache}", file=sys.stderr)
        with open(pkl_cache, "rb") as f:
            return pickle.load(f)

    print(f"Reading Excel: {xlsx_file}", file=sys.stderr)

    all_months = _load_excel_months(xlsx_file, 1, EXCEL_ALL_MONTHS)

    with open(pkl_cache, "wb") as f:
        pickle.dump(all_months, f)

    return all_months


# ==========================================================================
# PDF テキスト抽出の小道具
# ==========================================================================


def word_center_x(word) -> float:
    """PyMuPDF の word タプルの x 座標中心を返す。"""
    return (word[0] + word[2]) / 2.0


def word_center_y(word) -> float:
    """PyMuPDF の word タプルの y 座標中心を返す。"""
    return (word[1] + word[3]) / 2.0


# ==========================================================================
# PDF の月ブロック解析
# ==========================================================================


def find_month_headers(page) -> list[tuple[float, int]]:
    """ページ内の「n 月」ヘッダを左から順に (x 中心, 月) で返す。

    月名は表の左上（y が約 115 ～ 140）に描かれています。
    同じ x 位置に複数回抽出されてもいいように、x を 0.1 単位で丸めて
    重複を取り除いています。
    """
    headers = []
    for word in page.get_text("words"):
        text = word[4]
        if (
            re.fullmatch(r"(?:[1-9]|1[0-2])月", text)
            and MONTH_HEADER_Y_RANGE[0] <= word[1] <= MONTH_HEADER_Y_RANGE[1]
        ):
            headers.append((word_center_x(word), int(text[:-1])))

    # x 座標で重複を除去（同じ列に月名が 2 回出ることはない想定）
    unique = {round(x, 1): month for x, month in headers}
    return sorted(unique.items())


def find_weekday_columns(page, block_left: float) -> list[float]:
    """1 か月分のブロック内にある 10 個の曜日ヘッダの x 中心を返す。

    この帳票では「月 火 水 木 金 月 火 水 木 金」の順に並び、
    前半 5 個が本科、後半 5 個が専攻科です。
    曜日文字が並んでいるのは表の見出し行（y が約 132 ～ 142）です。
    """
    candidates = []
    for word in page.get_text("words"):
        if not (WEEKDAY_HEADER_Y_RANGE[0] <= word[1] <= WEEKDAY_HEADER_Y_RANGE[1]):
            continue
        text = word[4]
        if not text or any(ch not in WEEKDAY_CHARS for ch in text):
            continue

        x = word_center_x(word)
        # 対象の月ブロックの範囲内の曜日だけを採用する
        if not (block_left + WEEKDAY_LABEL_X_MIN <= x <= block_left + WEEKDAY_LABEL_X_MAX):
            continue

        if len(text) == 1:
            candidates.append((x, text))
        else:
            # 「月火水木金…」のように 1 語へ結合された曜日ヘッダを
            # 等幅とみなして 1 文字ずつに分割する PDF 形式に対応する。
            width = (word[2] - word[0]) / len(text)
            for index, char in enumerate(text):
                candidates.append((word[0] + width * (index + 0.5), char))

    candidates.sort()
    if len(candidates) < 10:
        raise ValueError(
            f"曜日ヘッダを 10 列取得できませんでした。（取得数={len(candidates)}）"
        )
    return [x for x, _ in candidates[:10]]


def find_date_positions(page, block_left: float) -> list[tuple[int, float]]:
    """月ブロック左端の列から「日付数字」を抽出する。

    日付は表中の各行の左端に描かれています。
    （日付の y 中心が空きセルによるゴミ数字を除外する目印にもなります）
    帰り値は (日付, その行の y 中心) のリストで、y 順にソート済みです。
    """
    positions = []
    for word in page.get_text("words"):
        if not word[4].isdigit():
            continue
        x = word_center_x(word)
        y = word_center_y(word)

        # ブロック左端の縦の列の中にあり、かつ日付と同じ行にある数字
        if not (
            block_left + DATE_LABEL_X_MIN <= x <= block_left + DATE_LABEL_X_MAX
            and DATE_LABEL_Y_MIN <= y <= DATE_LABEL_Y_MAX
        ):
            continue

        day = int(word[4])
        if 1 <= day <= 31:
            positions.append((day, y))

    # 同じ日付が 2 回拾われた場合に備え、日付単位で最初の行位置を残す
    unique = {}
    for day, y in sorted(positions, key=lambda item: item[1]):
        unique.setdefault(day, y)
    return sorted(unique.items())


def find_row_boundaries(page) -> list[float]:
    """表の横罫線の y 座標を行の境界として昇順で返す。

    この帳票は行ごとに高さが一定でない（土日祝の行が高い）ため、
    日付との距離だけでは行事の所属する行を正しく判定できません。
    描画オブジェクトから横罫線を拾い、行の範囲を確定させます。
    """
    ys = set()
    for drawing in page.get_drawings():
        for item in drawing["items"]:
            if item[0] == "l":
                p1, p2 = item[1], item[2]
                if abs(p1.y - p2.y) < 0.5 and abs(p2.x - p1.x) >= ROW_BOUNDARY_MIN_WIDTH:
                    ys.add(round(p1.y, 1))
            elif item[0] == "re":
                rect = item[1]
                if (
                    rect.width >= ROW_BOUNDARY_MIN_WIDTH
                    and rect.height < 1.5
                ):
                    ys.add(round(rect.y0, 1))
    return sorted(
        y for y in ys if ROW_BOUNDARY_Y_MIN <= y <= ROW_BOUNDARY_Y_MAX
    )


def _row_index(boundaries: list[float], y: float) -> Optional[int]:
    """y 座標が含まれる行の添字を返す（行の境界線は [ 開始, 終了 ) で判定）。"""
    for index in range(len(boundaries) - 1):
        if boundaries[index] <= y < boundaries[index + 1]:
            return index
    return None


def get_cell_value(page, x: float, y: float, words) -> Optional[int]:
    """座標 (x, y) のセル内に「回数」の数字があればその値を返す。

    セル内の数字は行の日付数字とほとんど同じ場所にあります。
    そこで「日付数字との距離が許容範囲内にある数字」だけを採用します。
    数字が 0 ～ 99 の場合は授業回数として有効とみなします。
    """
    for word in words:
        if not word[4].isdigit():
            continue
        wx = word_center_x(word)
        wy = word_center_y(word)
        if abs(wx - x) <= CELL_X_TOLERANCE and abs(wy - y) <= CELL_Y_TOLERANCE:
            value = int(word[4])
            return value if 0 <= value <= CLASS_SESSION_MAX else None
    return None


def _find_event_words(words, block_left: float) -> list[tuple[float, float, float, str]]:
    """月ブロックの行事欄（日付列と曜日列の間）にあるテキストを抽出する。

    帰り値は (y 中心, x 開始, x 終了, 文字列) のリストです。
    """
    events = []
    for word in words:
        text = word[4]
        if len(text) == 1 and text in set(WEEKDAY_CHARS + "土日"):
            continue  # 日付列の曜日文字は読み飛ばす
        x0, x1 = word[0], word[2]
        y = word_center_y(word)
        if not (
            block_left + EVENT_LABEL_X_MIN <= (x0 + x1) / 2 <= block_left + EVENT_LABEL_X_MAX
            and DATE_LABEL_Y_MIN <= y <= DATE_LABEL_Y_MAX
        ):
            continue
        events.append((y, x0, x1, text))
    return events


def _join_event_text(words_in_line: list[tuple[float, float, str]]) -> list[str]:
    """同じ行に並んだ語を、隣接する語同士だけ結合してテキスト群を返す。

    書体の切り替わりなどで 1 つの文字列が複数の語に分断されている PDF
    形式に対応します。ギャップが大きい語は別の行事として扱います。
    """
    lines = sorted(words_in_line)
    merged: list[str] = []
    for x0, x1, text in lines:
        if merged and x0 - last_x1 <= MAX_EVENT_WORD_GAP:
            merged[-1] += text
            last_x1 = x1
        else:
            merged.append(text)
            last_x1 = x1
    return merged


def _row_event_texts(words_in_row: list[tuple[float, float, float, str]]) -> list[str]:
    """同じ行に属する行事の語を、行内の表示行ごとに結合してテキスト群を返す。"""
    line_tolerance = 2.0
    lines: list[list] = []
    for word in sorted(words_in_row):
        if lines and abs(word[0] - lines[-1][0][0]) <= line_tolerance:
            lines[-1].append(word)
        else:
            lines.append([word])
    texts: list[str] = []
    for line in lines:
        texts.extend(
            _join_event_text([(x0, x1, text) for _, x0, x1, text in line])
        )
    return texts


def parse_month(page, month_header_x: float, month: int, year: int) -> list[list]:
    """1 か月分を、共通の 17 列の 2 次元配列に変換する。

    戻り値の各行は
        [日付(datetime), 行事×5, 本科 月〜金の回数, 専攻科 月〜金の回数, 未使用]
    という形になります。回数を表す列には、授業がなければ None が入ります。
    """
    words = page.get_text("words")

    # 月名の中心から、その月ブロックの左端の x を逆算する
    block_left = month_header_x - MONTH_TO_BLOCK_OFFSET

    # このブロック内の曜日ヘッダの x 座標（10 個）を取得
    weekday_x = find_weekday_columns(page, block_left)

    # このブロック内の日付と各行の y 位置を取得
    date_positions = find_date_positions(page, block_left)

    # 行事欄のテキストを、日付と同じ行（罫線で区切られた範囲）へ対応付ける
    event_words = _find_event_words(words, block_left)
    boundaries = find_row_boundaries(page)

    # 行の添字ごとに、その行の日付と y 中心、行事テキストをまとめる
    row_info: dict[int, dict] = {}
    if boundaries:
        for day, y in date_positions:
            index = _row_index(boundaries, y)
            if index is not None:
                row_info[index] = {"day": day, "y": y, "events": []}
        for ey, x0, x1, text in event_words:
            index = _row_index(boundaries, ey)
            if index is not None and index in row_info:
                row_info[index]["events"].append((ey, x0, x1, text))

    # 罫線が取れなかった場合は、日付との距離で対応付ける（フォールバック）
    if not row_info:
        ordered_info = [
            {"day": day, "y": y, "events": []} for day, y in date_positions
        ]
        for ey, x0, x1, text in event_words:
            nearest = min(
                ordered_info, key=lambda info: abs(info["y"] - ey), default=None
            )
            if nearest is not None and abs(nearest["y"] - ey) <= CELL_Y_TOLERANCE:
                nearest["events"].append((ey, x0, x1, text))
    else:
        ordered_info = [row_info[index] for index in sorted(row_info)]
    ordered_info.sort(key=lambda info: info["y"])

    rows = []
    for info in ordered_info:
        day = info["day"]
        y = info["y"]
        try:
            date = datetime.datetime(year, month, day)
        except ValueError:
            # 存在しない日付（例: 2 月 31 日）は読み飛ばす
            continue

        row: list = [None] * TOTAL_COLUMNS
        row[0] = date

        # 行事欄: 同じ行にあるテキストを表示行ごとに結合し、順に行事列へ入れる。
        same_line = _row_event_texts(info["events"])
        for col_index, text in zip(range(EVENT_COL_START, EVENT_COL_END), same_line):
            row[col_index] = text

        # 本科（前半 5 列）と専攻科（後半 5 列）について、
        # それぞれ月曜～金曜のセルに回数が書かれているかを確認する。
        for course in range(2):
            for weekday in range(NUM_WEEKDAYS):
                cell_x = weekday_x[weekday + course * NUM_WEEKDAYS]
                target_column = DATA_COL_OFFSET + weekday + course * NUM_WEEKDAYS
                row[target_column] = get_cell_value(page, cell_x, y, words)

        rows.append(row)

    return rows


def parse_schedule_pdf(pdf_source, academic_year: int) -> list[list[list]]:
    """年間行事予定表 PDF を読み、学年度を通した月別の表を返す。

    返り値は「月ごとの表」のリスト。月の並びは 4 月 → 3 月です。
    ページ構成（PAGE_MONTHS）と PDF の記載月が一致しているかも検査します。
    """
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

                # 1 ～ 3 月は学年度の翌年にあたる
                year = academic_year if month >= FIRST_YEAR_MONTH else academic_year + 1
                all_months.append(parse_month(page, header_x, month, year))

    return all_months


# ==========================================================================
# URL / ファイル解決
# ==========================================================================


class PdfLinkParser(HTMLParser):
    """HTML に記載された最初の PDF リンクを取り出す。"""

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
    """URL を取得し、本文・最終URL・Content-Type を返す。"""
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
    """Web ページの最初の PDF（または PDF の直URL）をメモリへ読み込む。"""
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
    """CLI の指定から PyMuPDF に渡す入力元と表示名を返す。"""
    if page_url:
        pdf_data, pdf_url = load_pdf_from_url(page_url)
        return pdf_data, pdf_url

    path = Path(pdf_file or "r8schedule_20260624_1.pdf")
    if not path.is_file():
        raise FileNotFoundError(f"PDF ファイルがありません: {path}")
    return path, str(path)


def infer_academic_year(source_name: str) -> Optional[int]:
    """ファイル名・URL から学年度（開始年）を推測する。

    例: r7schedule_20251030.pdf → 2025 （令和7年度）
         r8schedule_20260624_1.pdf → 2026
         2025schedule.pdf → 2025
    """
    base = Path(source_name).name

    m = re.search(r"[rＲ](\d{1,2})\D", base)
    if m:
        return REIWA_TO_SEIREKI_OFFSET + int(m.group(1))

    m = re.search(r"令和(\d{1,2})年度", base)
    if m:
        return REIWA_TO_SEIREKI_OFFSET + int(m.group(1))

    m = re.search(r"20(\d{2})", base)
    if m:
        return 2000 + int(m.group(1))

    return None


def current_academic_year() -> int:
    """現在日時から学年度（開始年）を計算する（1 ～ 3 月は前年度扱い）。"""
    now = datetime.datetime.now()
    return now.year if now.month >= FIRST_YEAR_MONTH else now.year - 1


# ==========================================================================
# 科目データ
# ==========================================================================


@dataclass
class Subject:
    """subjects.json の 1 行（1 科目分）を表すデータクラス。

    subjects.json の各要素は次の並びになっています。
        [グループ番号, 科目名, 時限, 教室名, 科（0=本科 / 1=専攻科）]
    例）  [1, "情報セキュリティI", 2, "講1-2", 0]
          → グループ 1、2 限、教室「講1-2」、本科 の科目
    """

    group: int  # グループ番号（1 始まり）。列番号計算の起点になる
    name: str  # 科目名
    period: int  # 時限（1 ～ 5）。TIME_SLOTS の添え字に 1 引いて使う
    room: str  # 教室名（出力の [教室名] の部分に使う）
    course: int  # 科（0 = 本科, 1 = 専攻科）

    def target_column(self) -> int:
        """この科目の講義回数が入っている列番号（0 始まり）を返す。

        列構成は「本科: 6 ～ 10 列」→「専攻科: 11 ～ 15 列」。
        グループ番号 group は 1 から始まるので 1 を引いて加算し、
        course が 1 なら 5 列（本科分）だけ後ろにずらします。
        """
        return (self.group - 1) + DATA_COL_OFFSET + self.course * NUM_WEEKDAYS


def load_subjects(json_file: str, term: int) -> list[Subject]:
    """subjects.json から指定した期の科目一覧を読み込む。

    term=1 は前期（term1）、term=2 は後期（term2）を使います。
    """
    with open(json_file, "r", encoding="utf-8") as f:
        data = json.load(f)
    key = "term1" if term == 1 else "term2"
    return [Subject(*row) for row in data[key]]
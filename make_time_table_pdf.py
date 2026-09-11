#!/usr/bin/env python3

"""
【このプログラムの目的】
    年間行事予定表 PDF（4ページ、各ページに3か月分の表）から
    授業が行われた日を読み取り、「科目名・何回目・開始日時・終了日時」を
    CSV 形式で標準出力へ流すツールです。

    PDF には「日付 × 曜日列」のセルに授業回数の数字（1, 2, 3 ...）が
    印刷されています。その数字の有無を手がかりに、
    各科目の n 回目の講義日を特定しています。

【使い方】
    python3 make_time_table_pdf.py [オプション] > output.csv
    （ファイルは作成せず標準出力へ出すので、> でリダイレクトします）

【元のプログラムとの違い】
    ・変数名・関数名を日本語の意味が伝わる名前に変更
    ・レイアウト上の "マジックナンバー" を定数として冒頭に集約
    ・科目データを dataclass に整理
    ・新しい年度にも対応できるよう --year オプションを追加
"""

import argparse
import datetime
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

try:
    import pymupdf
except ImportError as e:
    raise SystemExit(
        "PyMuPDF が必要です。pip install pymupdf でインストールしてください。"
    ) from e

# ==========================================================================
# 設定値
# ==========================================================================

# 各時限の開始・終了時刻（配列の添え字 0 が 1 限、1 が 2 限 ...）
# subjects.json の「時限」欄は 1 始まりなので、参照時に 1 を引いています。
TIME_SLOTS = [
    ("8:50", "10:20"),  # 1 限
    ("10:30", "12:00"),  # 2 限
    ("13:00", "14:30"),  # 3 限
    ("14:40", "16:10"),  # 4 限
    ("13:00", "16:10"),  # 特殊な時限（3 限+4 限分の長さ）
]

# 授業回数として認識する数値の上限（回数セルの数字チェック幅として使用）
CLASS_SESSION_MAX = 99

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

# 月名の文字中心とブロック左端の x 差（この PDF の帳票レイアウト由来）
MONTH_TO_BLOCK_OFFSET = 54.6

# ある日付の行に「回数」数字があるかどうかを判定する許容距離
# （セル内の数字の中心が、行の日付数字の中心とほぼ同じ位置にあることを確認する）
CELL_X_TOLERANCE = 2.2
CELL_Y_TOLERANCE = 4.0

# 表の列構成（旧 Excel 版と互換の 17 列）
#   [0]        : 日付（datetime）
#   [1:6]      : 行事欄（このプログラムでは未使用）
#   [6:11]     : 本科の 月～金 の 5 列
#   [11:16]    : 専攻科の 月～金 の 5 列
#   [16]       : 末尾の未使用領域
DATA_COL_OFFSET = 6  # 本科の開始列番号
NUM_WEEKDAYS = 5  # 月～金の 5 曜日
TOTAL_COLUMNS = 17  # 1 行の列数

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
        if word[4] not in set(WEEKDAY_CHARS):
            continue

        x = word_center_x(word)
        # 対象の月ブロックの範囲内の曜日だけを採用する
        if block_left + WEEKDAY_LABEL_X_MIN <= x <= block_left + WEEKDAY_LABEL_X_MAX:
            candidates.append((x, word[4]))

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


def parse_month(page, month_header_x: float, month: int, year: int) -> list[list]:
    """1 か月分を、旧 Excel 版と同じく 17 列の 2 次元配列に変換する。

    戻り値の各行は
        [日付(datetime), 未使用×5, 本科 月～金の回数, 専攻科 月～金の回数, 未使用]
    という形になります。回数を表す列には、授業がなければ None が入ります。
    """
    words = page.get_text("words")

    # 月名の中心から、その月ブロックの左端の x を逆算する
    block_left = month_header_x - MONTH_TO_BLOCK_OFFSET

    # このブロック内の曜日ヘッダの x 座標（10 個）を取得
    weekday_x = find_weekday_columns(page, block_left)

    # このブロック内の日付と各行の y 位置を取得
    date_positions = find_date_positions(page, block_left)

    rows = []
    for day, y in date_positions:
        try:
            date = datetime.datetime(year, month, day)
        except ValueError:
            # 存在しない日付（例: 2 月 31 日）は読み飛ばす
            continue

        row: list = [None] * TOTAL_COLUMNS
        row[0] = date

        # 本科（前半 5 列）と専攻科（後半 5 列）について、
        # それぞれ月曜～金曜のセルに回数が書かれているかを確認する。
        for course in range(2):
            for weekday in range(NUM_WEEKDAYS):
                cell_x = weekday_x[weekday + course * NUM_WEEKDAYS]
                target_column = DATA_COL_OFFSET + weekday + course * NUM_WEEKDAYS
                row[target_column] = get_cell_value(page, cell_x, y, words)

        rows.append(row)

    return rows


def parse_schedule_pdf(pdf_file, academic_year: int) -> list[list[list]]:
    """年間行事予定表 PDF を読み、学年度を通した月別の表を返す。

    返り値は「月ごとの表」のリスト。月の並びは 4 月 → 3 月です。
    ページ構成（PAGE_MONTHS）と PDF の記載月が一致しているかも検査します。
    """

    pdf_path = Path(pdf_file)
    if not pdf_path.is_file():
        raise FileNotFoundError(f"PDF ファイルがありません: {pdf_file}")

    all_months = []
    with pymupdf.open(pdf_path) as doc:
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
# CSV 出力
# ==========================================================================


def generate_ical_csv(all_months, subjects, start: int, end: int) -> None:
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
                if row[target_column] is None:
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


# ==========================================================================
# エントリポイント
# ==========================================================================


def main() -> None:
    parser = argparse.ArgumentParser(
        description="年間行事予定表 PDF から時間割用データを作成するプログラム"
    )
    parser.add_argument(
        "-p",
        "--pdf",
        default="r8schedule_20260624_1.pdf",
        help="年間行事予定表 PDF ファイル",
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
        default=2026,
        type=int,
        help="学年度の開始年（例: 2026 年度の PDF なら 2026）",
    )
    args = parser.parse_args()

    if args.start < 1 or args.end < args.start:
        parser.error("--start と --end の指定が不正です。")

    print("Reading schedule PDF.", file=sys.stderr)
    all_months = parse_schedule_pdf(args.pdf, args.year)

    # 前期は 4 ～ 8 月、後期は 9 ～ 2 月（旧 Excel 版と同じ対象範囲）
    # ※ all_months の並びは「4月,5月,...,3月」なのでスライスで表せる
    if args.term == 1:
        term_months = all_months[:5]  # 4 月 ～ 8 月
    else:
        term_months = all_months[5:11]  # 9 月 ～ 2 月

    subjects = load_subjects(args.subjects, args.term)
    generate_ical_csv(term_months, subjects, args.start, args.end)


if __name__ == "__main__":
    main()


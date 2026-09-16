# make_time_table 時間割と行事予定のCSVファイル作成ツール

## 概要
年間行事予定表（PDF または Excel ファイル）から CSV ファイルを作成するツール群です。

| ツール | 入力 | 出力 |
|---|---|---|
| `make_time_table.py` | 年間行事予定表（PDF / Excel / URL） | 科目の時間割（講義日・時刻）CSV |
| `make_schedule.py` | 年間行事予定表（Excel） | 行事予定（終日イベント）CSV |

## make_time_table.py（時間割の作成）

年間行事予定表から授業が行われた日を読み取り、
「科目名・何回目・開始日時・終了日時」を CSV 形式で標準出力へ流します。
科目の定義は `subjects.json` で行います。

`--xlsx` / `--pdf` / `--url` のいずれも指定しない場合は、沖縄高専の年間行事予定表の
Web ページ（PDF）を自動的に読み込みます。

### 使い方（PDF・ローカルファイル）
```
python3 make_time_table.py -p r8schedule_20260624_1.pdf -t 1 > 2021time_table.csv
```

### 使い方（Excel）
```
python3 make_time_table.py -x 2026schedule.xlsx -t 1 > 2026time_table.csv
```
※一度実行すると `make_time_schedule_1_pdflayout.pkl`（前期）または `make_time_schedule_2_pdflayout.pkl`（後期）に
Excel のデータがシリアライズされ、次回からは高速に読み込まれます。Excel ファイルが更新されている場合は
自動的に再作成されます。

### 使い方（URL から取得・沖縄高専デフォルト）
```
python3 make_time_table.py -t 1 > 2026time_table.csv
```
`--url` を指定せずに実行すると、沖縄高専の年間行事予定表の Web ページに自動的にアクセスし、
最初にリンクされている PDF を使用します。

Web ページ内で最初にリンクされた PDF を使う場合（相対リンクにも対応）:
```
python3 make_time_table.py -u https://example.jp/schedule.html -t 1 > 2026time_table.csv
```

`--url` の後の URL を省略すると、沖縄高専の年間行事予定ページを使います。
```
python3 make_time_table.py --url -t 1 > 2026time_table.csv
```

PDF の URL を直接指定することもできます。

**オプション:**
| オプション | 説明 | デフォルト値 |
|---|---|---|
| `-x, --xlsx` | ローカルの年間行事予定表 Excel ファイル（.xlsx） | （指定なし = URL から取得） |
| `-p, --pdf` | ローカルの年間行事予定表 PDF ファイル | （指定なし = URL から取得） |
| `-u, --url [URL]` | 最初にリンクされた PDF を取得する Web ページ（PDF 直URLも可） | オプション指定時にURLを省略すると沖縄高専の年間行事予定ページ |
| `-s, --start` | 開始回 | 1 |
| `-e, --end` | 終了回 | 15 |
| `-t, --term` | 前後期（1:前期, 2:後期） | 1 |
| `-j, --subjects` | 科目定義JSONファイル | subjects.json |
| `-y, --year` | 学年度の開始年（PDF のみ使用） | 2026 |

- `--xlsx` / `--pdf` / `--url` は同時には指定できません。
- この 3 つをすべて省略した場合は、沖縄高専の年間行事予定表の Web ページから PDF を取得します。
- 「最初のPDF」は HTML 内の `<a href="...pdf">` を記載順に見た最初のリンクです。

## make_schedule.py（行事予定表 CSV の作成）

年間行事予定表 Excel ファイルに記載された行事名を読み取り、
「行事名・日付・終日イベント（TRUE）」の形式で CSV を標準出力へ流します。

### 使い方
```
python3 make_schedule.py 2026schedule.xlsx > 2026schedule.csv
```

**引数（位置引数）:**
| 引数 | 説明 | デフォルト値 |
|---|---|---|
| 第1引数 | 年間行事予定表 Excel ファイル（.xlsx） | schedule.xlsx |
| 第2引数 | キャッシュファイル名（.pkl） | schedule.pkl |

※一度実行すると `schedule.pkl` に Excel のデータがシリアライズされます。
Excel ファイルが更新されている場合は自動的に再作成されます。

**出力形式:**
```
Subject,Start Date,All Day Event
入学式,2025/04/03,TRUE
```
各行が 1 件の行事を表し、`All Day Event` は常に `TRUE`（終日イベント）です。
行事名からは「※…」の注記と空白文字が除去されます。

例:
```
python3 make_schedule.py 2025schedule.xlsx > 2025schedule.csv
```

## 科目定義ファイル (subjects.json)

`make_time_table.py` で使用する科目の定義ファイルです。
各科目は以下の形式で定義されています：

```json
[曜日, 科目名, 時限, 教室, 本科/専攻科]
```

| インデックス | 項目 | 値の例 | 説明 |
|---|---|---|---|
| 0 | 曜日 | 1〜5 | 1:月, 2:火, 3:水, 4:木, 5:金 |
| 1 | 科目名 | "情報セキュリティI" | 科目名 |
| 2 | 時限 | 1〜5 | 1:1限, 2:2限, 3:3限, 4:4限, 5:5限 |
| 3 | 教室 | "講1-2" | 教室名 |
| 4 | 本科/専攻科 | 0 or 1 | 0:本科, 1:専攻科 |

**時限の時間帯:**
| 時限 | 時間 |
|---|---|
| 1限 | 8:50〜10:20 |
| 2限 | 10:30〜12:00 |
| 3限 | 13:00〜14:30 |
| 4限 | 14:40〜16:10 |
| 5限 | 13:00〜16:10 |

**設定例:**
```json
{
  "term1": [
    [1, "コンピュータ概論", 2, "1号館201", 0],
    [2, "アルゴリズム入門", 2, "2号館305", 0],
    [3, "卒業研究", 5, "研究室A", 0],
    [4, "卒業研究", 5, "研究室A", 0],
    [5, "応用プログラミング", 3, "演習室B", 1]
  ],
  "term2": [
    [1, "ネットワーク基礎", 2, "1号館201", 0],
    [2, "データベース設計", 3, "演習室B", 0],
    [2, "セキュリティ概論", 4, "2号館305", 1],
    [5, "ソフトウェア工学", 2, "1号館201", 0],
    [5, "グラフィックス", 3, "演習室B", 1]
  ]
}
```

## 依存関係
- `pymupdf`（PDF 読み取り用 / make_time_table.py）: `pip install pymupdf`
- `openpyxl`（Excel 読み取り用 / make_time_table.py, make_schedule.py）: `pip install openpyxl`
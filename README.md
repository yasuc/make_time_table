# make_time_table 時間割と行事予定のCSVファイル作成ツール

## 概要
ExcelファイルまたはPDFファイルから行事予定表や時間割のCSVファイルを作成するツールです。

## 使い方

### 行事予定表の作成例（Excel）
```
python3 make_schedule.py 2021schedule.xlsx > 2021schedule.csv
```

### 時間割の作成例（Excel）
```
python3 make_time_table.py 2021schedule.xlsx > 2021time_table.csv
```

※一度実行するとschedule.pklファイルに行事予定表のデータがシリアライズされる。

### 時間割の作成例（PDF）
```
python3 make_time_table_pdf.py -p r8schedule_20260624_1.pdf -t 1 > 2021time_table.csv
```

**オプション:**
| オプション | 説明 | デフォルト値 |
|---|---|---|
| `-p, --pdf` | 年間行事予定表PDFファイル | r8schedule_20260624_1.pdf |
| `-s, --start` | 開始回 | 1 |
| `-e, --end` | 終了回 | 15 |
| `-t, --term` | 前後期（1:前期, 2:後期） | 1 |
| `-j, --subjects` | 科目定義JSONファイル | subjects.json |

**依存関係:**
- PyMuPDF (`pip install pymupdf`)


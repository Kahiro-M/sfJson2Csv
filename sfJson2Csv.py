import json
import csv
import sys

# ---------- ユーティリティ ----------
def normalize(s: str) -> str:
    return s.replace("\ufeff", "").strip()

# ---------- 入力 ----------
REPORT_JSON = "report.json"
HEADER_CSV = "output_header.csv"
OUTPUT_CSV = "output.csv"

if len(sys.argv)==1:
    REPORT_JSON = "report.json"
    HEADER_CSV = "output_header.csv"
    OUTPUT_CSV = "output.csv"
elif len(sys.argv)==2:
    REPORT_JSON = sys.argv[1]
elif len(sys.argv)==3:
    REPORT_JSON = sys.argv[1]
    HEADER_CSV = sys.argv[2]
elif len(sys.argv)==4:
    REPORT_JSON = sys.argv[1]
    HEADER_CSV = sys.argv[2]
    OUTPUT_CSV = sys.argv[3]
else:
    print("Usage: python sfJson2Csv.py [report.json] [header.csv] [output.csv]")
    print("  [report.json] : Salesforce レポートの JSON エクスポートファイル（省略時: report.json）")
    print("  [header.csv]  : 出力する列ラベルを1行目に持つ CSV ファイル（省略時: output_header.csv）")
    print("  [output.csv]  : 出力 CSV ファイル名（省略時: output.csv）")
    sys.exit(1)

# 1. JSON 読み込み
with open(REPORT_JSON, encoding="utf-8") as f:
    data = json.load(f)

# 2. ヘッダCSV（列ラベル）読み込み
with open(HEADER_CSV, encoding="utf-8-sig", newline="") as f:
    reader = csv.reader(f)
    header_labels = [normalize(h) for h in next(reader)]

# 3. 列ラベル → index マップ作成
detail_columns = data["reportMetadata"]["detailColumns"]
detail_info = data["reportExtendedMetadata"]["detailColumnInfo"]

column_index = {}
for i, key in enumerate(detail_columns):
    label = normalize(detail_info[key]["label"])
    column_index[label] = i
# 4. 出力対象 index を header から解決
target_indexes = []
for label in header_labels:
    if label not in column_index:
        available = ", ".join(column_index.keys())
        raise ValueError(f"列ラベルがレポートに存在しません: {label}\n利用可能: {available}")
    target_indexes.append(column_index[label])

# 5. 明細行抽出
rows = []

for fm in data["factMap"].values():
    if "rows" not in fm:
        continue
    for r in fm["rows"]:
        if not isinstance(r, dict):
            continue
        cells = r.get("dataCells")
        if not isinstance(cells, list):
            continue

        row = []
        for idx in target_indexes:
            cell = cells[idx] if idx < len(cells) else {}
            # value があれば value、なければ label
            # value = cell.get("value") if cell.get("value") not in (None, "") else cell.get("label")
            value = cell.get("label")
            row.append(value)

        rows.append(row)

# 6. CSV 出力（Excel対応）
with open(OUTPUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
    writer = csv.writer(f)
    writer.writerow(header_labels)
    writer.writerows(rows)

print(f"出力完了: {OUTPUT_CSV}")

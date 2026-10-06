# /// script
# requires-python = ">=3.10"
# dependencies = ["xlrd"]
# ///
"""ETL：內政部「十五歲以上人口按年齡及教育程度(96).xls」縣市表 -> output/kaohsiung_gradschool.csv

只讀名稱以「(縣市)」結尾的工作表；高雄與全國的研究所／碩士／博士人數（畢業、肄業）。
執行：uv run etl/etl.py
"""
import csv
import re
import sys
from datetime import datetime
from pathlib import Path

import xlrd

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "Data" / "十五歲以上人口按年齡及教育程度(96).xls"
OUT = ROOT / "output" / "kaohsiung_gradschool.csv"
LOG = ROOT / "etl" / "validation_log.txt"
FIELDS = ["year", "area", "education", "completion", "count", "population_15"]
COMPLETIONS = ("畢業", "肄業")
SUBTOTALS = {"總計", "臺灣地區", "臺灣省", "福建省"}  # 小計列／總計列，不是縣市
SPLIT_FROM = 100  # 100 年起才輸出博士、碩士分列（98–99 年博士偏低）


def norm(v):
    """移除所有空白（含全形空白）。"""
    return re.sub(r"\s+", "", str(v))


def to_int(v):
    """儲存格 -> int；「—」視為 0，去除千分位逗號。"""
    if isinstance(v, (int, float)):
        if v != int(v):
            raise ValueError(f"非整數: {v!r}")
        return int(v)
    s = norm(v).replace(",", "")
    if s in ("—", "－", "-", "–", ""):
        return 0
    return int(s)


def open_workbook():
    return xlrd.open_workbook(str(SRC))


def sheets_by_year(wb, kind):
    """kind='縣市' 或 '年齡'；回傳 {民國年: sheet}，名稱先 strip。"""
    out = {}
    for sh in wb.sheets():
        name = sh.name.strip()
        m = re.fullmatch(rf"(\d+)\.\d+\({kind}\)", name)
        if not m:
            continue
        y = int(m.group(1))
        if y in out:
            raise ValueError(f"重複的工作表年份 {y} ({kind})")
        out[y] = sh
    return out


def parse_county_sheet(sh, year):
    """依表頭文字找欄位，回傳 (rows, edu_cols)。

    rows: [(區域名稱(已去空白), 從 B 欄起的數值 list；比率欄為 None)]；edu_cols: {教育程度: 畢業欄號(A=0)}。
    """
    grad_row = next((r for r in range(sh.nrows) if norm(sh.cell_value(r, 3)) == "畢業"), None)
    if grad_row is None:
        raise ValueError(f"{year}: 找不到畢業/肄業表頭列")
    head = [[norm(sh.cell_value(r, c)) for c in range(sh.ncols)] for r in range(grad_row)]
    if not any(row[0] == "區域別" for row in head):
        raise ValueError(f"{year}: A 欄表頭不是「區域別」")
    if not any(row[1] == "總計" for row in head):
        raise ValueError(f"{year}: B 欄表頭不是「總計」")
    if not any(row[2].startswith("識字") for row in head):
        raise ValueError(f"{year}: C 欄表頭不是「識字」")

    label_cols = {}
    for row in head:
        for c, text in enumerate(row):
            if text in ("研究所", "博士", "碩士"):
                label_cols.setdefault(text, []).append(c)
    for k, cols in label_cols.items():
        if len(set(cols)) != 1:
            raise ValueError(f"{year}: 表頭「{k}」出現在多個欄位 {cols}")
    label_cols = {k: v[0] for k, v in label_cols.items()}

    if year <= 97:
        if set(label_cols) != {"研究所"}:
            raise ValueError(f"{year}: 預期只有研究所，實際表頭 {label_cols}")
        edu_cols = {"研究所": label_cols["研究所"]}
        expected = {"研究所": 3}
    else:
        # 98 年起「研究所」可能只是博士、碩士的上層群組標題，只取博士、碩士
        if not {"博士", "碩士"} <= set(label_cols):
            raise ValueError(f"{year}: 預期有博士與碩士，實際表頭 {label_cols}")
        edu_cols = {"博士": label_cols["博士"], "碩士": label_cols["碩士"]}
        expected = {"博士": 3, "碩士": 5}
    if edu_cols != expected:
        raise ValueError(f"{year}: 表頭欄位 {edu_cols} 與預期位置 {expected} 不一致")
    for edu, c in edu_cols.items():
        if (norm(sh.cell_value(grad_row, c)), norm(sh.cell_value(grad_row, c + 1))) != COMPLETIONS:
            raise ValueError(f"{year}: {edu} 欄位下不是 畢業/肄業")

    # 「不識字率」等比率欄不是人數，值記為 None（不轉整數、不參與加總）
    rate = {c for c in range(sh.ncols) if any(row[c].endswith("率") for row in head)}
    rows = []
    for r in range(grad_row + 1, sh.nrows):
        name = norm(sh.cell_value(r, 0))
        if not name or name.startswith("說明") or name.startswith("註"):
            continue  # 空白列、說明列
        if sh.cell_type(r, 1) == xlrd.XL_CELL_EMPTY:
            continue
        rows.append((name, [None if c in rate else to_int(sh.cell_value(r, c)) for c in range(1, sh.ncols)]))
    names = [n for n, _ in rows]
    if len(set(names)) != len(names):
        raise ValueError(f"{year}: 區域名稱重複")
    return rows, edu_cols


def extract_year(sh, year):
    """單一年份 -> {area: {"pop": int, (edu, completion): count}}"""
    rows, edu_cols = parse_county_sheet(sh, year)
    by_name = dict(rows)  # 值從 B 欄起算，index 0 = 總計
    if "總計" not in by_name:
        raise ValueError(f"{year}: 找不到「總計」列")
    kaohsiung_names = ["高雄市"] + (["高雄縣"] if year <= 98 else [])
    for n in kaohsiung_names:
        if n not in by_name:
            raise ValueError(f"{year}: 找不到「{n}」列")
    if year >= 99 and "高雄縣" in by_name:
        raise ValueError(f"{year}: 99 年起不應再有「高雄縣」")

    def collect(names):
        res = {"pop": sum(by_name[n][0] for n in names)}
        for edu, c in edu_cols.items():
            for i, comp in enumerate(COMPLETIONS):
                res[(edu, comp)] = sum(by_name[n][c - 1 + i] for n in names)  # c-1：陣列從 B 欄起
        return res

    return {"高雄": collect(kaohsiung_names), "全國": collect(["總計"])}


def build_rows(wb):
    county = sheets_by_year(wb, "縣市")
    out = []
    for year in sorted(county):
        data = extract_year(county[year], year)
        for area in ("高雄", "全國"):
            d = data[area]
            pop = d["pop"]
            for comp in COMPLETIONS:
                if year <= 97:
                    grad = d[("研究所", comp)]
                    parts = []
                else:
                    phd, ms = d[("博士", comp)], d[("碩士", comp)]
                    grad = phd + ms
                    parts = [("博士", phd), ("碩士", ms)] if year >= SPLIT_FROM else []
                for edu, cnt in parts + [("研究所", grad)]:
                    out.append(dict(year=year, area=area, education=edu, completion=comp,
                                    count=cnt, population_15=pop))
    out.sort(key=lambda r: (r["year"], r["area"] != "高雄",
                            ("博士", "碩士", "研究所").index(r["education"]),
                            COMPLETIONS.index(r["completion"])))
    return out


def write_csv(rows):
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        w.writerows(rows)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    wb = open_workbook()
    rows = build_rows(wb)
    write_csv(rows)
    print(f"已輸出 {len(rows)} 列 -> {OUT.relative_to(ROOT)}")

    import validate  # 同資料夾；uv run 會把腳本目錄放進 sys.path
    results = validate.run_all(wb, rows)
    lines = [f"ETL 執行時間：{datetime.now():%Y-%m-%d %H:%M:%S}",
             f"輸出：{OUT.relative_to(ROOT)}（{len(rows)} 列）", ""]
    lines += validate.format_results(results)
    LOG.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("\n".join(lines))
    print(f"驗證紀錄 -> {LOG.relative_to(ROOT)}")
    return 0 if all(ok for _, ok, _ in results) else 1


if __name__ == "__main__":
    sys.exit(main())

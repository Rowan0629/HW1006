# /// script
# requires-python = ">=3.10"
# dependencies = ["xlrd"]
# ///
"""驗證 output/kaohsiung_gradschool.csv 與原始 .xls。全部通過 exit 0，否則 exit 1。
執行：uv run etl/validate.py
"""
import csv
import sys
from collections import defaultdict

import etl

YEARS = list(range(96, 115))


def load_csv():
    with open(etl.OUT, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    for r in rows:
        r["year"], r["count"], r["population_15"] = int(r["year"]), int(r["count"]), int(r["population_15"])
    return rows


def check_keys(wb, rows):
    keys = [(r["year"], r["area"], r["education"], r["completion"]) for r in rows]
    dup = len(keys) - len(set(keys))
    years = sorted({r["year"] for r in rows})
    missing = [(y, a) for y in YEARS for a in ("高雄", "全國")
               if not any(r["year"] == y and r["area"] == a for r in rows)]
    expected = sum(2 * 2 * (3 if y >= etl.SPLIT_FROM else 1) for y in YEARS)
    bad_edu = {r["education"] for r in rows} - {"研究所", "碩士", "博士"}
    bad_area = {r["area"] for r in rows} - {"高雄", "全國"}
    ok = (years == YEARS and len(years) == 19 and not dup and not missing
          and len(rows) == expected and not bad_edu and not bad_area)
    return ok, f"{len(years)} 年、{len(rows)} 列（預期 {expected}），重複鍵 {dup}，缺少 {missing}"


def check_phd_ms(wb, rows):
    d = {(r["year"], r["area"], r["education"], r["completion"]): r["count"] for r in rows}
    bad, n = [], 0
    for y in range(etl.SPLIT_FROM, 115):
        for a in ("高雄", "全國"):
            for c in etl.COMPLETIONS:
                n += 1
                if d[(y, a, "博士", c)] + d[(y, a, "碩士", c)] != d[(y, a, "研究所", c)]:
                    bad.append((y, a, c))
    early = [r for r in rows if r["year"] < etl.SPLIT_FROM and r["education"] != "研究所"]
    return (not bad and not early), (f"{n} 組（100–114 年 × 2 地區 × 2）博士+碩士=研究所，不符 {bad}；"
                                     f"96–99 年多餘博碩列 {len(early)}")


def check_county_sum(wb, rows):
    county = etl.sheets_by_year(wb, "縣市")
    bad, notes = [], []
    for y in YEARS:
        parsed, _ = etl.parse_county_sheet(county[y], y)
        total = dict(parsed)["總計"]
        leaf = [(n, v) for n, v in parsed if n not in etl.SUBTOTALS]
        cols = [i for i in range(len(total)) if total[i] is not None]  # 排除比率欄（不識字率）
        diff = [i + 1 for i in cols if sum(v[i] for _, v in leaf) != total[i]]  # 欄號（B=1）
        if diff:
            bad.append((y, diff))
        if y <= 98:
            naive = sum(v[0] for n, v in parsed if n != "總計")
            notes.append((y, naive - total[0]))
    msg = (f"19 年逐年、全部人數欄（B 欄起，不含不識字率）縣市列加總 = 總計列，不符 {bad}。"
           "處理方式：排除「臺灣地區」「臺灣省」「福建省」小計列（及總計列）後加總；")
    if notes:
        msg += ("96–98 年若把所有非總計列直接加總，十五歲以上人口會多算 "
                + "、".join(f"{y}年 {d:,} 人" for y, d in notes)
                + "（這些小計列與其下縣市重複），故必須排除")
    return not bad, msg


def age_total(sh, year):
    hdr = next((r for r in range(sh.nrows) if etl.norm(sh.cell_value(r, 0)) == "年齡別"), None)
    if hdr is None or etl.norm(sh.cell_value(hdr, 2)) != "總計":
        raise ValueError(f"{year}: 年齡表表頭不符（C 欄應為總計）")
    for r in range(hdr + 1, sh.nrows):  # 第一個「計」列 = 全部年齡、男女合計
        if etl.norm(sh.cell_value(r, 1)) == "計":
            return etl.to_int(sh.cell_value(r, 2))
    raise ValueError(f"{year}: 年齡表找不到合計列")


def check_age_total(wb, rows):
    county = etl.sheets_by_year(wb, "縣市")
    age = etl.sheets_by_year(wb, "年齡")
    nat = {r["year"]: r["population_15"] for r in rows if r["area"] == "全國"}
    bad = []
    for y in YEARS:
        a = age_total(age[y], y)
        c = dict(etl.parse_county_sheet(county[y], y)[0])["總計"][0]
        if not (a == c == nat[y]):
            bad.append((y, a, c, nat[y]))
    return not bad, f"19 年逐年 年齡表總計 = 縣市表總計 = CSV 全國 population_15；114 年 {nat[114]:,}；不符 {bad}"


def check_kh_continuity(wb, rows):
    pop = {r["year"]: r["population_15"] for r in rows if r["area"] == "高雄"}
    chg = {y: (pop[y] / pop[y - 1] - 1) * 100 for y in YEARS[1:]}
    detail = "；".join(f"{y - 1}→{y}年 {chg[y]:+.2f}%" for y in (98, 99, 100))
    others = [abs(v) for y, v in chg.items() if y not in (98, 99)]
    ok = abs(chg[99]) <= 2.0 and abs(chg[98]) <= 2.0
    return ok, f"{detail}（門檻 ±2%）；其他年度最大 {max(others):.2f}%"


def check_reference(wb, rows):
    pop = {(r["year"], r["area"]): r["population_15"] for r in rows}
    cnt = defaultdict(int)
    for r in rows:
        cnt[(r["year"], r["area"], r["education"])] += r["count"]  # 畢業 + 肄業
    expect = [
        ("高雄15歲以上人口 96", pop[(96, "高雄")], 2305635),
        ("高雄15歲以上人口 98", pop[(98, "高雄")], 2343528),
        ("高雄15歲以上人口 100", pop[(100, "高雄")], 2381300),
        ("高雄15歲以上人口 114", pop[(114, "高雄")], 2424297),
        ("高雄研究所 96", cnt[(96, "高雄", "研究所")], 83746),
        ("高雄研究所 114", cnt[(114, "高雄", "研究所")], 214075),
        ("高雄博士 100", cnt[(100, "高雄", "博士")], 10344),
        ("高雄碩士 100", cnt[(100, "高雄", "碩士")], 107237),
        ("高雄博士 114", cnt[(114, "高雄", "博士")], 18544),
        ("高雄碩士 114", cnt[(114, "高雄", "碩士")], 195531),
    ]
    bad = [(k, got, exp) for k, got, exp in expect if got != exp]
    return not bad, f"{len(expect)} 個對照值，不符 {bad}"


CHECKS = [
    ("1. 19 年 × 高雄/全國齊全、鍵不重複", check_keys),
    ("2. 100 年起 博士+碩士=研究所", check_phd_ms),
    ("3. 各縣市列加總 = 全國總計", check_county_sum),
    ("4. 縣市表總計 = 年齡表總計（19 年）", check_age_total),
    ("5. 高雄 98→99 年人口變動 ≤ 2%", check_kh_continuity),
    ("6. 對照值", check_reference),
]


def run_all(wb, rows):
    out = []
    for name, fn in CHECKS:
        try:
            ok, detail = fn(wb, rows)
        except Exception as e:  # 驗證本身出錯視為不通過
            ok, detail = False, f"執行錯誤：{e!r}"
        out.append((name, ok, detail))
    return out


def format_results(results):
    lines = [f"[{'PASS' if ok else 'FAIL'}] {name}\n       {detail}" for name, ok, detail in results]
    n_ok = sum(ok for _, ok, _ in results)
    lines.append("")
    lines.append(f"結果：{n_ok}/{len(results)} 項通過 -> {'全部通過' if n_ok == len(results) else '有項目未通過'}")
    return lines


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    results = run_all(etl.open_workbook(), load_csv())
    print("\n".join(format_results(results)))
    return 0 if all(ok for _, ok, _ in results) else 1


if __name__ == "__main__":
    sys.exit(main())

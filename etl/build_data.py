# /// script
# requires-python = ">=3.10"
# ///
"""把 output/kaohsiung_gradschool.csv 轉成網頁用的 docs/data.js（window.GRAD_DATA = {...}），並核對。

不用 fetch，雙擊 docs/index.html 就能讀到資料。
執行：uv run etl/build_data.py
"""
import csv
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "output" / "kaohsiung_gradschool.csv"
OUT = ROOT / "docs" / "data.js"

META = {
    "title": "高雄歷年碩博士人數趨勢",
    "source": "內政部（戶政司），《十五歲以上人口按年齡及教育程度(96)》縣市表",
    "period": "民國 96 年底至 114 年底（西元 2007–2025 年，每年一筆）",
    "unit": "人",
    "mergeYear": 99,   # 高雄縣市合併：96–98 年為高雄縣 + 高雄市兩列相加，99 年起為合併後的高雄市
    "splitYear": 100,  # 博士、碩士分列穩定起點；96–99 年只有研究所合計
    "notes": [
        "96–99 年只有「研究所」（博士 + 碩士）；100 年起才分列博士與碩士。98–99 年博士人數偏低，不用於博碩分開的趨勢。",
        "高雄 96–98 年為高雄縣與高雄市兩列相加，99 年起為合併後的高雄市，99 年前後的差異不一定全是真實的人口變化。",
        "全國取縣市表「總計」列。",
        "研究所合計 = 博士 + 碩士；畢業與肄業分開保留，可自行切換或合併。",
    ],
    "limitations": [
        "資料是依戶籍登記的教育程度註記編製，不是抽樣調查，也不是在學學生數；「碩博士人數」是具該學歷的人口，不是當年的研究生人數。",
        "戶籍所在地不等於實際居住地。",
        "縣市表沒有年齡層與性別，不能拆分。",
        "圖表只描述高低與差異，不代表因果。",
    ],
}


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    with open(SRC, encoding="utf-8-sig", newline="") as f:
        records = [
            {"year": int(r["year"]), "area": r["area"], "education": r["education"],
             "completion": r["completion"], "count": int(r["count"]),
             "population_15": int(r["population_15"])}
            for r in csv.DictReader(f)
        ]
    years = sorted({r["year"] for r in records})
    meta = dict(META, years=[years[0], years[-1]], areas=["高雄", "全國"],
                educations=["研究所", "碩士", "博士"], completions=["畢業", "肄業"])
    data = {"meta": meta, "records": records}

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text("window.GRAD_DATA = " + json.dumps(data, ensure_ascii=False, separators=(",", ":")) + ";\n",
                   encoding="utf-8")

    # 核對：把剛寫出的 data.js 讀回來比對
    text = OUT.read_text(encoding="utf-8")
    assert text.startswith("window.GRAD_DATA = ") and text.endswith(";\n")
    back = json.loads(text[len("window.GRAD_DATA = "):-2])
    assert back == data, "data.js 讀回內容與來源不一致"
    recs = back["records"]
    pop = lambda y, a: {r["population_15"] for r in recs if r["year"] == y and r["area"] == a}
    grad = sum(r["count"] for r in recs
               if r["year"] == 114 and r["area"] == "高雄" and r["education"] == "研究所")
    checks = [
        ("高雄 114 年十五歲以上人口", pop(114, "高雄"), {2424297}),
        ("高雄 114 年研究所（畢業+肄業）", grad, 214075),
        ("全國 114 年十五歲以上人口", pop(114, "全國"), {20617242}),
    ]
    ok = True
    for name, got, exp in checks:
        good = got == exp
        ok &= good
        print(f"[{'PASS' if good else 'FAIL'}] {name}: {got} (預期 {exp})")
    print(f"{OUT.relative_to(ROOT)}：{len(recs)} 筆，{OUT.stat().st_size:,} bytes")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())

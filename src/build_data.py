#!/usr/bin/env python3
# Notion「三類法」(2026-10-17版) の訓練プログラム本文 → data.js（Web版）＋ ①承認シート用TSV を生成
import re, json, sys, os
HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "program_notion_2026-10-17.txt")
OUT_JS = os.path.join(HERE, "..", "data.js")
OUT_TSV = os.path.join(HERE, "out", "sheet1_program_v0.2.tsv")

BUI_MAP = {"頸椎（首）": "頸", "頸部": "頸", "肩関節": "肩", "腰部": "背腰", "背部・腰部": "背腰",
           "股関節": "股", "膝": "膝", "膝関節": "膝", "足関節": "足"}
KEYMAP = {"術前術後の検査": "kensa", "検査": "kensa",
          "対象にした病位": "taisho", "対象にした隣関節": "taisho", "対象": "taisho",
          "選択できる技術": "houhou", "施術の選択方法": "houhou", "技術": "houhou", "方法": "houhou",
          "指定される経穴": "keiketsu", "経穴": "keiketsu",
          "指定される番手": "bante", "番手": "bante", "指定される時間": "bante", "時間": "bante",
          "技術の合格基準": "gokaku", "合格": "gokaku"}

rows = []
cls = bui = buiLabel = jutsu = None
cur = None
def flush():
    global cur
    if cur: rows.append(cur); cur = None

for raw in open(SRC, encoding="utf-8"):
    line = raw.rstrip("\n").strip()
    if not line: continue
    m = re.match(r"^# (陰陽|六道|三焦)分類", line)
    if m: flush(); cls = m.group(1); continue
    m = re.match(r"^## (陰陽|六道|三焦)分類：(.+)$", line)
    if m:
        flush(); cls = m.group(1); buiLabel = m.group(2).strip(); bui = BUI_MAP[buiLabel]; continue
    m = re.match(r"^\*\*(手技|鍼術|灸術)\*\*$", line)
    if m: flush(); jutsu = m.group(1); continue
    m = re.match(r"^M([1-5])【(.+?)】(.*)$", line)
    if m:
        flush()
        cur = {"id": f"{cls}-{bui}-{jutsu}-M{m.group(1)}", "cls": cls, "bui": bui, "buiLabel": buiLabel,
               "jutsu": jutsu, "m": int(m.group(1)), "title": m.group(2),
               "kensa": "", "taisho": "", "houhou": "", "keiketsu": "", "bante": "", "gokaku": ""}
        rest = m.group(3).strip()
        if rest:  # 圧縮形式: 検査：…／技術：…／合格：…
            for part in rest.split("／"):
                k, _, v = part.partition("：")
                k = k.strip(); v = v.strip()
                if k not in KEYMAP: raise SystemExit(f"unknown key [{k}] in {line}")
                cur[KEYMAP[k]] = v
        continue
    # 詳細形式: キー：値
    k, sep, v = line.partition("：")
    if sep and k.strip() in KEYMAP and cur is not None:
        cur[KEYMAP[k.strip()]] = v.strip(); continue
    raise SystemExit(f"unparsed line: {line}")
flush()

# 検証：各分類×部位×術×M が揃っているか
from collections import Counter
c = Counter((r["cls"], r["bui"], r["jutsu"]) for r in rows)
bad = {k: v for k, v in c.items() if v != 5}
if bad: raise SystemExit(f"incomplete groups: {bad}")
exp = {"陰陽": 5, "六道": 6, "三焦": 6}
for cl, n in exp.items():
    got = len({r["bui"] for r in rows if r["cls"] == cl})
    if got != n: raise SystemExit(f"{cl}: 部位数 {got} != {n}")
print("rows:", len(rows), {cl: sum(1 for r in rows if r["cls"] == cl) for cl in exp})

# data.js
with open(OUT_JS, "w", encoding="utf-8") as f:
    f.write("// 三類法 訓練プログラム（Notion「三類法」2026-10-17版 植田完成版の転記）。src/build_data.py で生成。手で編集しない。\n")
    f.write("const DB = " + json.dumps(rows, ensure_ascii=False, indent=1) + ";\n")

# ① 承認シート v0.2 TSV
os.makedirs(os.path.dirname(OUT_TSV), exist_ok=True)
head = ["ID", "分類", "部位", "術", "M", "タイトル", "術前術後の検査", "対象（病位／隣関節）", "施術の選択方法／選択できる技術",
        "指定される経穴", "指定される番手／時間", "技術の合格基準", "出典", "植田さん判定（承認／修正あり／却下）", "修正内容メモ"]
with open(OUT_TSV, "w", encoding="utf-8") as f:
    f.write("\t".join(head) + "\n")
    for r in rows:
        vals = [r["id"], r["cls"], r["buiLabel"], r["jutsu"], f"M{r['m']}", r["title"], r["kensa"], r["taisho"] or "―",
                r["houhou"], r["keiketsu"] or "―", r["bante"] or "―", r["gokaku"], "Notion 三類法（2026-10-17版）", "", ""]
        f.write("\t".join(v.replace("\t", " ") for v in vals) + "\n")
print("wrote", OUT_JS, OUT_TSV)

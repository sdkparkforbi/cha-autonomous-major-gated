# -*- coding: utf-8 -*-
"""블록 → HWPX (hwpxgen). index.html을 최대한 반영."""
import pathlib
from hwpxgen import (HwpxDoc, CH_TITLE, CH_H1, CH_H2, CH_NORMAL, CH_SMALL, CH_WHITE,
                     CH_CELL, PA_CENTER, PA_LEFT)
from parse_proposal import parse

HERE = pathlib.Path(__file__).parent

d = HwpxDoc("_seed.hwpx").page(landscape=False, margin_lr_mm=18, margin_tb_mm=15)
CW = d.content_w
NAVY = d.fill("15294D"); LIGHT = d.fill("EEF3FB"); TEAL = d.fill("0E8A7D")
BOLD = d.char(CH_NORMAL, "1F3864", bold=True)
TEALC = d.char(CH_SMALL, "0E8A7D", bold=True)
CMAP = {"": CH_NORMAL, "b": BOLD, "em": BOLD}

def san(t):
    return t.replace("📦", "").replace(" ", " ")

def runs_to(lines, cmap=CMAP):
    out = []
    for st, t in lines:
        if t == "\n":
            out.append((" ", CH_NORMAL)); continue
        if t == "":
            continue
        out.append((san(t), cmap.get(st, CH_NORMAL)))
    return out or [(" ", CH_NORMAL)]

def para(runs, paraPr="21", cmap=CMAP):
    return d._para(runs_to(runs, cmap), paraPr)

def line(text, ch=CH_NORMAL, paraPr="21"):
    return d._para([(san(text), ch)], paraPr)

def widths(n):
    w = CW // n
    return [w] * (n - 1) + [CW - w * (n - 1)]

def htable(headers, rows, ncols):
    ws = widths(ncols)
    trows = []; r = 0
    if headers:
        trows.append([d.cell([(san(h), CH_WHITE)], c, 0, ws[c], bf=NAVY) for c, h in enumerate(headers)])
        r = 1
    for cells in rows:
        first = cells[0][0]
        is_total = any(("소계" in t or "합계" in t) for t, _ in cells)
        is_grand = "합계" in first
        rc = []; col = 0
        for (t, cs) in cells:
            w = sum(ws[col:col + cs])
            bf = (NAVY if is_grand else LIGHT) if is_total else "3"
            ch = CH_WHITE if is_grand else CH_CELL
            rc.append(d.cell([(san(t), ch)], col, r, w, bf=bf, colspan=cs))
            col += cs
        trows.append(rc); r += 1
    return d.table(trows, ws, treat_as_char=False)

def hbox(box):
    s = line("【 " + box["tag"] + " 】", BOLD, PA_LEFT)
    for blk in box["blocks"]:
        if blk["type"] == "p":
            s += para(blk["runs"])
        else:
            if blk["caption"]:
                s += line(blk["caption"], CH_SMALL)
            s += htable(blk["headers"], blk["rows"], blk["ncols"])
    return s

def hfigure(b):
    s = line("[그림] GATED 파이프라인: 비정형 텍스트 → 인과 지식그래프 → 그래프 추론", BOLD)
    s += line("STEP 1 대량 수집(뉴스 API·KCI) → STEP 2 관련성 필터링(GPT) → "
              "STEP 3 관계 추출·지식그래프 → STEP 4 GATED 추론(허브·예측)", CH_NORMAL)
    s += line("추출 triplet 예시: (무전공 유형Ⅰ→촉진→인기학과 쏠림) / (전공탐색학기→억제→중도탈락)", CH_SMALL)
    s += line(b["caption"], CH_SMALL)
    return s

def build():
    body = ""; TITLE = "자율전공제는 어떻게 성공하는가"
    for b in parse():
        ty = b["type"]
        if ty == "title":
            TITLE = b["title"]
            body += d._para([(b["kicker"], TEALC)], PA_CENTER)
            body += d._para([(b["sub"], CH_H2)], PA_CENTER)
            body += d._para([(" | ".join(b["chips"]), CH_SMALL)], PA_CENTER)
            for spans in b["metas"]:
                body += d._para([("   ".join(spans), CH_SMALL)], PA_CENTER)
            body += line(" ")
        elif ty == "note":
            txt = "".join(t for _, t in b["runs"] if t != "\n")
            body += d.table([[d.cell([(san(txt), CH_WHITE)], 0, 0, CW, bf=TEAL)]], [CW], treat_as_char=False)
        elif ty == "h1":
            body += line(b["text"], CH_H1)
        elif ty == "h2":
            body += line("▍ " + b["text"], CH_H2)
        elif ty == "p":
            body += para(b["runs"])
        elif ty == "ul":
            for item in b["items"]:
                body += para([("", "• ")] + item)
        elif ty == "box":
            body += hbox(b)
        elif ty == "table":
            if b["caption"]:
                body += line(b["caption"], CH_SMALL)
            body += htable(b["headers"], b["rows"], b["ncols"])
        elif ty == "figure":
            body += hfigure(b)
        elif ty == "foot":
            body += d._para([(" · ".join(b["spans"]), CH_SMALL)], PA_CENTER)

    out = HERE / "자율전공_GATED_연구제안서.hwpx"
    d.save(str(out), TITLE, body=body)
    print("HWPX:", out, out.stat().st_size, "bytes")

if __name__ == "__main__":
    build()

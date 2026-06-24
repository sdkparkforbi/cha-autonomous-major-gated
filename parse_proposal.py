# -*- coding: utf-8 -*-
"""index.html → 구조화 블록 리스트 (docx/hwpx 공통 입력)."""
import re, pathlib
from bs4 import BeautifulSoup, NavigableString

HERE = pathlib.Path(__file__).parent

def _norm(s):
    return re.sub(r"\s+", " ", s).strip()

def inline_runs(el):
    """(style, text) 런 리스트. style: '' | 'b' | 'em'. <br>는 '\n'."""
    runs = []
    for node in el.children:
        if isinstance(node, NavigableString):
            runs.append(("", str(node)))
        elif node.name in ("strong", "b"):
            runs.append(("b", node.get_text("")))
        elif node.name == "em":
            runs.append(("em", node.get_text("")))
        elif node.name == "br":
            runs.append(("", "\n"))
        elif node.name == "span":
            runs.append(("", node.get_text("")))
        else:
            runs.extend(inline_runs(node))
    # normalize whitespace per run, keep \n
    out = []
    for st, t in runs:
        parts = t.split("\n")
        for i, p in enumerate(parts):
            p = re.sub(r"[ \t]+", " ", p)
            if p.strip() or st == "":
                out.append((st, p))
            if i < len(parts) - 1:
                out.append(("", "\n"))
    # drop leading/trailing empty
    while out and out[0][1].strip() == "" and out[0][1] != "\n":
        out.pop(0)
    return out

def parse_table(t):
    cap = _norm(t.caption.get_text(" ")) if t.caption else ""
    headers = []
    thead = t.find("thead")
    if thead:
        headers = [_norm(th.get_text(" ")) for th in thead.find_all("th")]
    rows = []
    body = t.find("tbody") or t
    for tr in body.find_all("tr"):
        cells = []
        for td in tr.find_all(["td", "th"]):
            cs = int(td.get("colspan", 1))
            cells.append((_norm(td.get_text(" ")), cs))
        rows.append(cells)
    ncols = len(headers) if headers else max(sum(c for _, c in r) for r in rows)
    return {"type": "table", "caption": cap, "headers": headers, "rows": rows, "ncols": ncols}

def parse_box(div):
    classes = div.get("class", [])
    kind = next((c for c in ("key", "hyp", "gap", "dropin") if c in classes), "key")
    tag, blocks = "", []
    for ch in div.children:
        if isinstance(ch, NavigableString):
            continue
        cl = ch.get("class", [])
        if ch.name == "span" and "tag" in cl:
            tag = _norm(ch.get_text(" "))
        elif ch.name == "p":
            blocks.append({"type": "p", "runs": inline_runs(ch)})
        elif ch.name == "table":
            blocks.append(parse_table(ch))
    return {"type": "box", "kind": kind, "tag": tag, "blocks": blocks}

def parse_cover(h):
    def txt(sel):
        e = h.select_one(sel)
        return _norm(e.get_text(" ")) if e else ""
    chips = [_norm(c.get_text(" ")) for c in h.select(".chip")]
    metas = []
    for m in h.select(".meta"):
        spans = [_norm(s.get_text(" ")) for s in m.find_all("span")]
        metas.append(spans)
    return {"type": "title", "kicker": txt(".kicker"),
            "title": _norm(h.find("h1").get_text(" ")), "sub": txt(".sub"),
            "chips": chips, "metas": metas}

def parse(html_path=HERE / "index.html"):
    soup = BeautifulSoup(pathlib.Path(html_path).read_text(encoding="utf-8"), "html.parser")
    page = soup.select_one(".page")
    blocks = []
    for el in page.children:
        if isinstance(el, NavigableString):
            continue
        cl = el.get("class", [])
        if el.name == "header" and "cover" in cl:
            blocks.append(parse_cover(el))
        elif el.name == "div" and "readme" in cl:
            blocks.append({"type": "note", "runs": inline_runs(el)})
        elif el.name == "div" and "box" in cl:
            blocks.append(parse_box(el))
        elif el.name == "div" and "figwrap" in cl:
            cap = el.select_one(".figcap")
            blocks.append({"type": "figure", "caption": _norm(cap.get_text(" ")) if cap else ""})
        elif el.name == "div" and "foot" in cl:
            spans = [_norm(s.get_text(" ")) for s in el.find_all("span")]
            blocks.append({"type": "foot", "spans": spans})
        elif el.name == "h2":
            for sp in el.find_all("span", class_="num"):
                sp.extract()
            blocks.append({"type": "h1", "text": _norm(el.get_text(" "))})
        elif el.name == "h3":
            blocks.append({"type": "h2", "text": _norm(el.get_text(" "))})
        elif el.name == "p":
            blocks.append({"type": "p", "runs": inline_runs(el), "lead": "lead" in cl})
        elif el.name == "ul":
            items = [inline_runs(li) for li in el.find_all("li", recursive=False)]
            blocks.append({"type": "ul", "items": items})
        elif el.name == "table":
            blocks.append(parse_table(el))
    return blocks

if __name__ == "__main__":
    bs = parse()
    from collections import Counter
    print("blocks:", len(bs), dict(Counter(b["type"] for b in bs)))
    for b in bs:
        if b["type"] in ("h1", "h2"):
            print(" ", b["type"], b["text"])
        elif b["type"] == "table":
            print("   table:", b["caption"][:40], "| cols", b["ncols"], "| rows", len(b["rows"]))
        elif b["type"] == "box":
            print("   box[%s]:" % b["kind"], b["tag"][:30])

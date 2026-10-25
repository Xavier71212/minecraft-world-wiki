#!/usr/bin/env python3
"""Word 條目翻譯輔助工具（給 Claude 用）。

把中文條目的文字段落抽成簡單標記，翻譯後再套回同一份 Word 檔的副本，
表格、圖片、段落樣式、超連結目標與顏色都保持不變。

用法：
  python tools/docx_i18n.py extract 中文.docx > 原文.json
  python tools/docx_i18n.py apply   中文.docx 譯文.json 譯文.docx [語言代碼，預設 en；托拉語用 art-x-tora]

JSON 是一個陣列，每個元素 {"i": 段落編號, "t": 標記文字}；"i": "title" 為標題。
標記：
  [[目標.docx|文字]]   超連結（目標沿用原文的連結目標；顏色沿用原文同目標的連結）
  **文字**             粗體
  //文字//             斜體（英文書名用）
  ==文字==             黃色螢光筆（尚無條目的專有名詞）
  {{RRGGBB|文字}}      指定字色
只含圖片的段落不會出現在 JSON 裡，套回時原封不動。
譯文檔的「摘要」欄位會記下中文原檔的 SHA-1，建站時用來判斷翻譯是否需要更新。
"""
import copy
import hashlib
import json
import re
import sys

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

FONT_LATIN = "Microsoft JhengHei"
FONT_EA = "微軟正黑體"
TOKEN_RE = re.compile(r"(\[\[[^\]|]+\|[^\]]*\]\]|\*\*.+?\*\*|//.+?//|==.+?==|\{\{[0-9A-Fa-f]{6}\|.+?\}\})")


def text_paragraphs(doc):
    """依文件順序列出（編號, 段落元素），包含表格儲存格中的段落。"""
    out = []
    n = 0
    for el in doc.element.body.iterchildren():
        tag = el.tag.split("}")[1]
        if tag == "p":
            out.append((n, el)); n += 1
        elif tag == "tbl":
            for p in el.iter(qn("w:p")):
                out.append((n, p)); n += 1
    return out


def is_title(doc, p):
    ppr = p.find(qn("w:pPr"))
    if ppr is None or ppr.find(qn("w:pStyle")) is None:
        return False
    sid = ppr.find(qn("w:pStyle")).get(qn("w:val"))
    try:
        st = doc.styles.get_by_id(sid, 1)
        return st is not None and st.name.lower() in ("title", "標題")
    except Exception:
        return sid.lower() == "title"


def has_drawing(p):
    return p.find(".//" + qn("w:drawing")) is not None


def rpr_flags(r):
    rpr = r.find(qn("w:rPr"))
    f = {}
    if rpr is None:
        return f
    b = rpr.find(qn("w:b"))
    if b is not None and b.get(qn("w:val")) not in ("0", "false"):
        f["b"] = True
    i = rpr.find(qn("w:i"))
    if i is not None and i.get(qn("w:val")) not in ("0", "false"):
        f["i"] = True
    c = rpr.find(qn("w:color"))
    if c is not None and c.get(qn("w:val")) not in (None, "auto", "000000"):
        f["c"] = c.get(qn("w:val")).upper()
    h = rpr.find(qn("w:highlight"))
    if h is not None and h.get(qn("w:val")) not in (None, "none"):
        f["h"] = True
    return f


def run_text(r):
    s = ""
    for c in r.iterchildren():
        t = c.tag.split("}")[1]
        if t == "t":
            s += c.text or ""
        elif t == "tab":
            s += "\t"
    return s


def wrap(s, f):
    if not s:
        return ""
    if f.get("h"):
        s = f"=={s}=="
    if f.get("c"):
        s = "{{" + f["c"] + "|" + s + "}}"
    if f.get("i"):
        s = f"//{s}//"
    if f.get("b"):
        s = f"**{s}**"
    return s


def para_markup(p, part):
    pieces = []  # (flags, text) 合併相同格式
    out = []

    def flush():
        while pieces:
            f, t = pieces.pop(0)
            out.append(wrap(t, f))

    for c in p.iterchildren():
        tag = c.tag.split("}")[1]
        if tag == "r":
            t = run_text(c)
            if not t:
                continue
            f = rpr_flags(c)
            if pieces and pieces[-1][0] == f:
                pieces[-1] = (f, pieces[-1][1] + t)
            else:
                pieces.append((f, t))
        elif tag == "hyperlink":
            flush()
            rid = c.get(qn("r:id"))
            tgt = part.rels[rid].target_ref if rid and rid in part.rels else ""
            t = "".join(run_text(r) for r in c.iter(qn("w:r")))
            out.append(f"[[{tgt}|{t}]]")
    flush()
    return "".join(out)


def extract(path):
    doc = Document(path)
    items = []
    for n, p in text_paragraphs(doc):
        if has_drawing(p):
            continue
        m = para_markup(p, doc.part)
        if not m.strip():
            continue
        items.append({"i": "title" if is_title(doc, p) else n, "t": m})
    return items


# ---------------------------------------------------------------- 套回譯文

def set_font(rpr):
    f = rpr.find(qn("w:rFonts"))
    if f is None:
        f = OxmlElement("w:rFonts"); rpr.insert(0, f)
    for a in list(f.attrib):
        if a.endswith("Theme"):
            del f.attrib[a]
    f.set(qn("w:ascii"), FONT_LATIN); f.set(qn("w:hAnsi"), FONT_LATIN)
    f.set(qn("w:eastAsia"), FONT_EA); f.set(qn("w:cs"), FONT_LATIN)


def strip_fmt(rpr):
    for t in ("w:b", "w:bCs", "w:i", "w:iCs", "w:color", "w:highlight", "w:u", "w:rStyle"):
        for e in rpr.findall(qn(t)):
            rpr.remove(e)
    return rpr


def make_run(text, base_rpr, *, b=False, i=False, color=None, hl=False, rpr=None):
    r = OxmlElement("w:r")
    rp = copy.deepcopy(rpr if rpr is not None else base_rpr)
    if rp is None:
        rp = OxmlElement("w:rPr")
    set_font(rp)
    if rpr is None:
        strip_fmt(rp)
        if b:
            rp.append(OxmlElement("w:b"))
        if i:
            rp.append(OxmlElement("w:i"))
        if color:
            c = OxmlElement("w:color"); c.set(qn("w:val"), color); rp.append(c)
        if hl:
            h = OxmlElement("w:highlight"); h.set(qn("w:val"), "yellow"); rp.append(h)
    # rPr 子元素順序：rFonts, b, i, color, highlight 大致符合 schema；Word 可容忍
    r.append(rp)
    t = OxmlElement("w:t"); t.text = text; t.set(qn("xml:space"), "preserve")
    r.append(t)
    return r


def parse(markup):
    """回傳 tokens：(kind, text, extra)。kind: text/link/b/i/hl/color（可巢狀一層）。"""
    toks = []
    for part in TOKEN_RE.split(markup):
        if not part:
            continue
        if part.startswith("[[") and "|" in part:
            tgt, txt = part[2:-2].split("|", 1)
            toks.append(("link", txt, tgt))
        elif part.startswith("**") and part.endswith("**") and len(part) > 4:
            toks.append(("b", part[2:-2], None))
        elif part.startswith("//") and part.endswith("//") and len(part) > 4:
            toks.append(("i", part[2:-2], None))
        elif part.startswith("==") and part.endswith("==") and len(part) > 4:
            toks.append(("hl", part[2:-2], None))
        elif part.startswith("{{") and part.endswith("}}"):
            col, txt = part[2:-2].split("|", 1)
            toks.append(("color", txt, col.upper()))
        else:
            toks.append(("text", part, None))
    return toks


def apply(src, trans, out, title=None, lang="en"):
    doc = Document(src)
    part = doc.part
    # 原文中每個連結目標對應的 rId 與 rPr（保留顏色）
    link_info = {}
    for h in doc.element.body.iter(qn("w:hyperlink")):
        rid = h.get(qn("r:id"))
        if rid and rid in part.rels:
            tgt = part.rels[rid].target_ref
            r = h.find(qn("w:r"))
            rp = r.find(qn("w:rPr")) if r is not None else None
            link_info.setdefault(tgt, (rid, rp))
    by_i = {str(x["i"]): x["t"] for x in trans}
    paras = dict((str(n), p) for n, p in text_paragraphs(doc))
    for n, p in text_paragraphs(doc):
        key = "title" if is_title(doc, p) else str(n)
        if key not in by_i or has_drawing(p):
            continue
        base = None
        for r in p.iter(qn("w:r")):
            if r.getparent().tag == qn("w:p") and run_text(r):
                base = r.find(qn("w:rPr")); break
        if base is None:
            for r in p.iter(qn("w:r")):
                base = r.find(qn("w:rPr")); break
        base = copy.deepcopy(base) if base is not None else None
        for c in list(p):
            if c.tag in (qn("w:r"), qn("w:hyperlink"), qn("w:proofErr"), qn("w:bookmarkStart"), qn("w:bookmarkEnd")):
                p.remove(c)
        for kind, txt, extra in parse(by_i[key]):
            # 巢狀：**[[...]]** 之類不支援；粗體內的連結請拆開寫
            if kind == "link":
                if extra in link_info:
                    rid, rp = link_info[extra]
                else:
                    rid = part.relate_to(extra, RT.HYPERLINK, is_external=True)
                    rp = None
                h = OxmlElement("w:hyperlink"); h.set(qn("r:id"), rid)
                if rp is not None:
                    h.append(make_run(txt, base, rpr=rp))
                else:
                    rr = make_run(txt, base, color="0563C1")
                    u = OxmlElement("w:u"); u.set(qn("w:val"), "none"); rr.find(qn("w:rPr")).append(u)
                    h.append(rr)
                p.append(h)
            elif kind == "b":
                if txt.startswith("//") and txt.endswith("//") and len(txt) > 4:
                    p.append(make_run(txt[2:-2], base, b=True, i=True))  # **//書名//** 粗斜體
                else:
                    p.append(make_run(txt, base, b=True))
            elif kind == "i":
                p.append(make_run(txt, base, i=True))
            elif kind == "hl":
                p.append(make_run(txt, base, hl=True))
            elif kind == "color":
                p.append(make_run(txt, base, color=extra))
            else:
                p.append(make_run(txt, base))
    sha = hashlib.sha1(open(src, "rb").read()).hexdigest()
    cp = doc.core_properties
    cp.comments = f"source-sha1:{sha}"
    if title:
        cp.title = title
    else:
        for x in trans:
            if x["i"] == "title":
                cp.title = re.sub(r"(\*\*|//|==)", "", x["t"])
    cp.language = lang
    doc.save(out)


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    if cmd == "extract":
        print(json.dumps(extract(sys.argv[2]), ensure_ascii=False, indent=1))
    elif cmd == "apply":
        apply(sys.argv[2], json.load(open(sys.argv[3], encoding="utf-8")), sys.argv[4],
              lang=sys.argv[5] if len(sys.argv) > 5 else "en")
    else:
        print(__doc__)

#!/usr/bin/env python3
"""Minecraft World Wiki 建站程式。

把 content/（本機「已發布」資料夾的鏡像）裡的 Word 條目轉成百科風靜態網站，輸出到 _site/。
- 子資料夾 = 分類，會自動產生分類頁。
- 條目裡的超連結依目標狀態上色：已發布 → 藍色連結；已有條目但未發布（列在 entries.txt）→ 紅字；其他 → 一般文字。
- 文件中的圖片（旗幟、標誌、路線圖）會自動取出。
"""
import hashlib
import html
import json
import os
import re
import shutil
import subprocess
import sys
import urllib.parse
from datetime import datetime, timezone
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

ROOT = Path(__file__).resolve().parent
CONTENT = ROOT / "content"
STATIC = ROOT / "static"
OUT = ROOT / "_site"
CONFIG = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))

E = html.escape

CATEGORY_CLASS = {
    "國家與地區": "cat-region",
    "交通": "cat-transport",
    "國際組織": "cat-org",
    "宗教": "cat-religion",
    "世界總覽": "cat-history",
    "歷史": "cat-history",
    "典籍": "cat-history",
    "人物": "cat-history",
}

BULLET_RE = re.compile(r"^\s*[•●・‧▪◦]\s*")


def slug(name):
    """檔名安全的網址片段（保留中文，GitHub Pages 可正常提供）。"""
    s = re.sub(r'[\\/:*?"<>|#%]', "-", name).strip()
    return s or "entry"


def url_quote(path):
    return urllib.parse.quote(path, safe="/.-_~")


# ---------------------------------------------------------------- 收集條目

def collect():
    entries = {}  # stem -> info
    for f in sorted(CONTENT.rglob("*.docx")):
        if f.name.startswith("~$"):
            continue
        rel = f.relative_to(CONTENT)
        stem = f.stem.strip()
        if stem in entries:
            print(f"警告：條目名稱重複，略過 {rel}", file=sys.stderr)
            continue
        entries[stem] = {
            "stem": stem,
            "file": f,
            "cats": list(rel.parent.parts),
            "url": f"wiki/{url_quote(slug(stem))}.html",
            "out": f"wiki/{slug(stem)}.html",
        }
    existing = set()
    ex_file = ROOT / "entries.txt"
    if ex_file.exists():
        existing = {l.strip() for l in ex_file.read_text(encoding="utf-8").splitlines() if l.strip()}
    return entries, existing


def git_dates():
    """每個檔案最後一次提交的時間（最近更新用）。"""
    dates = {}
    try:
        out = subprocess.run(
            ["git", "log", "--name-only", "--format=@@%cI", "--", "content"],
            cwd=ROOT, capture_output=True, text=True, check=True,
            env={**os.environ, "GIT_CONFIG_PARAMETERS": "'core.quotepath=false'"},
        ).stdout
        cur = None
        for line in out.splitlines():
            if line.startswith("@@"):
                cur = line[2:]
            elif line.strip() and cur:
                dates.setdefault(line.strip(), cur)
    except Exception as e:  # 本機沒有 git 時不影響建站
        print("無法讀取 git 紀錄：", e, file=sys.stderr)
    return dates


# ---------------------------------------------------------------- docx 轉 HTML

class Converter:
    def __init__(self, entries, existing, media_dir):
        self.entries = entries
        self.existing = existing
        self.media_dir = media_dir
        self.media_dir.mkdir(parents=True, exist_ok=True)

    def link_state(self, target, text):
        stem = None
        if target:
            t = urllib.parse.unquote(target).replace("\\", "/").split("#")[0]
            name = t.rsplit("/", 1)[-1]
            if name.lower().endswith(".docx"):
                stem = name[:-5].strip()
        for cand in (stem, text.strip()):
            if cand and cand in self.entries:
                return "pub", cand
        for cand in (stem, text.strip()):
            if cand and cand in self.existing:
                return "unpub", cand
        return "none", None

    def image(self, part, rid):
        try:
            img = part.related_parts[rid]
        except KeyError:
            return None
        blob = img.blob
        ext = Path(img.partname).suffix.lower() or ".png"
        if ext in (".emf", ".wmf"):
            return None
        name = hashlib.sha1(blob).hexdigest()[:16] + ext
        dest = self.media_dir / name
        if not dest.exists():
            dest.write_bytes(blob)
        return name

    def runs_html(self, el, part, base, imgs):
        out = []
        for c in el.iterchildren():
            tag = c.tag.split("}")[1]
            if tag == "r":
                out.append(self.run_html(c, part, base, imgs))
            elif tag == "hyperlink":
                text = "".join(x.text or "" for x in c.iter(qn("w:t")))
                rid = c.get(qn("r:id"))
                target = None
                if rid and rid in part.rels:
                    target = part.rels[rid].target_ref
                inner = "".join(self.run_html(r, part, base, imgs) for r in c.iter(qn("w:r")))
                state, stem = self.link_state(target, text)
                if state == "pub":
                    out.append(f'<a href="{base}{self.entries[stem]["url"]}">{inner}</a>')
                elif state == "unpub":
                    out.append(f'<span class="unpub" title="條目尚未發布">{inner}</span>')
                elif target and re.match(r"https?://", target):
                    out.append(f'<a href="{E(target)}" rel="noopener">{inner}</a>')
                else:
                    out.append(inner)
            elif tag in ("smartTag", "ins", "fldSimple"):
                out.append(self.runs_html(c, part, base, imgs))
        return "".join(out)

    def run_html(self, r, part, base, imgs):
        parts = []
        rpr = r.find(qn("w:rPr"))
        bold = rpr is not None and rpr.find(qn("w:b")) is not None and rpr.find(qn("w:b")).get(qn("w:val")) not in ("0", "false")
        ital = rpr is not None and rpr.find(qn("w:i")) is not None and rpr.find(qn("w:i")).get(qn("w:val")) not in ("0", "false")
        for c in r.iterchildren():
            tag = c.tag.split("}")[1]
            if tag == "t":
                parts.append(E(c.text or ""))
            elif tag == "tab":
                parts.append("　")
            elif tag == "br":
                parts.append("<br>")
            elif tag == "drawing":
                for blip in c.iter(qn("a:blip")):
                    name = self.image(part, blip.get(qn("r:embed")))
                    if name:
                        imgs.append(name)
        s = "".join(parts)
        if s and bold:
            s = f"<b>{s}</b>"
        if s and ital:
            s = f"<i>{s}</i>"
        return s

    def convert(self, info, base):
        doc = Document(str(info["file"]))
        part = doc.part
        body = doc.element.body
        blocks = []  # (kind, data)
        infobox = None
        seen_heading = False
        for el in body.iterchildren():
            tag = el.tag.split("}")[1]
            if tag == "p":
                p = Paragraph(el, doc)
                style = (p.style.name if p.style is not None else "") or ""
                imgs = []
                h = self.runs_html(el, part, base, imgs).strip()
                text = p.text.strip()
                if style in ("Title", "標題") or style.lower() == "title":
                    continue
                if style.startswith("Heading") or style.startswith("標題"):
                    lvl = 3 if ("3" in style or "4" in style) else 2
                    if text:
                        blocks.append(("h", (lvl, text)))
                        seen_heading = True
                    continue
                for name in imgs:
                    blocks.append(("img", name))
                if not text:
                    continue
                if BULLET_RE.match(text):
                    h = BULLET_RE.sub("", re.sub(r"^(\s|&nbsp;)*", "", h), count=1)
                    h = re.sub(r"^(<[^>]+>)*\s*[•●・‧▪◦]\s*", lambda m: (m.group(1) or ""), h)
                    blocks.append(("li", h.strip()))
                else:
                    blocks.append(("p", h))
            elif tag == "tbl":
                rows = []
                for tr in el.iter(qn("w:tr")):
                    cells = []
                    for tc in tr.findall(qn("w:tc")):
                        imgs = []
                        parts = []
                        for pe in tc.iter(qn("w:p")):
                            ph = self.runs_html(pe, part, base, imgs).strip()
                            if ph:
                                parts.append(ph)
                        cells.append({"html": "<br>".join(parts), "imgs": imgs,
                                      "text": "".join(x.text or "" for x in tc.iter(qn("w:t"))).strip()})
                    rows.append(cells)
                if infobox is None and not seen_heading and not any(k in ("p", "li") for k, _ in blocks):
                    infobox = rows
                else:
                    blocks.append(("table", rows))
        return blocks, infobox


# ---------------------------------------------------------------- 版面

def cat_class(cats):
    return CATEGORY_CLASS.get(cats[0], "cat-history") if cats else "cat-history"


def page(title, body, base, *, sidebar="", description="", page_class=""):
    site = CONFIG["site_title"]
    full = site if title == site else f"{title} - {site}"
    return f"""<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{E(full)}</title>
<meta name="description" content="{E(description or CONFIG['tagline'])}">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&family=Noto+Serif+TC:wght@500;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="{base}static/style.css">
<link rel="icon" href="{base}static/favicon.svg" type="image/svg+xml">
</head>
<body class="{page_class}" data-base="{base}">
<a class="skip" href="#main">跳到內容</a>
<header class="topbar">
  <div class="wrap topbar-in">
    <a class="brand" href="{base}index.html">
      <span class="mark" aria-hidden="true">世</span>
      <span class="brand-text"><span class="brand-name">{E(site)}</span><span class="brand-sub">{E(CONFIG['short_tagline'])}</span></span>
    </a>
    <form class="search" role="search" action="{base}search.html" method="get">
      <label for="q" class="sr">搜尋條目</label>
      <input id="q" name="q" type="search" placeholder="搜尋國家、城市、事件…" autocomplete="off">
      <button type="submit">搜尋</button>
      <ul class="suggest" hidden></ul>
    </form>
    <nav class="toplinks" aria-label="工具">
      <a href="{base}random.html">隨機條目</a>
      <a href="{base}recent.html">最近更新</a>
      <a href="{base}about.html">關於本站</a>
    </nav>
  </div>
</header>
<div class="wrap layout">
  <aside class="side">{sidebar}</aside>
  <main id="main" class="main">{body}</main>
</div>
<footer class="foot"><div class="wrap">{E(CONFIG['footer'])}</div></footer>
<script src="{base}static/site.js"></script>
</body>
</html>
"""


def nav_sidebar(tree, base):
    items = "".join(
        f'<a href="{base}category/{url_quote(slug(c))}/index.html">{E(c)}</a>' for c in tree.get("dirs", {})
    )
    regions = tree.get("dirs", {}).get("國家與地區", {}).get("dirs", {})
    reg = "".join(
        f'<a href="{base}category/{url_quote(slug("國家與地區"))}/{url_quote(slug(r))}/index.html">{E(r)}</a>'
        for r in regions
    )
    html_ = f'<div class="side-group"><div class="side-label">導覽</div><a href="{base}index.html">首頁</a>{items}</div>'
    if reg:
        html_ += f'<div class="side-group"><div class="side-label">大陸與海洋</div>{reg}</div>'
    return html_


def build_tree(entries):
    tree = {"dirs": {}, "entries": []}
    for d in sorted(p for p in CONTENT.rglob("*") if p.is_dir()):
        node = tree
        for part in d.relative_to(CONTENT).parts:
            node = node["dirs"].setdefault(part, {"dirs": {}, "entries": []})
    for info in entries.values():
        node = tree
        for part in info["cats"]:
            node = node["dirs"].setdefault(part, {"dirs": {}, "entries": []})
        node["entries"].append(info["stem"])
    return tree


def count_entries(node):
    return len(node["entries"]) + sum(count_entries(n) for n in node["dirs"].values())


def cat_url(cats):
    return "category/" + "/".join(url_quote(slug(c)) for c in cats) + "/index.html"


def render_blocks(blocks, base):
    out = []
    i = 0
    toc = []
    hid = 0
    while i < len(blocks):
        kind, data = blocks[i]
        if kind == "h":
            lvl, text = data
            hid += 1
            anchor = f"s{hid}"
            toc.append((lvl, text, anchor))
            out.append(f'<h{lvl} id="{anchor}">{E(text)}</h{lvl}>')
        elif kind == "p":
            out.append(f"<p>{data}</p>")
        elif kind == "li":
            items = []
            while i < len(blocks) and blocks[i][0] == "li":
                items.append(f"<li>{blocks[i][1]}</li>")
                i += 1
            cls = ' class="cols"' if toc and toc[-1][1] == "參見" else ""
            out.append(f"<ul{cls}>{''.join(items)}</ul>")
            continue
        elif kind == "img":
            # 圖片後面緊接的短段落視為圖說
            cap = ""
            if i + 1 < len(blocks) and blocks[i + 1][0] == "p" and len(re.sub("<[^>]+>", "", blocks[i + 1][1])) <= 40:
                cap = blocks[i + 1][1]
                i += 1
            out.append(
                f'<figure class="fig"><a href="{base}media/{data}"><img src="{base}media/{data}" alt="{E(re.sub("<[^>]+>", "", cap))}" loading="lazy"></a>'
                + (f"<figcaption>{cap}</figcaption>" if cap else "") + "</figure>"
            )
        elif kind == "table":
            rows = []
            for r in data:
                cells = ""
                for c in r:
                    pics = "".join('<img src="' + base + "media/" + n + '" alt="">' for n in c["imgs"])
                    cells += "<td>" + c["html"] + pics + "</td>"
                rows.append(f"<tr>{cells}</tr>")
            out.append(f'<div class="tablewrap"><table class="wikitable"><tbody>{"".join(rows)}</tbody></table></div>')
        i += 1
    return "\n".join(out), toc


def render_infobox(rows, title, cats, base):
    if not rows:
        return ""
    cap = rows[0][0]["text"] if rows[0] and len(rows[0]) == 1 else title
    start = 1 if rows[0] and len(rows[0]) == 1 else 0
    eng = ""
    trs = []
    for r in rows[start:]:
        imgs = [n for c in r for n in c["imgs"]]
        if imgs:
            labels = " ".join(c["text"] for c in r).split()
            figs = []
            for k, n in enumerate(imgs):
                lab = labels[k] if k < len(labels) else ""
                figs.append(
                    f'<figure><img src="{base}media/{n}" alt="{E(cap + lab)}"><figcaption>{E(lab)}</figcaption></figure>'
                )
            trs.append(f'<tr><td colspan="2" class="ib-img"><div class="ib-figs">{"".join(figs)}</div></td></tr>')
        elif len(r) >= 2:
            k, v = r[0], r[1]
            if k["text"] in ("英文名稱", "英文站名") and not eng:
                eng = v["text"]
            trs.append(f'<tr><th scope="row">{k["html"]}</th><td>{v["html"]}</td></tr>')
        elif len(r) == 1 and r[0]["text"]:
            trs.append(f'<tr><th colspan="2" class="ib-sub">{r[0]["html"]}</th></tr>')
    sub = f"<span>{E(eng)}</span>" if eng else ""
    return (
        f'<table class="infobox"><caption class="{cat_class(cats)}">{E(cap)}{sub}</caption>'
        f'<tbody>{"".join(trs)}</tbody></table>'
    )


def first_text(blocks, limit=120):
    for k, d in blocks:
        if k == "p":
            t = re.sub("<[^>]+>", "", d)
            t = html.unescape(t).strip()
            if t:
                return t if len(t) <= limit else t[:limit].rstrip() + "……"
    return ""


def breadcrumb(cats, base, current=None):
    parts = []
    for i in range(len(cats)):
        parts.append(f'<a href="{base}{cat_url(cats[:i + 1])}">{E(cats[i])}</a>')
    if current:
        parts.append(E(current))
    return '<nav class="crumbs" aria-label="位置">' + "　›　".join(parts) + "</nav>" if parts else ""


# ---------------------------------------------------------------- 主程式

def main():
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir()
    shutil.copytree(STATIC, OUT / "static")
    (OUT / ".nojekyll").write_text("")

    entries, existing = collect()
    tree = build_tree(entries)
    dates = git_dates()
    conv = Converter(entries, existing, OUT / "media")
    index = []

    # 條目頁
    for stem, info in entries.items():
        base = "../"
        try:
            blocks, infobox = conv.convert(info, base)
        except Exception as e:
            print(f"無法轉換 {info['file'].name}：{e}", file=sys.stderr)
            continue
        info["summary"] = first_text(blocks)
        rel = info["file"].relative_to(CONTENT).as_posix()
        info["updated"] = dates.get(f"content/{rel}", "")
        body_html, toc = render_blocks(blocks, base)
        ib = render_infobox(infobox, stem, info["cats"], base)
        info["thumb"] = None
        if infobox:
            for r in infobox:
                for c in r:
                    if c["imgs"] and not info["thumb"]:
                        info["thumb"] = c["imgs"][0]
        toc_html = ""
        if len(toc) >= 2:
            toc_html = '<div class="side-group"><div class="side-label">目錄</div><a href="#top">（頂部）</a>' + "".join(
                f'<a class="toc-l{l}" href="#{a}">{E(t)}</a>' for l, t, a in toc
            ) + "</div>"
        cats_bar = ""
        if info["cats"]:
            cats_bar = '<div class="catbar">分類：' + "｜".join(
                f'<a href="{base}{cat_url(info["cats"][:i + 1])}">{E(info["cats"][i])}</a>' for i in range(len(info["cats"]))
            ) + "</div>"
        upd = ""
        if info["updated"]:
            upd = f'<p class="updated">本條目最後更新於 {info["updated"][:10]}。</p>'
        body = (
            f'<article class="article" id="top">{breadcrumb(info["cats"][:-1] if info["cats"] and info["cats"][-1] == stem else info["cats"], base, stem)}'
            f'<h1 class="title">{E(stem)}</h1>'
            f'<div class="article-body">{ib}{body_html}</div>{cats_bar}{upd}</article>'
        )
        out = OUT / info["out"]
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(page(stem, body, base, sidebar=toc_html or nav_sidebar(tree, base), description=info["summary"]), encoding="utf-8")
        index.append({"t": stem, "u": info["url"], "s": info["summary"], "c": " › ".join(info["cats"])})

    # 分類頁
    def walk(node, cats):
        if cats:
            depth = len(cats) + 1
            base = "../" * depth
            main_entry = cats[-1] if cats[-1] in entries else None
            parts = [f'<article class="article">{breadcrumb(cats[:-1], base, cats[-1])}<h1 class="title">{E(cats[-1])}</h1>']
            if main_entry:
                parts.append(
                    f'<p class="hatnote">主條目：<a href="{base}{entries[main_entry]["url"]}">{E(main_entry)}</a></p>'
                )
            if node["dirs"]:
                parts.append("<h2>子分類</h2><ul class=\"cols\">" + "".join(
                    f'<li><a href="{base}{cat_url(cats + [d])}">{E(d)}</a>（{count_entries(n)}）</li>'
                    for d, n in node["dirs"].items()
                ) + "</ul>")
            listed = [e for e in sorted(node["entries"]) if e != main_entry]
            if listed:
                parts.append(f"<h2>此分類中的條目（{len(listed)}）</h2><ul class=\"cols\">" + "".join(
                    f'<li><a href="{base}{entries[e]["url"]}">{E(e)}</a></li>' for e in listed
                ) + "</ul>")
            if not node["dirs"] and not listed and not main_entry:
                parts.append('<p class="empty">此分類目前還沒有已發布的條目。</p>')
            parts.append("</article>")
            out = OUT / "category" / Path(*[slug(c) for c in cats]) / "index.html"
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_text(page(cats[-1], "".join(parts), base, sidebar=nav_sidebar(tree, base)), encoding="utf-8")
        for d, n in node["dirs"].items():
            walk(n, cats + [d])

    walk(tree, [])

    # 首頁
    base = ""
    featured = None
    for t in CONFIG.get("featured", []):
        if t in entries and entries[t].get("summary") is not None:
            featured = entries[t]
            break
    dated = sorted((e for e in entries.values() if e.get("summary")), key=lambda e: e.get("updated", ""), reverse=True)
    if not featured and dated:
        featured = dated[0]
    feat_html = '<p class="empty">網站剛剛開張，第一批條目即將發布。</p>'
    if featured:
        thumb = f'<img class="feat-img" src="media/{featured["thumb"]}" alt="">' if featured.get("thumb") else ""
        feat_html = (
            f'<div class="feat">{thumb}<p><a href="{featured["url"]}"><b>{E(featured["stem"])}</b></a>'
            f'{E(featured["summary"][len(featured["stem"]):] if featured["summary"].startswith(featured["stem"]) else "：" + featured["summary"])}'
            f'<a href="{featured["url"]}">（閱讀全文）</a></p></div>'
        )
    dyk = CONFIG.get("did_you_know", [])
    dyk_items = []
    for item in dyk:
        def rep(m):
            name = m.group(1)
            if name in entries:
                return f'<a href="{entries[name]["url"]}">{E(name)}</a>'
            return E(name)
        dyk_items.append("<li>……" + re.sub(r"\[\[(.+?)\]\]", rep, E(item).replace("&#x27;", "'")) + "</li>")
    dyk_html = f'<ul class="dyk">{"".join(dyk_items)}</ul>' if dyk_items else '<p class="empty">敬請期待。</p>'

    regions = tree["dirs"].get("國家與地區", {"dirs": {}})
    reg_rows = []
    for r, n in regions["dirs"].items():
        links = []
        for c, cn in n["dirs"].items():
            if c in entries:
                links.append(f'<a href="{entries[c]["url"]}">{E(c)}</a>')
            elif count_entries(cn):
                links.append(f'<a href="{cat_url(["國家與地區", r, c])}">{E(c)}</a>')
            else:
                links.append(f'<span class="plain">{E(c)}</span>')
        for e in sorted(n["entries"]):
            if e not in n["dirs"]:
                links.append(f'<a href="{entries[e]["url"]}">{E(e)}</a>')
        rh = f'<a href="{cat_url(["國家與地區", r])}">{E(r)}</a>'
        reg_rows.append(f'<tr><th scope="row">{rh}</th><td>{"・".join(links) or "—"}</td></tr>')
    reg_html = (
        f'<section class="box"><h2 class="box-h">國家與地區</h2><div class="tablewrap"><table class="regions"><tbody>{"".join(reg_rows)}</tbody></table></div></section>'
        if reg_rows else ""
    )
    recent = [e for e in dated[:8]]
    recent_html = ""
    if recent:
        recent_html = '<section class="box"><h2 class="box-h">最近更新</h2><ul class="recent">' + "".join(
            f'<li><a href="{e["url"]}">{E(e["stem"])}</a><span>{e.get("updated", "")[:10]}</span></li>' for e in recent
        ) + "</ul></section>"
    total = len(index)
    home = f"""
<section class="welcome">
  <h1>歡迎來到 {E(CONFIG['site_title'])}</h1>
  <p>{E(CONFIG['tagline'])}</p>
  <p class="count">目前收錄 <b>{total}</b> 個條目</p>
</section>
<figure class="box mapfig">
  <a href="static/map-large.jpg"><img src="static/map.jpg" alt="世界地圖與各國勢力範圍"></a>
  <figcaption>{E(CONFIG['map_caption'])}點擊地圖可放大檢視。</figcaption>
</figure>
<div class="grid2">
  <section class="box"><h2 class="box-h head-featured">特色條目</h2><div class="box-b">{feat_html}</div></section>
  <section class="box"><h2 class="box-h head-dyk">你知道嗎？</h2><div class="box-b">{dyk_html}</div></section>
</div>
{reg_html}
{recent_html}
"""
    (OUT / "index.html").write_text(page(CONFIG["site_title"], home, base, sidebar=nav_sidebar(tree, base), page_class="home"), encoding="utf-8")

    # 最近更新
    rec_all = "".join(
        f'<li><a href="{e["url"]}">{E(e["stem"])}</a><span>{e.get("updated", "")[:10]}</span></li>' for e in dated[:100]
    ) or '<li class="empty">目前沒有條目。</li>'
    (OUT / "recent.html").write_text(page("最近更新", f'<article class="article"><h1 class="title">最近更新</h1><ul class="recent">{rec_all}</ul></article>', base, sidebar=nav_sidebar(tree, base)), encoding="utf-8")

    # 搜尋、隨機、關於、404
    (OUT / "search-index.json").write_text(json.dumps(index, ensure_ascii=False), encoding="utf-8")
    (OUT / "search.html").write_text(page("搜尋", '<article class="article"><h1 class="title">搜尋結果</h1><p id="search-status" class="muted">搜尋中……</p><ul id="search-results" class="results"></ul></article>', base, sidebar=nav_sidebar(tree, base), page_class="search-page"), encoding="utf-8")
    (OUT / "random.html").write_text(page("隨機條目", '<article class="article"><h1 class="title">隨機條目</h1><p id="random-status" class="muted">正在帶你前往隨機條目……</p></article>', base, sidebar=nav_sidebar(tree, base), page_class="random-page"), encoding="utf-8")
    about = "".join(f"<p>{E(p)}</p>" for p in CONFIG["about"])
    (OUT / "about.html").write_text(page("關於本站", f'<article class="article"><h1 class="title">關於本站</h1>{about}</article>', base, sidebar=nav_sidebar(tree, base)), encoding="utf-8")
    (OUT / "404.html").write_text(page("找不到頁面", '<article class="article"><h1 class="title">找不到這個頁面</h1><p>這個條目可能尚未發布，或網址有誤。可以試試上方的搜尋。</p></article>', "/" + CONFIG.get("base_path", "").strip("/") + "/" if CONFIG.get("base_path") else "/", sidebar=""), encoding="utf-8")

    print(f"完成：{total} 個條目，輸出到 {OUT}")


if __name__ == "__main__":
    main()

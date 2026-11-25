# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import datetime as dt
import html
import json
import os
import re
import shutil
import sys
import threading
import traceback
import webbrowser
from pathlib import Path

try:
    import customtkinter as ctk
except Exception:  # CustomTkinter is not needed for CLI maintenance commands.
    ctk = None


APP_NAME = "YUKIZ BLOG 投稿DAKE"
DEFAULT_SITE_PATH = Path("C:/Users/yukiz/devlop/yukizblog-site")
DEFAULT_BASE_URL = os.environ.get("YUKIZBLOG_BASE_URL", "https://yukizblog.com/")
TOKYO = dt.timezone(dt.timedelta(hours=9), "JST")

STATUS_READY = ""
STATUS_PREVIEW = "置く前に、少し見る。"
STATUS_DONE = "静かに置きました。"
STATUS_ERROR = "静かに置けませんでした。"
STATUS_WORKING = ("置いています.", "置いています..", "置いています...")

POST_FILE_RE = re.compile(r"^(\d{8})-(\d{3})\.html$")
PARTICLE_RE = re.compile(
    r"(?:について|として|から|まで|より|だけ|ほど|など|とか|でも|では|には|へは|とは|"
    r"の|を|が|は|に|へ|で|と|も|や)"
)
SPLIT_RE = re.compile(r"[、。！？!?・／/\\\s　（）()「」『』［］\[\]【】:：;；,]+|" + PARTICLE_RE.pattern)

BANNED_TITLE_WORDS = (
    "完全版",
    "まとめ",
    "おすすめ",
    "解説",
    "方法",
    "とは",
    "した話",
    "する話",
)
STOP_WORDS = {
    "これ",
    "それ",
    "あれ",
    "ここ",
    "そこ",
    "どこ",
    "こと",
    "もの",
    "ため",
    "よう",
    "さん",
    "今日",
    "昨日",
    "明日",
    "自分",
    "少し",
    "とても",
    "かなり",
    "そして",
    "しかし",
    "だから",
}
ONE_CHAR_WORDS = {"空", "夜", "朝", "雨", "風", "青", "白", "黒", "光", "影", "火", "水", "机"}
IMPRESSION_WORDS = (
    "余白",
    "静か",
    "空",
    "夜",
    "朝",
    "青",
    "白",
    "影",
    "光",
    "風",
    "雨",
    "理由",
    "整理",
    "流れ",
    "温度",
    "冷たさ",
    "あたたかさ",
    "熾火",
    "薄い",
    "広い",
    "整う",
)


def now_tokyo() -> dt.datetime:
    return dt.datetime.now(TOKYO)


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = get_app_dir()


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig")


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\n") as handle:
        handle.write(text)


def json_dump(data: object) -> str:
    return json.dumps(data, ensure_ascii=False, indent=2) + "\n"


def normalize_body(text: str) -> str:
    text = (text or "").replace("\ufeff", "").replace("\r\n", "\n").replace("\r", "\n")
    lines: list[str] = []
    last_was_blank = False
    for raw_line in text.split("\n"):
        line = re.sub(r"[ \t　]+", " ", raw_line).strip()
        if not line:
            if lines and not last_was_blank:
                lines.append("")
            last_was_blank = True
            continue
        lines.append(line)
        last_was_blank = False
    while lines and not lines[-1]:
        lines.pop()
    return "\n".join(lines).strip()


def body_paragraphs(body_text: str) -> list[str]:
    return [line.strip() for line in normalize_body(body_text).split("\n") if line.strip()]


def plain_body(body_text: str) -> str:
    return "".join(body_paragraphs(body_text))


def first_meaningful_sentence(body_text: str) -> str:
    normalized = normalize_body(body_text)
    for part in re.split(r"[。！？!?]+|\n+", normalized):
        sentence = part.strip(" 　、。")
        if len(sentence) >= 2:
            return sentence
    return plain_body(body_text)


def strip_html_tags(value: str) -> str:
    return re.sub(r"<[^>]+>", "", value or "")


def clean_word(word: str) -> str:
    word = html.unescape(word or "")
    word = re.sub(r"[「」『』（）()\[\]【】.,、。！？!?;；:：]+", "", word)
    word = word.strip(" 　/／・-")
    word = re.sub(r"^[のをがはにへでともや]+", "", word)
    word = re.sub(r"[のをがはにへでともや]+$", "", word)
    return word.strip()


def is_word_candidate(word: str) -> bool:
    if not word:
        return False
    if PARTICLE_RE.search(word):
        return False
    if any(banned in word for banned in BANNED_TITLE_WORDS):
        return False
    if word in STOP_WORDS:
        return False
    if len(word) == 1 and word not in ONE_CHAR_WORDS:
        return False
    if len(word) > 14:
        return False
    if re.fullmatch(r"[ぁ-んー]+", word) and word not in IMPRESSION_WORDS:
        return False
    return bool(re.search(r"[一-龥々〆ヵヶぁ-んァ-ヴーA-Za-z0-9]", word))


def candidate_words(text: str) -> list[str]:
    words: list[str] = []
    pieces = [piece for piece in SPLIT_RE.split(text) if piece]
    tokens = re.findall(
        r"[A-Za-z0-9][A-Za-z0-9_-]{1,}|[ァ-ヴー]{2,}|[一-龥々〆ヵヶ][一-龥々〆ヵヶぁ-んァ-ヴーA-Za-z0-9]{0,13}",
        text,
    )
    for raw in pieces + tokens:
        word = clean_word(raw)
        if is_word_candidate(word) and word not in words:
            words.append(word)
    return words


def pick_impression_word(text: str) -> str | None:
    found: list[tuple[int, str]] = []
    for word in IMPRESSION_WORDS:
        index = text.find(word)
        if index >= 0:
            found.append((index, word))
    if not found:
        return None
    return sorted(found, key=lambda item: item[0])[0][1]


def clip_chars(text: str, min_len: int, max_len: int, ellipsis: bool = False) -> str:
    text = re.sub(r"\s+", "", text or "").strip()
    if len(text) <= max_len:
        return text
    cut = max_len
    for mark in ("。", "！", "？", ".", "!", "?"):
        pos = text.rfind(mark, min_len, max_len + 1)
        if pos >= min_len:
            cut = pos + 1
            break
    result = text[:cut].rstrip("、, ")
    if ellipsis and cut < len(text) and not result.endswith(("。", "！", "？", ".", "!", "?", "…")):
        result += "…"
    return result


def generate_internal_title(body_text: str, date_text: str | None = None) -> str:
    body = normalize_body(body_text)
    sentence = first_meaningful_sentence(body)
    words = candidate_words(sentence)

    if len(words) < 2:
        for word in candidate_words(body):
            if word not in words:
                words.append(word)
            if len(words) >= 2:
                break

    impression = pick_impression_word(body)
    if impression and impression not in words and len(words) < 3:
        words.append(impression)
    elif len(words) == 1:
        words.append("余白")
    elif len(words) == 2 and not any(word in IMPRESSION_WORDS for word in words):
        words.append("静か")

    if words:
        title = " / ".join(words[:3])
        if not any(banned in title for banned in BANNED_TITLE_WORDS):
            return title

    fallback = clip_chars(sentence or body, 20, 35, ellipsis=False)
    if fallback:
        return fallback
    return date_text or now_tokyo().strftime("%Y.%m.%d")


def title_needs_regeneration(title: str) -> bool:
    if not title:
        return True
    if any(banned in title for banned in BANNED_TITLE_WORDS):
        return True
    parts = [clean_word(part) for part in title.split("/") if clean_word(part)]
    if not parts or len(parts) > 3:
        return True
    return any(PARTICLE_RE.search(part) for part in parts)


def generate_meta_description(body_text: str) -> str:
    text = plain_body(body_text)
    if not text:
        return "YUKIZ BLOG に静かに置かれた言葉。"
    return clip_chars(text, 60, 90, ellipsis=True)


def generate_excerpt(body_text: str) -> str:
    text = plain_body(body_text)
    return clip_chars(text, 24, 72, ellipsis=True)


def generate_air_tags(body_text: str, internal_title: str) -> list[str]:
    tags: list[str] = []
    for part in internal_title.split("/"):
        word = clean_word(part)
        if word and word not in tags:
            tags.append(word)
    impression = pick_impression_word(body_text)
    if impression and impression not in tags:
        tags.append(impression)
    return tags[:5]


def display_date(date_text: str) -> str:
    try:
        return dt.date.fromisoformat(date_text).strftime("%Y.%m.%d")
    except ValueError:
        return date_text.replace("-", ".")


def join_url(base_url: str, rel_url: str) -> str:
    base = (base_url or DEFAULT_BASE_URL).rstrip("/") + "/"
    return base + rel_url.lstrip("/")


def post_filename(post: dict) -> str:
    return Path(str(post.get("file") or post.get("url") or "")).name


def post_url_from_file(filename: str) -> str:
    return f"posts/{filename}"


def post_sort_key(post: dict) -> tuple[str, int, str, str]:
    filename = post_filename(post)
    match = POST_FILE_RE.match(filename)
    sequence = int(match.group(2)) if match else 0
    return (
        str(post.get("date", "")),
        sequence,
        str(post.get("created_at", "")),
        filename,
    )


def next_sequence(posts: list[dict], posts_dir: Path, date_code: str) -> int:
    sequence = 0
    for post in posts:
        filename = post_filename(post)
        match = POST_FILE_RE.match(filename)
        if match and match.group(1) == date_code:
            sequence = max(sequence, int(match.group(2)))
    if posts_dir.exists():
        for path in posts_dir.glob(f"{date_code}-*.html"):
            match = POST_FILE_RE.match(path.name)
            if match:
                sequence = max(sequence, int(match.group(2)))
    return sequence + 1


def parse_existing_post_file(path: Path, site_path: Path) -> dict | None:
    if path.name == "index.html" or path.name.startswith("_"):
        return None
    text = read_text(path)
    date_match = re.search(r"<time[^>]*datetime=[\"']([^\"']+)[\"'][^>]*>", text, re.I)
    title_match = re.search(r"<title>(.*?)</title>", text, re.I | re.S)
    desc_match = re.search(r"<meta\s+name=[\"']description[\"']\s+content=[\"'](.*?)[\"']", text, re.I | re.S)
    body_match = re.search(r"<div\s+class=[\"']entry-body[\"']\s*>(.*?)</div>", text, re.I | re.S)
    paragraphs: list[str] = []
    if body_match:
        for paragraph in re.findall(r"<p[^>]*>(.*?)</p>", body_match.group(1), re.I | re.S):
            plain = html.unescape(strip_html_tags(paragraph)).strip()
            if plain:
                paragraphs.append(plain)
    if not paragraphs:
        return None

    stat = path.stat()
    stamp = dt.datetime.fromtimestamp(stat.st_mtime, TOKYO).isoformat(timespec="seconds")
    date_text = date_match.group(1)[:10] if date_match else dt.datetime.fromtimestamp(stat.st_mtime, TOKYO).date().isoformat()
    body = normalize_body("\n".join(paragraphs))
    internal_title = generate_internal_title(body, display_date(date_text))
    rel = path.relative_to(site_path).as_posix()
    return {
        "id": path.stem,
        "date": date_text,
        "file": rel,
        "url": rel,
        "internal_title": internal_title,
        "page_title": f"{internal_title}｜YUKIZ BLOG",
        "meta_description": html.unescape(desc_match.group(1)).strip() if desc_match else generate_meta_description(body),
        "og_title": html.unescape(title_match.group(1)).strip() if title_match else f"{internal_title}｜YUKIZ BLOG",
        "og_description": html.unescape(desc_match.group(1)).strip() if desc_match else generate_meta_description(body),
        "excerpt": generate_excerpt(body),
        "body": body,
        "air_tags": generate_air_tags(body, internal_title),
        "prev": None,
        "next": None,
        "created_at": stamp,
        "updated_at": stamp,
    }


def normalize_post(post: dict) -> dict:
    body = normalize_body(str(post.get("body", "")))
    date_text = str(post.get("date") or now_tokyo().date().isoformat())
    filename = post_filename(post) or f"{date_text.replace('-', '')}-001.html"
    rel = post.get("file") or post.get("url") or post_url_from_file(filename)
    rel = str(rel).replace("\\", "/")
    internal_title = str(post.get("internal_title") or "")
    regenerated_title = False
    if not internal_title or title_needs_regeneration(internal_title):
        internal_title = generate_internal_title(body, display_date(date_text))
        regenerated_title = True
    page_title = f"{internal_title}｜YUKIZ BLOG"
    meta_description = str(post.get("meta_description") or generate_meta_description(body))
    created_at = str(post.get("created_at") or now_tokyo().isoformat(timespec="seconds"))
    updated_at = str(post.get("updated_at") or created_at)
    if regenerated_title:
        updated_at = now_tokyo().isoformat(timespec="seconds")
    return {
        "id": str(post.get("id") or Path(filename).stem),
        "date": date_text,
        "file": rel,
        "url": str(post.get("url") or rel).replace("\\", "/"),
        "internal_title": internal_title,
        "page_title": page_title,
        "meta_description": meta_description,
        "og_title": page_title,
        "og_description": meta_description,
        "excerpt": str(post.get("excerpt") or generate_excerpt(body)),
        "body": body,
        "air_tags": generate_air_tags(body, internal_title),
        "prev": post.get("prev"),
        "next": post.get("next"),
        "created_at": created_at,
        "updated_at": updated_at,
    }


def load_posts(site_path: Path) -> list[dict]:
    data_path = site_path / "data" / "posts.json"
    posts: list[dict] = []
    if data_path.exists():
        raw = json.loads(read_text(data_path))
        if isinstance(raw, list):
            posts = [normalize_post(item) for item in raw if isinstance(item, dict)]
        elif isinstance(raw, dict) and isinstance(raw.get("posts"), list):
            posts = [normalize_post(item) for item in raw["posts"] if isinstance(item, dict)]

    known_urls = {str(post.get("url")) for post in posts}
    posts_dir = site_path / "posts"
    if posts_dir.exists():
        for path in sorted(posts_dir.glob("*.html")):
            parsed = parse_existing_post_file(path, site_path)
            if parsed and parsed["url"] not in known_urls:
                posts.append(normalize_post(parsed))
                known_urls.add(parsed["url"])
    return posts


def recompute_navigation(posts: list[dict], updated_at: str | None = None) -> list[dict]:
    normalized = [normalize_post(post) for post in posts]
    chronological = sorted(normalized, key=post_sort_key)
    stamp = updated_at or now_tokyo().isoformat(timespec="seconds")
    for index, post in enumerate(chronological):
        prev_url = chronological[index - 1]["url"] if index > 0 else None
        next_url = chronological[index + 1]["url"] if index < len(chronological) - 1 else None
        if post.get("prev") != prev_url or post.get("next") != next_url:
            post["updated_at"] = stamp
        post["prev"] = prev_url
        post["next"] = next_url
    return sorted(chronological, key=post_sort_key, reverse=True)


def render_paragraphs(body_text: str) -> str:
    lines = []
    for paragraph in body_paragraphs(body_text):
        lines.append(f"        <p>{html.escape(paragraph)}</p>")
    return "\n".join(lines)


def nav_href(url: str | None, site_path: Path | None = None, preview: bool = False) -> str | None:
    if not url:
        return None
    if preview and site_path:
        return (site_path / url).resolve().as_uri()
    return post_filename({"url": url})


def nav_item(label: str, url: str | None, site_path: Path | None = None, preview: bool = False) -> str:
    href = nav_href(url, site_path, preview)
    if not href:
        return f'      <span class="muted">{html.escape(label)}</span>'
    return f'      <a href="{html.escape(href, quote=True)}">{html.escape(label)}</a>'


def render_article_html(
    post: dict,
    *,
    stylesheet_href: str = "../style.css",
    script_src: str | None = "../main.js",
    home_href: str = "../",
    list_href: str = "./",
    base_url: str = DEFAULT_BASE_URL,
    site_path: Path | None = None,
    preview: bool = False,
) -> str:
    page_title = html.escape(str(post["page_title"]), quote=True)
    description = html.escape(str(post["meta_description"]), quote=True)
    og_title = html.escape(str(post["og_title"]), quote=True)
    og_description = html.escape(str(post["og_description"]), quote=True)
    canonical = html.escape(join_url(base_url, str(post["url"])), quote=True)
    script = f'\n  <script src="{html.escape(script_src, quote=True)}"></script>' if script_src else ""
    return f"""<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{page_title}</title>
  <meta name="description" content="{description}">
  <meta property="og:type" content="article">
  <meta property="og:site_name" content="YUKIZ BLOG">
  <meta property="og:title" content="{og_title}">
  <meta property="og:description" content="{og_description}">
  <link rel="canonical" href="{canonical}">
  <link rel="stylesheet" href="{html.escape(stylesheet_href, quote=True)}">
</head>
<body>
  <main class="site-place" data-fade>
    <header class="site-header">
      <a class="site-name" href="{html.escape(home_href, quote=True)}">YUKIZ BLOG</a>
    </header>

    <article class="entry">
      <time class="entry-date" datetime="{html.escape(str(post["date"]), quote=True)}">{html.escape(display_date(str(post["date"])))}</time>

      <div class="entry-body">
{render_paragraphs(str(post["body"]))}
      </div>
    </article>

    <nav class="page-nav" aria-label="置かれた言葉を巡る">
{nav_item("前へ", post.get("prev"), site_path, preview)}
      <a href="{html.escape(list_href, quote=True)}">巡る</a>
{nav_item("次へ", post.get("next"), site_path, preview)}
    </nav>

    <footer class="site-footer">Yukihiko Kikuta</footer>
  </main>{script}
</body>
</html>
"""


def fragment_lines(post: dict) -> list[str]:
    paragraphs = body_paragraphs(str(post.get("body", "")))
    if not paragraphs:
        paragraphs = [str(post.get("excerpt", ""))]
    return [clip_chars(paragraph, 24, 54, ellipsis=True) for paragraph in paragraphs[:2] if paragraph][:2]


def render_posts_index_html(posts: list[dict]) -> str:
    fragments: list[str] = []
    for post in sorted(posts, key=post_sort_key, reverse=True):
        lines = "\n".join(f"        <p>{html.escape(line)}</p>" for line in fragment_lines(post))
        fragments.append(
            f"""      <a class="fragment" href="{html.escape(post_filename(post), quote=True)}">
        <time datetime="{html.escape(str(post["date"]), quote=True)}">{html.escape(display_date(str(post["date"])))}</time>
{lines}
      </a>"""
        )
    joined = "\n\n".join(fragments)
    return f"""<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>reasons｜YUKIZ BLOG</title>
  <meta name="description" content="静かな断片を置いていく場所。">
  <meta property="og:type" content="website">
  <meta property="og:site_name" content="YUKIZ BLOG">
  <meta property="og:title" content="reasons｜YUKIZ BLOG">
  <meta property="og:description" content="静かな断片を置いていく場所。">
  <link rel="stylesheet" href="../style.css">
</head>
<body>
  <main class="site-place" data-fade>
    <header class="site-header">
      <a class="site-name" href="../">YUKIZ BLOG</a>
      <a class="small-link" href="../">戻る</a>
    </header>

    <section class="fragments" aria-label="静かな並び">
{joined}
    </section>

    <footer class="site-footer">Yukihiko Kikuta</footer>
  </main>

  <script src="../main.js"></script>
</body>
</html>
"""


def render_main_js(posts: list[dict]) -> str:
    fragments = [str(post["url"]) for post in sorted(posts, key=post_sort_key, reverse=True)]
    fragment_json = json.dumps(fragments, ensure_ascii=False, indent=2)
    return f"""(function () {{
  var wander = document.getElementById("wander-link");
  var fragments = {fragment_json};

  if (wander && fragments.length) {{
    wander.addEventListener("click", function (event) {{
      event.preventDefault();
      var index = Math.floor(Math.random() * fragments.length);
      window.location.href = fragments[index];
    }});
  }}
}})();
"""


def render_sitemap(posts: list[dict], base_url: str = DEFAULT_BASE_URL) -> str:
    today = now_tokyo().date().isoformat()
    entries = [
        ("", today),
        ("posts/", today),
    ]
    for post in sorted(posts, key=post_sort_key, reverse=True):
        entries.append((str(post["url"]), str(post.get("updated_at", today))[:10]))
    urls = []
    for rel, lastmod in entries:
        urls.append(
            f"""  <url>
    <loc>{html.escape(join_url(base_url, rel))}</loc>
    <lastmod>{html.escape(lastmod)}</lastmod>
  </url>"""
        )
    return """<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
""" + "\n".join(urls) + "\n</urlset>\n"


def render_robots(base_url: str = DEFAULT_BASE_URL) -> str:
    return f"""User-agent: *
Allow: /
Sitemap: {join_url(base_url, "sitemap.xml")}
"""


def render_default_home() -> str:
    return """<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>YUKIZ BLOG</title>
  <meta name="description" content="人であり、そこで、理由を置く。">
  <link rel="stylesheet" href="style.css">
</head>
<body class="home">
  <main class="home-place" data-fade>
    <p class="site-name">YUKIZ BLOG</p>

    <h1 class="home-words">
      人であり、<br>
      そこで、<br>
      理由を置く。
    </h1>

    <nav class="quiet-nav" aria-label="入口">
      <a href="posts/">入る</a>
      <a href="posts/" id="wander-link">巡る</a>
    </nav>
  </main>

  <script src="main.js"></script>
</body>
</html>
"""


def render_default_style() -> str:
    return """:root {
  --bg: #F6F7F9;
  --text: #1C1C1C;
  --sub: #667085;
  --accent: #2F6FED;
  --line: rgba(28, 28, 28, 0.12);
  --soft-line: rgba(28, 28, 28, 0.07);
  --font-ja: "Zen Kaku Gothic New", "Noto Sans JP", "BIZ UDPGothic", "Yu Gothic UI", sans-serif;
  --font-latin: "Segoe UI", Arial, Verdana, var(--font-ja);
}

* { box-sizing: border-box; }
html { min-height: 100%; background: var(--bg); }
body {
  min-height: 100vh;
  margin: 0;
  color: var(--text);
  background: var(--bg);
  font-family: var(--font-ja);
  font-size: 16px;
  line-height: 2;
  letter-spacing: 0;
}
a { color: inherit; text-decoration: none; transition: color 560ms ease, opacity 560ms ease; }
a:hover, a:focus-visible { color: rgba(47, 111, 237, 0.62); }
.home-place, .site-place { width: min(100% - 48px, 720px); margin: 0 auto; }
.home-place { min-height: 100vh; display: flex; flex-direction: column; justify-content: center; padding: 13vh 0 16vh; }
.site-place { padding: 76px 0 64px; }
.site-name { margin: 0 0 54px; color: var(--sub); font-size: 0.8rem; font-weight: 400; }
.home-words { margin: 0; font-size: 3.35rem; font-weight: 400; line-height: 1.62; }
.quiet-nav, .site-header, .page-nav { color: var(--sub); font-size: 0.84rem; }
.quiet-nav { display: flex; gap: 1.5rem; margin-top: 88px; }
.site-header { display: flex; justify-content: space-between; gap: 24px; align-items: baseline; margin-bottom: 92px; }
.site-header .site-name { margin: 0; }
.small-link { color: var(--sub); font-size: 0.78rem; }
.fragments { display: grid; gap: 62px; }
.fragment { display: block; padding: 0 0 44px; border-bottom: 1px solid var(--soft-line); }
.fragment time, .entry-date { display: block; margin-bottom: 1.32rem; color: var(--sub); font-size: 0.95rem; }
.fragment p, .entry-body p { margin: 0; font-size: 1.34rem; line-height: 1.94; }
.entry { max-width: 620px; padding-top: 10px; }
.entry-body { margin-top: 64px; }
.entry-body p { font-size: 1.48rem; line-height: 2; }
.page-nav { display: grid; grid-template-columns: 1fr auto 1fr; gap: 16px; align-items: center; margin-top: 112px; padding-top: 24px; border-top: 1px solid var(--line); color: #7B8798; }
.page-nav a { color: #7B8798; transition: color 560ms ease, opacity 560ms ease; }
.page-nav a:hover, .page-nav a:focus-visible { color: rgba(47, 111, 237, 0.62); opacity: 0.78; text-decoration: none; }
.page-nav a:first-child, .page-nav span:first-child { justify-self: start; }
.page-nav a:nth-child(2) { justify-self: center; }
.page-nav a:last-child, .page-nav span:last-child { justify-self: end; }
.page-nav .muted { color: rgba(102, 112, 133, 0.42); }
.site-footer { margin-top: 90px; color: var(--sub); font-family: var(--font-latin); font-size: 0.74rem; letter-spacing: 0; }
[data-fade] { animation: breathe-in 900ms ease both; }
@keyframes breathe-in { from { opacity: 0; } to { opacity: 1; } }
@media (max-width: 560px) {
  .home-place, .site-place { width: min(100% - 32px, 720px); }
  .home-words { font-size: 2.24rem; line-height: 1.58; }
  .fragment p, .entry-body p { font-size: 1.18rem; line-height: 1.95; }
}
"""


def ensure_home_index(site_path: Path) -> None:
    path = site_path / "index.html"
    if not path.exists():
        write_text(path, render_default_home())
        return
    text = read_text(path)
    updated = text.replace('href="posts/">enter</a>', 'href="posts/">入る</a>')
    if updated != text:
        write_text(path, updated)


def ensure_style(site_path: Path) -> None:
    path = site_path / "style.css"
    if not path.exists():
        write_text(path, render_default_style())
        return
    text = read_text(path)
    if ".site-footer" not in text:
        text = text.rstrip() + """

.site-footer {
  margin-top: 90px;
  color: var(--sub);
  font-family: var(--font-latin);
  font-size: 0.74rem;
  font-weight: 400;
  letter-spacing: 0;
}
"""
        write_text(path, text)


def make_post_from_body(site_path: Path, posts: list[dict], body_text: str, preview: bool = False) -> dict:
    body = normalize_body(body_text)
    if not body:
        raise ValueError("本文が空です。")
    now = now_tokyo()
    date_code = now.strftime("%Y%m%d")
    sequence = next_sequence(posts, site_path / "posts", date_code)
    filename = f"{date_code}-{sequence:03d}.html"
    internal_title = generate_internal_title(body, now.strftime("%Y.%m.%d"))
    meta_description = generate_meta_description(body)
    stamp = now.isoformat(timespec="seconds")
    post_id = f"yb-{date_code}-{sequence:03d}"
    if preview:
        post_id = f"preview-{date_code}-{sequence:03d}"
    return {
        "id": post_id,
        "date": now.date().isoformat(),
        "file": post_url_from_file(filename),
        "url": post_url_from_file(filename),
        "internal_title": internal_title,
        "page_title": f"{internal_title}｜YUKIZ BLOG",
        "meta_description": meta_description,
        "og_title": f"{internal_title}｜YUKIZ BLOG",
        "og_description": meta_description,
        "excerpt": generate_excerpt(body),
        "body": body,
        "air_tags": generate_air_tags(body, internal_title),
        "prev": None,
        "next": None,
        "created_at": stamp,
        "updated_at": stamp,
    }


def write_site_outputs(site_path: Path, posts: list[dict], rebuild_articles: bool = True) -> list[dict]:
    site_path.mkdir(parents=True, exist_ok=True)
    (site_path / "posts").mkdir(parents=True, exist_ok=True)
    (site_path / "data").mkdir(parents=True, exist_ok=True)

    ensure_home_index(site_path)
    ensure_style(site_path)

    posts = recompute_navigation(posts)
    write_text(site_path / "data" / "posts.json", json_dump(posts))
    write_text(site_path / "posts" / "index.html", render_posts_index_html(posts))
    write_text(site_path / "main.js", render_main_js(posts))
    write_text(site_path / "sitemap.xml", render_sitemap(posts))
    write_text(site_path / "robots.txt", render_robots())

    if rebuild_articles:
        for post in posts:
            write_text(site_path / str(post["file"]), render_article_html(post))
    return posts


def refresh_site(site_path: Path) -> list[dict]:
    posts = load_posts(site_path)
    return write_site_outputs(site_path, posts, rebuild_articles=True)


def place_body(site_path: Path, body_text: str) -> dict:
    posts = load_posts(site_path)
    post = make_post_from_body(site_path, posts, body_text)
    posts.append(post)
    write_site_outputs(site_path, posts, rebuild_articles=True)
    return post


def create_preview_html(site_path: Path, body_text: str) -> Path:
    posts = load_posts(site_path)
    draft = make_post_from_body(site_path, posts, body_text, preview=True)
    all_posts = recompute_navigation([dict(post) for post in posts] + [draft])
    draft = next(post for post in all_posts if post["id"] == draft["id"])

    preview_dir = APP_DIR / "_preview"
    preview_dir.mkdir(parents=True, exist_ok=True)
    preview_path = preview_dir / "yukizblog-preview.html"
    stylesheet = (site_path / "style.css").resolve().as_uri()
    home = (site_path / "index.html").resolve().as_uri()
    listing = (site_path / "posts" / "index.html").resolve().as_uri()
    html_text = render_article_html(
        draft,
        stylesheet_href=stylesheet,
        script_src=None,
        home_href=home,
        list_href=listing,
        site_path=site_path,
        preview=True,
    )
    write_text(preview_path, html_text)
    return preview_path


class YukizBlogPostApp(ctk.CTk):  # type: ignore[misc]
    def __init__(self, site_path: Path) -> None:
        super().__init__()
        self.site_path = site_path
        self.is_working = False
        self.status_index = 0
        self.last_post_path: Path | None = None

        self.title(APP_NAME)
        self.geometry("960x780")
        self.minsize(780, 640)
        self.configure(fg_color="#F6F7F9")

        self.font_family = "BIZ UDPGothic"
        self.title_font = ctk.CTkFont(family=self.font_family, size=25, weight="normal")
        self.sub_font = ctk.CTkFont(family=self.font_family, size=13)
        self.body_font = ctk.CTkFont(family=self.font_family, size=19)
        self.button_font = ctk.CTkFont(family=self.font_family, size=14)
        self.footer_font = ctk.CTkFont(family=self.font_family, size=11)
        self.link_font = ctk.CTkFont(family=self.font_family, size=11)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=56, pady=(46, 24))
        header.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(header, text=APP_NAME, font=self.title_font, text_color="#20242A")
        title.grid(row=0, column=0, sticky="w")

        subtitle = ctk.CTkLabel(header, text="言葉を、静かに置く。", font=self.sub_font, text_color="#667085")
        subtitle.grid(row=1, column=0, sticky="w", pady=(8, 0))

        body_wrap = ctk.CTkFrame(
            self,
            fg_color="#FFFFFF",
            corner_radius=20,
            border_width=1,
            border_color="#E8EDF4",
        )
        body_wrap.grid(row=1, column=0, sticky="nsew", padx=56, pady=(0, 24))
        body_wrap.grid_columnconfigure(0, weight=1)
        body_wrap.grid_rowconfigure(0, weight=1)

        self.body_text = ctk.CTkTextbox(
            body_wrap,
            fg_color="#FFFFFF",
            text_color="#1C1C1C",
            border_width=0,
            corner_radius=18,
            font=self.body_font,
            wrap="word",
            spacing1=8,
            spacing2=10,
            spacing3=16,
            padx=34,
            pady=32,
        )
        self.body_text.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        self.body_text.bind("<KeyPress>", self.on_body_keypress)
        try:
            self.body_text._textbox.configure(
                insertbackground="#5E7895",
                selectbackground="#DDE8F1",
                selectforeground="#1C1C1C",
            )
        except Exception:
            pass
        self.body_text.focus_set()

        lower = ctk.CTkFrame(self, fg_color="transparent")
        lower.grid(row=2, column=0, sticky="ew", padx=56, pady=(0, 28))
        lower.grid_columnconfigure(2, weight=1)
        lower.grid_columnconfigure(3, weight=0)

        self.place_button = ctk.CTkButton(
            lower,
            text="置く",
            font=self.button_font,
            fg_color="#587898",
            hover_color="#4D6B88",
            text_color="#FFFFFF",
            corner_radius=14,
            width=108,
            height=40,
            command=self.start_place,
        )
        self.place_button.grid(row=0, column=0, sticky="w")

        self.preview_button = ctk.CTkButton(
            lower,
            text="プレビュー",
            font=self.button_font,
            fg_color="#EEF3F7",
            hover_color="#E4ECF3",
            text_color="#344054",
            corner_radius=14,
            width=124,
            height=40,
            command=self.preview,
        )
        self.preview_button.grid(row=0, column=1, sticky="w", padx=(18, 0))

        self.status_label = ctk.CTkLabel(lower, text=STATUS_READY, font=self.sub_font, text_color="#667085")
        self.status_label.grid(row=1, column=0, columnspan=3, sticky="w", pady=(16, 0))

        self.view_post_link = ctk.CTkLabel(
            lower,
            text="置いた記事を見る",
            font=self.link_font,
            text_color="#8A9AAC",
        )
        self.view_post_link.grid(row=1, column=3, sticky="e", pady=(16, 0))
        self.view_post_link.bind("<Button-1>", lambda _event: self.open_last_post())
        self.view_post_link.bind("<Enter>", lambda _event: self.view_post_link.configure(text_color="#6F88A6"))
        self.view_post_link.bind("<Leave>", lambda _event: self.view_post_link.configure(text_color="#8A9AAC"))
        self.view_post_link.grid_remove()

        footer = ctk.CTkLabel(self, text="Yukihiko Kikuta", font=self.footer_font, text_color="#8A94A6")
        footer.grid(row=3, column=0, sticky="s", pady=(0, 26))

    def get_body(self) -> str:
        return self.body_text.get("1.0", "end")

    def on_body_keypress(self, _event: object | None = None) -> None:
        if self.is_working:
            return
        self.view_post_link.grid_remove()
        if self.status_label.cget("text") in {STATUS_DONE, STATUS_PREVIEW}:
            self.status_label.configure(text=STATUS_READY)

    def set_buttons(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        self.place_button.configure(state=state)
        self.preview_button.configure(state=state)

    def animate_status(self) -> None:
        if not self.is_working:
            return
        self.status_label.configure(text=STATUS_WORKING[self.status_index % len(STATUS_WORKING)])
        self.status_index += 1
        self.after(430, self.animate_status)

    def start_place(self) -> None:
        if self.is_working:
            return
        body = self.get_body()
        if not normalize_body(body):
            self.status_label.configure(text=f"{STATUS_ERROR} 本文が空です。")
            return
        self.view_post_link.grid_remove()
        self.is_working = True
        self.status_index = 0
        self.set_buttons(False)
        self.animate_status()
        thread = threading.Thread(target=self.place_worker, args=(body,), daemon=True)
        thread.start()

    def place_worker(self, body: str) -> None:
        try:
            post = place_body(self.site_path, body)
        except Exception as exc:
            detail = str(exc).strip() or exc.__class__.__name__
            traceback.print_exc()
            self.after(0, lambda: self.finish_error(detail))
            return
        self.after(0, lambda: self.finish_success(post))

    def finish_success(self, post: dict) -> None:
        self.is_working = False
        self.set_buttons(True)
        self.body_text.delete("1.0", "end")
        self.last_post_path = self.site_path / str(post["file"])
        self.status_label.configure(text=STATUS_DONE)
        self.view_post_link.grid()
        self.body_text.focus_set()

    def finish_error(self, detail: str) -> None:
        self.is_working = False
        self.set_buttons(True)
        detail = detail.splitlines()[0][:80]
        self.status_label.configure(text=f"{STATUS_ERROR} {detail}")

    def open_last_post(self) -> None:
        if self.last_post_path:
            webbrowser.open(self.last_post_path.resolve().as_uri())

    def preview(self) -> None:
        if self.is_working:
            return
        body = self.get_body()
        if not normalize_body(body):
            self.status_label.configure(text=f"{STATUS_ERROR} 本文が空です。")
            return
        try:
            self.view_post_link.grid_remove()
            preview_path = create_preview_html(self.site_path, body)
            webbrowser.open(preview_path.resolve().as_uri())
            self.status_label.configure(text=STATUS_PREVIEW)
        except Exception as exc:
            detail = (str(exc).strip() or exc.__class__.__name__).splitlines()[0][:80]
            self.status_label.configure(text=f"{STATUS_ERROR} {detail}")


def run_self_test() -> dict:
    root = APP_DIR / "_selftest"
    if root.exists():
        shutil.rmtree(root, ignore_errors=True)
    site = root / "site"
    first = place_body(site, "役所調査の帰り。\n空が広かった。")
    second = place_body(site, "プリンターのガラス面を拭く。\n少し整う。")
    preview = create_preview_html(site, "夜のコピー機が、少し青く見えた。")

    data = json.loads(read_text(site / "data" / "posts.json"))
    first_file = site / first["file"]
    second_file = site / second["file"]
    index_html = read_text(site / "posts" / "index.html")
    article_html = read_text(second_file)

    assert len(data) == 2
    assert first_file.exists()
    assert second_file.exists()
    assert (site / "sitemap.xml").exists()
    assert (site / "robots.txt").exists()
    assert "前へ" in article_html and "巡る" in article_html and "次へ" in article_html
    assert index_html.find("プリンター") < index_html.find("役所調査")
    assert preview.exists()

    return {
        "ok": True,
        "generated": [first["file"], second["file"]],
        "preview": str(preview),
        "site": str(site),
    }


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=APP_NAME)
    parser.add_argument("--site", default=str(DEFAULT_SITE_PATH), help="YUKIZ BLOG site path")
    parser.add_argument("--refresh-site", action="store_true", help="rebuild site index/json/navigation/sitemap")
    parser.add_argument("--self-test", action="store_true", help="run non-GUI generation test")
    parser.add_argument("--body", help="place text from command line")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    site_path = Path(args.site)

    if args.self_test:
        print(json_dump(run_self_test()), end="")
        return 0

    if args.refresh_site:
        posts = refresh_site(site_path)
        print(json_dump({"ok": True, "site": str(site_path), "posts": len(posts)}), end="")
        return 0

    if args.body is not None:
        post = place_body(site_path, args.body)
        print(json_dump({"ok": True, "post": post}), end="")
        return 0

    if ctk is None:
        print("customtkinter が見つかりません。requirements.txt を確認してください。", file=sys.stderr)
        return 1

    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
    app = YukizBlogPostApp(site_path)
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

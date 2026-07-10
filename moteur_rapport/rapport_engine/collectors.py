from __future__ import annotations

import html
import re
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from typing import Iterable


USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125 Safari/537.36"
)


@dataclass
class SearchItem:
    kind: str
    title: str
    url: str
    snippet: str
    source: str
    published: str | None = None
    query: str | None = None
    verified: bool = False
    verification_note: str = ""


def clean_text(value: str | None) -> str:
    if not value:
        return ""
    value = re.sub(r"<[^>]+>", " ", value)
    value = html.unescape(value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def host_from_url(url: str) -> str:
    try:
        return urllib.parse.urlparse(url).netloc.replace("www.", "")
    except Exception:
        return "source inconnue"


def bing_rss_url(query: str, count: int) -> str:
    params = urllib.parse.urlencode({"q": query, "format": "rss", "count": str(count), "cc": "FR"})
    return f"https://www.bing.com/search?{params}"


def fetch_url(url: str, timeout: int) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return response.read()


def parse_pub_date(value: str | None) -> str | None:
    if not value:
        return None
    try:
        return parsedate_to_datetime(value).date().isoformat()
    except Exception:
        return clean_text(value)[:32] or None


def search_bing_rss(query: str, kind: str, timeout: int, count: int) -> list[SearchItem]:
    url = bing_rss_url(query, count)
    raw = fetch_url(url, timeout)
    root = ET.fromstring(raw)
    items: list[SearchItem] = []
    for node in root.findall(".//item"):
        title = clean_text(node.findtext("title"))
        link = clean_text(node.findtext("link"))
        snippet = clean_text(node.findtext("description"))
        published = parse_pub_date(node.findtext("pubDate"))
        if not title or not link:
            continue
        items.append(
            SearchItem(
                kind=kind,
                title=title,
                url=link,
                snippet=snippet,
                source=host_from_url(link),
                published=published,
                query=query,
            )
        )
    return items


def absolute_url(url: str) -> str:
    if url.startswith("http"):
        return url
    return urllib.parse.urljoin("https://www.free-work.com", url)


def title_from_slug(url: str) -> str:
    slug = urllib.parse.urlparse(url).path.rstrip("/").split("/")[-1]
    slug = slug.replace("-", " ")
    return slug[:1].upper() + slug[1:]


def extract_meta_description(page: str) -> str:
    patterns = [
        r'<meta[^>]+name=["\']description["\'][^>]+content=["\']([^"\']+)["\']',
        r'<meta[^>]+property=["\']og:description["\'][^>]+content=["\']([^"\']+)["\']',
    ]
    for pattern in patterns:
        match = re.search(pattern, page, flags=re.I | re.S)
        if match:
            return clean_text(match.group(1))
    text = clean_text(page)
    return text[:700]


def extract_page_title(page: str, fallback: str) -> str:
    match = re.search(r"<title[^>]*>(.*?)</title>", page, flags=re.I | re.S)
    if match:
        title = clean_text(match.group(1))
        title = re.sub(r"\s*\|\s*Free-Work.*$", "", title).strip()
        if title:
            return title
    return fallback


def extract_published(page: str) -> str | None:
    match = re.search(r"Publi[ée]e?\s+le\s+(\d{2}/\d{2}/\d{4})", page, flags=re.I)
    if not match:
        return None
    day, month, year = match.group(1).split("/")
    return f"{year}-{month}-{day}"


def collect_freework_jobs(pages: list[str], timeout: int, max_links: int = 40) -> list[SearchItem]:
    urls: list[str] = []
    for page_url in pages:
        try:
            raw = fetch_url(page_url, timeout).decode("utf-8", "ignore")
        except Exception:
            continue
        for href in re.findall(r'href=["\']([^"\']*/job-mission/[^"\']*)["\']', raw, flags=re.I):
            url = absolute_url(href)
            if url not in urls:
                urls.append(url)
            if len(urls) >= max_links:
                break

    items: list[SearchItem] = []
    for url in urls[:max_links]:
        fallback = title_from_slug(url)
        try:
            page = fetch_url(url, timeout).decode("utf-8", "ignore")
            title = extract_page_title(page, fallback)
            snippet = extract_meta_description(page)
            published = extract_published(page)
            verified = True
            note = "GET OK"
        except Exception as exc:
            title = fallback
            snippet = "Annonce Free-Work detectee via page publique ; details a verifier sur la plateforme."
            published = None
            verified = False
            note = f"Acces limite ou echec verification: {type(exc).__name__}"
        items.append(
            SearchItem(
                kind="opportunity",
                title=title,
                url=url,
                snippet=snippet,
                source="free-work.com",
                published=published,
                query="Free-Work direct public pages",
                verified=verified,
                verification_note=note,
            )
        )
        time.sleep(0.2)
    return items


def dedupe(items: Iterable[SearchItem]) -> list[SearchItem]:
    seen: set[str] = set()
    result: list[SearchItem] = []
    for item in items:
        key = urllib.parse.urlsplit(item.url)._replace(query="", fragment="").geturl().lower()
        if key in seen:
            continue
        seen.add(key)
        result.append(item)
    return result


def verify_item(item: SearchItem, timeout: int) -> SearchItem:
    try:
        req = urllib.request.Request(item.url, headers={"User-Agent": USER_AGENT}, method="HEAD")
        with urllib.request.urlopen(req, timeout=timeout) as response:
            item.verified = 200 <= response.status < 400
            item.verification_note = f"HTTP {response.status}"
            return item
    except Exception:
        try:
            fetch_url(item.url, min(timeout, 8))
            item.verified = True
            item.verification_note = "GET OK"
        except Exception as exc:
            item.verified = False
            item.verification_note = f"Acces limite ou echec verification: {type(exc).__name__}"
    return item


def collect(config: dict) -> tuple[list[SearchItem], list[SearchItem], dict]:
    search_cfg = config["search"]
    timeout = int(search_cfg.get("timeout_seconds", 18))
    count = int(search_cfg.get("results_per_query", 8))
    article_items: list[SearchItem] = []
    opportunity_items: list[SearchItem] = []
    errors: list[str] = []

    for query in search_cfg["article_queries"]:
        try:
            article_items.extend(search_bing_rss(query, "article", timeout, count))
            time.sleep(0.4)
        except Exception as exc:
            errors.append(f"Article query failed: {query} ({type(exc).__name__})")

    for query in search_cfg["opportunity_queries"]:
        try:
            opportunity_items.extend(search_bing_rss(query, "opportunity", timeout, count))
            time.sleep(0.4)
        except Exception as exc:
            errors.append(f"Opportunity query failed: {query} ({type(exc).__name__})")

    try:
        opportunity_items.extend(
            collect_freework_jobs(search_cfg.get("freework_pages", []), timeout)
        )
    except Exception as exc:
        errors.append(f"Free-Work direct collection failed ({type(exc).__name__})")

    article_items = dedupe(article_items)
    opportunity_items = dedupe(opportunity_items)

    for item in article_items[:40] + opportunity_items[:40]:
        verify_item(item, timeout)

    meta = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "article_queries": len(search_cfg["article_queries"]),
        "opportunity_queries": len(search_cfg["opportunity_queries"]),
        "raw_articles": len(article_items),
        "raw_opportunities": len(opportunity_items),
        "errors": errors,
    }
    return article_items, opportunity_items, meta

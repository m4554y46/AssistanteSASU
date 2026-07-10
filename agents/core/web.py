from __future__ import annotations

import html
import json
import re
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from dataclasses import dataclass, asdict
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from typing import Iterable


USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125 Safari/537.36"
)


@dataclass
class SearchResult:
    title: str
    url: str
    snippet: str
    source: str
    published: str | None = None
    query: str | None = None


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


def bing_rss_url(query: str, count: int = 8) -> str:
    params = urllib.parse.urlencode({"q": query, "format": "rss", "count": str(count), "cc": "FR"})
    return f"https://www.bing.com/search?{params}"


def fetch_url(url: str, timeout: int = 15) -> bytes:
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


def search_bing_rss(query: str, timeout: int = 15, count: int = 8) -> list[SearchResult]:
    url = bing_rss_url(query, count)
    raw = fetch_url(url, timeout)
    root = ET.fromstring(raw)
    items: list[SearchResult] = []
    for node in root.findall(".//item"):
        title = clean_text(node.findtext("title"))
        link = clean_text(node.findtext("link"))
        snippet = clean_text(node.findtext("description"))
        published = parse_pub_date(node.findtext("pubDate"))
        if not title or not link:
            continue
        items.append(
            SearchResult(
                title=title,
                url=link,
                snippet=snippet,
                source=host_from_url(link),
                published=published,
                query=query,
            )
        )
    return items


def dedupe(items: Iterable[SearchResult]) -> list[SearchResult]:
    seen: set[str] = set()
    result: list[SearchResult] = []
    for item in items:
        key = urllib.parse.urlsplit(item.url)._replace(query="", fragment="").geturl().lower()
        if key in seen:
            continue
        seen.add(key)
        result.append(item)
    return result


def multi_search(queries: list[str], timeout: int = 15, count: int = 8) -> list[SearchResult]:
    all_items: list[SearchResult] = []
    for query in queries:
        try:
            all_items.extend(search_bing_rss(query, timeout, count))
            time.sleep(0.4)
        except Exception:
            pass
    return dedupe(all_items)


# --- Free-Work scraping (from the original moteur_rapport) ---

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
    return clean_text(page)[:700]


def extract_page_title(page: str, fallback: str) -> str:
    match = re.search(r"<title[^>]*>(.*?)</title>", page, flags=re.I | re.S)
    if match:
        title = clean_text(match.group(1))
        title = re.sub(r"\s*\|\s*Free-Work.*$", "", title).strip()
        if title:
            return title
    return fallback


def extract_freework_tjm(page: str) -> int | None:
    """Extract TJM from Free-Work mission page HTML (JSON-LD + text patterns)."""
    # Try JSON-LD first (Free-Work embeds structured data)
    ld_matches = re.findall(r'<script[^>]*type="application/ld\+json"[^>]*>(.*?)</script>', page, re.I | re.S)
    for ld_str in ld_matches:
        try:
            data = json.loads(ld_str)
            if isinstance(data, dict):
                min_sal = data.get("minDailySalary") or data.get("baseSalary", {}).get("minValue")
                max_sal = data.get("maxDailySalary") or data.get("baseSalary", {}).get("maxValue")
                if max_sal and isinstance(max_sal, (int, float)):
                    return int(max_sal)
                if min_sal and isinstance(min_sal, (int, float)):
                    return int(min_sal)
        except json.JSONDecodeError:
            pass

    # Text patterns fallback
    patterns = [
        r"TJM\s*(?:moyen|moyenne)?\s*[:\s]*(\d{3,4})\s*(?:EUR|€)",
        r"(\d{3,4})\s*(?:EUR|€)\s*/jour",
        r"minDailySalary[:\"]+(\d{3,4})",
        r"maxDailySalary[:\"]+(\d{3,4})",
        r"dailySalary[:\"]+(\d{3,4})",
        r"taux journalier[:\s]+(\d{3,4})",
    ]
    for pat in patterns:
        for match in re.finditer(pat, page, flags=re.I):
            val = int(match.group(1))
            if 300 <= val <= 1500:
                return val
    return None


def extract_freework_duration(page: str) -> str | None:
    match = re.search(r"Durée\s*[:\s]*([^\n<]+)", page, flags=re.I)
    if match:
        return match.group(1).strip()
    return None


def extract_freework_location(page: str) -> str | None:
    match = re.search(r"(?:Localisation|Lieu)\s*[:\s]*([^\n<]+)", page, flags=re.I)
    if match:
        return match.group(1).strip()
    return None


def collect_freework_jobs(pages: list[str], timeout: int, max_links: int = 50) -> list[SearchResult]:
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

    items: list[SearchResult] = []
    for url in urls[:max_links]:
        fallback = title_from_slug(url)
        try:
            page = fetch_url(url, timeout).decode("utf-8", "ignore")
            title = extract_page_title(page, fallback)
            snippet = extract_meta_description(page)
            tjm = extract_freework_tjm(page)
            duree = extract_freework_duration(page)
            location = extract_freework_location(page)
            extra = ""
            if tjm:
                extra += f" | TJM: {tjm}EUR"
            if duree:
                extra += f" | Duree: {duree}"
            if location:
                extra += f" | Localisation: {location}"
            snippet += extra
        except Exception:
            title = fallback
            snippet = "Offre Free-Work"
        items.append(
            SearchResult(
                title=title,
                url=url,
                snippet=snippet,
                source="free-work.com",
                query="Free-Work direct",
            )
        )
        time.sleep(0.3)
    return items

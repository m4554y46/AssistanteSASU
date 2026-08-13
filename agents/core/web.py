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

BLOCKED_DOMAINS = [
    "pinterest", "amazon", "ebay", "etsy", "aliexpress", "walmart",
    "shopify", "boulanger", "fnac", "cdiscount", "ikea", "leroymerlin",
    "decathlon", "booking", "tripadvisor",
    "larousse", "lerobert", "dictionnaire", "wiktionary", "cnrtl",
    "linternaute", "wikihow", "wikipedia",
    "fiverr", "freelance.com", "upwork", "peopleperhour",
    "facebook", "instagram", "twitter", "x.com", "tiktok",
    "youtube", "leboncoin",
    "wordreference", "linguee", "reverso",
    "cambridge", "merriam", "oxford", "collins",
    "allocine", "mozzartbet", "bet365", "parionssport", "poker",
    "synonymo", "aujourdhui",
    "producthunt", "keejob", "optioncarriere", "acc.tn", "lci.com.tn",
    "indeed", "glassdoor", "monster", "kijiji",
]


@dataclass
class SearchResult:
    title: str
    url: str
    snippet: str
    source: str
    published: str | None = None
    query: str | None = None
    company: str | None = None
    location: str | None = None
    duration: str | None = None
    tjm: int | None = None


def _is_latin(text: str) -> bool:
    for ch in text:
        cp = ord(ch)
        if cp > 0x024F and cp < 0x1E00:
            return False
        if cp > 0x1EFF and cp < 0x2000:
            return False
        if cp > 0x2E7F:
            return False
    return True


def clean_text(value: str | None) -> str:
    if not value:
        return ""
    value = html.unescape(value)
    # Decode les sequences \uXXXX presentes dans le JSON-LD des pages Free-Work
    value = re.sub(r"\\u([0-9a-fA-F]{4})", lambda m: chr(int(m.group(1), 16)), value)
    value = re.sub(r"<[^>]*>?", " ", value)
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
        domain = host_from_url(link)
        if any(b in domain for b in BLOCKED_DOMAINS):
            continue
        if not _is_latin(title) or (snippet and not _is_latin(snippet)):
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
                # baseSalary = {"value": 590, "unitText": "DAY"}
                base = data.get("baseSalary", {})
                if isinstance(base, dict):
                    val = base.get("value") or (base.get("value", {}) or {}).get("value")
                    unit = base.get("unitText")
                    if val and isinstance(val, (int, float)) and unit and "DAY" in str(unit).upper():
                        return int(val)
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
    match = re.search(r"Dur[ée]e\s*[:\s]*([^\n<]{1,40})", page, flags=re.I)
    if match:
        value = match.group(1).strip()
        # Ne garde que les formes courtes et propres (3 mois, 12 mois, 6 semaines...)
        m = re.search(r"^\d+\s*(mois|jours|semaines|semaine|jour|an|ans)(\s+renouvelable)?", value, re.I)
        if m:
            return m.group(0)
        return None
    return None


def extract_freework_location(page: str) -> str | None:
    match = re.search(r"(?:Localisation|Lieu)\s*[:\s]*([^\n<]{1,60})", page, flags=re.I)
    if match:
        value = match.group(1).strip()
        if value and len(value) > 2:
            return value
    return None


def extract_freework_company(page: str) -> str | None:
    ld_matches = re.findall(r'<script[^>]*type="application/ld\+json"[^>]*>(.*?)</script>', page, re.I | re.S)
    for ld_str in ld_matches:
        try:
            data = json.loads(ld_str)
            if isinstance(data, dict):
                org = data.get("hiringOrganization") or {}
                name = org.get("name")
                if name:
                    return clean_text(name)
        except json.JSONDecodeError:
            pass
    match = re.search(r"-\s*(.+?)\s*\|", extract_page_title(page, ""))
    if match:
        return match.group(1).strip()
    return None


def _load_job_posting_ld(page: str) -> dict | None:
    """Extrait le JSON-LD de type JobPosting d'une page Free-Work.
    C'est la source la plus fiable : toutes les infos y sont structurees et propres."""
    ld_matches = re.findall(r'<script[^>]*type="application/ld\+json"[^>]*>(.*?)</script>', page, re.I | re.S)
    for ld_str in ld_matches:
        try:
            data = json.loads(ld_str)
            if isinstance(data, dict) and data.get("@type") == "JobPosting":
                return data
        except json.JSONDecodeError:
            continue
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
        except Exception:
            continue
        # Source fiable : JSON-LD JobPosting
        ld = _load_job_posting_ld(page)
        # La duree ("3 mois", "12 mois") n'est pas fiable dans le JSON-LD
        # (index Nuxt), on l'extrait du texte affiche dans les deux cas.
        clean_page = clean_text(page)
        duree = extract_freework_duration(clean_page)
        location_regex = extract_freework_location(clean_page)
        if ld:
            title = clean_text(ld.get("title") or fallback)
            snippet = clean_text(ld.get("description") or "")
            company = clean_text((ld.get("hiringOrganization") or {}).get("name") or "")
            base = ld.get("baseSalary") or {}
            if isinstance(base, dict):
                val = base.get("value") or {}
                if isinstance(val, dict):
                    val = val.get("value")
                unit = base.get("unitText") or ""
                tjm = int(val) if isinstance(val, (int, float)) and "DAY" in str(unit).upper() else None
            else:
                tjm = None
            loc = ld.get("jobLocation") or {}
            address = loc.get("address") or {}
            location = clean_text(address.get("addressLocality") or "") or location_regex
            duree = duree
        else:
            title = extract_page_title(page, fallback)
            snippet = extract_meta_description(page)
            tjm = extract_freework_tjm(page)
            location = location_regex
            company = extract_freework_company(page)
            snippet = clean_text(snippet)
            duree = clean_text(duree) if duree else None
            location = clean_text(location) if location else None
            company = clean_text(company) if company else None
        if not _is_latin(title) or (snippet and not _is_latin(snippet)):
            continue
        items.append(
            SearchResult(
                title=title,
                url=url,
                snippet=snippet,
                source="free-work.com",
                query="Free-Work direct",
                company=company,
                location=location,
                duration=duree,
                tjm=tjm,
            )
        )
        time.sleep(0.3)
    return items

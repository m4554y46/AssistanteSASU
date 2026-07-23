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
    "opencare.com", "doctor.webmd.com", "deltadental.com", "seattlemet.com",
    "zocdoc.com", "healthgrades.com", "ratemds.com",
]

BLOCKED_CONTENT = [
    "dentist", "dentiste", "dental", "tooth", "teeth", "orthodontist",
    "plumber", "plombier", "electrician", "electricien",
    "lawyer", "avocat", "attorney", "notaire",
    "doctor", "medecin", "physician", "hospital", "hopital", "clinic", "clinique",
    "restaurant", "pizza", "sushi", "bakery", "boulangerie",
    "real estate", "immobilier", "apartment", "appartement", "condo",
    "mechanic", "garage auto", "car repair",
    "best near me", "top rated", "top 10", "top 5", "best of",
    "yelp.com", "opencare.com", "healthgrades",
    "insurance", "assurance auto", "assurance habitation",
    "pest control", "exterminator",
    "landscaping", "jardinier",
    "moving company", "demenageur",
    "chiropractor", "chiropracteur",
    "optician", "opticien", "optometrist",
    "daycare", "nounou", "babysitter",
    "gym", "fitness", "yoga studio",
    "nail salon", "coiffeur", "barber",
    "vet", "veterinaire", "pet grooming",
    "locksmith", "serrurier",
    "roofer", "couvreur", "contractor",
    "carpet cleaning", "nettoyage",
]

def _has_blocked_content(text: str) -> bool:
    lower = text.lower()
    for term in BLOCKED_CONTENT:
        if term in lower:
            return True
    return False


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
    value = re.sub(r"<[^>]*>?", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def host_from_url(url: str) -> str:
    try:
        return urllib.parse.urlparse(url).netloc.replace("www.", "")
    except Exception:
        return "source inconnue"


CURATED_FEEDS = [
    ("https://www.producttalk.org/feed/", "Product Talk"),
    ("https://www.svpg.com/feed/", "SVPG"),
    ("https://medium.com/feed/product-coalition", "Product Coalition"),
    ("https://medium.com/feed/mind-the-product", "Mind the Product"),
    ("https://www.agilealliance.org/feed/", "Agile Alliance"),
    ("https://martinfowler.com/feed.atom", "Martin Fowler"),
    ("https://medium.com/feed/leading-agile", "Leading Agile"),
]


def bing_html_url(query: str, count: int) -> str:
    params = urllib.parse.urlencode({"q": query, "cc": "FR"})
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


def extract_rss_feed(feed_url: str, source_name: str, kind: str, timeout: int) -> list[SearchItem]:
    items: list[SearchItem] = []
    try:
        raw = fetch_url(feed_url, timeout)
        root = ET.fromstring(raw)
        is_atom = root.tag.endswith("feed")
        if is_atom:
            ns = {"atom": "http://www.w3.org/2005/Atom"}
            for node in root.findall("atom:entry", ns):
                title_el = node.find("atom:title", ns)
                link_el = node.find("atom:link", ns)
                summary_el = node.find("atom:summary", ns)
                content_el = node.find("atom:content", ns)
                published_el = node.find("atom:published", ns) or node.find("atom:updated", ns)
                title = clean_text(title_el.text) if title_el is not None and title_el.text else None
                link = (link_el.get("href") or "").strip() if link_el is not None else None
                snippet = clean_text((summary_el or content_el).text) if (summary_el is not None or content_el is not None) else ""
                published = parse_pub_date(published_el.text) if published_el is not None and published_el.text else None
                if not title or not link:
                    continue
                domain = host_from_url(link)
                if any(b in domain for b in BLOCKED_DOMAINS):
                    continue
                if not _is_latin(title) or (snippet and not _is_latin(snippet)):
                    continue
                if _has_blocked_content(title) or (snippet and _has_blocked_content(snippet)):
                    continue
                items.append(
                    SearchItem(
                        kind=kind,
                        title=title,
                        url=link,
                        snippet=snippet[:500] if snippet else "",
                        source=source_name,
                        published=published,
                        query=source_name,
                    )
                )
        else:
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
                if _has_blocked_content(title) or (snippet and _has_blocked_content(snippet)):
                    continue
                items.append(
                    SearchItem(
                        kind=kind,
                        title=title,
                        url=link,
                        snippet=snippet[:500] if snippet else "",
                        source=source_name,
                        published=published,
                        query=source_name,
                    )
                )
    except Exception:
        pass
    return items


def bing_rss_url(query: str, count: int = 8) -> str:
    params = urllib.parse.urlencode({"q": query, "format": "rss", "count": str(count), "cc": "FR"})
    return f"https://www.bing.com/search?{params}"


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
        domain = host_from_url(link)
        if any(b in domain for b in BLOCKED_DOMAINS):
            continue
        if not _is_latin(title) or (snippet and not _is_latin(snippet)):
            continue
        if _has_blocked_content(title) or (snippet and _has_blocked_content(snippet)):
            continue
        items.append(
            SearchItem(
                kind=kind,
                title=title,
                url=link,
                snippet=snippet,
                source=domain,
                published=published,
                query=query,
            )
        )
    return items


def search_bing_html(query: str, kind: str, timeout: int, count: int) -> list[SearchItem]:
    url = bing_html_url(query, count)
    raw = fetch_url(url, timeout).decode("utf-8", "ignore")
    items: list[SearchItem] = []
    seen_urls: set[str] = set()
    for idx, href in enumerate(re.findall(r'<a[^>]+href="(https?://[^"]+)"[^>]*>(.*?)</a>', raw, flags=re.I | re.S)):
        link, title_block = href
        if not link or any(x in link for x in ["bing.com", "go.microsoft", "creativecommons"]):
            continue
        if link in seen_urls:
            continue
        seen_urls.add(link)
        title = clean_text(title_block)
        if not title or len(title) < 10:
            continue
        domain = host_from_url(link)
        if any(b in domain for b in BLOCKED_DOMAINS):
            continue
        if not _is_latin(title):
            continue
        caption = ""
        for cap_match in re.finditer(r'<p[^>]*>(.*?)</p>', raw[idx:idx+3000], flags=re.I | re.S):
            cap_text = clean_text(cap_match.group(1))
            if cap_text and len(cap_text) > 20:
                caption = cap_text[:500]
                break
        items.append(
            SearchItem(
                kind=kind,
                title=title,
                url=link,
                snippet=caption,
                source=domain,
                published=None,
                query=query,
            )
        )
        if len(items) >= count:
            break
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
        if not _is_latin(title) or (snippet and not _is_latin(snippet)):
            continue
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


def search_github_sources(queries: list[str], timeout: int, count: int) -> list[SearchItem]:
    items: list[SearchItem] = []
    for query in queries:
        if "site:github.com" not in query:
            query = f"site:github.com {query}"
        try:
            # RSS-based search works better than HTML for site: queries
            items.extend(search_bing_rss(query, "github", timeout, count))
            time.sleep(0.3)
        except Exception:
            try:
                items.extend(search_bing_html(query, "github", timeout, count))
                time.sleep(0.3)
            except Exception:
                pass
    seen = set()
    unique = []
    for item in items:
        key = (item.title, item.url)
        if key not in seen:
            seen.add(key)
            unique.append(item)
    return unique


def collect_tokenforge(config: dict) -> list[SearchItem]:
    search_cfg = config["search"]
    timeout = int(search_cfg.get("timeout_seconds", 18))
    count = int(search_cfg.get("results_per_query", 8))
    queries = search_cfg.get("tokenforge_queries", [])
    items: list[SearchItem] = []

    items.extend(search_github_sources(queries, timeout, count))

    non_github = [q for q in queries if "github" not in q.lower()]
    for query in non_github:
        try:
            items.extend(search_bing_rss(query, "article", timeout, count))
            time.sleep(0.4)
        except Exception:
            pass

    return dedupe(items)


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

    for feed_url, source_name in CURATED_FEEDS:
        try:
            article_items.extend(extract_rss_feed(feed_url, source_name, "article", timeout))
        except Exception as exc:
            errors.append(f"RSS feed failed: {source_name} ({type(exc).__name__})")

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

from __future__ import annotations

import re
from dataclasses import asdict
from datetime import date

from .collectors import SearchItem


def text_of(item: SearchItem) -> str:
    return f"{item.title} {item.snippet} {item.source}".lower()


def keyword_hits(text: str, keywords: list[str]) -> int:
    hits = 0
    for kw in keywords:
        if kw.lower() in text:
            hits += 1
    return hits


def source_quality(source: str) -> float:
    source = source.lower()
    premium = [
        "openai.com", "mckinsey.com", "atlassian.com", "scrumguides.org",
        "producttalk.org", "svpg.com", "teamtopologies.com", "basecamp.com",
        "scaledagileframework.com", "gartner.com", "hbr.org"
    ]
    platforms = ["free-work.com", "linkedin.com", "welcometothejungle.com", "indeed.com", "malt.fr"]
    if any(p in source for p in premium):
        return 1.4
    if any(p in source for p in platforms):
        return 1.2
    if source.endswith(".com") or source.endswith(".fr") or source.endswith(".org"):
        return 0.9
    return 0.5


def date_bonus(published: str | None) -> float:
    if not published:
        return 0.2
    match = re.match(r"(20\d{2})-(\d{2})-(\d{2})", published)
    if not match:
        return 0.2
    try:
        d = date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    except ValueError:
        return 0.2
    age = (date.today() - d).days
    if age <= 14:
        return 1.2
    if age <= 45:
        return 0.8
    if age <= 120:
        return 0.4
    return 0.1


def score_item(item: SearchItem, positive: list[str], negative: list[str]) -> dict:
    text = text_of(item)
    pos = keyword_hits(text, positive)
    neg = keyword_hits(text, negative)
    raw = 2.0 + min(pos, 8) * 0.45 + source_quality(item.source) + date_bonus(item.published)
    if item.verified:
        raw += 0.4
    raw -= min(neg, 4) * 0.65
    if item.kind == "opportunity":
        raw += opportunity_specific_adjustment(item, text)
    score = max(0.5, min(9.6, round(raw, 1)))
    return {
        **asdict(item),
        "score": score,
        "keyword_hits": pos,
        "negative_hits": neg,
    }


def selection_reason(item: SearchItem, score: float, hits: int) -> str:
    if item.kind == "opportunity":
        return (
            "Retenue car elle combine signaux de mission freelance, proximite avec Product/IA/delivery "
            f"et source identifiable ({item.source}). Score {score}/10 apres filtrage."
        )
    return (
        "Retenu car le contenu est exploitable pour une conversation client ou une prise de parole, "
        f"avec {hits} signaux thematiques detectes et une source identifiable ({item.source})."
    )


def recommendation(item: SearchItem, score: float) -> str:
    if item.kind == "opportunity":
        if score >= 8.5:
            return "POSTULER / QUALIFIER EN PRIORITE"
        if score >= 7.4:
            return "QUALIFIER RAPIDEMENT"
        return "SURVEILLER"
    if score >= 8.5:
        return "Transformer en angle d'offre ou post LinkedIn cette semaine."
    if score >= 7.4:
        return "Conserver pour enrichir le discours commercial et les rendez-vous."
    return "A surveiller, sans action immediate."


def opportunity_specific_adjustment(item: SearchItem, text: str) -> float:
    url = item.url.lower()
    role_terms = [
        "product owner", "product manager", "chef de projet", "consultant ia",
        "delivery manager", "change manager", "responsable ia"
    ]
    mission_terms = ["freelance", "mission", "job-mission", "tjm", "teletravail", "durée", "démarrage"]
    generic_terms = [
        "embaucher", "trouver des freelances", "annuaire", "plateforme malt",
        "plus de 10 000 emplois", "freelances à", "freelances a", "emplois freelances en ligne"
    ]

    adjustment = 0.0
    if "/job-mission/" in url or "linkedin.com/jobs/view" in url:
        adjustment += 1.4
    elif "/jobs/" in url:
        adjustment += 0.4
    else:
        adjustment -= 1.4

    if any(term in text for term in role_terms):
        adjustment += 0.8
    else:
        adjustment -= 0.8

    if any(term in text or term in url for term in mission_terms):
        adjustment += 0.5

    if any(term in text or term in url for term in generic_terms):
        adjustment -= 2.5

    if "statics.free-work.com/users/documents" in url:
        adjustment -= 4.0

    return adjustment


def is_viable_opportunity(item: SearchItem) -> bool:
    text = text_of(item)
    url = item.url.lower()
    if "statics.free-work.com/users/documents" in url:
        return False
    generic = ["embaucher", "trouver des freelances", "annuaire", "emplois freelances en ligne"]
    if any(term in text for term in generic):
        return False
    role = ["product owner", "product manager", "chef de projet", "consultant ia", "delivery manager", "responsable ia"]
    has_role = any(term in text for term in role)
    has_mission_url = "/job-mission/" in url or "linkedin.com/jobs/view" in url or "/jobs/" in url
    return has_role and has_mission_url


def rank_items(items: list[SearchItem], positive: list[str], negative: list[str], limit: int) -> list[dict]:
    if items and items[0].kind == "opportunity":
        items = [item for item in items if is_viable_opportunity(item)]
    scored = [score_item(item, positive, negative) for item in items]
    scored.sort(key=lambda x: (x["score"], x["verified"], x["keyword_hits"]), reverse=True)
    return scored[:limit]


def extract_tokenforge_items(items: list[SearchItem], positive: list[str], limit: int) -> list[dict]:
    scored = []
    for item in items:
        text = text_of(item)
        hits = keyword_hits(text, positive)
        if hits < 1:
            continue
        source = source_quality(item.source)
        freshness = date_bonus(item.published)
        # Base: 4.0 + hits + source + freshness, capped at 9.5
        raw = 2.0 + min(hits, 10) * 0.6 + source + freshness
        if item.kind == "github":
            raw += 0.4
        if item.verified:
            raw += 0.3
        # Pricing news bonus
        pricing_kw = ["pricing", "tariff", "price", "cheaper", "cost", "reduction"]
        if any(kw in text for kw in pricing_kw):
            raw += 0.5
        # TokenForge core bonus
        core_kw = ["compression", "token", "proxy", "caching", "optimization", "finops"]
        if any(kw in text for kw in core_kw):
            raw += 0.5
        score = max(0.5, min(9.5, round(raw, 1)))
        scored.append({
            **asdict(item),
            "score": score,
            "keyword_hits": hits,
        })
    scored.sort(key=lambda x: (x["score"], x["keyword_hits"]), reverse=True)
    return scored[:limit]

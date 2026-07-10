from __future__ import annotations

import csv
import re
import sys
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

from core.common import (
    add_hyperlink,
    bullet,
    ensure_output_dir,
    load_config,
    para,
    save_json_report,
    set_run_font,
    setup_doc,
)
from core.web import SearchResult, multi_search, collect_freework_jobs
from docx.shared import Pt, RGBColor

NAVY = RGBColor(11, 37, 69)
BLACK = RGBColor(0, 0, 0)
GRAY = RGBColor(86, 95, 108)
RED = RGBColor(155, 28, 28)
GREEN = RGBColor(0, 128, 0)
PALEBLUE = "E8EEF5"
LIGHT = "F2F4F7"

BLOCKED_DOMAINS = [
    "dictionnaire", "larousse", "lerobert", "reverso", "wiktionary",
    "wikipedia.org", "cnrtl", "linternaute", "wikihow",
    "boulanger.com", "fnac.com", "amazon", "leboncoin", "ebay",
    "decathlon", "head.com", "hesge", "letudiant",
    "aujourdhui", "linternaute",
]

BLOCKED_ROLES = [
    "cobol", "mainframe", "as400", "rust", "developpeur back",
    "charge de recrutement", "ingenieur reseau", "analyste mainframe",
    "integrateur devops", "ingenieur devops",
]


def is_blocked(item: SearchResult) -> bool:
    if any(d in item.source.lower() for d in BLOCKED_DOMAINS):
        return True
    text = f"{item.title} {item.snippet}".lower()
    if any(role in text for role in BLOCKED_ROLES):
        return True
    return False


def calculer_score_mission(item: SearchResult, profil: dict) -> dict | None:
    text = f"{item.title} {item.snippet} {item.source}".lower()
    url = item.url.lower()

    positive = [
        "directeur de projet", "head of product", "product manager", "product owner",
        "product management", "transformation digitale", "transformation numérique",
        "mobile", "omnicanal", "retail", "luxe", "logistique", "ecommerce",
        "strategie produit", "roadmap", "delivery", "agile", "scrum",
        "management", "pilotage", "maîtrise d'ouvrage", "assistance maîtrise d'ouvrage",
        "AMOA", "MOA", "program manager", "consultant",
        "international", "worldwide", "ERP", "PIM", "CRM",
        "conseil", "audit", "accompagnement", "transformation",
    ]
    industries = [ind.lower() for ind in profil.get("industries", [])]

    pos_hits = [kw for kw in positive if kw in text]
    industry_hits = [ind for ind in industries if ind in text]

    # Must have at least 1 positive hit or a direct role match
    has_role = any(role in text for role in ["directeur de projet", "head of product", "product manager", "product owner", "program manager", "maîtrise d'ouvrage", "delivery manager"])
    if not pos_hits and not has_role:
        return None

    tjm_match = re.search(r"tjm[:\s]*(\d{3,4})|minDailySalary[:\"]+(\d{3,4})|maxDailySalary[:\"]+(\d{3,4})", text, re.I)
    tjm_extracted = 0
    for g in tjm_match.groups() if tjm_match else []:
        if g:
            tjm_extracted = max(tjm_extracted, int(g))

    raw_score = 3.0
    raw_score += min(len(pos_hits) * 0.7, 3.5)
    raw_score += min(len(industry_hits) * 0.6, 1.8)

    if "free-work.com" in item.source and "/job-mission/" in url:
        raw_score += 0.8
    if "linkedin.com/jobs" in url:
        raw_score += 0.5

    if tjm_extracted >= 750:
        raw_score += 2.0
        tjm_label = f"{tjm_extracted}EUR (PREMIUM)"
    elif tjm_extracted >= 650:
        raw_score += 1.2
        tjm_label = f"{tjm_extracted}EUR (bon)"
    elif tjm_extracted >= 500:
        raw_score += 0.3
        tjm_label = f"{tjm_extracted}EUR (moyen)"
    else:
        tjm_label = f"{tjm_extracted}EUR" if tjm_extracted > 0 else "non specific"

    score = max(0.5, min(9.5, round(raw_score, 1)))

    if score >= 7.0 and tjm_extracted >= 650:
        decision = "POSTULER EN PRIORITE"
    elif score >= 5.5:
        decision = "POSTULER"
    elif score >= 3.5:
        decision = "A QUALIFIER"
    else:
        decision = "SURVEILLER"

    fit_skills = pos_hits[:5]
    return {
        "title": item.title,
        "url": item.url,
        "snippet": item.snippet,
        "source": item.source,
        "score": score,
        "tjm": tjm_label,
        "tjm_value": tjm_extracted,
        "fit_skills": fit_skills,
        "decision": decision,
        "industry_match": industry_hits,
    }


def load_manual_missions() -> list[dict]:
    csv_path = Path(__file__).resolve().parent / "missions_manuelles.csv"
    missions = []
    if not csv_path.exists():
        return missions
    with csv_path.open(newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f, delimiter=";"):
            missions.append(row)
    return missions


def collect_missions(agent_cfg: dict, profil: dict) -> list[dict]:
    timeout = int(agent_cfg.get("timeout_seconds", 15))
    count = int(agent_cfg.get("results_per_query", 10))
    max_missions = int(agent_cfg.get("max_missions", 10))

    all_items: list[SearchResult] = []

    # 1. Free-Work direct scraping
    fw_pages = agent_cfg.get("freework_pages", [])
    print(f"  Scraping {len(fw_pages)} pages Free-Work...")
    try:
        fw_items = collect_freework_jobs(fw_pages, timeout)
        print(f"    -> {len(fw_items)} missions trouvees sur Free-Work")
        all_items.extend(fw_items)
    except Exception as e:
        print(f"    -> Free-Work scrape failed: {e}")

    # 2. Bing RSS search
    bing_queries = agent_cfg.get("bing_queries", [])
    print(f"  Recherche Bing: {len(bing_queries)} requetes...")
    try:
        bing_items = multi_search(bing_queries, timeout, count)
        print(f"    -> {len(bing_items)} resultats Bing")
        all_items.extend(bing_items)
    except Exception as e:
        print(f"    -> Bing search failed: {e}")

    # Score and rank web results
    scored = []
    for item in all_items:
        if is_blocked(item):
            continue
        s = calculer_score_mission(item, profil)
        if s:
            scored.append(s)

    # Add manual missions (always on top)
    manual = load_manual_missions()
    for m in manual:
        tjm_val = int(m.get("tjm", 0)) if str(m.get("tjm", "")).isdigit() else 0
        scored.insert(0, {
            "title": m.get("titre", "Mission manuelle"),
            "url": m.get("url", ""),
            "snippet": m.get("description", m.get("contexte", "")),
            "source": m.get("source", "manuel"),
            "score": 8.0,
            "tjm": f"{tjm_val}EUR" if tjm_val > 0 else "non specific",
            "tjm_value": tjm_val,
            "fit_skills": [],
            "decision": "POSTULER - Cible identifiee manuellement",
            "industry_match": [],
        })
    print(f"  Missions manuelles: {len(manual)}")

    scored.sort(key=lambda x: (x["score"], x["tjm_value"]), reverse=True)
    top = scored[:max_missions]
    print(f"  Missions retenues: {len(top)}")
    return top


def build_report(config: dict, missions: list[dict], output_path: Path):
    profil = config.get("profile", {})
    name = profil.get("name", "Michael ASSAYAG")
    title = profil.get("title", "")
    positioning = profil.get("positioning", "")
    tjm_cible = profil.get("tjm_cible", 750)
    targets = profil.get("target_roles", [])

    doc = setup_doc(
        "CHASSEUR DE MISSIONS - RAPPORT HEBDOMADAIRE",
        f"Cibles freelance identifiees pour {name} - {positioning}",
        "Assistant IA SASU",
    )

    # Executive summary
    para(doc, "1. Synthese executive", style="Heading 1")
    bullet(doc, f"Profil cible : {' | '.join(targets[:3])}")
    bullet(doc, f"TJM vise : {tjm_cible} EUR+")
    bullet(doc, f"Mission identifiees ce cycle : {len(missions)}")

    premium = [m for m in missions if m.get("tjm_value", 0) >= 700]
    bulletin = [m for m in missions if m["decision"] == "POSTULER EN PRIORITE" or m["decision"] == "POSTULER"]

    bullet(doc, f"  - Missions premium (TJM >= 700EUR) : {len(premium)}")
    bullet(doc, f"  - Recommandees en priorite : {len(bulletin)}")
    if premium:
        bullet(doc, f"  - TJM max detecte : {max(m['tjm_value'] for m in missions)} EUR")
    para(doc, "", after=6)

    # Missions detail
    para(doc, "2. Missions qualifiees", style="Heading 1")
    for idx, m in enumerate(missions, 1):
        decision_color = GREEN if "PRIORITE" in m["decision"] else (RGBColor(200, 120, 0) if m["decision"] == "POSTULER" else GRAY)
        para(doc, f"{idx}. {m['title']}", style="Heading 3")
        p = para(doc, after=2)
        set_run_font(p.add_run(f"Source : {m['source']} | Score : {m['score']}/10 | TJM : {m['tjm']} | "), size=9.5, color=GRAY, bold=True)
        add_hyperlink(p, "URL", m["url"])
        if m.get("snippet"):
            bullet(doc, f"Description : {m['snippet'][:500]}")
        if m.get("fit_skills"):
            bullet(doc, f"Fit competences : {', '.join(m['fit_skills'])}")
        if m.get("industry_match"):
            bullet(doc, f"Secteur concordant : {', '.join(m['industry_match'])}")
        p = para(doc, after=6)
        set_run_font(p.add_run(f"Decision : {m['decision']}"), bold=True, color=decision_color)

    # Recommandations
    para(doc, "3. Recommandations de la semaine", style="Heading 1")
    if bulletin:
        para(doc, "Missions a cibler cette semaine :", style="Heading 2")
        for m in bulletin[:3]:
            bullet(doc, f"{m['title']} - {m['source']} (TJM: {m['tjm']})")
    else:
        bullet(doc, "Aucune mission prioritaire detectee ce cycle.")

    if premium:
        bullet(doc, f"TJM moyen des missions premium : {sum(m['tjm_value'] for m in premium)/len(premium):,.0f} EUR")
    bullet(doc, f"Votre TJM cible ({tjm_cible} EUR) est tenable sur le marche actuel." if premium else "Elargir les criteres de recherche pour trouver plus d'opportunites.")

    para(doc, "4. Profil valorise", style="Heading 1")
    for client in profil.get("key_clients", []):
        bullet(doc, f"Reference : {client}")
    for skill in profil.get("key_skills", [])[:6]:
        bullet(doc, f"Competence cle : {skill}")

    doc.save(output_path)
    return output_path


def main() -> int:
    config = load_config()
    agent_cfg = config.get("chasseur_missions", {})
    if not agent_cfg.get("enabled", True):
        print("Agent Chasseur de Missions desactive.")
        return 0

    print(">>> Agent 1 : Chasseur de Missions")
    profil = config.get("profile", {})
    missions = collect_missions(agent_cfg, profil)

    output_dir = ensure_output_dir(config)
    today = date.today().isoformat()
    output_path = output_dir / f"Chasseur_Missions_{today}.docx"
    build_report(config, missions, output_path)

    run_log = output_dir / f"Chasseur_Missions_{today}_log.json"
    save_json_report(
        {
            "generated_at": datetime.now(timezone.utc).isoformat(),
            "missions_count": len(missions),
            "missions": missions,
        },
        run_log,
    )
    print(f"  Rapport : {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

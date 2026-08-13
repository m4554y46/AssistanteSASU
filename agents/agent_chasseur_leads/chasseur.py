from __future__ import annotations

import csv
import random
import re
import sys
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

from core.common import (
    accent_heading,
    add_callout,
    add_hyperlink,
    add_separator,
    bullet,
    ensure_output_dir,
    load_config,
    make_pro_table,
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
    "cabinets-conseil.com", "edcparis.edu", "consultport.com",
    "linkup-coaching.com", "scrum.org",
    "producthunt", "keejob", "optioncarriere", "acc.tn", "lci.com.tn",
    "indeed", "glassdoor", "monster", "kijiji", "emploitic", "tanitjobs",
    "kejobb", "paysage-tunisie",
]

BLOCKED_ROLES = [
    "cobol", "mainframe", "as400", "rust", "developpeur back",
    "charge de recrutement", "ingenieur reseau", "analyste mainframe",
    "integrateur devops", "ingenieur devops", "mission locale",
    "software engineer", "data scientist", "data engineer", "llm engineer",
    "tech lead", "full stack", "fullstack", "front end", "frontend",
    "back end", "backend", "developpeur", "developpement", "ingenieur",
    "java", "angular", "react", "node.js", "python", "sap", "erp developer",
    "devops", "sysadmin", "administrateur systeme", "testeur", "qa engineer",
    "designer", "graphiste", "ux designer", "art director",
    "data analyst", "analyste donnees", "dba", "développeur",
    "analyste metier", "analyste programmeur", "concepteur", "integrateur",
    "gestionnaire de service", "admin", "administrateur", "expert technique",
    "business analyst", "architecte",
    "telecom", "infrastructure", "production", "cybersecurite", "securite it",
    "compliance", "conformite", "audit", "risques it", "grc", "risque operationnel",
    "paie", "rh ", "ressources humaines", "recrutement", "drh",
]

BLOCKED_WORDS = [
    "definition", "fiche métier", "fiche metier",
    "coaching", "coach",
    "qu'est-ce qu'un", "quest-ce quun",
    "what is a", "what is an", "guide complet",
    "c'est quoi", "formation", "cours",
    "trouvez les meilleurs", "plateforme",
    "comparatif", "comparaison",
    "outils de", "logiciel de",
]


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


# Slogans et appels publicitaires qui ne sont jamais une vraie description de mission
SLUG_PHRASES = [
    "decouvrez", "trouvez votre prochain poste", "postulez facilement",
    "les meilleurs", "offres d'emploi", "tous les postes a pourvoir",
    "en une seule recherche", "consultez nos", "consulte nos",
    "the best new products", "every day", "everyone's talking about",
    "curation of the best", "browse", "sign up", "join now",
    "meilleures offres d'emploi", "de nouvelles offres", "offres du jour",
    "utilisez votre reseau professionnel",
]


def _clean_description(raw: str, max_len: int = 320) -> str:
    """Nettoie un snippet : supprime les slogans, ne garde que des phrases
    completes (fini les bouts de phrases tronques en plein milieu)."""
    if not raw:
        return ""
    text = raw.strip()
    if not text:
        return ""

    # Suppression des slogans / contenus publicitaires
    lower = text.lower()
    for phrase in SLUG_PHRASES:
        if phrase in lower:
            return ""

    # Decoupage en phrases : on ne garde que celles qui se terminent
    # par une ponctuation reelle (point, point d'exclamation, points de
    # suspension) et pas un point interne comme dans "Node.js".
    sentences = re.split(r"(?<=[.!?])\s+", text)
    kept = []
    for sent in sentences:
        s = sent.strip()
        if not s:
            continue
        if not re.search(r"[.!?]\s*$", s):
            break
        kept.append(s)
    text = " ".join(kept)
    if len(text) < 25:
        return ""
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) > max_len:
        text = text[:max_len].rstrip() + "…"
    return text


def _clean_title(raw: str, max_len: int = 120) -> str:
    """Nettoie un titre : enleve les prefixes 'Mission freelance 103025/',
    les codes departement, les localisations repetees et le nom de la
    plateforme en fin de chaine."""
    if not raw:
        return ""
    text = raw.strip()
    # Prefixe "Entreprise – Mission freelance <id>/" ou "Entreprise – Mission freelance "
    text = re.sub(r"^.*?[-–]\s*Mission\s*freelance\s*(\d*/)?", "", text)
    # Suffixe " | Free-work", " | Free-Work"
    text = re.sub(r"\s*[|]\s*free-?work.*$", "", text, flags=re.I)
    # Codes departement (63), region (63)
    text = re.sub(r"\s*\(\d{1,3}\)", "", text)
    # Dedoublonne les mots consecutifs identiques (Clermont-Ferrand Clermont-Ferrand)
    text = re.sub(r"\b(\w+(?:-\w+)*)\s+\1\b", r"\1", text)
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) > max_len:
        text = text[:max_len].rstrip() + "…"
    return text


def is_blocked(item: SearchResult) -> bool:
    if any(d in item.source.lower() for d in BLOCKED_DOMAINS):
        return True
    title_lower = item.title.lower()
    # Roles bloques : on regarde le TITRE du poste uniquement, pas la
    # description (sinon "Product Owner" avec mention "production" dans
    # le texte serait rejete a tort).
    if any(role in title_lower for role in BLOCKED_ROLES):
        return True
    # Mots "contenu" (formation, cours...) : titre uniquement, car une
    # vraie description de mission peut legitiment les mentionner.
    if any(w in title_lower for w in BLOCKED_WORDS):
        return True
    text = f"{item.title} {item.snippet}".lower()
    if any(phrase in text for phrase in SLUG_PHRASES):
        return True
    if not _is_latin(item.title) or (item.snippet and not _is_latin(item.snippet)):
        return True
    return False


def calculer_score_mission(item: SearchResult, profil: dict) -> dict | None:
    text = f"{item.title} {item.snippet} {item.source}".lower()
    url = item.url.lower()

    # Vrais roles : ceux qui correspondent au profil de Michael. Seuls ceux-la
    # donnent un "fit competences" presentable a un humain.
    role_keywords = [
        "directeur de projet", "chef de projet", "head of product",
        "product manager", "product owner", "product management",
        "program manager", "delivery manager", "maîtrise d'ouvrage",
        "assistance maîtrise d'ouvrage", "product lead", "chief product",
        "consultant", "consulting", "strategy", "strategie", "transformation",
        "scrum master", "coach agile", "agile coach", "manager de programme",
    ]
    # Signaux faibles : confirment qu'on est bien sur une mission (pas la qualifient)
    mission_signals = ["freelance", "mission", "tjm", "job-mission", "consultant"]

    industries = [ind.lower() for ind in profil.get("industries", [])]

    role_hits = [kw for kw in role_keywords if kw in text]
    industry_hits = [ind for ind in industries if ind in text]
    has_mission_signal = any(sig in text for sig in mission_signals)

    is_freework_direct = "free-work.com" in item.source and "/job-mission/" in url

    # Un vrai resultat doit avoir un role cible, ou etre une page Free-Work directe
    # (mission certaine par construction) - mais meme la, on exige un role ou un secteur.
    if not role_hits:
        if not (is_freework_direct and industry_hits):
            return None

    # Les pages type "annuaire/portail emploi" ne sont jamais des missions
    if not is_freework_direct and not re.search(r"/job|/mission|/emploi|/offre", url):
        return None

    tjm_match = re.search(r"tjm[:\s]*(\d{3,4})|minDailySalary[:\"]+(\d{3,4})|maxDailySalary[:\"]+(\d{3,4})", text, re.I)
    tjm_extracted = 0
    for g in tjm_match.groups() if tjm_match else []:
        if g:
            tjm_extracted = max(tjm_extracted, int(g))
    # Le TJM peut venir du champ structure Free-Work
    if item.tjm:
        tjm_extracted = max(tjm_extracted, item.tjm)

    raw_score = 3.0
    raw_score += min(len(role_hits) * 0.9, 3.2)
    raw_score += min(len(industry_hits) * 0.6, 1.8)
    if has_mission_signal:
        raw_score += 0.3

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
        tjm_label = f"{tjm_extracted}EUR" if tjm_extracted > 0 else "non communiqué"

    score = max(0.5, min(9.5, round(raw_score, 1)))

    if score >= 7.0 and tjm_extracted >= 650:
        decision = "POSTULER EN PRIORITE"
    elif score >= 5.5:
        decision = "POSTULER"
    elif score >= 3.5:
        decision = "A QUALIFIER"
    else:
        decision = "SURVEILLER"

    # Fit competences : uniquement les vrais roles, pas les mots generiques
    fit_skills = role_hits[:5]
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
        "company": item.company,
        "location": item.location,
        "duration": item.duration,
    }


def load_manual_missions() -> list[dict]:
    csv_path = Path(__file__).resolve().parent / "missions_manuelles.csv"
    missions = []
    if not csv_path.exists():
        return missions
    with csv_path.open(newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f, delimiter=";"):
            if not row.get("url", "").strip():
                continue
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
            "tjm": f"{tjm_val}EUR" if tjm_val > 0 else "non communiqué",
            "tjm_value": tjm_val,
            "fit_skills": [],
            "decision": "POSTULER - Cible identifiee manuellement",
            "industry_match": [],
            "company": m.get("entreprise", "").strip(),
            "location": m.get("localisation", "").strip(),
            "duration": m.get("duree", "").strip(),
        })
    print(f"  Missions manuelles: {len(manual)}")

    # Dedup by title (first 80 chars)
    seen_titles = set()
    deduped = []
    for m in scored:
        key = m["title"].strip().lower()[:80]
        if key not in seen_titles:
            seen_titles.add(key)
            deduped.append(m)
    print(f"  Apres deduplication: {len(deduped)}")

    deduped.sort(key=lambda x: (x["score"], x["tjm_value"]), reverse=True)
    top = deduped[:max_missions]
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
        config.get("author", "Virginie Benayoun"),
    )

    # Executive summary
    accent_heading(doc, "1. Synthese executive")

    premium = [m for m in missions if m.get("tjm_value", 0) >= 700]
    bulletin = [m for m in missions if m["decision"] == "POSTULER EN PRIORITE" or m["decision"] == "POSTULER"]

    summary_lines = (
        f"Profil cible : {' | '.join(targets[:3])}\n"
        f"TJM vise : {tjm_cible} EUR+\n"
        f"Missions identifiees ce cycle : {len(missions)}\n"
        f"Missions premium (TJM >= 700EUR) : {len(premium)}\n"
        f"Recommandees en priorite : {len(bulletin)}"
    )
    if premium:
        summary_lines += f"\nTJM max detecte : {max(m['tjm_value'] for m in missions)} EUR"
    add_callout(doc, summary_lines, title="Synthese")

    # Missions detail
    accent_heading(doc, "2. Missions qualifiees")
    for idx, m in enumerate(missions, 1):
        decision_color = GREEN if "PRIORITE" in m["decision"] else (RGBColor(200, 120, 0) if m["decision"] == "POSTULER" else GRAY)
        para(doc, f"{idx}. {_clean_title(m['title'])}", style="Heading 3")

        meta_parts = [f"Source : {m['source']}", f"Score : {m['score']}/10", f"TJM : {m['tjm']}"]
        p = para(doc, after=2)
        set_run_font(p.add_run(" | ".join(meta_parts)), size=9.5, color=GRAY, bold=True)

        if m.get("url"):
            p = para(doc, after=2)
            set_run_font(p.add_run("Lien : "), size=9.5, color=GRAY)
            add_hyperlink(p, m["url"], m["url"])

        detail_lines = []
        if m.get("company"):
            detail_lines.append(f"Entreprise : {m['company']}")
        if m.get("location"):
            detail_lines.append(f"Localisation : {m['location']}")
        if m.get("duration"):
            detail_lines.append(f"Duree : {m['duration']}")
        for line in detail_lines:
            bullet(doc, line)

        desc = _clean_description(m.get("snippet"))
        if desc:
            bullet(doc, f"Description : {desc}")

        if m.get("fit_skills"):
            bullet(doc, f"Fit competences : {', '.join(m['fit_skills'])}")
        if m.get("industry_match"):
            bullet(doc, f"Secteur concordant : {', '.join(m['industry_match'])}")
        p = para(doc, after=6)
        set_run_font(p.add_run(f"Decision : {m['decision']}"), bold=True, color=decision_color)

    # Recommandations
    add_separator(doc)
    accent_heading(doc, "3. Recommandations de la semaine")
    if bulletin:
        para(doc, "Missions a cibler cette semaine :", style="Heading 2")
        for m in bulletin[:3]:
            bullet(doc, f"{_clean_title(m['title'])} - {m['source']} (TJM: {m['tjm']})")
    else:
        bullet(doc, "Aucune mission prioritaire detectee ce cycle.")

    if premium:
        bullet(doc, f"TJM moyen des missions premium : {sum(m['tjm_value'] for m in premium)/len(premium):,.0f} EUR")
        bullet(doc, f"Ton TJM cible ({tjm_cible} EUR) reste coherent avec le marche actuel.")
    elif missions:
        bullet(doc, "Cette semaine, aucune mission avec un TJM annonce dans la fourchette visee, mais plusieurs pistes pertinentes a explorer.")
        bullet(doc, "Je continue de surveiller les nouvelles publications la semaine prochaine.")
    else:
        bullet(doc, "Cette semaine, rien de pertinent n'a ete publie sur les plateformes surveillees.")
        bullet(doc, "Je continue de surveiller les nouvelles publications la semaine prochaine.")

    add_separator(doc)
    accent_heading(doc, "4. Profil valorise")
    for client in profil.get("key_clients", []):
        bullet(doc, f"Reference : {client}")
    for skill in profil.get("key_skills", [])[:6]:
        bullet(doc, f"Competence cle : {skill}")

    doc.save(output_path)
    return output_path


INTRO_VARIATIONS = [
    "Je me permets de vous contacter pour vous proposer un profil qui pourrait correspondre a vos besoins.",
    "Je souhaitais vous presenter un profil senior en gestion de projet et transformation digitale.",
    "Dans le cadre de mon activite chez ASTRA MOMENTUM, je me permets de vous proposer un accompagnement sur vos projets produit et digitaux.",
]


def generate_prospection_pack(config: dict, missions: list[dict], output_dir: Path):
    agent_cfg = config.get("chasseur_missions", {})
    prospection_cfg = config.get("prospection", {})
    if not prospection_cfg.get("enabled", True):
        print("  Prospection pack desactive.")
        return

    min_score = prospection_cfg.get("min_score", 5.0)
    profil = config.get("profile", {})

    target_missions = [m for m in missions if m["score"] >= min_score and m.get("company") and m.get("url")]

    leads_path = Path(__file__).resolve().parent / "leads_manuels.csv"
    if leads_path.exists():
        with leads_path.open(newline="", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f, delimiter=";"):
                company = row.get("entreprise", "").strip()
                url = row.get("url", "").strip()
                if company and url:
                    target_missions.append({
                        "title": row.get("contexte", "Lead manuel"),
                        "url": url,
                        "snippet": "",
                        "source": row.get("source", ""),
                        "score": 10.0,
                        "tjm": "non communiqué",
                        "tjm_value": 0,
                        "fit_skills": [],
                        "decision": "CONTACTER - Lead identifie manuellement",
                        "industry_match": [],
                        "company": company,
                    })

    if not target_missions:
        print("  Aucune mission qualifiee pour le prospection pack.")
        return

    today = date.today().isoformat()
    doc = setup_doc(
        "PROSPECTION PACK - EMAILS DE CONTACT",
        f"Drafts d'emails pour les missions cibles - {today}",
        config.get("author", "Virginie Benayoun"),
    )

    accent_heading(doc, "Emails de prospection generes")
    add_callout(doc,
        f"Profil propose : {profil.get('name', 'Michael ASSAYAG')}\n"
        f"Positionnement : {profil.get('positioning', '')}\n"
        f"Contacts generes : {len(target_missions)}",
        title="Prospection Pack"
    )

    tracking_data = []
    random.shuffle(INTRO_VARIATIONS)
    intro_idx = 0

    for i, m in enumerate(target_missions, 1):
        company = m.get("company", "")
        if not company:
            continue

        if intro_idx >= len(INTRO_VARIATIONS):
            intro_idx = 0
        intro = INTRO_VARIATIONS[intro_idx]
        intro_idx += 1

        subject = f"ASTRA MOMENTUM - Accompagnement {company} / Produit & Digital"
        body = f"""Bonjour,

{intro}

Michael ASSAYAG, Head of Product & Digital Transformation chez ASTRA MOMENTUM, intervient aupres des directions produit et digitales pour leurs projets de transformation. Il a une double expertise :
- Product Management & Delivery (Kering - Gucci, Saint Laurent - 7 ans)
- Deploiement mobile worldwide (GeoPost / DPD - equipe 9+ pays)
- Innovation & strategie digitale (Saint-Gobain - AR, 3D, ERP)

Il intervient en freelance sur des missions de :
- Direction de produit / Head of Product (interim ou conseil)
- Transformation digitale et organisation produit
- Conseil aupres des dirigeants de PME/ETI

Je reste a votre disposition pour echanger sur vos besoins actuels ou a venir.

Bien cordialement,
Virginie Benayoun
Responsable commerciale - ASTRA MOMENTUM"""

        para(doc, f"Email {i} - {company}", style="Heading 2")
        p = para(doc, after=2)
        set_run_font(p.add_run(f"Mission : {m['title']}"), bold=True, size=10)
        if m.get("url"):
            p = para(doc, after=2)
            set_run_font(p.add_run("URL : "), size=9.5, color=GRAY)
            add_hyperlink(p, m["url"], m["url"])
        p = para(doc, after=2)
        set_run_font(p.add_run(f"Objet : {subject}"), size=9.5, color=GRAY)
        para(doc, body, after=6)
        para(doc, "\u2500" * 60, after=6)

        tracking_data.append({
            "date_generated": today,
            "company": company,
            "mission_title": m["title"],
            "mission_url": m.get("url", ""),
            "subject": subject,
            "sent": "",
            "date_sent": "",
        })

    output_path = output_dir / f"Prospection_Pack_{today}.docx"
    doc.save(output_path)
    print(f"  Prospection pack : {output_path}")

    tracking_path = output_dir / "prospection_tracking.csv"
    existing_rows = []
    if tracking_path.exists():
        with tracking_path.open(newline="", encoding="utf-8-sig") as f:
            existing_rows = list(csv.DictReader(f, delimiter=";"))

    seen = {(r.get("company", ""), r.get("mission_url", "")) for r in existing_rows}
    for row in tracking_data:
        key = (row["company"], row["mission_url"])
        if key not in seen:
            existing_rows.append(row)
            seen.add(key)

    with tracking_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["date_generated", "company", "mission_title", "mission_url", "subject", "sent", "date_sent"], delimiter=";")
        writer.writeheader()
        writer.writerows(existing_rows)
    print(f"  Tracking CSV : {tracking_path}")


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

    generate_prospection_pack(config, missions, output_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

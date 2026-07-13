from __future__ import annotations

import random
from datetime import date, datetime, timezone
from pathlib import Path

from core.web import SearchResult, multi_search

NOTES_DIR = Path(__file__).resolve().parent / "notes"

MARCHE_QUERIES = [
    '"product manager" freelance mission 2026 TJM',
    '"directeur de projet" freelance mission 2026',
    '"product owner" freelance mission 2026',
    'free-work mission "chef de projet" IT',
    'freelance IT TJM product manager 2026',
]

ECOSYSTEM_QUERIES = [
    'Indy logiciel comptabilité indépendant 2026',
    'Dougs expert comptable indépendant avis 2026',
    'meilleur compte professionnel freelance 2026 comparatif',
    'Shine compte pro indépendant 2026',
    'Qonto banque en ligne professionnelle 2026',
    'formation certifiante freelance management 2026',
    'mutuelle freelance independant comparatif 2026',
    'prevoyance assurance freelance independant 2026',
]

IA_PRODUCT_QUERIES = [
    'agentic IA product management transformation',
    'continuous discovery product management 2026',
    'innovation transformation digitale entreprise 2026',
]

RANDOM_TOPICS = [
    'facture électronique obligatoire freelance 2026',
    'Malt freelance plateforme mission 2026',
    'Deel plateforme paie internationale freelance',
    'TVA franchise indépendant 2026 seuil',
    'Alan mutuelle freelance avis 2026',
    'Spirica prévoyance indépendant 2026',
    'Blank compta freelance 2026',
    'frais kilométrique freelance SASU 2026',
]


def _seed() -> int:
    return int(date.today().strftime("%Y%U"))


BLOCKED_DOMAINS = [
    "pinterest", "amazon", "ebay", "etsy", "aliexpress", "walmart",
    "shopify", "boulanger", "fnac", "cdiscount", "ikea", "leroymerlin",
    "decathlon", "booking", "tripadvisor",
    "larousse", "lerobert", "dictionnaire", "wiktionary", "cnrtl",
    "linternaute", "wikihow", "wikipedia",
    "fiverr", "freelance.com", "upwork", "peopleperhour",
    "mission-emploi", "indeed", "hellowork",
    "allocine", "facebook", "instagram", "twitter", "x.com", "tiktok",
    "mozzartbet", "bet365", "parionssport", "poker",
    "youtube", "tiktok",
    "leboncoin", "aujourdhui",
    "wordreference", "producthunt", "linguee", "reverso",
    "cambridge", "merriam", "oxford", "collins",
    "synonymo",
]
BLOCKED_WORDS = [
    "login", "sign in", "se connecter", "s identifier", "password",
    "connexion", "subscribe", "inscription", "create account",
    "earplugs", "bouchons", "oreille", "shoes", "chaussure",
    "recrutement", "annonce recrute", "nous recrutons",
    "www.", "http://", "https://",
    "trouvez les meilleurs freelances",
    "freelance services marketplace",
    "postuler à cette mission",
    "définition", "definition", "définitions", "definitions",
    "mission locale", "recrute", "recrute", "recrutons",
]


def _search_titles(raw: list[SearchResult], max_items: int = 4) -> list[str]:
    fake_names = {chr(i) * 5 for i in range(97, 123)}
    titles = []
    for item in raw:
        if not item.title:
            continue
        source = item.source.lower()
        title = item.title.lower()
        snippet = (item.snippet or "").lower()
        text = f"{title} {snippet}"
        if any(d in source for d in BLOCKED_DOMAINS):
            continue
        if any(w in text for w in BLOCKED_WORDS):
            continue
        if item.title.strip() in fake_names:
            continue
        if len(item.title) < 10:
            continue
        titles.append(f"{item.title} ({item.source})")
    seen: set[str] = set()
    unique = []
    for t in titles:
        key = t.lower()[:80]
        if key not in seen:
            seen.add(key)
            unique.append(t)
    return unique[:max_items]


def _date_fr() -> str:
    jours = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"]
    mois = ["", "janvier", "février", "mars", "avril", "mai", "juin",
            "juillet", "août", "septembre", "octobre", "novembre", "décembre"]
    t = date.today()
    return f"{jours[t.weekday()]} {t.day} {mois[t.month]} {t.year}"


def _heure_aleatoire() -> str:
    rng = random.Random(_seed())
    heure = rng.choice([11, 12, 13])
    minute = rng.randint(0, 59)
    if heure == 11 and minute < 30:
        minute = 30
    if heure == 14:
        minute = min(minute, 0)
    return f"{heure:02d}:{minute:02d}"


def _duree() -> str:
    rng = random.Random(_seed() + 1)
    minutes = rng.randint(25, 35)
    return f"{minutes} min"


def _generer_market_notes(articles: list[str]) -> str:
    if not articles:
        return random.choice([
            "  - Rien de marquant sur le marche cette semaine, j'ai principalement vu passer des annonces Free-Work classiques.\n",
            "  - Semaine calme sur le marche freelance, rien de particulier a signaler.\n",
            "  - Pas de movement notable cette semaine sur le marche des missions IT.\n",
        ])

    notes = ""
    has_paris = any("paris" in a.lower() or "france" in a.lower() for a in articles)
    has_suisse = any("suisse" in a.lower() or "geneve" in a.lower() or "zurich" in a.lower() for a in articles)
    has_us = any("us " in a.lower() or "american" in a.lower() or "states" in a.lower() or "usa" in a.lower() for a in articles)

    if has_paris:
        notes += "  - Le marché parisien est stable cette semaine, rien de fou.\n"
    else:
        notes += "  - Pas de gros mouvements à Paris, les missions sortent au compte-gouttes.\n"

    if has_suisse:
        notes += "  - Côté Suisse, quelques offres notées mais les TJM sont souvent sous les 700 CHF.\n"
    else:
        notes += "  - Suisse : toujours compliqué de trouver du 100% remote depuis la France.\n"

    if has_us:
        notes += "  - US : j'ai vu passer des missions, mais le décalage horaire reste un frein.\n"
    else:
        notes += "  - US : rien de concret cette semaine, je n'ai pas creusé.\n"

    top = articles[:2]
    if top:
        notes += "  - " + " | ".join(top) + "\n"

    return notes


def _generer_ecosystem_notes(articles: list[str]) -> str:
    if not articles:
        return "  - Pas de nouvelles marquantes sur l'écosystème freelance cette semaine.\n"

    notes = ""
    tools = [a for a in articles if any(x in a.lower() for x in ["indy", "dougs", "shine", "qonto", "deel"])]
    if tools:
        notes += "  - Côté outils : " + " | ".join(tools[:2]) + "\n"
    else:
        notes += "  - Rien de neuf côté outils de compta/banque cette semaine.\n"

    other = [a for a in articles if a not in tools][:2]
    for a in other:
        notes += f"  - {a}\n"

    if not notes:
        notes += "  - Je n'ai pas trouvé d'articles pertinents cette semaine.\n"

    return notes


def _generer_topics_notes(articles: list[str]) -> str:
    if not articles:
        return ""
    sep = random.Random(_seed() + 2).choice(["\n", ""])
    lines = []
    for a in articles[:2]:
        lines.append(f"  - {a}")
    return "\n".join(lines) + "\n"


def _generer_actions() -> str:
    rng = random.Random(_seed() + 3)
    actions = [
        "  - [Michael] Vérifier les annonces Free-Work pour la semaine",
        "  - [Michael] Envoyer la newsletter ou le post LinkedIn prévu",
        "  - [Virginie] Préparer la veille tarifaire pour la semaine prochaine",
        "  - [Michael] Relancer les contacts en cours",
        "  - [Virginie] Mettre à jour le tableau de bord des factures",
        "  - [Michael] Suivre les réponses aux candidatures envoyées",
        "  - [Virginie] Chercher des ressources formation/conférence intéressantes",
    ]
    rng.shuffle(actions)
    selected = actions[:rng.randint(2, 4)]
    return "\n".join(selected) + "\n"


def _generer_remarque_libre() -> str:
    rng = random.Random(_seed() + 4)
    remarques = [
        "  - Rien de spécial à signaler par ailleurs.",
        "  - J'ai noté que les annonces Free-Work sont moins nombreuses que la semaine dernière.",
        "  - J'ai vu passer une offre intéressante sur LinkedIn, je te l'ai envoyée.",
        "  - Pense à checker tes notifications Malt, il y a peut-être des invitations.",
        "  - J'ai relu le CR de la semaine dernière, les actions sont toujours en cours.",
        "  - J'ai commencé à regarder les formations, je te fais une sélection demain.",
        "  - Pas de nouvelles de Dougs/Indy cette semaine, tout est calme.",
        "",
    ]
    return rng.choice(remarques)


def generate() -> str:
    aujourd_hui = _date_fr()
    heure = _heure_aleatoire()
    duree = _duree()

    rng = random.Random(_seed())
    timeout = 12
    count = 5

    print(f"  Recherche Bing pour le CR hebdo ({aujourd_hui})...")

    raw_marche = multi_search(MARCHE_QUERIES, timeout, count)
    raw_eco = multi_search(ECOSYSTEM_QUERIES, timeout, count)
    raw_ia = multi_search(IA_PRODUCT_QUERIES, timeout, count)
    raw_topics = multi_search(RANDOM_TOPICS, timeout, count)

    articles_marche = _search_titles(raw_marche, 3)
    articles_eco = _search_titles(raw_eco, 4)
    articles_ia = _search_titles(raw_ia, 2)
    articles_topics = _search_titles(raw_topics, 3)

    notes = f"""Client: ASTRA MOMENTUM
Projet: Point hebdo freelance
Date: {date.today().isoformat()}
Participants: Michael ASSAYAG, Virginie Benayoun
Durée: {duree}
---
Contexte:
Point hebdomadaire du {aujourd_hui} ({heure}). Passage en revue du marché freelance, de l'écosystème et des actions en cours.

Marché freelance:
{_generer_market_notes(articles_marche)}
Écosystème freelance (outils, banques, compta, formations):
{_generer_ecosystem_notes(articles_eco)}
Veille techno / Product / IA:
"""

    if articles_ia:
        notes += "".join(f"  - {a}\n" for a in articles_ia)
    else:
        notes += "  - Rien de neuf côté veille techno cette semaine.\n"

    extra = _generer_topics_notes(articles_topics)
    if extra.strip():
        notes += f"\nSujets divers:\n{extra}"

    notes += f"""
Décisions:
  - On continue la prospection ciblée sur les profils Head of Product et Directeur de projet IT
  - Priorité aux missions avec TJM >= 650€
  - Vérifier la prochaine échéance TVA
  - Relancer les contacts qui n'ont pas répondu la semaine dernière

Actions:
{_generer_actions()}
Remarques:
{_generer_remarque_libre()}
"""
    return notes


def main() -> int:
    NOTES_DIR.mkdir(parents=True, exist_ok=True)
    notes = generate()
    filename = f"point_hebdo_{date.today().isoformat()}.txt"
    path = NOTES_DIR / filename
    path.write_text(notes, encoding="utf-8")
    print(f"  Notes generees : {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

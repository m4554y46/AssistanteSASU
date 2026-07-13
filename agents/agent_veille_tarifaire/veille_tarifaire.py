from __future__ import annotations

import random
import re
from collections import defaultdict
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

from core.common import (
    accent_heading,
    add_callout,
    add_separator,
    apply_table_borders,
    bullet,
    ensure_output_dir,
    load_config,
    make_pro_table,
    para,
    save_json_report,
    set_run_font,
    set_table_widths,
    shade_cell,
    setup_doc,
)
from core.web import SearchResult, multi_search, collect_freework_jobs
from docx.shared import Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT

NAVY = RGBColor(11, 37, 69)
GRAY = RGBColor(86, 95, 108)
BLACK = RGBColor(0, 0, 0)
GREEN = RGBColor(0, 128, 0)
RED = RGBColor(155, 28, 28)
ORANGE = RGBColor(200, 120, 0)
PALEBLUE = "E8EEF5"

PROFILS = [
    {"nom": "Product Owner", "mots": ["product owner", "product-owner", "po"]},
    {"nom": "Product Manager", "mots": ["product manager", "product-manager", "pm"]},
    {"nom": "Chef de projet IT", "mots": ["chef de projet", "chef-de-projet", "cdp"]},
    {"nom": "Consultant IA / Data", "mots": ["consultant ia", "consultant ai", "data", "ia", "intelligence artificielle"]},
    {"nom": "Scrum Master / Agile Coach", "mots": ["scrum master", "scrum-master", "agile coach", "agile-coach"]},
]


def extraire_tjm(text: str) -> list[int]:
    patterns = [
        r"(?:TJM|tjm|taux journalier)\s*(?:moyen|moyenne)?\s*(?:de|:)?\s*(\d{3,4})\s*(?:EUR|€|euros)?",
        r"(\d{3,4})\s*(?:EUR|€|euros)?\s*(?:/|par)\s*jour",
        r"(?:entre\s+)?(\d{3,4})\s*(?:et|a)\s*(\d{3,4})\s*(?:EUR|€)",
    ]
    values = []
    for pat in patterns:
        for match in re.finditer(pat, text, re.I):
            if match.lastindex and match.lastindex >= 2:
                v1, v2 = int(match.group(1)), int(match.group(2))
                values.extend([v1, v2])
            else:
                v = int(match.group(1))
                values.append(v)
    return [v for v in values if 200 <= v <= 1500]


def collecter_donnees(agent_cfg: dict) -> dict[str, dict]:
    timeout = int(agent_cfg.get("timeout_seconds", 12))
    count = int(agent_cfg.get("results_per_query", 6))
    queries = agent_cfg.get("queries", [])
    freework_pages = agent_cfg.get("freework_pages", [])

    all_raw: list[SearchResult] = []

    if freework_pages:
        print(f"  Scraping {len(freework_pages)} pages Free-Work...")
        try:
            fw_items = collect_freework_jobs(freework_pages, timeout)
            print(f"    -> {len(fw_items)} offres trouvees")
            all_raw.extend(fw_items)
        except Exception as e:
            print(f"    -> Free-Work scrape failed: {e}")

    print(f"  Recherche Bing: {len(queries)} requetes...")
    try:
        raw = multi_search(queries, timeout, count)
        print(f"    -> {len(raw)} resultats")
        all_raw.extend(raw)
    except Exception as e:
        print(f"    -> Bing search failed: {e}")

    print(f"  Total resultats: {len(all_raw)}")

    profils_data = {p["nom"]: {"results": [], "tjms": []} for p in PROFILS}

    for item in all_raw:
        text = f"{item.title} {item.snippet}".lower()
        tjms = extraire_tjm(text + " " + item.url)
        for profil in PROFILS:
            if any(m in text for m in profil["mots"]):
                profils_data[profil["nom"]]["results"].append(item)
                profils_data[profil["nom"]]["tjms"].extend(tjms)

    return profils_data


def analyser(profils_data: dict, config: dict) -> list[dict]:
    sasu = config.get("sasu", {})
    tjm_cible = int(sasu.get("tjm_cible", 650))
    tjm_min = int(sasu.get("tjm_min", 550))

    analyses = []
    for nom, data in profils_data.items():
        tjms = data["tjms"]
        if not tjms:
            analyses.append({
                "profil": nom,
                "resultats": len(data["results"]),
                "tjms_trouves": 0,
                "tjm_moyen": 0,
                "tjm_min_trouve": 0,
                "tjm_max_trouve": 0,
                "recommandation": "Donnees insuffisantes pour ce profil",
                "alerte": False,
            })
            continue

        tjm_moyen = round(sum(tjms) / len(tjms), 0)
        tjm_min_t = min(tjms)
        tjm_max_t = max(tjms)

        if tjm_moyen < tjm_min:
            recommandation = f"Le marche est sous votre TJM plancher ({tjm_min} EUR). Verifiez si votre positionnement est correct."
            alerte = True
        elif tjm_moyen >= tjm_cible:
            recommandation = f"Votre TJM cible ({tjm_cible} EUR) est dans la fourchette haute du marche. Maintenez."
            alerte = False
        else:
            recommandation = f"Le marche est entre {tjm_min_t:,.0f} et {tjm_max_t:,.0f} EUR. Vous pouvez envisager une augmentation progressive."
            alerte = False

        analyses.append({
            "profil": nom,
            "resultats": len(data["results"]),
            "tjms_trouves": len(tjms),
            "tjm_moyen": tjm_moyen,
            "tjm_min_trouve": tjm_min_t,
            "tjm_max_trouve": tjm_max_t,
            "recommandation": recommandation,
            "alerte": alerte,
        })

    return analyses


def build_report(config: dict, analyses: list[dict], output_path: Path):
    sasu = config.get("sasu", {})
    tjm_cible = int(sasu.get("tjm_cible", 650))
    tjm_min = int(sasu.get("tjm_min", 550))

    doc = setup_doc(
        "VEILLE CONCURRENTIELLE TARIFAIRE",
        "Analyse des TJM pratiques sur le marche freelance IT",
        config.get("author", "Virginie Benayoun"),
    )

    accent_heading(doc, "1. Position actuelle")
    add_callout(doc,
        f"Votre TJM cible actuel : {tjm_cible} EUR\n"
        f"Votre TJM plancher : {tjm_min} EUR\n"
        f"Positionnement : {sasu.get('positioning', '')}",
        title="Position tarifaire"
    )

    add_separator(doc)
    accent_heading(doc, "2. Analyse par profil")

    table = doc.add_table(rows=1, cols=6)
    headers = ["Profil", "TJM moyen", "Min", "Max", "Echantillon", "Recommandation"]
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        shade_cell(cell, PALEBLUE)
        cell.text = h
        for p in cell.paragraphs:
            for r in p.runs:
                set_run_font(r, bold=True, color=NAVY, size=8.5)

    for a in analyses:
        cells = table.add_row().cells
        vals = [
            a["profil"],
            f"{a['tjm_moyen']:,.0f} EUR" if a["tjm_moyen"] > 0 else "N/A",
            f"{a['tjm_min_trouve']:,.0f} EUR" if a["tjm_min_trouve"] > 0 else "N/A",
            f"{a['tjm_max_trouve']:,.0f} EUR" if a["tjm_max_trouve"] > 0 else "N/A",
            str(a["tjms_trouves"]),
            a["recommandation"],
        ]
        for i, v in enumerate(vals):
            cells[i].text = v
            for p in cells[i].paragraphs:
                for r in p.runs:
                    set_run_font(r, size=8.5, color=RED if a["alerte"] and i == 5 else BLACK)

    set_table_widths(table, [1.3, 0.9, 0.7, 0.7, 0.7, 2.1])
    apply_table_borders(table)

    add_separator(doc)
    accent_heading(doc, "3. Synthese et recommandation tarifaire")
    avg_market = [a["tjm_moyen"] for a in analyses if a["tjm_moyen"] > 0]
    if avg_market:
        market_avg = round(sum(avg_market) / len(avg_market), 0)
        bullet(doc, f"TJM moyen observe sur le marche (tous profils confondus) : {market_avg:,.0f} EUR")

        if tjm_cible > market_avg * 1.15:
            bullet(doc, random.choice([
                f"Votre TJM cible ({tjm_cible} EUR) est significativement au-dessus de la moyenne du marche ({market_avg:,.0f} EUR). Assurez-vous que votre positionnement le justifie.",
                f"A {tjm_cible} EUR, vous etes au-dessus de la moyenne de marche ({market_avg:,.0f} EUR). Verifions que le positionnement tient la route.",
                f"Ecart significatif : votre TJM ({tjm_cible} EUR) depasse la moyenne ({market_avg:,.0f} EUR). C'est tenable si le positionnement est clair.",
            ]))
        elif tjm_cible >= market_avg * 0.9:
            bullet(doc, random.choice([
                f"Votre TJM cible ({tjm_cible} EUR) est dans la fourchette haute du marche. Positionnement premium maintenable.",
                f"A {tjm_cible} EUR, vous etes bien positionne dans le haut du marche ({market_avg:,.0f} EUR de moyenne). Pas d'inquietude.",
                f"Votre TJM ({tjm_cible} EUR) est coherent avec un positionnement premium par rapport a la moyenne ({market_avg:,.0f} EUR).",
            ]))
        else:
            bullet(doc, random.choice([
                f"Votre TJM cible ({tjm_cible} EUR) est dans la moyenne du marche ({market_avg:,.0f} EUR). Une augmentation est envisageable.",
                f"A {tjm_cible} EUR, vous etes dans la moyenne ({market_avg:,.0f} EUR). On pourrait envisager une montee progressive.",
                f"TJM ({tjm_cible} EUR) aligne avec la moyenne de marche ({market_avg:,.0f} EUR). Il y a de la marge pour augmenter.",
            ]))
    else:
        bullet(doc, random.choice([
            "Pas assez de donnees pour calculer une moyenne de marche.",
            "Echantillon insuffisant pour etablir une moyenne fiable ce mois-ci.",
            "Trop peu de donnees collectees pour degager une tendance de marche.",
        ]), size=10)

    add_separator(doc)
    accent_heading(doc, "4. Actions recommandees")
    alerts = [a for a in analyses if a["alerte"]]
    if alerts:
        for a in alerts:
            bullet(doc, f"ALERTE : {a['profil']} - {a['recommandation']}", size=10)
    else:
        bullet(doc, random.choice([
            "Aucune alerte tarifaire. Votre positionnement est coherent avec le marche.",
            "Pas d'alerte cette semaine : le positionnement tarifaire est en phase avec le marche.",
            "Tout est coherent au niveau tarifaire. Rien a signaler pour cette semaine.",
        ]))

    doc.save(output_path)
    return output_path


def main() -> int:
    config = load_config()
    agent_cfg = config.get("veille_tarifaire", {})
    if not agent_cfg.get("enabled", True):
        print("Agent Veille tarifaire desactive.")
        return 0

    print(">>> Agent 7 : Veille concurrentielle tarifaire")
    profils_data = collecter_donnees(agent_cfg)
    analyses = analyser(profils_data, config)

    output_dir = ensure_output_dir(config)
    today = date.today().isoformat()
    output_path = output_dir / f"Veille_Tarifaire_{today}.docx"
    build_report(config, analyses, output_path)

    run_log = output_dir / f"Veille_Tarifaire_{today}_log.json"
    save_json_report(
        {
            "generated_at": datetime.now(timezone.utc).isoformat(),
            "analyses": analyses,
        },
        run_log,
    )

    print(f"  Rapport genere : {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

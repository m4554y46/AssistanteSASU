from __future__ import annotations

import json
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
from docx.shared import Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT

NAVY = RGBColor(11, 37, 69)
GRAY = RGBColor(86, 95, 108)
BLACK = RGBColor(0, 0, 0)
PALEBLUE = "E8EEF5"
LIGHT = "F2F4F7"

BRIEFS_DIR = Path(__file__).resolve().parent / "briefs"


def list_briefs() -> list[Path]:
    if not BRIEFS_DIR.exists():
        BRIEFS_DIR.mkdir(parents=True, exist_ok=True)
        return []
    files = sorted(BRIEFS_DIR.glob("*.txt")) + sorted(BRIEFS_DIR.glob("*.json"))
    return files


def parse_brief(path: Path) -> dict:
    if path.suffix == ".json":
        data = json.loads(path.read_text(encoding="utf-8"))
        return {
            "client": data.get("client", "Client"),
            "contexte": data.get("contexte", ""),
            "besoin": data.get("besoin", ""),
            "perimetre": data.get("perimetre", ""),
            "duree_souhaitee": data.get("duree", "3 mois"),
            "tjm_estime": data.get("tjm", 650),
            "objectif": data.get("objectif", ""),
        }

    raw = path.read_text(encoding="utf-8")
    lines = raw.strip().split("\n")
    meta = {
        "client": "Client",
        "contexte": "",
        "besoin": "",
        "perimetre": "",
        "duree_souhaitee": "3 mois",
        "tjm_estime": 650,
        "objectif": "",
    }

    current_key = "contexte"
    text_buffer = ""
    for line in lines:
        l = line.strip().lower()
        if l.startswith("client:"):
            meta["client"] = line.split(":", 1)[1].strip()
        elif l.startswith("contexte:"):
            if text_buffer:
                meta[current_key] = text_buffer.strip()
            current_key = "contexte"
            text_buffer = line.split(":", 1)[1].strip() if ":" in line else ""
        elif l.startswith("besoin:"):
            if text_buffer:
                meta[current_key] = text_buffer.strip()
            current_key = "besoin"
            text_buffer = line.split(":", 1)[1].strip() if ":" in line else ""
        elif l.startswith("perimetre:"):
            if text_buffer:
                meta[current_key] = text_buffer.strip()
            current_key = "perimetre"
            text_buffer = line.split(":", 1)[1].strip() if ":" in line else ""
        elif l.startswith("duree:"):
            meta["duree_souhaitee"] = line.split(":", 1)[1].strip()
        elif l.startswith("tjm:"):
            try:
                meta["tjm_estime"] = float(line.split(":", 1)[1].strip())
            except ValueError:
                pass
        elif l.startswith("objectif:"):
            if text_buffer:
                meta[current_key] = text_buffer.strip()
            current_key = "objectif"
            text_buffer = line.split(":", 1)[1].strip() if ":" in line else ""
        else:
            if text_buffer:
                text_buffer += " " + line
            else:
                text_buffer = line

    if text_buffer:
        meta[current_key] = text_buffer.strip()
    return meta


def generer_proposition(brief: dict, config: dict) -> dict:
    sasu = config.get("sasu", {})
    company = sasu.get("company", "ASSISTANT SASU")
    name = sasu.get("name", "")
    positioning = sasu.get("positioning", "")
    tjm_cible = int(sasu.get("tjm_cible", 650))

    tjm_propose = int(brief.get("tjm_estime", tjm_cible))
    client = brief.get("client", "Client")
    contexte = brief.get("contexte", "")
    besoin = brief.get("besoin", "")
    perimetre = brief.get("perimetre", "")
    duree = brief.get("duree_souhaitee", "3 mois")

    # Generer contenu de proposition
    approche = (
        f"Notre approche repose sur trois piliers :\n\n"
        f"1. Diagnostic et cadrage : comprendre le contexte, les parties prenantes, les contraintes et les objectifs mesurables.\n"
        f"2. Accompagnement operationnel : mise en place de rituels, outils et processus adaptes au contexte du client.\n"
        f"3. Transfert de competence : autonomisation des equipes internes pour assurer la perennite des pratiques mises en place."
    )

    livrables = [
        "Compte-rendu de chaque atelier / point d'avancement",
        "Tableau de bord de suivi des indicateurs cles",
        "Document de cadrage et feuille de route",
        "Synthese et recommandations finales",
    ]

    planning = (
        f"La mission se deroulerait sur {duree}, avec un rythme suggere de 2 a 3 jours par semaine. "
        f"Un point d'etape hebdomadaire sera organise avec le sponsor du projet."
    )

    return {
        "client": client,
        "date": date.today().isoformat(),
        "titre": f"Proposition d'accompagnement - {client}",
        "contexte": contexte,
        "besoin": besoin,
        "perimetre": perimetre,
        "approche": approche,
        "livrables": livrables,
        "planning": planning,
        "tjm": tjm_propose,
        "duree": duree,
        "statut": "Brouillon",
        "company": company,
        "name": name,
        "positioning": positioning,
    }


def build_report(config: dict, propositions: list[dict], output_path: Path):
    doc = setup_doc(
        "PROPOSITIONS COMMERCIALES",
        "Projets commerciaux generes depuis les briefs",
        config.get("author", "Assistant IA SASU"),
    )

    if not propositions:
        para(doc, "Aucun brief trouve.", bold=True, color=RGBColor(155, 28, 28))
        brief_path = str(BRIEFS_DIR)
        bullet(doc, f"Deposez un fichier .txt dans : {brief_path}")
        bullet(doc, "Format attendu :")
        for l in ["Client: Nom", "Contexte: ...", "Besoin: ...", "Perimetre: ...", "Duree: 3 mois", "TJM: 650", "Objectif: ..."]:
            bullet(doc, l)
        doc.save(output_path)
        return

    for idx, prop in enumerate(propositions, 1):
        para(doc, f"{idx}. {prop['titre']}", style="Heading 1")
        p = para(doc, after=2)
        set_run_font(p.add_run(f"Client : {prop['client']} | Statut : {prop['statut']} | Date : {prop['date']}"), size=9.5, color=GRAY)

        para(doc, "Contexte", style="Heading 2")
        para(doc, prop.get("contexte", "Non renseigne"))

        para(doc, "Besoin identifie", style="Heading 2")
        para(doc, prop.get("besoin", "Non renseigne"))

        para(doc, "Perimetre", style="Heading 2")
        para(doc, prop.get("perimetre", "A definir avec le client"))

        para(doc, "Approche methodologique", style="Heading 2")
        para(doc, prop.get("approche", ""))

        para(doc, "Livrables", style="Heading 2")
        for l in prop.get("livrables", []):
            bullet(doc, l)

        para(doc, "Planning indicatif", style="Heading 2")
        para(doc, prop.get("planning", ""))

        para(doc, "Conditions", style="Heading 2")
        bullet(doc, f"Duree : {prop.get('duree', 'A definie')}")
        bullet(doc, f"TJM propose : {prop.get('tjm', 650):,.0f} EUR")
        bullet(doc, "Modalites : Facturation mensuelle, TVA 20%, paiement a 30 jours")

        para(doc, "", after=6)
        doc.add_page_break()

    doc.save(output_path)
    return output_path


def main() -> int:
    config = load_config()
    agent_cfg = config.get("propositions", {})
    if not agent_cfg.get("enabled", True):
        print("Agent Propositions desactive.")
        return 0

    print(">>> Agent 6 : Redacteur de propositions commerciales")
    briefs = list_briefs()

    propositions = []
    for brief_path in briefs:
        brief = parse_brief(brief_path)
        prop = generer_proposition(brief, config)
        propositions.append(prop)
        print(f"  OK {brief_path.name} -> Proposition generee")

    output_dir = ensure_output_dir(config)
    today = date.today().isoformat()
    output_path = output_dir / f"Propositions_Commerciales_{today}.docx"
    build_report(config, propositions, output_path)

    run_log = output_dir / f"Propositions_{today}_log.json"
    save_json_report(
        {
            "generated_at": datetime.now(timezone.utc).isoformat(),
            "propositions_count": len(propositions),
        },
        run_log,
    )

    print(f"  Rapport genere : {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

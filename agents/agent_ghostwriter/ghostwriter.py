from __future__ import annotations

import random
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

from core.common import (
    bullet,
    ensure_output_dir,
    load_config,
    para,
    save_json_report,
    set_run_font,
    setup_doc,
)
from core.web import multi_search
from docx.shared import Pt, RGBColor

NAVY = RGBColor(11, 37, 69)
GRAY = RGBColor(86, 95, 108)
BLACK = RGBColor(0, 0, 0)

# Cas reels anonymises pour generer du contenu credible
MISSIONS_REELLES = [
    {
        "titre": "Deploiement mobile a l'echelle mondiale",
        "client": "groupe logistique international (90 000 collaborateurs, 30 pays)",
        "contexte": "Notre cabinet a ete mandate pour structurer et piloter le deploiement d'applications mobiles a destination de millions d'utilisateurs, dans 30 pays, pour le compte d'un leader mondial de la livraison de colis.",
        "mission": "Head of Mobile App Deployment",
        "actions": [
            "Gouvernance multi-pays et coordination des parties prenantes (BU, editors, prestataires)",
            "Harmonisation des roadmaps produit entre les differentes entites (DPD, Chronopost, BRT...)",
            "Mise en place de rituels agiles pour orchestrer les releases simultanees",
            "Priorisation des fonctionnalites selon l'impact business et les contraintes reglementaires locales",
        ],
        "resultats": "Applications utilisees par des millions d'utilisateurs quotidiens, deployment coordonne sur 30 pays, reduction des conflits de priorite inter-BU.",
        "secteur": "Logistique / Transport",
    },
    {
        "titre": "Refonte de l'experience client omnicanal dans le luxe",
        "client": "groupe de luxe international (Gucci, Saint Laurent, Balenciaga, Bottega Veneta...)",
        "contexte": "Notre cabinet a accompagne un leader mondial du luxe dans la transformation de l'experience client sur l'ensemble des canaux : en boutique, a distance et en ligne.",
        "mission": "Senior Product Manager / Customer Experience",
        "actions": [
            "Conception et deploiement de solutions omnicanales utilisees par des millions de clients et vendeurs",
            "Coordination des equipes produit (POs, designers) pour livrer des fonctionnalites a forte valeur ajoutee",
            "Integration de l'IA et des donnees comportementales pour personnaliser l'experience client",
            "Pilotage de la feuille de route produit en accord avec la strategie des Maisons",
        ],
        "resultats": "Solutions utilisees par des millions d'utilisateurs chaque annee, experience client unifiee sur l'ensemble des canaux, augmentation du taux de conversion omnicanal.",
        "secteur": "Luxe / Retail",
    },
    {
        "titre": "Innovation digitale et transformation des points de vente",
        "client": "leader mondial des materiaux de construction (Saint-Gobain)",
        "contexte": "Notre equipe a pilote l'innovation digitale pour un groupe industriel present dans 75 pays, en concevant des solutions disruptives pour les points de vente et les canaux en ligne.",
        "mission": "Digital Innovation Manager",
        "actions": [
            "Developpement d'applications de Realite Augmentee et de modelisation 3D pour les clients",
            "Conception d'outils web 2.0 de devis en ligne personnalise et en temps reel",
            "Management d'une equipe de 9 collaborateurs sur des projets digitaux innovants",
            "Implementation de solutions ERP pour integrer les nouveaux canaux digitaux",
        ],
        "resultats": "Solutions de Realite Augmentee deployees en magasin, outils de devis en ligne adoptes par des milliers de clients, equipe digitalisee et formee.",
        "secteur": "Industrie / Construction",
    },
    {
        "titre": "Accompagnement produit post-LBO",
        "client": "groupe hotelier en croissance (Goldman Sachs LBO)",
        "contexte": "Notre cabinet est intervenu en tant que Product Manager freelance pour accompagner un groupe hotelier en pleine restructuration post-LBO, avec des objectifs de croissance et de rentabilite.",
        "mission": "Product Manager (freelance)",
        "actions": [
            "Refonte du site web et de l'application mobile pour augmenter le taux de conversion",
            "Mise en place et integration d'un PIM (Product Information Management)",
            "Deploiement de solutions de reservation en ligne pour les clients directs",
            "Management d'une equipe de Product Owners et de designers",
        ],
        "resultats": "Augmentation du taux de reservation directe, reduction des couts d'acquisition, plateforme digitale unifiee.",
        "secteur": "Hotellerie / Services",
    },
]

CONTENUS_POSSIBLES = [
    {
        "type": "retour_experience",
        "structure": "experience",
        "description": "Retour d'experience anonymise d'une mission reelle",
    },
    {
        "type": "article_opinion",
        "structure": "opinion",
        "description": "Article d'opinion sur une tendance du marche",
    },
    {
        "type": "conseil_pratique",
        "structure": "conseil",
        "description": "Conseil operationnel pour dirigeant ou DSI",
    },
    {
        "type": "analyse_secteur",
        "structure": "analyse",
        "description": "Analyse d'un secteur (luxe, logistique, retail)",
    },
]


def generer_article_opinion(config: dict) -> dict:
    company = config.get("profile", {}).get("company", "ASSISTANT SASU")
    sujet = random.choice([
        "Pourquoi le Product Operating Model est la cle d'une transformation IA reussie",
        "Ce que le deploiement mobile dans 30 pays nous a appris sur la gouvernance produit",
        "L'omnicanal dans le luxe : retour sur 7 ans de transformation chez un leader mondial",
        "Realite Augmentee en retail : etait-ce un gadget ou un investissement strategique ?",
    ])
    body = (
        f"Chez {company}, nous accompagnons les directions generales et les DSI "
        f"dans leurs transformations produit et digitales.\n\n"
        f"Notre constat est le meme d'un secteur a l'autre : la technologie n'est jamais "
        f"le verrou. Ce qui fait la difference, c'est la capacite a organiser la delivery, "
        f"a prioriser les investissements et a creer les conditions de l'adoption.\n\n"
        f"Apres avoir pilote des transformations a grande echelle dans le luxe (Groupe Kering), "
        f"l'industrie (Saint-Gobain), la logistique (groupe GeoPost) et les services (LBO hotelier), "
        f"nous avons identifie 3 piliers communs aux reussites produit :\n\n"
        f"1. Un Product Operating Model clair : qui decide quoi, sur quel horizon, avec quelle mesure d'impact.\n"
        f"2. Une discovery continue : les decisions produit ne se prennent pas au feeling mais a partir "
        f"d'interviews clients, de donnees d'usage et de tests structures.\n"
        f"3. Une gouvernance de la delivery : les equipes savent ce qu'elles doivent livrer, "
        f"pourquoi, et dans quel cadre.\n\n"
        f"Nous intervenons pour cadrer, auditer ou accompagner la mise en place de ces piliers. "
        f"En 5 jours ou en mission longue."
    )
    return {
        "titre": sujet,
        "corps": body,
        "accroche": sujet,
        "type": "Article d'opinion",
        "signature": f"\n\n---\n{company}\nAccompagnement produit, digital et IA",
    }


def generer_retour_experience(mission: dict, config: dict) -> dict:
    company = config.get("profile", {}).get("company", "ASSISTANT SASU")
    titre = f"Cas client : {mission['titre']}"
    body = (
        f"Contexte\n"
        f"{mission['contexte']}\n\n"
        f"Notre intervention\n"
        f"Mission confiee : {mission['mission']}\n"
        f"Secteur : {mission['secteur']}\n\n"
        f"Ce que nous avons fait :\n"
    )
    for a in mission["actions"]:
        body += f"- {a}\n"
    body += (
        f"\nResultats\n"
        f"{mission['resultats']}\n\n"
        f"Ce type de mission illustre notre approche : un cabinet de conseil "
        f"operationnel, capable d'intervenir a la fois sur la strategie et sur la delivery. "
        f"Nous travaillons avec des groupes internationaux comme avec des PME en forte croissance."
    )
    return {
        "titre": titre,
        "corps": body,
        "accroche": f"{mission['titre']} - {mission['secteur']}",
        "type": "Cas client",
        "signature": f"\n\n---\n{company}",
    }


def generer_conseil_dirigeant(config: dict) -> dict:
    company = config.get("profile", {}).get("company", "ASSISTANT SASU")
    titre = random.choice([
        "Comment evaluer la maturite produit de votre entreprise en 2 heures",
        "Les 5 signes que votre organisation produit a besoin d'un audit",
        "Audit IA : 3 questions a poser avant d'investir le premier euro",
        "La feuille de route produit en environnement complexe : notre methode",
    ])
    body = (
        f"Nous rencontrons regulierement des dirigeants de PME et d'ETI qui nous disent : "
        f"« Nous savons qu'il faut faire quelque chose, mais nous ne savons pas par ou commencer. »\n\n"
        f"Voici le cadre que nous utilisons pour un diagnostic rapide :\n\n"
        f"1. Cartographie des flux de valeur : quels sont les processus metiers critiques, "
        f"ou se situent les goulots d'etranglement, quel est le cout de la non-qualite digitale ?\n\n"
        f"2. Audit des competences produit : qui prend les decisions produit ? Sur quels criteres ? "
        f"Avec quelle mesure d'impact ?\n\n"
        f"3. Evaluation de la maturite IA : quels processus peuvent etre augmentes par l'IA, "
        f"quelles donnees sont disponibles, quel est le niveau de maturite de l'organisation ?\n\n"
        f"Notre cabinet realise ce diagnostic en 5 jours. Vous repartez avec un plan d'action "
        f"priorise et une estimation budgetaire. Sans engagement."
    )
    return {
        "titre": titre,
        "corps": body,
        "accroche": titre,
        "type": "Conseil pratique",
        "signature": f"\n\n---\n{company}\nAudit, conseil et accompagnement produit",
    }


def generer_analyse_tendance(config: dict) -> dict:
    company = config.get("profile", {}).get("company", "ASSISTANT SASU")
    titre = random.choice([
        "Transformation digitale dans le retail : les lecons du luxe applicables aux PME",
        "Mobile first en B2B : ce que la logistique nous apprend sur l'adoption",
        "IA generatives dans la boutique : entre hype et valeur reelle",
        "Product Management en contexte LBO : comment concilier croissance et rentabilite",
    ])
    body = (
        f"Chez {company}, nous intervenons dans des contextes varies : du groupe du CAC40 "
        f"a la PME en pleine scale-up. Cette diversite nous donne une vision unique des tendances "
        f"qui marchent vraiment.\n\n"
        f"Notre analyse du moment :\n\n"
        f"Le mobile n'est plus un canal, c'est le point d'entree principal de l'experience client. "
        f"Dans la logistique, le retail et les services, les utilisateurs exigent une experience "
        f"fluide et coherente, qu'ils soient en boutique, sur le web ou sur une application. "
        f"Les entreprises qui reussissent leur transformation mobile sont celles qui ont traite "
        f"le sujet comme un programme produit, pas comme un projet technique.\n\n"
        f"Fort de notre experience dans le deploiement mobile a l'echelle mondiale (30 pays), "
        f"nous aidons les directions a structurer leur strategie mobile et omnicanale, "
        f"a prioriser les investissements et a organiser la delivery."
    )
    return {
        "titre": titre,
        "corps": body,
        "accroche": titre,
        "type": "Analyse de tendance",
        "signature": f"\n\n---\n{company}\nStrategie et delivery digitale",
    }


def build_report(config: dict, contenus: list[dict], output_path: Path):
    company = config.get("profile", {}).get("company", "ASSISTANT SASU")
    doc = setup_doc(
        "NEWSLETTER - CONTENUS EDITORIAUX",
        f"Contenus pour le positionnement de {company}",
        "Assistant IA SASU",
    )

    para(doc, "1. Strategie de contenu", style="Heading 1")
    bullet(doc, f"Positionnement : {company} - cabinet de conseil en produit et transformation digitale")
    bullet(doc, f"Ton : expert, direct, sans bullshit - on parle de ce qu'on a vraiment fait")
    bullet(doc, f"Canal suggere : Newsletter Substack + Page LinkedIn entreprise")
    bullet(doc, f"Rythme : 1 publication / semaine")
    bullet(doc, f"Cible : Dirigeants de PME/ETI, DSI, CDO, directeurs produits")
    para(doc, "", after=8)

    para(doc, "2. Editoriaux de la semaine", style="Heading 1")
    for idx, c in enumerate(contenus, 1):
        para(doc, f"{idx}. {c['titre']}", style="Heading 3")
        p = para(doc, after=2)
        set_run_font(p.add_run(f"Type : {c['type']}"), size=9.5, color=GRAY)
        para(doc, c["corps"], before=4, after=6)
        if c.get("signature"):
            set_run_font(doc.add_paragraph().add_run(c["signature"]), size=9, color=GRAY, italic=True)
        doc.add_page_break()

    doc.save(output_path)
    return output_path


def main() -> int:
    config = load_config()
    agent_cfg = config.get("personal_branding", {})
    if not agent_cfg.get("enabled", True):
        print("Agent Personal Branding desactive.")
        return 0

    print(">>> Personal Branding : contenus newsletter")

    contenus = []

    contenus.append(generer_article_opinion(config))
    contenus.append(generer_retour_experience(random.choice(MISSIONS_REELLES), config))
    contenus.append(generer_conseil_dirigeant(config))
    contenus.append(generer_analyse_tendance(config))

    print(f"  {len(contenus)} contenus generes")
    for c in contenus:
        print(f"    - {c['type']}: {c['titre'][:80]}")

    output_dir = ensure_output_dir(config)
    today = date.today().isoformat()
    output_path = output_dir / f"Newsletter_Branding_{today}.docx"
    build_report(config, contenus, output_path)

    run_log = output_dir / f"Newsletter_Branding_{today}_log.json"
    save_json_report(
        {
            "generated_at": datetime.now(timezone.utc).isoformat(),
            "positionnement": f"Cabinet conseil produit & digital",
            "contenus": contenus,
        },
        run_log,
    )
    print(f"  Rapport : {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

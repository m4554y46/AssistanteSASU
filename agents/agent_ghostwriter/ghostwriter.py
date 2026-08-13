from __future__ import annotations

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
    para,
    save_json_report,
    set_run_font,
    setup_doc,
)
from core.web import SearchResult, multi_search
from docx.shared import Pt, RGBColor

NAVY = RGBColor(11, 37, 69)
GRAY = RGBColor(86, 95, 108)
BLACK = RGBColor(0, 0, 0)


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


MISSIONS_REELLES = [
    {
        "titre": "Deploiement mobile a l'echelle mondiale",
        "client": "groupe logistique international (90 000 collaborateurs, 30 pays)",
        "contexte": "Michael a ete mandate pour structurer et piloter le deploiement d'applications mobiles a destination de millions d'utilisateurs, dans 30 pays, pour le compte d'un leader mondial de la livraison de colis.",
        "mission": "Head of Mobile App Deployment",
        "actions": [
            "Coordination multi-pays entre les differentes entites (DPD, Chronopost, BRT...)",
            "Mise en place de rituels pour orchestrer les releases simultanees",
            "Priorisation des fonctionnalites selon l'impact business et les contraintes locales",
        ],
        "resultats": "Applications utilisees par des millions d'utilisateurs quotidiens, deployment coordonne sur 30 pays.",
        "secteur": "Logistique / Transport",
    },
    {
        "titre": "Refonte de l'experience client omnicanal dans le luxe",
        "client": "groupe de luxe international (Gucci, Saint Laurent, Balenciaga, Bottega Veneta...)",
        "contexte": "Michael a accompagne un leader mondial du luxe dans la transformation de l'experience client sur l'ensemble des canaux : en boutique, a distance et en ligne.",
        "mission": "Senior Product Manager / Customer Experience",
        "actions": [
            "Conception et deploiement de solutions omnicanales utilisees par des millions de clients et vendeurs",
            "Coordination des equipes produit (POs, designers) pour livrer des fonctionnalites a forte valeur ajoutee",
            "Integration de l'IA et des donnees comportementales pour personnaliser l'experience client",
            "Pilotage du planning en accord avec la strategie des Maisons",
        ],
        "resultats": "Solutions utilisees par des millions d'utilisateurs chaque annee, experience client unifiee sur tous les canaux.",
        "secteur": "Luxe / Retail",
    },
    {
        "titre": "Innovation digitale et transformation des points de vente",
        "client": "leader mondial des materiaux de construction (Saint-Gobain)",
        "contexte": "Michael a pilote l'innovation digitale pour un groupe industriel present dans 75 pays, en concevant des solutions pour les points de vente et les canaux en ligne.",
        "mission": "Digital Innovation Manager",
        "actions": [
            "Developpement d'applications de Realite Augmentee et de modelisation 3D pour les clients",
            "Conception d'outils web de devis en ligne personnalise et en temps reel",
            "Management d'une equipe de 9 personnes sur des projets digitaux",
            "Implementation de solutions ERP pour integrer les nouveaux canaux digitaux",
        ],
        "resultats": "Solutions de Realite Augmentee deployees en magasin, outils de devis en ligne adoptes par des milliers de clients.",
        "secteur": "Industrie / Construction",
    },
    {
        "titre": "Accompagnement produit post-LBO",
        "client": "groupe hotelier en croissance (Goldman Sachs LBO)",
        "contexte": "Michael est intervenu en tant que Product Manager freelance pour accompagner un groupe hotelier en pleine restructuration apres un LBO, avec des objectifs de croissance et de rentabilite.",
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

BLOCKED_DOMAINS = [
    "pinterest", "amazon", "ebay", "etsy", "aliexpress", "walmart",
    "shopify", "boulanger", "fnac", "cdiscount", "ikea", "leroymerlin",
    "decathlon", "booking", "tripadvisor",
    "larousse", "lerobert", "dictionnaire", "wiktionary", "cnrtl",
    "linternaute", "wikihow", "wikipedia",
    "lalanguefrancaise", "compagnie-fiduciaire", "agentprovocateur",
    "fiverr", "freelance.com", "upwork", "peopleperhour",
    "facebook", "instagram", "meta.com", "play.google.com",
    "twitter", "x.com", "tiktok",
    "youtube", "leboncoin",
    "wordreference", "linguee", "reverso",
    "cambridge", "merriam", "oxford", "collins",
    "allocine", "mozzartbet", "bet365", "parionssport", "poker",
    "synonymo",
    "aujourdhui", "transformation.co.uk",
    "transformation.gouv.fr", "blog-expertise.fr", "ideascale.com",
]

BLOCKED_WORDS = [
    "login", "sign in", "se connecter", "s identifier", "password",
    "connexion", "subscribe", "inscription", "create account",
    "earplugs", "bouchons", "oreille", "shoes",
    "recrutement", "nous recrutons", "annonce recrute",
    "definition", "definition",
    "trouvez les meilleurs", "freelance services marketplace",
    "mission locale",
    "dentist", "dentiste", "dental", "tooth", "plumber",
    "avocat", "lawyer", "notaire", "medecin", "doctor",
    "restaurant", "immobilier", "real estate",
    "cross dressing", "transgender", "lingerie", "swimwear", "hosiery",
    "university", "universite", "université", "harvard", "primerbank",
    "principes", "types", "types d'audit", "quest ce que", "qu'est-ce que",
    "ecole", "school", "faculte", "faculté",
]

# Liste blanche : seuls les articles de ces sources sont retenus dans la
# newsletter. Les resultats Bing hors de ces domaines sont presque toujours
# hors sujet (agregateurs, definitions, sites commerciaux).
TRUSTED_DOMAINS = [
    "svpg.com", "producttalk.org", "martinfowler.com", "agilealliance.org",
    "medium.com", "productcoalition", "mindtheproduct",
    "openai.com", "anthropic.com", "ai.googleblog.com", "blog.google",
    "deepmind.google", "mistral.ai", "huggingface.co", "hbr.org",
    "mckinsey.com", "bcg.com", "theverge.com", "techcrunch.com",
    "venturebeat.com", "stripe.com", "basecamp.com", "teamtopologies.com",
    "github.blog", "github.com", "docs.github.com", "towardsdatascience.com",
    "atlassian.com", "scaledagileframework.com", "productplan.com",
    "intercom.com", "productcoaching.com", "lennyrachitsky.com",
    "firstround.com", "a16z.com", "sequoiacap.com",
]

THEME_QUERIES = {
    "ia_agentic": [
        "agentic IA entreprise 2026",
        "agent IA autonome product management",
        "intelligence artificielle agentique entreprise",
    ],
    "product": [
        "product management methode innovation 2026",
        "continuous discovery produit 2026",
        "product operating model transformation",
    ],
    "transformation": [
        "transformation numerique entreprise 2026",
        "innovation digitale retail 2026",
        "mobile first strategie entreprise",
    ],
    "conseil": [
        "audit organisation produit methode",
        "gouvernance produit equipe 2026",
        "conseil direction transformation digitale",
        "formation certifiante product manager chef de projet 2026",
    ],
    "freelance": [
        "SASU freelance fiscalite independant 2026",
        "meilleur compte pro banque freelance independant 2026",
        "mutuelle prevoyance assurance freelance 2026",
    ],
}

MOIS_FR = ["", "janvier", "fevrier", "mars", "avril", "mai", "juin",
           "juillet", "aout", "septembre", "octobre", "novembre", "decembre"]


def _semaine() -> str:
    t = date.today()
    return f"{t.day} {MOIS_FR[t.month]} {t.year}"


def _filtrer(items: list[SearchResult]) -> list[SearchResult]:
    out = []
    for item in items:
        if not item.title or len(item.title) < 15:
            continue
        source = item.source.lower()
        title = item.title.lower()
        snippet = (item.snippet or "").lower()
        text = f"{title} {snippet}"
        if any(d in source for d in BLOCKED_DOMAINS):
            continue
        if not any(t in source for t in TRUSTED_DOMAINS):
            continue
        if any(w in text for w in BLOCKED_WORDS):
            continue
        if not _is_latin(item.title) or (item.snippet and not _is_latin(item.snippet)):
            continue
        out.append(item)
    seen = set()
    unique = []
    for it in out:
        k = it.title.lower()[:60]
        if k not in seen:
            seen.add(k)
            unique.append(it)
    return unique


def _rechercher(queries: list[str], timeout: int = 10, count: int = 5) -> list[SearchResult]:
    raw = multi_search(queries, timeout, count)
    return _filtrer(raw)[:4]


def generer_veille_semaine(company: str) -> dict:
    articles = _rechercher(THEME_QUERIES["ia_agentic"] + THEME_QUERIES["product"])
    titre = f"Veille de la semaine du {_semaine()}"

    if not articles:
        corps = "Pas d'article suffisamment pertinent cette semaine dans ma veille IA/Product."
        items = []
    else:
        items = [{"title": a.title, "url": a.url, "source": a.source} for a in articles]
        corps = "Voici les articles que j'ai releves cette semaine :"

    return {
        "titre": titre,
        "corps": corps,
        "accroche": titre,
        "type": "Veille hebdomadaire",
        "articles": items,
    }


def generer_retour_experience(mission: dict, company: str) -> dict:
    titre = f"Cas client : {mission['titre']}"
    actions = "\n".join(f"- {a}" for a in mission["actions"])
    body = (
        f"Contexte\n{mission['contexte']}\n\n"
        f"Notre intervention\n"
        f"Mission confiee : {mission['mission']}\n"
        f"Secteur : {mission['secteur']}\n\n"
        f"Ce que nous avons fait :\n{actions}\n\n"
        f"Resultats\n{mission['resultats']}"
    )
    return {
        "titre": titre,
        "corps": body,
        "accroche": f"{mission['titre']} - {mission['secteur']}",
        "type": "Cas client",
    }


def generer_astuce_pratique(company: str) -> dict:
    articles = _rechercher(THEME_QUERIES["conseil"])
    titre = "Conseil pratique de la semaine"

    if not articles:
        corps = "Un point souvent neglige mais qui fait la difference : la clarte du cadrage en amont. Avant de lancer un projet, posez-vous 3 questions : quel est le probleme, pour qui, et comment saurons-nous que c'est resolu ?"
    else:
        ref = articles[0]
        corps = f"Un point que je vois regulierement sur le terrain, et cet article le resume bien : {ref.title} ({ref.source})."

    return {
        "titre": titre,
        "corps": corps,
        "accroche": titre,
        "type": "Conseil pratique",
    }

    return {
        "titre": titre,
        "corps": corps,
        "accroche": titre,
        "type": "Conseil pratique",
    }


def generer_analyse_tendance(company: str) -> dict:
    queries = THEME_QUERIES["transformation"]
    articles = _rechercher(queries)
    titre = "Tendance de la semaine"

    if not articles:
        corps = "Pas de tendance marquante cette semaine dans le secteur. Je continue la veille."
        items = []
    else:
        items = [{"title": a.title, "url": a.url, "source": a.source} for a in articles[:2]]
        corps = "Signaux retenus cette semaine :"

    return {
        "titre": titre,
        "corps": corps,
        "accroche": titre,
        "type": "Analyse de tendance",
        "articles": items,
    }


def build_report(config: dict, contenus: list[dict], output_path: Path):
    company = config.get("profile", {}).get("company", "ASTRA MOMENTUM")
    doc = setup_doc(
        "NEWSLETTER - CONTENUS EDITORIAUX",
        f"Contenus pour le positionnement de {company}",
        config.get("author", "Virginie Benayoun"),
    )

    for idx, c in enumerate(contenus, 1):
        para(doc, f"{idx}. {c['titre']}", style="Heading 3")
        para(doc, c["corps"], before=4, after=6)
        for art in c.get("articles", []):
            p = para(doc, after=2)
            set_run_font(p.add_run(f"- {art['title']} "), size=10.5)
            add_hyperlink(p, art["url"], art["url"])
            p2 = para(doc, after=6)
            set_run_font(p2.add_run(f"   Source : {art['source']}"), size=9.5, color=GRAY)
        doc.add_page_break()

    doc.save(output_path)
    return output_path


def main() -> int:
    config = load_config()
    agent_cfg = config.get("personal_branding", {})
    if not agent_cfg.get("enabled", True):
        print("Agent Personal Branding desactive.")
        return 0

    company = config.get("profile", {}).get("company", "ASTRA MOMENTUM")
    print(">>> Personal Branding : contenus newsletter")

    contenus = []

    print("  1/4 Veille de la semaine...")
    contenus.append(generer_veille_semaine(company))

    print("  2/4 Cas client...")
    contenus.append(generer_retour_experience(MISSIONS_REELLES[-1], company))

    print("  3/4 Astuce pratique...")
    contenus.append(generer_astuce_pratique(company))

    print("  4/4 Tendance...")
    contenus.append(generer_analyse_tendance(company))

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
            "contenus": contenus,
        },
        run_log,
    )
    print(f"  Rapport : {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

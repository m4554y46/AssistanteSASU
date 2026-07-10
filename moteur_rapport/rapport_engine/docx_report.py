from __future__ import annotations

from datetime import date
from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


NAVY = RGBColor(11, 37, 69)
BLUE = RGBColor(46, 116, 181)
DARKBLUE = RGBColor(31, 77, 120)
GRAY = RGBColor(86, 95, 108)
BLACK = RGBColor(0, 0, 0)
RED = RGBColor(155, 28, 28)
PALEBLUE = "E8EEF5"
LIGHT = "F2F4F7"


def set_run_font(run, name="Calibri", size=None, color=None, bold=None, italic=None):
    run.font.name = name
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), name)
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), name)
    if size is not None:
        run.font.size = Pt(size)
    if color is not None:
        run.font.color.rgb = color
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def add_hyperlink(paragraph, text, url):
    part = paragraph.part
    rid = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rid)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    rpr.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    rpr.append(underline)
    run.append(rpr)
    text_node = OxmlElement("w:t")
    text_node.text = text
    run.append(text_node)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def shade_cell(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=80, start=120, bottom=80, end=120):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for key, value in {"top": top, "start": start, "bottom": bottom, "end": end}.items():
        node = tc_mar.find(qn(f"w:{key}"))
        if node is None:
            node = OxmlElement(f"w:{key}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_table_widths(table, widths):
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for row in table.rows:
        for idx, width in enumerate(widths):
            if idx >= len(row.cells):
                continue
            row.cells[idx].width = Inches(width)
            row.cells[idx].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(row.cells[idx])


def para(doc, text="", style=None, before=None, after=None, color=None, bold=False, italic=False, size=None, align=None):
    p = doc.add_paragraph(style=style)
    if align is not None:
        p.alignment = align
    if before is not None:
        p.paragraph_format.space_before = Pt(before)
    if after is not None:
        p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.10
    if text:
        r = p.add_run(text)
        set_run_font(r, size=size, color=color, bold=bold, italic=italic)
    return p


def bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    set_run_font(r, size=10.5)
    return p


def first_sentence(text: str, fallback: str) -> str:
    clean = " ".join((text or "").split())
    if not clean:
        return fallback
    parts = clean.split(". ")
    return (parts[0] + ".")[:420]


def setup_doc() -> Document:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)
    styles = doc.styles
    styles["Normal"].font.name = "Calibri"
    styles["Normal"]._element.rPr.rFonts.set(qn("w:ascii"), "Calibri")
    styles["Normal"]._element.rPr.rFonts.set(qn("w:hAnsi"), "Calibri")
    styles["Normal"].font.size = Pt(11)
    styles["Normal"].paragraph_format.space_after = Pt(6)
    for name, size, color, before, after in [
        ("Heading 1", 16, BLUE, 16, 8),
        ("Heading 2", 13, BLUE, 12, 6),
        ("Heading 3", 12, DARKBLUE, 8, 4),
    ]:
        st = styles[name]
        st.font.name = "Calibri"
        st._element.rPr.rFonts.set(qn("w:ascii"), "Calibri")
        st._element.rPr.rFonts.set(qn("w:hAnsi"), "Calibri")
        st.font.size = Pt(size)
        st.font.color.rgb = color
        st.font.bold = True
        st.paragraph_format.space_before = Pt(before)
        st.paragraph_format.space_after = Pt(after)
        st.paragraph_format.keep_with_next = True
    header = section.header.paragraphs[0]
    header.text = "Rapport Hebdo ASTRA MOMENTUM SASU | Veille, prospection et culture produit"
    set_run_font(header.runs[0], size=9, color=GRAY)
    footer = section.footer.paragraphs[0]
    footer.text = "Rapport hebdomadaire - confidentiel"
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(footer.runs[0], size=9, color=GRAY)
    return doc


def add_dashboard(doc, articles, opportunities, meta):
    para(doc, "1. Methode de selection et tableau de bord", style="Heading 1")
    para(doc, "Methode de travail appliquee", style="Heading 2")
    for text in [
        f"Recherche large : {meta['article_queries']} requetes de veille et {meta['opportunity_queries']} requetes de prospection executees.",
        f"Filtrage qualitatif : {meta['raw_articles']} resultats de veille et {meta['raw_opportunities']} resultats d'opportunites consolides avant arbitrage.",
        "Scoring : pertinence business, credibilite de la source, fraicheur, verification d'URL, potentiel commercial et adequation au positionnement de Michael ASSAYAG.",
        "Arbitrage final : seuls les signaux transformables en action commerciale, contenu LinkedIn, conversation client ou enrichissement methodologique sont retenus.",
    ]:
        bullet(doc, text)

    table = doc.add_table(rows=1, cols=4)
    headers = ["Signal", "Observation", "Action conseillee", "Priorite"]
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        shade_cell(cell, PALEBLUE)
        cell.text = h
        for p in cell.paragraphs:
            for r in p.runs:
                set_run_font(r, bold=True, color=NAVY, size=9.5)
    rows = [
        ["Veille IA/Product", f"{len(articles)} contenus retenus", "Transformer 1 signal en post LinkedIn", "Haute"],
        ["Prospection", f"{len(opportunities)} opportunites retenues", "Qualifier les 2 meilleurs fits", "Haute"],
        ["Methodologies", "Concepts selectionnes selon l'actualite et l'utilite commerciale", "Nourrir les rendez-vous clients", "Moyenne"],
    ]
    for row in rows:
        cells = table.add_row().cells
        for i, value in enumerate(row):
            cells[i].text = value
            for p in cells[i].paragraphs:
                for r in p.runs:
                    set_run_font(r, size=9.2)
    set_table_widths(table, [1.4, 2.2, 2.1, 0.8])


def add_article(doc, idx, item):
    para(doc, f"{idx}. {item['title']}", style="Heading 3")
    p = para(doc, after=2)
    set_run_font(p.add_run(f"Source : {item['source']} | Date : {item.get('published') or 'non precisee'} | Score : {item['score']}/10 | "), size=9.5, color=GRAY, bold=True)
    add_hyperlink(p, "URL", item["url"])
    bullet(doc, "Resume analytique : " + first_sentence(item.get("snippet", ""), "Contenu retenu pour sa pertinence thematique et commerciale."))
    bullet(doc, "Pourquoi cela doit interesser au developpement d'ASTRA MOMENTUM : signal exploitable pour renforcer son positionnement Product/IA et ouvrir une conversation client.")
    bullet(doc, "Commentaire de selection : " + item["selection_reason"])
    bullet(doc, "Recommandation : " + item["recommendation"])


def add_opportunity(doc, idx, item):
    para(doc, f"{idx}. {item['title']}", style="Heading 3")
    p = para(doc, after=2)
    set_run_font(p.add_run(f"Plateforme/source : {item['source']} | Date : {item.get('published') or 'non precisee'} | Score : {item['score']}/10 | "), size=9.5, color=GRAY, bold=True)
    add_hyperlink(p, "URL", item["url"])
    bullet(doc, "Resume mission : " + first_sentence(item.get("snippet", ""), "Opportunite detectee via recherche publique. Details a confirmer sur la plateforme."))
    bullet(doc, "Fit competences : proximite avec product management, gestion de projet digital, IA, delivery ou transformation.")
    bullet(doc, "Risques / points a verifier : accessibilite de l'annonce, TJM, statut freelance, modalites remote et niveau d'autonomie attendu.")
    bullet(doc, "Angle de reponse : valoriser cadrage, priorisation valeur/complexite, coordination parties prenantes, adoption et mesure d'impact.")
    bullet(doc, "Decision : " + item["recommendation"])


def add_methods(doc, methods):
    para(doc, "4. Enrichissement culturel et methodologique", style="Heading 1")
    for method in methods:
        para(doc, method["name"], style="Heading 3")
        bullet(doc, "Definition : " + method["definition"])
        bullet(doc, "Interet business : " + method["business_value"])
        bullet(doc, "Cas d'utilisation client : " + method["client_use"])
        bullet(doc, "Usage commercial pour la société : " + method["commercial_use"])
        p = para(doc, after=4)
        set_run_font(p.add_run("Source : "), size=9.5, color=GRAY, bold=True)
        add_hyperlink(p, method["url"], method["url"])


def add_actions(doc, articles, opportunities):
    para(doc, "5. Recommandations operationnelles", style="Heading 1")
    para(doc, "Priorites pour la semaine suivante", style="Heading 2")
    top_opp = opportunities[0]["title"] if opportunities else "la meilleure opportunite identifiee"
    top_article = articles[0]["title"] if articles else "le signal IA/Product le plus fort"
    for text in [
        f"Qualifier en priorite : {top_opp}.",
        f"Transformer le signal '{top_article}' en prise de parole LinkedIn.",
        "Mettre a jour l'offre 'Audit IA et Product Discovery' avec une promesse courte, un livrable et un format 5 jours.",
    ]:
        bullet(doc, text)
    para(doc, "Brouillon court de post LinkedIn", style="Heading 2")
    para(
        doc,
        "Les agents IA ne creent pas de valeur par magie. Ils deviennent utiles quand ils sont relies a un outcome clair, a une gouvernance simple et a un processus metier que l'on comprend vraiment. En 2026, le sujet n'est plus seulement de tester des outils IA : c'est de choisir les bons cas d'usage, de mesurer l'impact et d'organiser l'adoption.",
        italic=True,
        color=GRAY,
    )


def add_sources(doc, articles, opportunities, methods, meta):
    para(doc, "6. Sources et URLs verifiables", style="Heading 1")
    if meta.get("errors"):
        para(doc, "Limites rencontrees", style="Heading 2")
        for err in meta["errors"][:8]:
            bullet(doc, err)
    for item in articles + opportunities:
        p = para(doc, after=2)
        set_run_font(p.add_run(f"{item['title']} - "), size=9.3, color=GRAY, bold=True)
        add_hyperlink(p, item["url"], item["url"])
    for method in methods:
        p = para(doc, after=2)
        set_run_font(p.add_run(f"{method['name']} - "), size=9.3, color=GRAY, bold=True)
        add_hyperlink(p, method["url"], method["url"])


def build_report(config: dict, articles: list[dict], opportunities: list[dict], methods: list[dict], meta: dict, output_path: Path) -> Path:
    doc = setup_doc()
    today = date.today().isoformat()
    para(doc, "RAPPORT HEBDOMADAIRE ", size=23, color=NAVY, bold=True, after=4)
    para(doc, "Veille approfondie, prospection qualifiee et recommandations Product/IA", size=13.5, color=GRAY, after=14)
    for label, value in [
        ("Destinataire", f"{config['recipient']} - President SASU"),
        ("Autrice", config["author"]),
        ("Positionnement", config["positioning"]),
        ("Date du rapport", today),
    ]:
        p = para(doc, after=2)
        set_run_font(p.add_run(label + " : "), bold=True, color=BLACK)
        set_run_font(p.add_run(value), color=BLACK)

    call = doc.add_table(rows=1, cols=1)
    set_table_widths(call, [6.5])
    shade_cell(call.cell(0, 0), LIGHT)
    p = call.cell(0, 0).paragraphs[0]
    set_run_font(p.add_run("Lecture executive\n"), bold=True, color=NAVY, size=11.5)
    set_run_font(
        p.add_run(
            "Ce rapport est le resultat d'une recherche hebdomadaire dynamique : collecte de sources, verification des URLs, filtrage, scoring et arbitrage final. La selection vise les contenus et opportunites directement utiles pour le positionnement de Michael ASSAYAG en product management digital, IA et transformation."
        ),
        size=10.5,
        color=BLACK,
    )

    add_dashboard(doc, articles, opportunities, meta)
    para(doc, "2. Veille technologique, product, agile et IA", style="Heading 1")
  
# --- AJOUT : texte explicatif du score ---
para(doc, 
     "Ce score est une évaluation en s'appuyant sur l'expertise du marché freelance Product/IA. Il repose sur une grille de lecture professionnelle incluant la pertinence thématique, la fraîcheur de la source (<1 ou 2 mois) , sa crédibilité et le potentiel commercial concret. ",
     before=4, after=8, color=GRAY, size=9.5, italic=True)





  
    for idx, item in enumerate(articles, 1):
        add_article(doc, idx, item)
    para(doc, "3. Prospection commerciale freelance", style="Heading 1")
    for idx, item in enumerate(opportunities, 1):
        add_opportunity(doc, idx, item)
    add_methods(doc, methods)
    add_actions(doc, articles, opportunities)
    add_sources(doc, articles, opportunities, methods, meta)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    return output_path

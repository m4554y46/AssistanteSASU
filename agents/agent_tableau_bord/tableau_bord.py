from __future__ import annotations

import csv
import json
import random
from collections import defaultdict
from datetime import date, datetime, timedelta, timezone
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
    para,
    save_json_report,
    set_run_font,
    set_table_widths,
    shade_cell,
    setup_doc,
)
from docx.shared import Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT

NAVY = RGBColor(11, 37, 69)
GRAY = RGBColor(86, 95, 108)
BLACK = RGBColor(0, 0, 0)
RED = RGBColor(155, 28, 28)
GREEN = RGBColor(0, 128, 0)
ORANGE = RGBColor(200, 120, 0)
PALEBLUE = "E8EEF5"

MONTHS_FR = [
    "", "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
    "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
]


def load_invoices(factures_dir: str) -> list[dict]:
    path = Path(factures_dir)
    invoices = []

    csv_files = list(path.glob("*.csv")) + list(path.glob("*.CSV"))
    for csv_file in csv_files:
        with csv_file.open(newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f, delimiter=";")
            for row in reader:
                invoices.append(row)

    json_files = list(path.glob("*.json")) + list(path.glob("*.JSON"))
    for jf in json_files:
        data = json.loads(jf.read_text(encoding="utf-8"))
        if isinstance(data, list):
            invoices.extend(data)
        elif isinstance(data, dict):
            invoices.append(data)

    return invoices


def parse_date(d: str | None) -> date | None:
    if not d:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(d.strip(), fmt).date()
        except ValueError:
            pass
    return None


def get_month_key(d: date) -> str:
    return f"{d.year}-{d.month:02d}"


def next_tva_date() -> tuple[date, int]:
    today = date.today()
    # TVA due aux mois de Jan, Avr, Jul, Oct (trimestrielle par défaut SASU)
    quarters = [date(today.year, 1, 15), date(today.year, 4, 15),
                date(today.year, 7, 15), date(today.year, 10, 15)]
    for qd in quarters:
        if qd > today:
            return qd, (qd - today).days
    return date(today.year + 1, 1, 15), (date(today.year + 1, 1, 15) - today).days


def build_report(config: dict, invoices: list[dict], output_path: Path):
    sasu = config.get("sasu", {})
    obj_mensuel = float(config.get("tableau_bord", {}).get("objectif_mensuel", 10000))
    obj_annuel = float(config.get("tableau_bord", {}).get("objectif_annuel", 120000))
    tjm_cible = float(sasu.get("tjm_cible", 650))

    doc = setup_doc(
        "TABLEAU DE BORD MENSUEL SASU",
        "Suivi de gestion : CA, TJM, taux d'occupation et projections",
        config.get("author", "Virginie Benayoun"),
    )

    if not invoices:
        para(doc, random.choice([
            "Aucune facture renseignée pour le mois en cours.",
            "Pas de facture saisie pour ce mois-ci.",
            "Aucune donnee de facturation pour la periode en cours.",
        ]), bold=True, color=RED)
        doc.save(output_path)
        return

    # Parse invoices
    parsed = []
    for inv in invoices:
        d = parse_date(inv.get("date", ""))
        if not d:
            continue
        try:
            montant = float(inv.get("montant_ht", inv.get("montant", 0)))
        except (ValueError, TypeError):
            continue
        try:
            tjm = float(inv.get("tjm", 0))
        except (ValueError, TypeError):
            tjm = 0
        try:
            jours = float(inv.get("jours", inv.get("nb_jours", 0)))
        except (ValueError, TypeError):
            jours = 0
        parsed.append({
            "date": d,
            "month_key": get_month_key(d),
            "month_name": MONTHS_FR[d.month],
            "year": d.year,
            "montant_ht": montant,
            "tjm": tjm,
            "jours": jours,
            "client": inv.get("client", "Inconnu"),
            "mission": inv.get("mission", ""),
        })

    if not parsed:
        para(doc, random.choice([
            "Impossible de lire les factures : verifie le format des fichiers.",
            "Erreur de lecture du fichier de factures : format peut-être invalide.",
            "Le fichier de factures n'a pas pu etre lu. Verifier le format et les colonnes attendues.",
        ]), bold=True, color=RED)
        doc.save(output_path)
        return

    today = date.today()
    current_month = get_month_key(today)

    # Monthly aggregation
    monthly: dict[str, dict] = defaultdict(lambda: {"ca": 0.0, "jours": 0.0, "tjm_sum": 0.0, "tjm_count": 0, "missions": set()})
    for inv in parsed:
        m = monthly[inv["month_key"]]
        m["ca"] += inv["montant_ht"]
        m["jours"] += inv["jours"]
        if inv["tjm"] > 0:
            m["tjm_sum"] += inv["tjm"]
            m["tjm_count"] += 1
        if inv["mission"]:
            m["missions"].add(inv["mission"])

    # This month
    this_month_data = monthly.get(current_month, {"ca": 0.0, "jours": 0.0, "tjm_sum": 0.0, "tjm_count": 0, "missions": set()})
    this_month_ca = this_month_data["ca"]
    this_month_days = this_month_data["jours"]
    this_month_tjm = round(this_month_data["tjm_sum"] / this_month_data["tjm_count"], 0) if this_month_data["tjm_count"] > 0 else 0

    # YTD
    ytd_months = [k for k in sorted(monthly.keys()) if k <= current_month]
    ytd_ca = sum(monthly[m]["ca"] for m in ytd_months)
    ytd_tjm_sum = sum(monthly[m]["tjm_sum"] for m in ytd_months)
    ytd_tjm_count = sum(monthly[m]["tjm_count"] for m in ytd_months)
    ytd_tjm = round(ytd_tjm_sum / ytd_tjm_count, 0) if ytd_tjm_count > 0 else 0
    ytd_days = sum(monthly[m]["jours"] for m in ytd_months)

    # Annual projection
    months_elapsed = max(len(ytd_months), 1)
    monthly_avg_ca = round(ytd_ca / months_elapsed, 0)
    projected_annual = round(monthly_avg_ca * 12, 0)
    ca_remaining = obj_annuel - ytd_ca

    # Occupancy rate (22 working days / month average)
    occupancy = round(ytd_days / (months_elapsed * 22) * 100, 1) if months_elapsed > 0 else 0

    # TVA alert
    next_tva, days_to_tva = next_tva_date()

    # --- BUILD REPORT ---
    accent_heading(doc, f"1. Synthese du mois de {MONTHS_FR[today.month]} {today.year}")
    para(doc, "Indicateurs cles", style="Heading 2")

    # Dashboard table
    table = doc.add_table(rows=1, cols=3)
    for i, h in enumerate(["Indicateur", "Valeur", "Statut"]):
        cell = table.rows[0].cells[i]
        shade_cell(cell, PALEBLUE)
        cell.text = h
        for p in cell.paragraphs:
            for r in p.runs:
                set_run_font(r, bold=True, color=NAVY, size=9.5)

    indicators = [
        ("CA du mois", f"{this_month_ca:,.0f} €"),
        ("CA cumulé (YTD)", f"{ytd_ca:,.0f} €"),
        ("TJM moyen (YTD)", f"{ytd_tjm:,.0f} €"),
        ("TJM cible", f"{tjm_cible:,.0f} €"),
        ("Jours facturés (YTD)", f"{ytd_days}"),
        ("Taux d'occupation", f"{occupancy}%"),
        ("Moyenne mensuelle", f"{monthly_avg_ca:,.0f} €"),
        ("Objectif mensuel", f"{obj_mensuel:,.0f} €"),
        ("Projection annuelle", f"{projected_annual:,.0f} €"),
        ("Objectif annuel", f"{obj_annuel:,.0f} €"),
        ("Reste à facturer (objectif)", f"{max(ca_remaining, 0):,.0f} €"),
    ]

    for label, value in indicators:
        cells = table.add_row().cells
        cells[0].text = label
        cells[1].text = value
        for p in cells[0].paragraphs:
            for r in p.runs:
                set_run_font(r, bold=True, size=9.2)
        for p in cells[1].paragraphs:
            for r in p.runs:
                set_run_font(r, size=9.2)

        # Status color
        status_cell = cells[2]
        ok = True
        if "CA du mois" in label:
            ok = this_month_ca >= obj_mensuel
        elif "cumulé" in label:
            ok = ytd_ca >= obj_annuel * months_elapsed / 12
        elif "TJM" in label and "cible" not in label and "moyen" in label:
            ok = ytd_tjm >= tjm_cible
        elif "occupation" in label:
            ok = occupancy >= 70
        elif "Projection" in label:
            ok = projected_annual >= obj_annuel
        status_cell.text = "✓ OK" if ok else "⚠ Alerte"
        for p in status_cell.paragraphs:
            for r in p.runs:
                set_run_font(r, bold=True, size=9.2, color=GREEN if ok else RED)

    set_table_widths(table, [2.2, 1.8, 1.0])
    apply_table_borders(table)

    # Section 2: Mensual detail
    add_separator(doc)
    accent_heading(doc, "2. Detail mensuel")
    m_table = doc.add_table(rows=1, cols=5)
    for i, h in enumerate(["Mois", "CA (€)", "Jours", "TJM moyen", "Missions"]):
        cell = m_table.rows[0].cells[i]
        shade_cell(cell, PALEBLUE)
        cell.text = h
        for p in cell.paragraphs:
            for r in p.runs:
                set_run_font(r, bold=True, color=NAVY, size=8.5)

    for mk in sorted(monthly.keys(), reverse=True):
        m = monthly[mk]
        cells = m_table.add_row().cells
        year, month_num = mk.split("-")
        month_label = f"{MONTHS_FR[int(month_num)]} {year}"
        cells[0].text = month_label
        cells[1].text = f"{m['ca']:,.0f}"
        cells[2].text = f"{m['jours']}"
        tjm_m = round(m["tjm_sum"] / m["tjm_count"], 0) if m["tjm_count"] > 0 else 0
        cells[3].text = f"{tjm_m:,.0f}"
        cells[4].text = ", ".join(m["missions"])[:80]
        for cell in cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    set_run_font(r, size=8.5)

    set_table_widths(m_table, [1.5, 1.2, 0.6, 1.0, 2.2])
    apply_table_borders(m_table)

    # Section 3: TVA Alert
    add_separator(doc)
    accent_heading(doc, "3. Alertes et echeances")
    if days_to_tva <= 30:
        bullet(doc, f"⚠ TVA à déclarer dans {days_to_tva} jours (le {next_tva.isoformat()})", size=10.5)
    else:
        bullet(doc, f"Prochaine échéance TVA : {next_tva.isoformat()} (dans {days_to_tva} jours)", size=10.5)

    if projected_annual < obj_annuel:
        delta = obj_annuel - projected_annual
        bullet(doc, f"⚠ Projection annuelle ({projected_annual:,.0f} €) sous l'objectif ({obj_annuel:,.0f} €) de {delta:,.0f} €", size=10.5)
        bullet(doc, f"Recommandation : générer {monthly_avg_ca + delta / max(12 - months_elapsed, 1):,.0f} €/mois pour atteindre l'objectif", size=10.5)
    else:
        bullet(doc, f"✓ Projection annuelle ({projected_annual:,.0f} €) au-dessus de l'objectif ({obj_annuel:,.0f} €)", size=10.5)

    # Section 4: Prospection gap
    add_separator(doc)
    accent_heading(doc, "4. Analyse et recommandations")
    if occupancy < 50:
        bullet(doc, "⚠ Faible taux d'occupation - intensifier la prospection (cf. Agent Chasseur de Leads)")
    elif occupancy < 75:
        bullet(doc, "Taux d'occupation modéré - maintenir les efforts de prospection")
    else:
        bullet(doc, "✓ Bon taux d'occupation - focus sur la qualité des missions et la montée en TJM")

    if this_month_ca < obj_mensuel * 0.8:
        bullet(doc, f"⚠ Mois en dessous de l'objectif ({this_month_ca:,.0f} € vs {obj_mensuel:,.0f} € visés)")
    bullet(doc, f"Objectif recommandé pour le prochain mois : {monthly_avg_ca:,.0f} € (maintien du rythme actuel)")

    doc.save(output_path)
    return output_path


def main() -> int:
    config = load_config()
    agent_cfg = config.get("tableau_bord", {})
    if not agent_cfg.get("enabled", True):
        print("Agent Tableau de bord désactivé.")
        return 0

    print(">>> Agent 3 : Tableau de bord mensuel SASU")
    factures_dir = agent_cfg.get("factures_dir", "")
    Path(factures_dir).mkdir(parents=True, exist_ok=True)

    invoices = load_invoices(factures_dir)
    print(f"  Factures chargées: {len(invoices)}")

    output_dir = ensure_output_dir(config)
    today = date.today().isoformat()
    output_path = output_dir / f"Tableau_Bord_SASU_{today}.docx"
    build_report(config, invoices, output_path)

    run_log = output_dir / f"Tableau_Bord_SASU_{today}_log.json"
    save_json_report(
        {
            "generated_at": datetime.now(timezone.utc).isoformat(),
            "invoices_count": len(invoices),
            "month": today[:7],
        },
        run_log,
    )

    print(f"  Rapport généré : {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

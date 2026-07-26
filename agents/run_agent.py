from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

AGENTS = {
    "missions": ("agent_chasseur_leads.chasseur", "Chasseur de Missions"),
    "pipeline": ("agent_pipeline_crm.pipeline", "Suivi Pipeline CRM"),
    "tableau": ("agent_tableau_bord.tableau_bord", "Tableau de bord mensuel"),
    "cr": ("agent_cr_reunion.cr_reunion", "Compte-rendu de reunion"),
    "branding": ("agent_ghostwriter.ghostwriter", "Personal Branding"),
    "proposition": ("agent_propositions.propositions", "Offres consulting"),
    "tarif": ("agent_veille_tarifaire.veille_tarifaire", "Veille tarifaire"),
}

# Ces agents sont conservés mais ne génèrent plus de contenu (plus de valeur réelle).
# Réactiver quand l'activité le justifiera.
_DISABLED_AGENTS = {"pipeline", "tableau"}


def usage():
    print("Usage: run_agent.py [agent_name|all|list]")
    print("  Agents disponibles :")
    for key, (_, desc) in AGENTS.items():
        print(f"    {key:15s} - {desc}")


def main() -> int:
    if len(sys.argv) < 2:
        usage()
        return 1

    cmd = sys.argv[1].lower()

    if cmd == "list":
        usage()
        return 0

    if cmd == "all":
        for key, (mod_path, desc) in AGENTS.items():
            if key in _DISABLED_AGENTS:
                print(f"\n  [{key}] Desactive (plus de valeur) - fichier conserve pour reactivation future")
                continue
            print(f"\n{'='*60}")
            print(f"  Lancement : {desc}")
            print(f"{'='*60}")
            try:
                mod = __import__(mod_path, fromlist=["main"])
                rc = mod.main()
                if rc != 0:
                    print(f"  [!] Agent {desc} termine avec code {rc}")
            except Exception as e:
                print(f"  [ERR] {desc} : {e}")
        return 0

    if cmd in _DISABLED_AGENTS:
        print(f"Agent [{cmd}] desactive - plus de valeur pour le moment. Fichier conserve.")
        return 0

    if cmd in AGENTS:
        mod_path, desc = AGENTS[cmd]
        print(f"  Lancement : {desc}")
        try:
            mod = __import__(mod_path, fromlist=["main"])
            return mod.main()
        except Exception as e:
            print(f"  [ERR] {desc} : {e}")
            return 1

    print(f"Agent inconnu: {cmd}")
    usage()
    return 1


if __name__ == "__main__":
    raise SystemExit(main())

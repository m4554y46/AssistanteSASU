from __future__ import annotations

from datetime import date
from pathlib import Path

NOTES_DIR = Path(__file__).resolve().parent / "notes"


def generate() -> str:
    return ""


def main() -> int:
    NOTES_DIR.mkdir(parents=True, exist_ok=True)
    print("  Aucune note manuelle trouvee - generation automatique desactivee (plus de valeur reelle).")
    print("  Deposer un fichier .txt dans agents/agent_cr_reunion/notes/ pour generer un CR.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

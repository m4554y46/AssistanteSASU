from __future__ import annotations

import json
import random
import sys
from datetime import date
from pathlib import Path

from rapport_engine.collectors import collect, collect_tokenforge
from rapport_engine.docx_report import build_report
from rapport_engine.scoring import extract_tokenforge_items, rank_items


ROOT = Path(__file__).resolve().parent
CONFIG_PATH = ROOT / "config.json"


def load_config() -> dict:
    with CONFIG_PATH.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def choose_methods(config: dict) -> list[dict]:
    methods = list(config["method_library"])
    seed = int(date.today().strftime("%Y%U"))
    random.Random(seed).shuffle(methods)
    return methods[: int(config.get("max_methods", 4))]


def main() -> int:
    config = load_config()
    print("Collecte dynamique des articles et opportunites...")
    articles_raw, opportunities_raw, meta = collect(config)
    print(f"Resultats bruts: {len(articles_raw)} articles, {len(opportunities_raw)} opportunites")

    keywords = config["keywords"]
    articles = rank_items(
        articles_raw,
        keywords["article_positive"],
        keywords["negative"],
        int(config.get("max_articles", 7)),
    )
    opportunities = rank_items(
        opportunities_raw,
        keywords["opportunity_positive"],
        keywords["negative"],
        int(config.get("max_opportunities", 5)),
    )
    methods = choose_methods(config)

    print("Collecte TokenForge...")
    tokenforge_raw = collect_tokenforge(config)
    print(f"  {len(tokenforge_raw)} resultats TokenForge bruts")
    tokenforge_items = extract_tokenforge_items(
        tokenforge_raw,
        keywords.get("tokenforge_positive", []),
        int(config.get("max_tokenforge", 8)),
    )
    print(f"  {len(tokenforge_items)} resultats TokenForge retenus")

    output_dir = Path(config["output_dir"])
    output_path = output_dir / f"{config['report_prefix']}_{date.today().isoformat()}.docx"
    build_report(config, articles, opportunities, methods, meta, output_path, tokenforge_items)

    run_log = output_dir / f"{config['report_prefix']}_{date.today().isoformat()}_runlog.json"
    run_log.write_text(
        json.dumps(
            {
                "meta": meta,
                "articles": articles,
                "opportunities": opportunities,
                "methods": methods,
                "tokenforge_items": tokenforge_items,
                "output": str(output_path),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(f"Rapport genere: {output_path}")
    print(f"Journal de collecte: {run_log}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

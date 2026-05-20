#!/usr/bin/env python3
"""Auto-scrape new EuroMillions draws and bump .version if any inserted.

Used by .github/workflows/scrape-draws.yml as a scheduled GitHub Action.
Runs the same scraper used by /api/actualizar-recentes, but against the
checked-out repo DB so changes can be committed back.
"""
import os
import sys
import datetime
from pathlib import Path

os.environ.pop("VERCEL", None)

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from euromilhoes import DatabaseManager, HistoricoScraper


MESES_PT = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
            "Jul", "Ago", "Set", "Out", "Nov", "Dez"]


def bump_version() -> tuple[int, str]:
    version_file = ROOT / ".version"
    raw = version_file.read_text(encoding="utf-8").strip().split("\n")
    new_num = int(raw[0]) + 1
    today = datetime.date.today()
    new_date = f"{today.day:02d}.{MESES_PT[today.month - 1]}.{today.year}"
    version_file.write_text(f"{new_num}\n{new_date}\n", encoding="utf-8")
    return new_num, new_date


def main() -> int:
    db = DatabaseManager()
    ultimo = db.ultimo_sorteio()
    if not ultimo:
        print("BD vazia, nada para atualizar.")
        return 0

    print(f"Ultimo sorteio na BD: {ultimo['data']}")

    scraper = HistoricoScraper()
    resultado = scraper.scrape_desde(ultimo["data"], db)

    print(f"Encontrados: {resultado['encontrados']}, Inseridos: {resultado['inseridos']}")
    for s in resultado["sorteios"]:
        print(f"  {s['data']}: {s['numeros']} + {s['estrelas']}")

    if resultado["inseridos"] > 0:
        new_num, new_date = bump_version()
        print(f"Version bumped -> {new_num} ({new_date})")
    else:
        print("Sem sorteios novos. Version inalterada.")

    return 0


if __name__ == "__main__":
    sys.exit(main())

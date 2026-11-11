#!/usr/bin/env python3
"""Auto-scrape new EuroMillions draws and their prizes; bump .version if any
draw was inserted.

Used by .github/workflows/scrape-draws.yml as a scheduled GitHub Action.
Runs the same scrapers used by /api/actualizar-recentes and
/api/scrape-premios, but against the checked-out repo DB so changes can be
committed back. Writes made by the live app do not survive, because on Vercel
the DB lives in /tmp — this script is the only durable write path.
"""
import os
import sqlite3
import sys
import time
import datetime
from pathlib import Path

os.environ.pop("VERCEL", None)

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from euromilhoes import (DB_PATH, SCRAPE_DELAY, DatabaseManager,
                         HistoricoScraper, PremiosScraper)


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


# Quantos sorteios sem prémios se tentam por execução. Normalmente são 0 ou
# 1; o limite só conta se alguma execução anterior tiver falhado, e existe
# para a job não passar do timeout de 10 minutos do workflow.
MAX_PREMIOS_POR_EXECUCAO = 30


def recolher_premios_em_falta(db: DatabaseManager) -> int:
    """Preenche os prémios dos sorteios que ainda não os têm, do mais recente
    para o mais antigo.

    Os prémios são publicados algumas horas depois do sorteio, por isso a
    primeira execução depois de um sorteio costuma não os encontrar. Cada
    execução volta a tentar o que faltou nas anteriores, o que torna isto
    auto-reparável sem precisar de guardar estado.
    """
    with sqlite3.connect(DB_PATH) as conn:
        faltam = [r[0] for r in conn.execute(
            "SELECT s.data FROM sorteios s "
            "WHERE NOT EXISTS (SELECT 1 FROM premios p WHERE p.data = s.data) "
            "ORDER BY s.data DESC LIMIT ?", (MAX_PREMIOS_POR_EXECUCAO,))]

    if not faltam:
        print("Premios: nenhum em falta.")
        return 0

    print(f"Premios em falta: {len(faltam)} (maximo {MAX_PREMIOS_POR_EXECUCAO} por execucao)")
    scraper = PremiosScraper()
    recolhidos = 0
    for data in faltam:
        try:
            dados = scraper.scrape_premios(data)
            if dados and db.inserir_premios(data, dados):
                recolhidos += 1
                print(f"  {data}: premios recolhidos")
            else:
                print(f"  {data}: ainda sem premios publicados")
        except Exception as e:
            print(f"  {data}: erro - {str(e)[:120]}")
        time.sleep(SCRAPE_DELAY)

    print(f"Premios: {recolhidos} de {len(faltam)} recolhidos.")
    return recolhidos


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

    # Os premios vem sempre depois dos sorteios: um sorteio inserido agora
    # pode ainda nao os ter publicados, e fica para a execucao seguinte.
    recolher_premios_em_falta(db)

    return 0


if __name__ == "__main__":
    sys.exit(main())

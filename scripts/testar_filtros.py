#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Verifica que as quatro cópias da lógica dos filtros continuam a concordar.

PORQUE É QUE ISTO EXISTE
  A mesma regra está escrita em quatro sítios, de propósito:

    1. FilterEngine.verificar()            — curto-circuita, corre milhões
                                             de vezes durante a geração
    2. FilterEngine.avaliar_todos()        — avalia os nove independentemente,
                                             para a tabela histórica
    3. app._compute_filter_histogram()     — versão otimizada que percorre
                                             C(50,5) para o painel interativo
    4. a fórmula fechada do filtro F em StatisticsAnalyzer.historico_filtros()
                                             — a taxa "ao acaso" das cores

  Já aconteceu alterar-se uma regra e a nº 4 ficar para trás: a coluna
  "ao acaso" das cores ficou 23 pontos abaixo da taxa real, o que faria
  parecer que o filtro tinha um poder preditivo enorme. Este teste apanha
  isso de imediato.

CORRER
  python scripts/testar_filtros.py

  Devolve 0 se tudo concordar, 1 se alguma cópia se desviou. Demora alguns
  minutos, porque percorre as 2.118.760 combinações mais de uma vez — é para
  correr quando se muda uma regra, não a cada alteração.
"""
import os
import random
import sys
import time
from itertools import combinations
from pathlib import Path

os.environ.pop("VERCEL", None)

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
os.chdir(ROOT)

from euromilhoes import (COLOR_WINDOW_DRAWS, DatabaseManager, FilterEngine,
                         StatisticsAnalyzer)

# Os filtros E e F dependem do contexto (sorteio anterior e cores), por isso
# não têm valor fixo em REJEITADAS_AO_ACASO e são verificados à parte.
SEM_CONTEXTO = "ABCDGHI"

falhas: list[str] = []


def verificar(descricao: str, obtido, esperado) -> None:
    ok = obtido == esperado
    if not ok:
        falhas.append(f"{descricao}: obtido {obtido}, esperado {esperado}")
    print(f"   {'OK  ' if ok else 'FALHA'}  {descricao:<46} {obtido}")


def main() -> int:
    t0 = time.time()
    db = DatabaseManager()
    stats = StatisticsAnalyzer(db)
    cores = stats.classificar_cores(db.ultimos_n_sorteios(COLOR_WINDOW_DRAWS))
    ultimo = db.ultimo_sorteio()
    ultimo_nums = ultimo["numeros"] if ultimo else []

    print("1. avaliar_todos() contra as constantes REJEITADAS_AO_ACASO")
    print("   (percorre as %d combinações)" % FilterEngine.TOTAL_COMBINACOES)
    rejeitadas = dict.fromkeys(SEM_CONTEXTO, 0)
    total = 0
    for chave in combinations(range(1, 51), 5):
        total += 1
        veredicto = FilterEngine.avaliar_todos(list(chave))
        for f in SEM_CONTEXTO:
            if not veredicto[f]:
                rejeitadas[f] += 1
    verificar("total de combinações", total, FilterEngine.TOTAL_COMBINACOES)
    for f in SEM_CONTEXTO:
        verificar(f"filtro {f} rejeita", rejeitadas[f],
                  FilterEngine.REJEITADAS_AO_ACASO[f])

    print("")
    print("2. _compute_filter_histogram() de app.py contra as mesmas constantes")
    import app as appmod
    hist = appmod._compute_filter_histogram(cores, ultimo_nums)
    verificar("total do histograma", hist["total"], FilterEngine.TOTAL_COMBINACOES)
    for f in SEM_CONTEXTO:
        verificar(f"filtro {f} elimina", hist["per_filter"][f]["eliminadas"],
                  FilterEngine.REJEITADAS_AO_ACASO[f])

    print("")
    print("3. fórmula fechada do filtro F contra força bruta")
    # Chama o código real — não repetir aqui a fórmula, senão este teste passa
    # a ser mais uma cópia e deixa de detetar o desvio que devia apanhar.
    fechada = round(FilterEngine.fracao_cores_ao_acaso(cores)
                    * FilterEngine.TOTAL_COMBINACOES)
    bruta = sum(1 for chave in combinations(range(1, 51), 5)
                if FilterEngine.avaliar_todos(list(chave), None, cores)["F"])
    verificar("fracao_cores_ao_acaso() contra força bruta", fechada, bruta)
    verificar("igual ao histograma de app.py",
              hist["per_filter"]["F"]["aceites"], bruta)

    print("")
    print("4. verificar() contra avaliar_todos() em chaves aleatórias")
    motor = FilterEngine(cores, ultimo_nums)
    random.seed(0)
    divergencias = 0
    amostra = 300_000
    for _ in range(amostra):
        chave = sorted(random.sample(range(1, 51), 5))
        passa_rapido = motor.verificar(chave)
        veredicto = FilterEngine.avaliar_todos(
            chave, ultimo_nums, cores,
            soma_min=motor.soma_min, soma_max=motor.soma_max)
        passa_completo = all(v for v in veredicto.values() if v is not None)
        if passa_rapido != passa_completo:
            divergencias += 1
            if divergencias <= 3:
                print(f"      {chave}: verificar={passa_rapido} "
                      f"avaliar_todos={passa_completo} {veredicto}")
    verificar(f"divergências em {amostra:,} chaves".replace(",", "."),
              divergencias, 0)

    print("")
    print("Demorou %.0f segundos." % (time.time() - t0))
    if falhas:
        print("")
        print("%d PROBLEMA(S) — alguma cópia da lógica ficou para trás:" % len(falhas))
        for f in falhas:
            print("   " + f)
        print("")
        print("Se a alteração a uma regra foi intencional, é preciso:")
        print("  - alterar as quatro cópias (ver o cabeçalho deste ficheiro);")
        print("  - recalcular REJEITADAS_AO_ACASO com os valores que este")
        print("    teste imprimiu na secção 1.")
        return 1
    print("Todas as cópias da lógica concordam.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

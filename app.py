#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Flask web server for the EuroMilhões key generator.
Run with:  python app.py
Then open: http://localhost:5051
"""

import os
import json
import datetime
import threading
from itertools import combinations
from collections import Counter as _Counter
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file

# Import core logic from euromilhoes.py
from euromilhoes import (
    DatabaseManager, StatisticsAnalyzer, FilterEngine,
    KeyGenerator, ExcelExporter, EuromilhoesScraper, HistoricoScraper,
    VERSION, PADROES_EQUILIBRADOS, BI, BP, AI, AP, HISTORICO_PATH,
    ExcelImporter, EXCEL_SOURCE_PATH,
    TODOS_PADROES_BIBPAIAP, classificar_padrao_bibpaiap, classificar_padrao_cores,
    corrigir_data_sorteio,
    PrizeChecker, PremiosScraper, PRIZE_TIERS, CUSTO_POR_APOSTA,
)

_IS_VERCEL = bool(os.environ.get("VERCEL"))

app = Flask(__name__)

# ── Shared singletons ────────────────────────────────────────────────────────
db      = DatabaseManager()
stats   = StatisticsAnalyzer(db)
exporter = ExcelExporter()
_scrape_lock = threading.Lock()


def _cores_serializable(cores: dict) -> dict:
    """Convert sets to sorted lists for JSON serialisation."""
    return {k: sorted(list(v)) for k, v in cores.items()}


def _get_cores():
    ultimos_9 = db.ultimos_n_sorteios(9)
    return stats.classificar_cores(ultimos_9), ultimos_9


def _sync_historico_json():
    """Re-export all DB draws to historico_completo.json so the file stays current."""
    todos = db.todos_sorteios()
    data = {
        "gerado_em": datetime.date.today().isoformat(),
        "total": len(todos),
        "sorteios": todos,
    }
    HISTORICO_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


# ════════════════════════════════════════════════════════════════════════════
# MAIN PAGE
# ════════════════════════════════════════════════════════════════════════════
@app.route("/")
def index():
    return render_template("index.html", version=VERSION)


# ════════════════════════════════════════════════════════════════════════════
# API – STATUS
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/status")
def api_status():
    ultimo = db.ultimo_sorteio()
    return jsonify({
        "total_sorteios": db.total_sorteios(),
        "ultimo_sorteio": ultimo,
        "version": VERSION,
    })


# ════════════════════════════════════════════════════════════════════════════
# API – GENERATE KEYS
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/gerar", methods=["POST"])
def api_gerar():
    body = request.get_json(silent=True) or {}
    quantidade = max(1, min(int(body.get("quantidade", 10)), 100))

    # Support per-filter toggles (filtro_A … filtro_H).
    # Fall back to legacy keys regra31/progressao for G and H.
    cfg = {
        "soma_range": body.get("soma_range", "padrao"),
        "filtro_A":   bool(body.get("filtro_A", True)),
        "filtro_B":   bool(body.get("filtro_B", True)),
        "filtro_C":   bool(body.get("filtro_C", True)),
        "filtro_D":   bool(body.get("filtro_D", True)),
        "filtro_E":   bool(body.get("filtro_E", True)),
        "filtro_F":   bool(body.get("filtro_F", True)),
        "filtro_G":   bool(body.get("filtro_G", body.get("regra31",   True))),
        "filtro_H":   bool(body.get("filtro_H", body.get("progressao", True))),
    }

    # #33/#34 Exclusion/inclusion lists
    excluir = set(int(n) for n in body.get("excluir", []) if 1 <= int(n) <= 50) if body.get("excluir") else set()
    incluir = set(int(n) for n in body.get("incluir", []) if 1 <= int(n) <= 50) if body.get("incluir") else set()

    cores, _ = _get_cores()
    ultimo   = db.ultimo_sorteio()
    ultimo_nums = ultimo["numeros"] if ultimo else []

    fe  = FilterEngine(cores, ultimo_nums, config=cfg)
    gen = KeyGenerator(fe, db)

    chaves_geradas = []
    total_tentativas = 0
    max_outer = quantidade * 200  # safety limit
    outer_tries = 0

    while len(chaves_geradas) < quantidade and outer_tries < max_outer:
        outer_tries += 1
        chave = gen.gerar_chave()
        if not chave:
            break
        total_tentativas += chave["tentativas"]
        nums_set = set(chave["numeros"])
        # Check exclusion: none of the excluded numbers should be present
        if excluir and nums_set & excluir:
            continue
        # Check inclusion: all included numbers must be present
        if incluir and not incluir.issubset(nums_set):
            continue
        chaves_geradas.append(chave)

    # Save to generation history
    hist_entry = {
        "id": str(int(datetime.datetime.now().timestamp() * 1000)),
        "data": datetime.datetime.now().isoformat(),
        "quantidade": len(chaves_geradas),
        "config": cfg,
    }
    try:
        hist = json.loads(db.get_metadata("historico_geracoes") or "[]")
        hist.insert(0, hist_entry)
        if len(hist) > 20:
            hist = hist[:20]
        db.set_metadata("historico_geracoes", json.dumps(hist, ensure_ascii=False))
    except Exception:
        pass

    return jsonify({
        "chaves": chaves_geradas,
        "cores": _cores_serializable(cores),
        "total_geradas": len(chaves_geradas),
        "total_pedidas": quantidade,
        "total_tentativas": total_tentativas,
        "filtros": fe.resumo_filtros(),
        "config": cfg,
    })


# ════════════════════════════════════════════════════════════════════════════
# API – ÚLTIMO SORTEIO + CORES
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/ultimo-sorteio")
def api_ultimo_sorteio():
    ultimo = db.ultimo_sorteio()
    cores, ultimos_9 = _get_cores()
    # Colors computed from the 9 draws BEFORE the last one, so the UI can show
    # which colour each drawn number had at the moment the draw happened.
    ultimos_10 = db.ultimos_n_sorteios(10)
    anteriores = ultimos_10[1:] if len(ultimos_10) > 1 else ultimos_10
    cores_antes = stats.classificar_cores(anteriores)

    # BI-BP-AI-AP pattern and colour pattern for the last draw
    padrao_comb = classificar_padrao_bibpaiap(ultimo["numeros"]) if ultimo else None
    cores_antes_sets = {k: set(v) for k, v in _cores_serializable(cores_antes).items()}
    padrao_cores_str = classificar_padrao_cores(ultimo["numeros"], cores_antes_sets) if ultimo else None

    return jsonify({
        "ultimo": ultimo,
        "cores": _cores_serializable(cores),
        "cores_antes": _cores_serializable(cores_antes),
        "ultimos_9": ultimos_9,
        "padrao_combinatorio": padrao_comb,
        "padrao_cores": padrao_cores_str,
    })


# ════════════════════════════════════════════════════════════════════════════
# API – ALL DRAWS (paginated)
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/sorteios")
def api_sorteios():
    page  = max(1, int(request.args.get("page", 1)))
    per   = min(50, max(5, int(request.args.get("per", 20))))
    todos = db.todos_sorteios()
    total = len(todos)
    start = (page - 1) * per
    end   = start + per
    return jsonify({
        "sorteios": todos[start:end],
        "total": total,
        "page": page,
        "per": per,
        "pages": (total + per - 1) // per,
    })


# ════════════════════════════════════════════════════════════════════════════
# API – ALL DRAWS GROUPED BY YEAR + MONTH
# ════════════════════════════════════════════════════════════════════════════
_MESES = ["", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
          "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

@app.route("/api/sorteios-agrupados")
def api_sorteios_agrupados():
    todos = db.todos_sorteios()
    dados: dict = {}
    for s in todos:
        year      = s["data"][:4]
        month_num = int(s["data"][5:7])
        month_key = s["data"][:7]          # "YYYY-MM"
        if year not in dados:
            dados[year] = {}
        if month_key not in dados[year]:
            dados[year][month_key] = {
                "nome": _MESES[month_num],
                "sorteios": [],
            }
        dados[year][month_key]["sorteios"].append(s)

    anos_ordenados = sorted(dados.keys(), reverse=True)
    return jsonify({
        "total": len(todos),
        "anos": anos_ordenados,
        "dados": dados,
    })


# ════════════════════════════════════════════════════════════════════════════
# API – INSERT DRAW MANUALLY
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/sorteio", methods=["POST"])
def api_inserir_sorteio():
    body = request.get_json(silent=True) or {}
    try:
        data   = body["data"]
        nums   = sorted([int(n) for n in body["numeros"]])
        stars  = sorted([int(s) for s in body["estrelas"]])
        datetime.date.fromisoformat(data)
        assert len(nums) == 5 and len(set(nums)) == 5
        assert all(1 <= n <= 50 for n in nums)
        assert len(stars) == 2 and len(set(stars)) == 2
        assert all(1 <= s <= 12 for s in stars)
    except (KeyError, ValueError, AssertionError) as e:
        return jsonify({"erro": f"Dados inválidos: {e}"}), 400

    data = corrigir_data_sorteio(data)
    ok = db.inserir_sorteio(data, nums, stars, fonte="web-manual")
    if ok:
        _sync_historico_json()
    return jsonify({"ok": ok, "mensagem": "Guardado." if ok else "Já existia na BD."})


# ════════════════════════════════════════════════════════════════════════════
# API – DELETE DRAW
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/sorteio/<data>", methods=["DELETE"])
def api_eliminar_sorteio(data):
    ok = db.eliminar_sorteio(data)
    if ok:
        _sync_historico_json()
    return jsonify({"ok": ok, "mensagem": "Eliminado." if ok else "Sorteio não encontrado."})


# ════════════════════════════════════════════════════════════════════════════
# API – UPDATE MISSING DRAWS (since last DB entry)
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/actualizar-recentes", methods=["POST"])
def api_actualizar_recentes():
    if not _scrape_lock.acquire(blocking=False):
        return jsonify({"erro": "Já existe um scraping em curso."}), 429

    try:
        ultimo = db.ultimo_sorteio()
        if not ultimo:
            return jsonify({"erro": "BD vazia. Importa primeiro o histórico."}), 400

        desde_data = ultimo["data"]
        scraper = HistoricoScraper()
        resultado = scraper.scrape_desde(desde_data, db)

        if resultado["inseridos"] > 0:
            _sync_historico_json()

        return jsonify({
            "ok": True,
            "desde": desde_data,
            "encontrados": resultado["encontrados"],
            "inseridos": resultado["inseridos"],
            "sorteios": resultado["sorteios"],
        })
    except Exception as e:
        return jsonify({"erro": f"Erro no scraping: {e}"}), 502
    finally:
        _scrape_lock.release()


# ════════════════════════════════════════════════════════════════════════════
# API – STATISTICS
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/estatisticas")
def api_estatisticas():
    freq_nums   = stats.frequencia_numeros()
    freq_stars  = stats.frequencia_estrelas()
    somas       = stats.estatisticas_somas()
    padroes     = stats.analise_padroes()
    atrasados   = stats.numeros_atrasados(15)
    quentes     = stats.sequencias_quentes(10)
    total       = db.total_sorteios()
    quentes_frios = stats.quentes_frios_completo(15)
    gaps        = stats.analise_gaps()
    tendencia   = stats.tendencia_somas(30)

    # Serialise pattern keys (tuples → strings)
    padroes_serial = {str(list(k)): v for k, v in padroes.items()}

    return jsonify({
        "total_sorteios": total,
        "frequencia_numeros":  freq_nums,
        "frequencia_estrelas": freq_stars,
        "somas":               somas,
        "padroes":             padroes_serial,
        "atrasados":           atrasados,
        "quentes":             quentes,
        "quentes_frios":       quentes_frios,
        "gaps":                gaps,
        "tendencia_somas":     tendencia,
    })


# ════════════════════════════════════════════════════════════════════════════
# API – STRATEGY INFO
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/estrategias")
def api_estrategias():
    cores, _ = _get_cores()
    return jsonify({
        "version": VERSION,
        "quadrantes": {
            "BI": {"descricao": "Baixos Ímpares (1–25)", "numeros": sorted(BI)},
            "BP": {"descricao": "Baixos Pares (2–24)",   "numeros": sorted(BP)},
            "AI": {"descricao": "Altos Ímpares (27–49)", "numeros": sorted(AI)},
            "AP": {"descricao": "Altos Pares (26–50)",   "numeros": sorted(AP)},
        },
        "padroes_equilibrados": PADROES_EQUILIBRADOS,
        "total_combinacoes_universo": 2_118_760,
        "total_combinacoes_aceites":  1_333_800,
        "filtros": [
            {"id": "A", "nome": "Soma",               "descricao": "Entre 80–190 (padrão) ou 95–160 (apertado). ~93% dos sorteios históricos."},
            {"id": "B", "nome": "Consecutivos",        "descricao": "Máx 1 par consecutivo, 0 triplos. Pares: ~42% histórico, triplos: <1%."},
            {"id": "C", "nome": "Dígito Final",        "descricao": "Máx 2 números com mesmo dígito final. 3+ iguais: <4% histórico."},
            {"id": "D", "nome": "Dezenas",              "descricao": "Mín 3 dezenas diferentes (1-10, 11-20, 21-30, 31-40, 41-50)."},
            {"id": "E", "nome": "Anti-Repetição",      "descricao": "Máx 2 números em comum com o sorteio anterior."},
            {"id": "F", "nome": "Cores (9 sorteios)", "descricao": "VERMELHOS 1–3 | VERDES 1–3 | AZUIS 0–2 | CASTANHOS 0."},
            {"id": "G", "nome": "Regra do 31",         "descricao": "Mín 1 número > 31. Reduz partilha de jackpot com jogadores de aniversários."},
            {"id": "H", "nome": "Progressão Aritmét.", "descricao": "Rejeita sequências como 5,10,15,20,25. Muito populares → má escolha."},
        ],
        "estrelas": "O sistema memoriza o uso de cada estrela (1–12) e selecciona sempre as 2 menos usadas.",
        "cores_actuais": _cores_serializable(cores),
    })


# ════════════════════════════════════════════════════════════════════════════
# API – FILTER STATS (histogram for interactive toggle)
# ════════════════════════════════════════════════════════════════════════════
_filter_cache = {"key": None, "data": None}


def _compute_filter_histogram(cores, ultimo_nums):
    """Iterate all C(50,5) combos once; build 256-entry histogram keyed by
    an 8-bit mask (bit i = combo passes filter i).  The frontend can then
    compute the accepted count for ANY filter combination instantly."""

    # Lookup arrays for O(1) per-number checks (index 0..50)
    is_v = [False] * 51
    is_g = [False] * 51
    is_a = [False] * 51
    is_c = [False] * 51
    is_u = [False] * 51
    for n in cores.get("vermelhos", set()):
        is_v[n] = True
    for n in cores.get("verdes", set()):
        is_g[n] = True
    for n in cores.get("azuis", set()):
        is_a[n] = True
    for n in cores.get("castanhos", set()):
        is_c[n] = True
    for n in (ultimo_nums or []):
        is_u[n] = True
    has_ult = bool(ultimo_nums)

    hist = [0] * 256

    for combo in combinations(range(1, 51), 5):
        n0, n1, n2, n3, n4 = combo
        mask = 0

        # A: sum 80-190
        s = n0 + n1 + n2 + n3 + n4
        if 80 <= s <= 190:
            mask = 1

        # B: max 1 consecutive pair, 0 triples
        d01 = n1 - n0
        d12 = n2 - n1
        d23 = n3 - n2
        d34 = n4 - n3
        p = (d01 == 1) + (d12 == 1) + (d23 == 1) + (d34 == 1)
        if p <= 1 and not (
            (d01 == 1 and d12 == 1)
            or (d12 == 1 and d23 == 1)
            or (d23 == 1 and d34 == 1)
        ):
            mask |= 2

        # C: max 2 same last digit
        f = [0] * 10
        f[n0 % 10] += 1
        f[n1 % 10] += 1
        f[n2 % 10] += 1
        f[n3 % 10] += 1
        f[n4 % 10] += 1
        if max(f) <= 2:
            mask |= 4

        # D: min 3 decades
        if len({(n0 - 1) // 10, (n1 - 1) // 10, (n2 - 1) // 10,
                (n3 - 1) // 10, (n4 - 1) // 10}) >= 3:
            mask |= 8

        # E: max 2 from last draw
        if not has_ult or (is_u[n0] + is_u[n1] + is_u[n2] + is_u[n3] + is_u[n4]) <= 2:
            mask |= 16

        # F: colors
        qv = is_v[n0] + is_v[n1] + is_v[n2] + is_v[n3] + is_v[n4]
        qg = is_g[n0] + is_g[n1] + is_g[n2] + is_g[n3] + is_g[n4]
        qa = is_a[n0] + is_a[n1] + is_a[n2] + is_a[n3] + is_a[n4]
        qc = is_c[n0] + is_c[n1] + is_c[n2] + is_c[n3] + is_c[n4]
        if 1 <= qv <= 3 and 1 <= qg <= 3 and 0 <= qa <= 2 and qc == 0:
            mask |= 32

        # G: min 1 number > 31
        if (n0 > 31) + (n1 > 31) + (n2 > 31) + (n3 > 31) + (n4 > 31) >= 1:
            mask |= 64

        # H: not perfect arithmetic progression
        if not (d01 == d12 == d23 == d34):
            mask |= 128

        hist[mask] += 1

    # Per-filter individual stats (from the histogram)
    total = sum(hist)
    per_filter = {}
    for bit, fid in enumerate("ABCDEFGH"):
        accepted = sum(hist[m] for m in range(256) if m & (1 << bit))
        per_filter[fid] = {"aceites": accepted, "eliminadas": total - accepted}

    return {"histogram": hist, "per_filter": per_filter, "total": total}


@app.route("/api/filter-stats")
def api_filter_stats():
    cores, _ = _get_cores()
    ultimo = db.ultimo_sorteio()
    ultimo_nums = ultimo["numeros"] if ultimo else []

    # Simple cache: avoid recomputing if data hasn't changed
    cache_key = (tuple(sorted(cores.get("vermelhos", set()))),
                 tuple(sorted(cores.get("verdes", set()))),
                 tuple(sorted(cores.get("azuis", set()))),
                 tuple(sorted(cores.get("castanhos", set()))),
                 tuple(ultimo_nums))
    if _filter_cache["key"] == cache_key and _filter_cache["data"]:
        return jsonify(_filter_cache["data"])

    result = _compute_filter_histogram(cores, ultimo_nums)
    _filter_cache["key"] = cache_key
    _filter_cache["data"] = result
    return jsonify(result)


# ════════════════════════════════════════════════════════════════════════════
# API – BI-BP-AI-AP PATTERNS (full table)
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/padroes-bibpaiap")
def api_padroes_bibpaiap():
    return jsonify({
        "padroes": TODOS_PADROES_BIBPAIAP,
        "total_combinacoes": 2_118_760,
    })


# ════════════════════════════════════════════════════════════════════════════
# API – HISTORICAL PATTERN BREAKDOWN
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/historico-padroes")
def api_historico_padroes():
    return jsonify(stats.historico_padroes())


# ════════════════════════════════════════════════════════════════════════════
# API – EXPORT EXCEL (last generated set)
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/export-excel", methods=["POST"])
def api_export_excel():
    body   = request.get_json(silent=True) or {}
    chaves = body.get("chaves", [])
    if not chaves:
        return jsonify({"erro": "Sem chaves para exportar."}), 400

    filtros = body.get("filtros", None)
    config  = body.get("config",  None)

    cores, _ = _get_cores()
    # Convert colour lists back to sets
    cores_sets = {k: set(v) for k, v in cores.items()}

    filepath = exporter.exportar(chaves, cores_sets, filtros=filtros, config=config)
    return send_file(str(filepath), as_attachment=True,
                     download_name=filepath.name,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ════════════════════════════════════════════════════════════════════════════
# API – HISTORICO FILE STATUS
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/historico-status")
def api_historico_status():
    if not HISTORICO_PATH.exists():
        return jsonify({"existe": False})
    try:
        data = json.loads(HISTORICO_PATH.read_text(encoding="utf-8"))
        return jsonify({
            "existe":       True,
            "gerado_em":    data.get("gerado_em"),
            "total":        data.get("total", 0),
            "tamanho_kb":   round(HISTORICO_PATH.stat().st_size / 1024, 1),
        })
    except Exception as e:
        return jsonify({"existe": True, "erro": str(e)})


# ════════════════════════════════════════════════════════════════════════════
# API – IMPORT EXISTING FILE INTO DB (no re-scraping)
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/importar-ficheiro", methods=["POST"])
def api_importar_ficheiro():
    if not HISTORICO_PATH.exists():
        return jsonify({"erro": "Ficheiro historico_completo.json não encontrado."}), 404
    try:
        data     = json.loads(HISTORICO_PATH.read_text(encoding="utf-8"))
        sorteios = data.get("sorteios", [])
        inseridos, ja_existiam = 0, 0
        for s in sorteios:
            ok = db.inserir_sorteio(s["data"], s["numeros"], s["estrelas"], "historico")
            if ok:
                inseridos += 1
            else:
                ja_existiam += 1
        return jsonify({
            "total_ficheiro": len(sorteios),
            "inseridos":      inseridos,
            "ja_existiam":    ja_existiam,
        })
    except Exception as e:
        return jsonify({"erro": str(e)}), 500


# ════════════════════════════════════════════════════════════════════════════
# API – IMPORT FROM LOCAL EXCEL FILE
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/importar-excel", methods=["POST"])
def api_importar_excel():
    """
    Read 'Euromilhões _ Todos os sorteios.xlsx' from the program folder
    and import all draws into the SQLite database.
    """
    try:
        importer = ExcelImporter()
        result   = importer.importar(db)
        return jsonify({
            "ok":           True,
            "inseridos":    result["inseridos"],
            "ja_existiam":  result["ja_existiam"],
            "erros":        result["erros"],
            "total_lido":   result["total_lido"],
            "total_bd":     db.total_sorteios(),
        })
    except FileNotFoundError as e:
        return jsonify({"erro": str(e)}), 404
    except Exception as e:
        return jsonify({"erro": str(e)}), 500


@app.route("/api/excel-status")
def api_excel_status():
    """Check whether the source Excel file exists."""
    exists = EXCEL_SOURCE_PATH.exists()
    return jsonify({
        "existe":   exists,
        "ficheiro": EXCEL_SOURCE_PATH.name,
        "tamanho_kb": round(EXCEL_SOURCE_PATH.stat().st_size / 1024, 1) if exists else None,
    })


# ════════════════════════════════════════════════════════════════════════════
# API – PRÓXIMO SORTEIO
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/proximo-sorteio")
def api_proximo_sorteio():
    hoje = datetime.date.today()
    weekday = hoje.weekday()
    # Next Tue (1) or Fri (4)
    dias_ate = []
    for target in [1, 4]:  # Tue, Fri
        diff = (target - weekday) % 7
        if diff == 0:
            diff = 7  # if today is draw day, next one
        dias_ate.append(diff)
    prox = hoje + datetime.timedelta(days=min(dias_ate))
    dia_semana = "Terça-feira" if prox.weekday() == 1 else "Sexta-feira"
    return jsonify({"data": prox.isoformat(), "dia_semana": dia_semana})


# ════════════════════════════════════════════════════════════════════════════
# API – FAVORITOS (save / load / delete)
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/favoritos", methods=["GET"])
def api_get_favoritos():
    favs = db.get_metadata("favoritos")
    return jsonify(json.loads(favs) if favs else [])

@app.route("/api/favoritos", methods=["POST"])
def api_save_favorito():
    body = request.get_json(silent=True) or {}
    favs = json.loads(db.get_metadata("favoritos") or "[]")
    fav = {
        "id": str(int(datetime.datetime.now().timestamp() * 1000)),
        "chaves": body.get("chaves", []),
        "data": datetime.datetime.now().isoformat(),
        "nome": body.get("nome", ""),
    }
    favs.insert(0, fav)
    if len(favs) > 50:
        favs = favs[:50]
    db.set_metadata("favoritos", json.dumps(favs, ensure_ascii=False))
    return jsonify({"ok": True, "id": fav["id"]})

@app.route("/api/favoritos/<fav_id>", methods=["DELETE"])
def api_delete_favorito(fav_id):
    favs = json.loads(db.get_metadata("favoritos") or "[]")
    favs = [f for f in favs if f["id"] != fav_id]
    db.set_metadata("favoritos", json.dumps(favs, ensure_ascii=False))
    return jsonify({"ok": True})


# ════════════════════════════════════════════════════════════════════════════
# API – HISTÓRICO DE GERAÇÕES
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/historico-geracoes", methods=["GET"])
def api_get_historico_geracoes():
    hist = db.get_metadata("historico_geracoes")
    return jsonify(json.loads(hist) if hist else [])


# ════════════════════════════════════════════════════════════════════════════
# API – CHECK KEYS AGAINST LAST DRAW
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/check-keys", methods=["POST"])
def api_check_keys():
    body = request.get_json(silent=True) or {}
    chaves = body.get("chaves", [])
    ultimo = db.ultimo_sorteio()
    if not ultimo:
        return jsonify({"erro": "Sem sorteios na BD."}), 400

    results = []
    for ch in chaves:
        nums_match = set(ch["numeros"]) & set(ultimo["numeros"])
        stars_match = set(ch["estrelas"]) & set(ultimo["estrelas"])
        results.append({
            "numeros": ch["numeros"],
            "estrelas": ch["estrelas"],
            "nums_acertados": sorted(nums_match),
            "stars_acertados": sorted(stars_match),
            "total_nums": len(nums_match),
            "total_stars": len(stars_match),
        })
    return jsonify({"ultimo": ultimo, "resultados": results})


# ════════════════════════════════════════════════════════════════════════════
# API – JSON API (public read-only data export)
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/dados.json")
def api_dados_json():
    """Public JSON API: returns all draws + basic stats for external tools."""
    todos = db.todos_sorteios()
    freq = stats.frequencia_numeros()
    freq_e = stats.frequencia_estrelas()
    return jsonify({
        "total": len(todos),
        "ultimo": todos[0] if todos else None,
        "sorteios": todos[:100],  # Last 100 draws
        "frequencia_numeros": freq,
        "frequencia_estrelas": freq_e,
        "version": VERSION,
    })


# ════════════════════════════════════════════════════════════════════════════
# API – DECADE TIMELINE (distribution by decade over time)
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/decadas-timeline")
def api_decadas_timeline():
    """Returns decade distribution per year for timeline visualization."""
    todos = db.todos_sorteios()
    timeline = {}
    for s in todos:
        year = s["data"][:4]
        if year not in timeline:
            timeline[year] = [0, 0, 0, 0, 0]
        for n in s["numeros"]:
            timeline[year][(n - 1) // 10] += 1
    result = []
    for year in sorted(timeline.keys()):
        total = sum(timeline[year])
        result.append({
            "ano": year,
            "dezenas": timeline[year],
            "total": total,
            "pcts": [round(d / total * 100, 1) if total else 0 for d in timeline[year]],
        })
    return jsonify(result)


# ════════════════════════════════════════════════════════════════════════════
# API – PRIZE CHECKER (verify multiple-combination bets)
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/verificar-aposta", methods=["POST"])
def api_verificar_aposta():
    data = request.get_json(force=True)
    numeros = data.get("numeros", [])
    estrelas = data.get("estrelas", [])
    data_sorteio = data.get("data_sorteio")

    if not data_sorteio:
        return jsonify({"erro": "Data do sorteio é obrigatória"}), 400

    import sqlite3
    # Get the draw
    with sqlite3.connect(db.db_path) as conn:
        row = conn.execute(
            "SELECT n1,n2,n3,n4,n5,e1,e2 FROM sorteios WHERE data = ?",
            (data_sorteio,)
        ).fetchone()
    if not row:
        return jsonify({"erro": f"Sorteio de {data_sorteio} não encontrado"}), 404

    sorteio_nums = [row[0], row[1], row[2], row[3], row[4]]
    sorteio_stars = [row[5], row[6]]

    # Get prize data if available
    premios = db.obter_premios(data_sorteio)

    resultado = PrizeChecker.verificar_aposta(
        numeros, estrelas, sorteio_nums, sorteio_stars, premios
    )
    if "erro" in resultado:
        return jsonify(resultado), 400
    return jsonify(resultado)


# ════════════════════════════════════════════════════════════════════════════
# API – PRIZES (get/scrape prize data for a draw)
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/premios/<data>")
def api_premios(data):
    premios = db.obter_premios(data)
    if not premios:
        # Try scraping on-demand
        scraper = PremiosScraper()
        prize_data = scraper.scrape_premios(data)
        if prize_data:
            db.inserir_premios(data, prize_data)
            premios = db.obter_premios(data)
    if not premios:
        return jsonify({"erro": "Prémios não disponíveis para esta data", "data": data}), 404
    return jsonify(premios)


# ════════════════════════════════════════════════════════════════════════════
# API – PRIZE STATISTICS (aggregate prize stats)
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/estatisticas-premios")
def api_estatisticas_premios():
    todos = db.todos_premios()
    if not todos:
        return jsonify({
            "total_sorteios_com_premios": 0,
            "maior_jackpot": None,
            "media_jackpot": 0,
            "evolucao_jackpots": [],
            "medias_por_tier": {},
        })

    jackpots = [(p["data"], p["jackpot"]) for p in todos if p.get("jackpot") and p["jackpot"] > 0]
    maior = max(jackpots, key=lambda x: x[1]) if jackpots else (None, 0)

    medias_tier = {}
    for t in range(1, 14):
        vals = [p[f"t{t}_prize"] for p in todos if p.get(f"t{t}_prize") and p[f"t{t}_prize"] > 0]
        tier_info = None
        for key, info in PRIZE_TIERS.items():
            if info["tier"] == t:
                tier_info = info
                break
        medias_tier[t] = {
            "nome": tier_info["name"] if tier_info else f"Tier {t}",
            "media": round(sum(vals) / len(vals), 2) if vals else 0,
            "min": round(min(vals), 2) if vals else 0,
            "max": round(max(vals), 2) if vals else 0,
            "total_sorteios": len(vals),
        }

    # Jackpot evolution (chronological)
    evolucao = [{"data": d, "jackpot": j} for d, j in sorted(jackpots, key=lambda x: x[0])]

    return jsonify({
        "total_sorteios_com_premios": len(todos),
        "maior_jackpot": {"data": maior[0], "valor": maior[1]} if maior[0] else None,
        "media_jackpot": round(sum(j for _, j in jackpots) / len(jackpots), 2) if jackpots else 0,
        "evolucao_jackpots": evolucao[-100:],  # Last 100
        "medias_por_tier": medias_tier,
        "total_sem_premios": len(db.datas_sem_premios()),
    })


# ════════════════════════════════════════════════════════════════════════════
# API – SCRAPE PRIZES (bulk import missing prize data)
# ════════════════════════════════════════════════════════════════════════════
_scrape_premios_status = {"running": False, "progresso": 0, "total": 0, "msg": ""}

@app.route("/api/scrape-premios", methods=["POST"])
def api_scrape_premios():
    global _scrape_premios_status
    if _scrape_premios_status["running"]:
        return jsonify({"erro": "Scrape já em curso", "status": _scrape_premios_status}), 409

    datas = db.datas_sem_premios()
    if not datas:
        return jsonify({"msg": "Todos os sorteios já têm dados de prémios", "total": 0})

    # Limit to 50 at a time to avoid timeouts
    datas = datas[:50]
    _scrape_premios_status = {"running": True, "progresso": 0, "total": len(datas), "msg": "A iniciar..."}

    def run_scrape():
        global _scrape_premios_status
        scraper = PremiosScraper()
        sucesso = 0
        falha = 0
        for i, data in enumerate(datas):
            _scrape_premios_status["progresso"] = i + 1
            _scrape_premios_status["msg"] = f"A processar {data}..."
            try:
                prize_data = scraper.scrape_premios(data)
                if prize_data:
                    db.inserir_premios(data, prize_data)
                    sucesso += 1
                else:
                    falha += 1
            except Exception:
                falha += 1
            import time
            time.sleep(1.5)
        _scrape_premios_status = {
            "running": False, "progresso": len(datas), "total": len(datas),
            "msg": f"Concluído: {sucesso} OK, {falha} falharam",
            "sucesso": sucesso, "falha": falha,
        }

    threading.Thread(target=run_scrape, daemon=True).start()
    return jsonify({"msg": f"Scrape iniciado para {len(datas)} sorteios", "total": len(datas)})

@app.route("/api/scrape-premios-status")
def api_scrape_premios_status():
    return jsonify(_scrape_premios_status)


# ════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    import os
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")
    print(f"\n  EuroMilhoes v{VERSION} - Servidor Web")
    print("  ---------------------------------------")
    print("  Acede em:  http://localhost:5051\n")
    app.run(host="0.0.0.0", port=5051, debug=False)

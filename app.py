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
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file, Response, stream_with_context

# Import core logic from euromilhoes.py
from euromilhoes import (
    DatabaseManager, StatisticsAnalyzer, FilterEngine,
    KeyGenerator, ExcelExporter, EuromilhoesScraper, HistoricoScraper,
    VERSION, PADROES_EQUILIBRADOS, BI, BP, AI, AP, HISTORICO_PATH,
    ExcelImporter, EXCEL_SOURCE_PATH,
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

    cores, _ = _get_cores()
    ultimo   = db.ultimo_sorteio()
    ultimo_nums = ultimo["numeros"] if ultimo else []

    fe  = FilterEngine(cores, ultimo_nums, config=cfg)
    gen = KeyGenerator(fe, db)

    chaves_geradas = []
    total_tentativas = 0

    for _ in range(quantidade):
        chave = gen.gerar_chave()
        if chave:
            chaves_geradas.append(chave)
            total_tentativas += chave["tentativas"]

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
    return jsonify({
        "ultimo": ultimo,
        "cores": _cores_serializable(cores),
        "ultimos_9": ultimos_9,
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

    ok = db.inserir_sorteio(data, nums, stars, fonte="web-manual")
    return jsonify({"ok": ok, "mensagem": "Guardado." if ok else "Já existia na BD."})


# ════════════════════════════════════════════════════════════════════════════
# API – WEB SCRAPE LATEST DRAW
# ════════════════════════════════════════════════════════════════════════════
@app.route("/api/actualizar-web", methods=["POST"])
def api_actualizar_web():
    if not _scrape_lock.acquire(blocking=False):
        return jsonify({"erro": "Já existe um scraping em curso."}), 429

    try:
        scraper = EuromilhoesScraper()
        result  = scraper.fetch_ultimo_sorteio()
        if not result:
            return jsonify({"erro": "Não foi possível obter dados. Tenta inserção manual."}), 502

        ok = db.inserir_sorteio(
            result["data"], result["numeros"], result["estrelas"], fonte="web-auto"
        )
        return jsonify({
            "ok": ok,
            "sorteio": result,
            "mensagem": "Guardado com sucesso." if ok else "Sorteio já existia na BD.",
        })
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
            {"id": "D", "nome": "Décadas",             "descricao": "Mín 3 décadas diferentes (0-9, 10-19, 20-29, 30-39, 40-49)."},
            {"id": "E", "nome": "Anti-Repetição",      "descricao": "Máx 2 números em comum com o sorteio anterior."},
            {"id": "F", "nome": "Cores (9 sorteios)", "descricao": "VERMELHOS 1–3 | VERDES 1–3 | AZUIS 0–1 | CASTANHOS 0."},
            {"id": "G", "nome": "Regra do 31",         "descricao": "Mín 2 números > 31. Reduz partilha de jackpot com jogadores de aniversários."},
            {"id": "H", "nome": "Progressão Aritmét.", "descricao": "Rejeita sequências como 5,10,15,20,25. Muito populares → má escolha."},
        ],
        "estrelas": "O sistema memoriza o uso de cada estrela (1–12) e selecciona sempre as 2 menos usadas.",
        "cores_actuais": _cores_serializable(cores),
    })


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
# API – SCRAPE FULL HISTORY (SSE streaming)
# ════════════════════════════════════════════════════════════════════════════
_historico_lock = threading.Lock()

@app.route("/api/scrape-historico-completo")
def api_scrape_historico_completo():
    """
    Server-Sent Events endpoint.
    Client connects and receives JSON events line by line.
    Disabled on Vercel (serverless timeout too short for full scrape).
    """
    if _IS_VERCEL:
        return jsonify({"erro": "Funcionalidade indisponível no Vercel. Use a versão local."}), 501

    if not _historico_lock.acquire(blocking=False):
        return jsonify({"erro": "Já existe um scraping em curso."}), 429

    def generate():
        try:
            scraper = HistoricoScraper()
            for evento in scraper.scrape_completo(db, HISTORICO_PATH):
                yield f"data: {json.dumps(evento, ensure_ascii=False)}\n\n"
                # Keep connection alive during long pauses
                if evento.get("tipo") == "progresso":
                    yield ": keep-alive\n\n"
        except Exception as exc:
            yield f"data: {json.dumps({'tipo': 'erro_fatal', 'msg': str(exc)})}\n\n"
        finally:
            _historico_lock.release()

    return Response(
        stream_with_context(generate()),
        content_type="text/event-stream",
        headers={
            "Cache-Control":    "no-cache",
            "X-Accel-Buffering": "no",
            "Connection":        "keep-alive",
        },
    )


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
if __name__ == "__main__":
    import os
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")
    print(f"\n  EuroMilhoes v{VERSION} - Servidor Web")
    print("  ---------------------------------------")
    print("  Acede em:  http://localhost:5051\n")
    app.run(host="0.0.0.0", port=5051, debug=False)

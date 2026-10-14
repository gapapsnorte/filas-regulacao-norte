"""Atualiza o painel das filas (todas as filas da Região Norte).

Roda no GitHub Actions (.github/workflows/atualizar.yml), que grava os dados e publica.
Uso local: python coleta/painel.py             (coleta e grava, sem publicar)
           python coleta/painel.py --se-mudou  (só coleta se o Mapa Social recarregou ou se o dia virou)

Gera:
  dados.json                resumo de todas as filas (painel)
  devolvidas.json           devolvidas por fila × unidade solicitante (abas de devolvidas)
  analises/<id>.json        detalhes agregados de cada fila, usados pelo botão "Gerar análise"
  detalhe/<id>.json         linhas do Mapa Social de cada fila (código SISREG, datas, status, risco, CID, unidade)
  detalhe/devolvidas.json   todas as devolvidas da região, linha a linha
  atraso.json               histórico do atraso do Mapa Social (uma linha por recarga), acumulado entre publicações
"""
import hashlib
import json
import os
import sys
import time
import urllib.error
import urllib.request
from collections import defaultdict
from datetime import date, datetime, timedelta
from fnmatch import fnmatchcase
from pathlib import Path

import mapa_social as m

AQUI = Path(__file__).resolve().parent
PAINEL = Path(os.environ.get("PAINEL_DIR", AQUI.parent))
PUBLICADO = os.environ.get("PAINEL_URL", "https://gapapsnorte.github.io/filas-regulacao-norte/")
R = "[regiao_saude_solicitante]={'%s'}" % m.REGIAO
# Elegíveis (fila) = pendentes + reenviadas ao regulador. "PENDENTE / FILA DE ESPERA" fica fora do painel.
PEND, REENV, DEVOL = ("SOLICITAÇÃO / PENDENTE / REGULADOR", "SOLICITAÇÃO / REENVIADA / REGULADOR",
                      "SOLICITAÇÃO / DEVOLVIDA / REGULADOR")
FILA = '[status_solicitacao_solicitacao]={"%s","%s"}' % (PEND, REENV)
DEV = '[status_solicitacao_solicitacao]={"%s"}' % DEVOL
FILA_DEV = '[status_solicitacao_solicitacao]={"%s","%s","%s"}' % (PEND, REENV, DEVOL)
MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
CORES = {"Prioridade 1": "vermelho", "Prioridade 2": "amarelo", "Prioridade 3": "verde", "Prioridade 4": "azul"}
N_PAINEL = 24          # meses mostrados no painel (e janela dos agendamentos no detalhe)
MIN_FILAS = 500        # abaixo disso a coleta é considerada falha (não publica)
META_PADRAO = {"vermelho": 30}   # referência geral quando o protocolo não tem meta cadastrada
ANOS_SERIE = 5         # série da análise: de janeiro de (ano - 5) até o mês anterior
SEM_SAIDA_DIAS = 365   # alerta "sem saída há mais de 1 ano"


def _mes_inicial_painel(hoje):
    a, mm = hoje.year, hoje.month - N_PAINEL
    while mm < 1:
        a, mm = a - 1, mm + 12
    return date(a, mm, 1)


def coletar(hoje):
    ini = hoje.year - ANOS_SERIE
    P, PA, U = "descrição_procedimento_solicitacao", "descrição_procedimento_agendamento", "nome_unidade_solicitante"
    ano_det = _mes_inicial_painel(hoje).year
    return m.consultar([
        ["fila", [P, "status_solicitacao_solicitacao", "classificacao_risco"],
         ["Count({<%s>} distinct codigo_solicitacao)" % R, "Date(Min({<%s>} data_solicitacao))" % R]],
        ["in_sol", [P, "Calendar.Ano_Solicitacao", "Calendar.Mes_Solicitacao"],
         ['Count({<%s,%s,[Calendar.Ano_Solicitacao]={">=%d"}>} distinct codigo_solicitacao)' % (R, FILA, ini)]],
        ["in_ag", [PA, "Calendar.Ano_Solicitacao", "Calendar.Mes_Solicitacao"],
         ['Count({<%s,[Calendar.Ano_Solicitacao]={">=%d"}>} distinct codigo_solicitacao)' % (R, ini)]],
        ["out", [PA, "Ano_Agendamento", "Mes_Agendamento"],
         ['Count({<%s,[Ano_Agendamento]={">=%d"}>} distinct codigo_solicitacao)' % (R, ini)]],
        ["ult_ag", [PA], ["Date(Max({<%s>} data_agendamento))" % R]],
        ["cid", [P, "CID_solicitacao"], ["Count({<%s,%s>} distinct codigo_solicitacao)" % (R, FILA)]],
        ["dev_un", [P, U], ["Count({<%s,%s>} distinct codigo_solicitacao)" % (R, DEV)]],
        ["un_sol", [P, U], ["Count({<%s>} distinct codigo_solicitacao)" % R]],
        ["un_ag", [PA, U], ["Count({<%s>} distinct codigo_solicitacao)" % R]],
        ["tipo", [P, "tipo_procedimento_solicitacao"], ["Count({<%s>} distinct codigo_solicitacao)" % R]],
        ["ultima", [], ["Date(Max(data_solicitacao))"]],
        ["det_sol", [P, "codigo_solicitacao", "data_solicitacao", "status_solicitacao_solicitacao",
                     "classificacao_risco", "CID_solicitacao", U, "data_devolucao"],
         ["Count({<%s,%s>} codigo_solicitacao)" % (R, FILA_DEV)]],
        ["det_ag", [PA, "codigo_solicitacao", "Calendar.Data_Solicitacao", "data_agendamento",
                    "status_solicitacao_agendamento", "classificacao_risco", U],
         ['Count({<%s,[Ano_Agendamento]={">=%d"}>} codigo_solicitacao)' % (R, ano_det)]],
    ], layout=True)


def _br(s):
    return datetime.strptime(s, "%d/%m/%Y").date()


def ident(p):
    return hashlib.sha1(p.encode("utf-8")).hexdigest()[:12]


def _valido(p):
    return p not in ("-", "", None)


def referencias():
    try:
        return json.loads((PAINEL / "referencias.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def nivel(u, niveis):
    for padrao, n, _ in niveis:
        if fnmatchcase(u, padrao):
            return n
    return "OUTROS"


def _curto(st):
    """'SOLICITAÇÃO / PENDENTE / REGULADOR' -> 'PENDENTE / REGULADOR'"""
    return st.split(" / ", 1)[1] if " / " in st else st


def montar(b, hoje, ref=None):
    ref = ref or {}
    metas = ref.get("metas", {})
    fim = (hoje.year, hoje.month - 1) if hoje.month > 1 else (hoje.year - 1, 12)
    meses, a, mm = [], hoje.year - ANOS_SERIE, 1
    while (a, mm) <= fim:
        meses.append((a, mm))
        a, mm = (a, mm + 1) if mm < 12 else (a + 1, 1)
    idx = {k: i for i, k in enumerate(meses)}
    n = len(meses)

    procs = defaultdict(lambda: {"in": [0] * n, "out": [0] * n, "cores": {}, "fila": 0, "devolvidas": 0,
                                 "futuros": 0, "tipo": "", "cid": {}, "dev_un": {}, "unidades": set(),
                                 "ult_ag": None})
    for p, st, risco, q, antiga in b["fila"]["rows"]:
        if not _valido(p):
            continue
        d, q = procs[p], int(q)
        if st == DEVOL:
            d["devolvidas"] += q
        elif st in (PEND, REENV):
            d["fila"] += q
            c = d["cores"].setdefault(CORES.get(risco, "sem risco"), {"n": 0, "antiga": None})
            c["n"] += q
            if _valido(antiga):
                dt = _br(antiga)
                c["antiga"] = min(c["antiga"], dt) if c["antiga"] else dt
    for chave, campo in (("in_sol", "in"), ("in_ag", "in"), ("out", "out")):
        for p, ano, mes, q in b[chave]["rows"]:
            if not _valido(p) or mes not in MESES:
                continue
            k = (int(ano), MESES.index(mes) + 1)
            if k in idx:
                procs[p][campo][idx[k]] += int(q)
            elif campo == "out" and k > fim:
                procs[p]["futuros"] += int(q)
    for p, ult in b["ult_ag"]["rows"]:
        if _valido(p) and _valido(ult):
            procs[p]["ult_ag"] = _br(ult)
    for p, cid, q in b["cid"]["rows"]:
        if _valido(p):
            procs[p]["cid"][cid if _valido(cid) else "SEM CID"] = int(q)
    for p, u, q in b["dev_un"]["rows"]:
        if _valido(p) and _valido(u):
            procs[p]["dev_un"][u] = int(q)
    for chave in ("un_sol", "un_ag"):
        for p, u, q in b[chave]["rows"]:
            if _valido(p) and _valido(u) and int(q):
                procs[p]["unidades"].add(u)
    for p, t, _ in b["tipo"]["rows"]:
        if p in procs:
            procs[p]["tipo"] = t

    rot = ["%s/%s" % (MESES[mm - 1], str(a)[2:]) for a, mm in meses]
    resumo, detalhes = [], {}
    for p, d in procs.items():
        if not (d["fila"] or d["devolvidas"] or sum(d["in"]) or sum(d["out"])):
            continue
        cores = {}
        for cor, c in d["cores"].items():
            cores[cor] = {"n": c["n"], "antiga": c["antiga"].isoformat() if c["antiga"] else None,
                          "dias": (hoje - c["antiga"]).days if c["antiga"] else None}
        ent, sai = d["in"][-N_PAINEL:], d["out"][-N_PAINEL:]
        u12 = [i - o for i, o in zip(ent, sai)][-12:]
        pos = sum(1 for s in u12 if s > 0)
        alertas = []
        if pos >= 9:
            alertas.append("Fila crescendo: saldo positivo em %d dos últimos 12 meses" % pos)
        meta = metas.get(p) or META_PADRAO
        for cor in ("vermelho", "amarelo", "verde", "azul"):
            c, lim = cores.get(cor), meta.get(cor)
            if c and lim and c["dias"] is not None and c["dias"] > lim:
                alertas.append("%s esperando há %s dias (meta: %d)" % (
                    cor.capitalize(), f"{c['dias']:,}".replace(",", "."), lim))
        if d["fila"] >= 20 and sum(sai[-3:]) == 0 and d["futuros"] == 0:
            alertas.append("Fila parada: ninguém agendado nos últimos 3 meses nem marcado à frente")
        ult = d["ult_ag"]
        if d["fila"] > 0:
            if ult is None:
                alertas.append("Sem saída há mais de 1 ano: nenhum agendamento registrado no Mapa Social")
            elif (hoje - ult).days > SEM_SAIDA_DIAS:
                alertas.append("Sem saída há mais de 1 ano: último agendamento em %s" % ult.strftime("%d/%m/%Y"))
        esperas = [c["dias"] for c in cores.values() if c["dias"] is not None]
        i = ident(p)
        resumo.append({"id": i, "p": p, "tipo": d["tipo"], "fila": d["fila"], "cores": cores,
                       "devolvidas": d["devolvidas"], "in": ent, "out": sai, "saldo12": sum(u12),
                       "futuros": d["futuros"], "ult_ag": ult.isoformat() if ult else None,
                       "espera_max": max(esperas) if esperas else None, "alertas": alertas})
        detalhes[i] = {"p": p, "tipo": d["tipo"], "meses": rot, "in": d["in"], "out": d["out"], "cores": cores,
                       "cid": d["cid"], "devolvidas_unidade": d["dev_un"], "unidades": sorted(d["unidades"])}
    resumo.sort(key=lambda x: -x["fila"])
    meta = {"gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M"), "data": hoje.isoformat(),
            "solicitacoes_ate": b["ultima"]["rows"][0][0], "recarga_painel": b.get("_reload", "")}

    # Devolvidas por fila × unidade solicitante (abas de devolvidas)
    niveis, esf = ref.get("niveis", []), ref.get("esf", {})
    nomes = sorted({u for d in procs.values() for u in d["dev_un"]})
    pos_u = {u: k for k, u in enumerate(nomes)}
    devolvidas = dict(meta, unidades=[[u, nivel(u, niveis), esf.get(u)] for u in nomes],
                      filas={ident(p): sorted(([pos_u[u], q] for u, q in d["dev_un"].items()), key=lambda x: -x[1])
                             for p, d in procs.items() if d["dev_un"]})
    return dict(meta, meses=rot[-N_PAINEL:], filas=resumo), detalhes, devolvidas, meta


def montar_detalhe(b, hoje):
    """Linhas de cada fila: solicitações (pendentes, reenviadas, devolvidas) e agendamentos dos últimos 24 meses."""
    corte = _mes_inicial_painel(hoje)
    det = defaultdict(lambda: {"s": [], "a": []})
    vistos = set()
    for p, cod, dt, st, risco, cid, u, ddev, _ in b["det_sol"]["rows"]:
        if not _valido(p) or (p, cod) in vistos:   # uma linha por solicitação, como no programa local
            continue
        vistos.add((p, cod))
        det[ident(p)]["p"] = p
        det[ident(p)]["s"].append([cod, dt, _curto(st), CORES.get(risco, ""), cid if _valido(cid) else "",
                                   u if _valido(u) else "", ddev if _valido(ddev) else ""])
    vistos = set()
    for p, cod, dsol, dag, st, risco, u, _ in b["det_ag"]["rows"]:
        if not _valido(p) or not _valido(dag) or _br(dag) < corte:
            continue
        lin = [cod, dsol if _valido(dsol) else "", dag, _curto(st), CORES.get(risco, ""), u if _valido(u) else ""]
        chave = (p, *lin)
        if chave in vistos:
            continue
        vistos.add(chave)
        det[ident(p)]["p"] = p
        det[ident(p)]["a"].append(lin)
    nomes = {i: d["p"] for i, d in det.items()}
    lista_p = sorted({nomes[i] for i, d in det.items() if any(s[2].startswith("DEVOLVIDA") for s in d["s"])})
    pos_p = {p: k for k, p in enumerate(lista_p)}
    devolvidas = [[pos_p[nomes[i]]] + s[:2] + s[3:] for i, d in det.items() for s in d["s"] if s[2].startswith("DEVOLVIDA")]
    return det, {"p": lista_p, "r": devolvidas, "agendamentos_desde": corte.isoformat()}


def _limpar(pasta):
    pasta.mkdir(exist_ok=True)
    for velho in pasta.glob("*.json"):   # apaga arquivo a arquivo (rmtree falha com pasta sincronizada pelo OneDrive)
        velho.unlink()


def _grava(caminho, obj):
    caminho.write_text(json.dumps(obj, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")


def _saida_actions(**kv):
    if os.environ.get("GITHUB_OUTPUT"):
        with open(os.environ["GITHUB_OUTPUT"], "a", encoding="utf-8") as f:
            for k, v in kv.items():
                f.write(f"{k}={v}\n")


# Medições feitas antes de o histórico ir para o GitHub (programa local, 09/10; coleta de 13/10)
ATRASO_INICIAL = [
    {"medido_em": "09/10/2026 19:54", "recarga_painel": "2026-10-09T22:32:41.057Z", "solicitacoes_ate": "02/10/2026", "atraso_dias": 7},
    {"medido_em": "13/10/2026 21:10", "recarga_painel": "2026-10-13T22:33:05.260Z", "solicitacoes_ate": "07/10/2026", "atraso_dias": 6},
]
DIAS = ["seg", "ter", "qua", "qui", "sex", "sáb", "dom"]


def historico_atraso(meta, hoje):
    """Acrescenta a medição desta recarga ao histórico publicado (o gh-pages é recriado a cada publicação)."""
    hist = None
    for _ in range(3):
        try:
            with urllib.request.urlopen(PUBLICADO + "atraso.json", timeout=60) as r:
                hist = json.loads(r.read().decode("utf-8"))["medicoes"]
            break
        except urllib.error.HTTPError as e:
            if e.code == 404:
                hist = list(ATRASO_INICIAL)
                break
            print("atraso.json publicado ilegível (%s); tentando de novo." % e)
        except Exception as e:
            print("atraso.json publicado ilegível (%s); tentando de novo." % e)
        time.sleep(20)
    if hist is None:
        print("AVISO: histórico do atraso não pôde ser lido; recomeçando das medições iniciais.")
        hist = list(ATRASO_INICIAL)
    ultima = _br(meta["solicitacoes_ate"])
    if not any(h["recarga_painel"] == meta["recarga_painel"] for h in hist):
        hist.append({"medido_em": meta["gerado_em"], "recarga_painel": meta["recarga_painel"],
                     "solicitacoes_ate": meta["solicitacoes_ate"], "atraso_dias": (hoje - ultima).days})
    for h in hist:
        h["dia_semana"] = DIAS[_br(h["medido_em"][:10]).weekday()]
    print("Atraso do Mapa Social: " + "; ".join("%s (%s) %d d" % (h["medido_em"][:5], h["dia_semana"], h["atraso_dias"])
                                                 for h in hist[-7:]))
    return {"medicoes": hist}


def precisa_atualizar(hoje):
    """Compara a última recarga do Mapa Social com a do painel publicado."""
    recarga = m.consultar([], layout=True).get("_reload", "")
    try:
        with urllib.request.urlopen(PUBLICADO + "dados.json", timeout=60) as r:
            pub = json.loads(r.read().decode("utf-8"))
    except Exception as e:  # sem painel publicado legível: atualiza
        print("Painel publicado ilegível (%s); atualizando." % e)
        return True
    if pub.get("data") != hoje.isoformat():
        print("Dados publicados são de %s; atualizando." % pub.get("data"))
        return True
    if pub.get("recarga_painel") != recarga:
        print("Mapa Social recarregou (%s → %s); atualizando." % (pub.get("recarga_painel"), recarga))
        return True
    print("Sem novidade: Mapa Social não recarregou desde a última publicação (%s)." % recarga)
    return False


def main():
    hoje = date.today()
    t = time.monotonic()
    if "--se-mudou" in sys.argv and not precisa_atualizar(hoje):
        _saida_actions(publicar="nao")
        return
    b = coletar(hoje)
    print("[tempo] coleta: %.0f s" % (time.monotonic() - t), flush=True)
    t = time.monotonic()
    dados, detalhes, devolvidas, meta = montar(b, hoje, referencias())
    if len(dados["filas"]) < MIN_FILAS or not dados["solicitacoes_ate"] or dados["solicitacoes_ate"] == "-":
        print(f"ERRO: coleta incompleta ({len(dados['filas'])} filas, solicitações até {dados['solicitacoes_ate']!r}); nada foi gravado.")
        sys.exit(1)
    det, dev_det = montar_detalhe(b, hoje)
    print("[tempo] montagem: %.0f s" % (time.monotonic() - t), flush=True)
    t = time.monotonic()
    _grava(PAINEL / "atraso.json", historico_atraso(meta, hoje))
    _grava(PAINEL / "dados.json", dados)
    _grava(PAINEL / "devolvidas.json", devolvidas)
    _limpar(PAINEL / "analises")
    for i, d in detalhes.items():
        _grava(PAINEL / "analises" / f"{i}.json", dict(meta, **d))
    _limpar(PAINEL / "detalhe")
    for i, d in det.items():
        _grava(PAINEL / "detalhe" / f"{i}.json", dict(meta, agendamentos_desde=dev_det["agendamentos_desde"], **d))
    _grava(PAINEL / "detalhe" / "devolvidas.json", dict(meta, **dev_det))
    print("[tempo] gravação: %.0f s" % (time.monotonic() - t), flush=True)
    _saida_actions(publicar="sim")
    print(f"{dados['gerado_em']}: {len(dados['filas'])} filas, solicitações até {dados['solicitacoes_ate']}, "
          f"{sum(len(d['s']) for d in det.values())} solicitações e {sum(len(d['a']) for d in det.values())} "
          f"agendamentos no detalhe")


if __name__ == "__main__":
    main()

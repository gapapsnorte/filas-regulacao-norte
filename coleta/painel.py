"""Atualiza o painel das filas (todas as filas da Região Norte).

Roda no GitHub Actions (.github/workflows/atualizar.yml), que grava os dados e publica.
Uso local: python coleta/painel.py             (coleta e grava, sem publicar)
           python coleta/painel.py --se-mudou  (só coleta se o Mapa Social recarregou ou se o dia virou)

Gera:
  dados.json                resumo de todas as filas (painel)
  devolvidas.json           devolvidas por fila × unidade solicitante, série mensal pela data da devolução e, em d12,
                            [unidade, devolvidas, solicitações] dos últimos 12 meses por fila (razão de devolvidas)
  espera.json               registro diário da fila de espera (total e por fila), versionado em main como atraso.json
  detalhe/espera.json       solicitações em fila de espera, linha a linha (aba Sem oferta)
  analises/<id>.json        detalhes agregados de cada fila, usados pelo botão "Gerar análise"
  detalhe/<id>.json         linhas do Mapa Social de cada fila (código SISREG, datas, status, risco, CID, unidade;
                            inclui as solicitações em fila de espera, com esse status)
  detalhe/devolvidas.json   todas as devolvidas da região, linha a linha
  atraso.json               histórico do atraso do Mapa Social (uma medição por dia); fica no branch main, versionado
  devolvidas_reg.json       registro diário da resolução das devolvidas (códigos que entraram e que saíram do status
                            devolvida, comparando com a publicação anterior); versionado em main como atraso.json
"""
import hashlib
import json
import os
import statistics
import sys
import time
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
# Fila de espera = solicitações aguardando pactuação de oferta (não estão com o regulador). Fica fora da "fila"
# do painel, mas tem aba própria ("Sem oferta") e registro diário em espera.json.
ESPERA = "SOLICITAÇÃO / PENDENTE / FILA DE ESPERA"
ESP = '[status_solicitacao_solicitacao]={"%s"}' % ESPERA
FILA_DEV_ESP = '[status_solicitacao_solicitacao]={"%s","%s","%s","%s"}' % (PEND, REENV, DEVOL, ESPERA)
MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
CORES = {"Prioridade 1": "vermelho", "Prioridade 2": "amarelo", "Prioridade 3": "verde", "Prioridade 4": "azul"}
N_PAINEL = 24          # meses mostrados no painel (e janela dos agendamentos no detalhe)
MIN_FILAS = 500        # abaixo disso a coleta é considerada falha (não publica)
ANOS_SERIE = 5         # série da análise: de janeiro de (ano - 5) até o mês anterior
SEM_SAIDA_DIAS = 365   # alerta "sem saída há mais de 1 ano"


def _mes_inicial_painel(hoje):
    a, mm = hoje.year, hoje.month - N_PAINEL
    while mm < 1:
        a, mm = a - 1, mm + 12
    return date(a, mm, 1)


def janela12(hoje):
    """Os 12 meses completos anteriores ao mês corrente, como (ano, mês)."""
    a, mm = hoje.year, hoje.month - 12
    if mm < 1:
        a, mm = a - 1, mm + 12
    out = []
    for _ in range(12):
        out.append((a, mm))
        a, mm = (a, mm + 1) if mm < 12 else (a + 1, 1)
    return out


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
        ["esp_un", [P, U], ["Count({<%s,%s>} distinct codigo_solicitacao)" % (R, ESP)]],
        # Razão de devolvidas (aba Devolvidas): solicitações feitas por unidade × mês da solicitação, nos
        # últimos 12 meses completos — total (pendentes + reenviadas + devolvidas + agendadas) e só devolvidas
        ["sol12", [P, U, "Calendar.Ano_Solicitacao", "Calendar.Mes_Solicitacao"],
         ['Count({<%s,%s,[Calendar.Ano_Solicitacao]={">=%d"}>} distinct codigo_solicitacao)' % (R, FILA_DEV, hoje.year - 1)]],
        ["dev12", [P, U, "Calendar.Ano_Solicitacao", "Calendar.Mes_Solicitacao"],
         ['Count({<%s,%s,[Calendar.Ano_Solicitacao]={">=%d"}>} distinct codigo_solicitacao)' % (R, DEV, hoje.year - 1)]],
        ["ag12", [PA, U, "Calendar.Ano_Solicitacao", "Calendar.Mes_Solicitacao"],
         ['Count({<%s,[Calendar.Ano_Solicitacao]={">=%d"}>} distinct codigo_solicitacao)' % (R, hoje.year - 1)]],
        ["tipo", [P, "tipo_procedimento_solicitacao"], ["Count({<%s>} distinct codigo_solicitacao)" % R]],
        ["ultima", [], ["Date(Max(data_solicitacao))"]],
        ["det_sol", [P, "codigo_solicitacao", "data_solicitacao", "status_solicitacao_solicitacao",
                     "classificacao_risco", "CID_solicitacao", U, "data_devolucao"],
         ["Count({<%s,%s>} codigo_solicitacao)" % (R, FILA_DEV_ESP)]],
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
                                 "ult_ag": None, "espera": {"n": 0, "antiga": None, "cores": {}, "unidades": {}}})
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
        elif st == ESPERA:
            e = d["espera"]
            e["n"] += q
            c = e["cores"].setdefault(CORES.get(risco, "sem risco"), {"n": 0, "antiga": None})
            c["n"] += q
            if _valido(antiga):
                dt = _br(antiga)
                c["antiga"] = min(c["antiga"], dt) if c["antiga"] else dt
                e["antiga"] = min(e["antiga"], dt) if e["antiga"] else dt
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
    for p, u, q in b["esp_un"]["rows"]:
        if _valido(p) and _valido(u) and int(q):
            procs[p]["espera"]["unidades"][u] = int(q)
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
        if not (d["fila"] or d["devolvidas"] or sum(d["in"]) or sum(d["out"]) or d["espera"]["n"]):
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
        meta = metas.get(p) or {}   # só há meta quando uma norma a define (aba Metas_espera); sem norma, sem alerta de meta
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
        e = d["espera"]
        espera = None
        if e["n"]:
            espera = {"n": e["n"], "antiga": e["antiga"].isoformat() if e["antiga"] else None,
                      "dias": (hoje - e["antiga"]).days if e["antiga"] else None,
                      "cores": {cor: {"n": c["n"], "antiga": c["antiga"].isoformat() if c["antiga"] else None,
                                      "dias": (hoje - c["antiga"]).days if c["antiga"] else None}
                                for cor, c in e["cores"].items()},
                      "unidades": sorted(e["unidades"].items(), key=lambda x: -x[1])}
        i = ident(p)
        resumo.append({"id": i, "p": p, "meta_normativa": bool(metas.get(p)), "tipo": d["tipo"], "fila": d["fila"], "cores": cores,
                       "devolvidas": d["devolvidas"], "in": ent, "out": sai, "saldo12": sum(u12),
                       "futuros": d["futuros"], "ult_ag": ult.isoformat() if ult else None,
                       "espera_max": max(esperas) if esperas else None, "alertas": alertas, "espera": espera})
        detalhes[i] = {"p": p, "tipo": d["tipo"], "meses": rot, "in": d["in"], "out": d["out"], "cores": cores,
                       "cid": d["cid"], "devolvidas_unidade": d["dev_un"], "unidades": sorted(d["unidades"])}
    resumo.sort(key=lambda x: -x["fila"])
    ultima = b["ultima"]["rows"][0][0]
    meta = {"gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M"), "data": hoje.isoformat(),
            "solicitacoes_ate": ultima, "recarga_painel": b.get("_reload", ""),
            "atraso_dias": (hoje - _br(ultima)).days if _valido(ultima) else None}

    # Devolvidas por fila × unidade solicitante (abas de devolvidas)
    niveis, esf = ref.get("niveis", []), ref.get("esf", {})
    # Razão de devolvidas: das solicitações feitas pela unidade nos últimos 12 meses completos (pela data da
    # solicitação), quantas estão devolvidas hoje. Cada código conta uma vez; a que foi agendada depois de uma
    # devolução conta como não devolvida (o Mapa Social só guarda a situação atual).
    j12 = set(janela12(hoje))
    tot12, dev12 = defaultdict(int), defaultdict(int)
    for chave, alvo in (("sol12", tot12), ("ag12", tot12), ("dev12", dev12)):
        for p, u, ano, mes, q in b[chave]["rows"]:
            if _valido(p) and _valido(u) and mes in MESES and (int(ano), MESES.index(mes) + 1) in j12:
                alvo[(p, u)] += int(q)
    nomes = sorted({u for d in procs.values() for u in d["dev_un"]} | {u for (p, u) in tot12 if p in procs})
    pos_u = {u: k for k, u in enumerate(nomes)}
    ids_serie, celulas = serie_devolvidas(b, meses[-N_PAINEL:], pos_u, nomes)
    d12 = defaultdict(list)
    for (p, u), t in sorted(tot12.items()):
        if p in procs and t:
            d12[ident(p)].append([pos_u[u], min(dev12.get((p, u), 0), t), t])
    r12 = ["%s/%s" % (MESES[mm - 1], str(a)[2:]) for a, mm in janela12(hoje)]
    devolvidas = dict(meta, unidades=[[u, nivel(u, niveis), esf.get(u)] for u in nomes],
                      filas={ident(p): sorted(([pos_u[u], q] for u, q in d["dev_un"].items()), key=lambda x: -x[1])
                             for p, d in procs.items() if d["dev_un"]},
                      meses=rot[-N_PAINEL:], serie_ids=ids_serie, serie=celulas,
                      janela12=[r12[0], r12[-1]], d12=dict(d12))
    return dict(meta, meses=rot[-N_PAINEL:], filas=resumo), detalhes, devolvidas, meta


def serie_devolvidas(b, janela, pos_u, nomes):
    """Devolvidas de hoje contadas pela data da devolução, nos meses da janela do painel, por fila × unidade.
    Saída esparsa: [índice da fila em ids, índice da unidade, índice do mês, quantidade]. Uma solicitação conta
    uma vez, na sua devolução mais recente (o Mapa Social guarda, para algumas, mais de uma data)."""
    idx = {k: i for i, k in enumerate(janela)}
    ult = {}
    for p, cod, _dt, st, _risco, _cid, u, ddev, _ in b["det_sol"]["rows"]:
        if not _valido(p) or st != DEVOL or not _valido(ddev) or not _valido(u):
            continue
        d = _br(ddev)
        if (p, cod) not in ult or d > ult[(p, cod)][0]:
            ult[(p, cod)] = (d, u)
    cel, ids = defaultdict(int), {}
    for (p, _cod), (d, u) in ult.items():
        k = (d.year, d.month)
        if k not in idx:
            continue
        if u not in pos_u:
            pos_u[u] = len(nomes)
            nomes.append(u)
        fi = ids.setdefault(ident(p), len(ids))
        cel[(fi, pos_u[u], idx[k])] += 1
    return list(ids), sorted([list(k) + [n] for k, n in cel.items()])


ATEND_DIAS = 90   # janela dos atendimentos usados no "tempo até o atendimento"


def estatisticas_cores(b, hoje):
    """Por procedimento e cor: espera de quem está na fila, espera até o atendimento (agendados nos últimos
    90 dias) e entrada/saída mensais médias dos últimos 12 meses completos (base da simulação por cor)."""
    a_, m_ = hoje.year, hoje.month - 12
    if m_ < 1:
        a_, m_ = a_ - 1, m_ + 12
    ini = date(a_, m_, 1)                     # 12 meses completos: do mesmo mês do ano anterior até o mês passado
    lim_fim = date(hoje.year, hoje.month, 1)  # primeiro dia do mês atual (exclusivo)
    corte_at = hoje - timedelta(days=ATEND_DIAS)
    acc = defaultdict(lambda: defaultdict(lambda: {"fila": [], "atend": [], "ent": 0, "sai": 0}))
    todos = defaultdict(list)                 # dias de espera de toda a fila (pendentes + reenviadas) por procedimento
    vistos, entradas = set(), set()   # entradas: solicitações já contadas na entrada, para nenhuma contar duas vezes
    for p, cod, dt, st, risco, *_ in b["det_sol"]["rows"]:
        if not _valido(p) or (p, cod) in vistos or not _valido(dt):
            continue
        vistos.add((p, cod))
        cor = CORES.get(risco, "sem risco")
        d = _br(dt)
        if st in (PEND, REENV):
            acc[p][cor]["fila"].append((hoje - d).days)
            todos[p].append((hoje - d).days)
            if ini <= d < lim_fim:
                entradas.add((p, cod))
                acc[p][cor]["ent"] += 1
    vistos = set()
    for p, cod, dsol, dag, _st, risco, *_ in b["det_ag"]["rows"]:
        if not _valido(p) or not _valido(dag) or (p, cod, dag) in vistos:
            continue
        vistos.add((p, cod, dag))
        cor = CORES.get(risco, "sem risco")
        da = _br(dag)
        if ini <= da < lim_fim:
            acc[p][cor]["sai"] += 1           # saída = cada agendamento, como no OUTPUT da série
        if _valido(dsol):
            ds = _br(dsol)
            if ini <= ds < lim_fim and (p, cod) not in entradas:
                entradas.add((p, cod))        # entrada = cada solicitação uma vez, mesmo remarcada
                acc[p][cor]["ent"] += 1
            if corte_at <= da <= hoje and da >= ds:
                acc[p][cor]["atend"].append((da - ds).days)
    saida = {}
    for p, cores in acc.items():
        r = {}
        for cor, v in cores.items():
            f, a = v["fila"], v["atend"]
            r[cor] = {"fila_media": round(statistics.mean(f)) if f else None,
                      "fila_mediana": round(statistics.median(f)) if f else None,
                      "atend_media": round(statistics.mean(a)) if a else None,
                      "atend_mediana": round(statistics.median(a)) if a else None,
                      "atend_n": len(a), "ent": round(v["ent"] / 12, 1), "sai": round(v["sai"] / 12, 1)}
        t = todos.get(p, [])
        r["_fila"] = {"media": round(statistics.mean(t)) if t else None,
                      "mediana": round(statistics.median(t)) if t else None, "n": len(t)}
        saida[ident(p)] = r
    return saida


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
    # fila de espera, linha a linha: [procedimento, código, data da solicitação, cor, CID, unidade]
    esp_p = sorted({nomes[i] for i, d in det.items() if any("FILA DE ESPERA" in s[2] for s in d["s"])})
    pos_e = {p: k for k, p in enumerate(esp_p)}
    espera = [[pos_e[nomes[i]]] + s[:2] + s[3:6] for i, d in det.items() for s in d["s"] if "FILA DE ESPERA" in s[2]]
    return det, {"p": lista_p, "r": devolvidas, "agendamentos_desde": corte.isoformat(),
                 "espera": {"p": esp_p, "r": espera}}


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
    """Acrescenta a medição desta recarga a atraso.json (arquivo versionado em main; o workflow faz o commit).
    Uma medição por dia: recargas repetidas na mesma noite, sem solicitações novas, não entram."""
    arq = PAINEL / "atraso.json"
    try:
        hist = json.loads(arq.read_text(encoding="utf-8"))["medicoes"]
    except (OSError, ValueError, KeyError) as e:
        print("atraso.json ausente ou ilegível (%s); começando das medições iniciais." % e)
        hist = list(ATRASO_INICIAL)
    dia = meta["gerado_em"][:10]
    if not any(h["medido_em"][:10] == dia and h["solicitacoes_ate"] == meta["solicitacoes_ate"] for h in hist):
        hist.append({"medido_em": meta["gerado_em"], "recarga_painel": meta["recarga_painel"],
                     "solicitacoes_ate": meta["solicitacoes_ate"], "atraso_dias": meta["atraso_dias"]})
    for h in hist:
        h["dia_semana"] = DIAS[_br(h["medido_em"][:10]).weekday()]
    print("Atraso do Mapa Social: " + "; ".join("%s (%s) %d d" % (h["medido_em"][:5], h["dia_semana"], h["atraso_dias"])
                                                 for h in hist[-7:]))
    return {"medicoes": hist}


def historico_espera(dados, hoje):
    """Registro diário da fila de espera (espera.json, versionado em main; o workflow faz o commit): total e
    quantidade por fila. Uma entrada por dia (a última coleta do dia prevalece)."""
    arq = PAINEL / "espera.json"
    try:
        dias = json.loads(arq.read_text(encoding="utf-8"))["dias"]
    except (OSError, ValueError, KeyError) as e:
        print("espera.json ausente ou ilegível (%s); começando o registro hoje." % e)
        dias = []
    filas = {f["id"]: f["espera"]["n"] for f in dados["filas"] if f.get("espera")}
    hoje_s = hoje.isoformat()
    dias = [d for d in dias if d["data"] != hoje_s]
    dias.append({"data": hoje_s, "total": sum(filas.values()), "filas": filas})
    dias.sort(key=lambda d: d["data"])
    print("Fila de espera: %d solicitações em %d procedimentos (registro com %d dias)" % (dias[-1]["total"], len(filas), len(dias)))
    return {"dias": dias}


def _destino(st):
    if not st:
        return "ausente"
    if "REENVIADA" in st:
        return "reenviada"
    if "AGENDADA" in st:
        return "agendada"
    if "FILA DE ESPERA" in st:
        return "fila de espera"
    if "PENDENTE" in st:
        return "pendente"
    return "outro"


def _indice(lista, v):
    """Posição de v na lista; acrescenta ao fim se ainda não estiver (listas p/u do registro)."""
    if v not in lista:
        lista.append(v)
    return lista.index(v)


def historico_devolvidas(dev_det, det, hoje, anterior=None):
    """Registro diário da resolução das devolvidas (devolvidas_reg.json, versionado em main; o workflow faz o commit).
    Compara as devolvidas de hoje com as da publicação anterior (detalhe/devolvidas.json no site): códigos que
    ENTRARAM no status devolvida e códigos que SAÍRAM dele, com o destino (reenviada, agendada, pendente, fila de
    espera ou ausente do Mapa Social). O Mapa Social só mostra a situação atual; este registro é a linha de base
    para medir quantas devoluções cada unidade corrige e em quanto tempo. Procedimentos e unidades ficam em listas
    (p, u) e as linhas usam os índices: entradas [pi, ui, código, data da devolução], saídas [..., destino]."""
    arq = PAINEL / "devolvidas_reg.json"
    try:
        reg = json.loads(arq.read_text(encoding="utf-8"))
        assert isinstance(reg.get("dias"), list)
    except (OSError, ValueError, AssertionError) as e:
        print("devolvidas_reg.json ausente ou ilegível (%s); começando o registro hoje." % e)
        reg = {"inicio": hoje.isoformat(), "estoque_inicial": len(dev_det["r"]), "p": [], "u": [], "dias": []}
    if anterior is None:
        try:
            with urllib.request.urlopen(PUBLICADO + "detalhe/devolvidas.json", timeout=180) as r:
                anterior = json.loads(r.read().decode("utf-8"))
        except Exception as e:   # sem comparação hoje: o registro fica como está (nada é inventado)
            print("Registro de devolvidas: não li a publicação anterior (%s); dia sem comparação." % e)
            return reg
    if not anterior.get("r") or not anterior.get("p"):
        print("Registro de devolvidas: publicação anterior sem devolvidas; dia sem comparação.")
        return reg
    atual = {r[1]: (dev_det["p"][r[0]], r[5], r[6]) for r in dev_det["r"]}
    antes = {r[1]: (anterior["p"][r[0]], r[5], r[6]) for r in anterior["r"]}
    status = {}
    for d in det.values():
        for sl in d["s"]:
            status[sl[0]] = sl[2]
        for ag in d["a"]:
            status.setdefault(ag[0], "AGENDADA")
    entradas = [[_indice(reg["p"], p), _indice(reg["u"], u), cod, dd] for cod, (p, u, dd) in atual.items() if cod not in antes]
    saidas = [[_indice(reg["p"], p), _indice(reg["u"], u), cod, dd, _destino(status.get(cod))] for cod, (p, u, dd) in antes.items() if cod not in atual]
    hoje_s = hoje.isoformat()
    dia = next((d for d in reg["dias"] if d["data"] == hoje_s), None)
    if dia is None:
        dia = {"data": hoje_s, "anterior": anterior.get("data"), "estoque": len(atual), "entradas": [], "saidas": []}
        reg["dias"].append(dia)
    else:   # segunda coleta do mesmo dia: acrescenta o que mudou desde a anterior, sem repetir códigos
        dia["estoque"] = len(atual)
    for chave, novos in (("entradas", entradas), ("saidas", saidas)):
        vistos = {x[2] for x in dia[chave]}
        dia[chave].extend(x for x in novos if x[2] not in vistos)
    reg["dias"].sort(key=lambda d: d["data"])
    dest = defaultdict(int)
    for x in dia["saidas"]:
        dest[x[4]] += 1
    print("Registro de devolvidas: %d no estoque, %d entradas e %d saídas hoje (%s); %d dias registrados" % (
        len(atual), len(dia["entradas"]), len(dia["saidas"]), ", ".join("%s %d" % kv for kv in sorted(dest.items())) or "sem saídas", len(reg["dias"])))
    return reg


def precisa_atualizar(hoje):
    """Compara a última recarga do Mapa Social com a do painel publicado."""
    r = m.consultar([["ultima", [], ["Date(Max(data_solicitacao))"]]], layout=True)
    linhas = r.get("ultima", {}).get("rows") or [["-"]]
    recarga, ultima = r.get("_reload", ""), linhas[0][0]
    try:
        with urllib.request.urlopen(PUBLICADO + "dados.json", timeout=60) as resp:
            pub = json.loads(resp.read().decode("utf-8"))
    except Exception as e:  # sem painel publicado legível: atualiza
        print("Painel publicado ilegível (%s); atualizando." % e)
        return True
    if pub.get("data") != hoje.isoformat():
        print("Dados publicados são de %s; atualizando." % pub.get("data"))
        return True
    if pub.get("recarga_painel") == recarga:
        print("Sem novidade: Mapa Social não recarregou desde a última publicação (%s)." % recarga)
        return False
    # Premissa: toda carga com solicitações novas avança a data da última solicitação (~3,5 mil por dia útil)
    if pub.get("solicitacoes_ate") == ultima:
        print("Mapa Social recarregou (%s → %s), mas sem solicitações novas (última: %s); nada a fazer." % (
            pub.get("recarga_painel"), recarga, ultima))
        return False
    print("Mapa Social recarregou (%s → %s) com solicitações até %s; atualizando." % (pub.get("recarga_painel"), recarga, ultima))
    return True


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
    est = estatisticas_cores(b, hoje)
    for f in dados["filas"]:
        cs = dict(est.get(f["id"], {}))
        f["esp_fila"] = cs.pop("_fila", None)   # espera de quem está na fila hoje, fila inteira (média, mediana, n)
        f["cor_stats"] = cs
    dados["atend_dias"] = ATEND_DIAS
    if len(dados["filas"]) < MIN_FILAS or not dados["solicitacoes_ate"] or dados["solicitacoes_ate"] == "-":
        print(f"ERRO: coleta incompleta ({len(dados['filas'])} filas, solicitações até {dados['solicitacoes_ate']!r}); nada foi gravado.")
        sys.exit(1)
    det, dev_det = montar_detalhe(b, hoje)
    print("[tempo] montagem: %.0f s" % (time.monotonic() - t), flush=True)
    t = time.monotonic()
    hist = historico_atraso(meta, hoje)
    _grava(PAINEL / "atraso.json", hist)
    _grava(PAINEL / "espera.json", historico_espera(dados, hoje))
    _grava(PAINEL / "devolvidas_reg.json", historico_devolvidas(dev_det, det, hoje))
    ultimos = [h["atraso_dias"] for h in hist["medicoes"][-7:]]
    for d in (meta, dados, devolvidas):
        d["atraso_ultimos"] = ultimos
    _grava(PAINEL / "dados.json", dados)
    _grava(PAINEL / "devolvidas.json", devolvidas)
    _limpar(PAINEL / "analises")
    for i, d in detalhes.items():
        _grava(PAINEL / "analises" / f"{i}.json", dict(meta, **d))
    _limpar(PAINEL / "detalhe")
    for i, d in det.items():
        _grava(PAINEL / "detalhe" / f"{i}.json", dict(meta, agendamentos_desde=dev_det["agendamentos_desde"], **d))
    _grava(PAINEL / "detalhe" / "espera.json", dict(meta, **dev_det.pop("espera")))
    _grava(PAINEL / "detalhe" / "devolvidas.json", dict(meta, **dev_det))
    print("[tempo] gravação: %.0f s" % (time.monotonic() - t), flush=True)
    _saida_actions(publicar="sim")
    print(f"{dados['gerado_em']}: {len(dados['filas'])} filas, solicitações até {dados['solicitacoes_ate']}, "
          f"{sum(len(d['s']) for d in det.values())} solicitações e {sum(len(d['a']) for d in det.values())} "
          f"agendamentos no detalhe")


if __name__ == "__main__":
    main()

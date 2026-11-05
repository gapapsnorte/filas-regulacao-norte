"""Atualiza o painel das filas (todas as filas da Região Norte).

Roda no GitHub Actions (.github/workflows/atualizar.yml), que grava os dados e publica.
Uso local: python coleta/painel.py             (coleta e grava, sem publicar)
           python coleta/painel.py --se-mudou  (só coleta se o Mapa Social recarregou ou se o dia virou)

Gera:
  dados.json                resumo de todas as filas (painel)
  devolvidas.json           devolvidas por fila × unidade solicitante e, em d12,
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
  fila_reg.json             saídas da fila por mês e por fila, com o destino de cada código, comparando com a publicação
                            anterior (base do saldo inexplicado); versionado em main como atraso.json
  detalhe/fila.json         códigos que estão na fila hoje, para a comparação da próxima coleta
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
# Além do nº de filas, a coleta precisa trazer a fila, as devolvidas e o detalhe (um cubo vazio deixaria tudo zerado e,
# no registro diário, marcaria ~47 mil devolvidas como "canceladas"). Valores bem abaixo do observado em out/2026
# (579 filas com fila, 94 mil na fila, 769 filas com devolvidas, 47 mil devolvidas).
MIN_COM_FILA, MIN_NA_FILA, MIN_COM_DEV, MIN_DEVOLVIDAS = 300, 20000, 300, 10000
MAX_QUEDA_TOTAL, MIN_FILA_GRANDE, CODIGO_TRAVA = 0.15, 500, 3   # queda em relação à publicação anterior: ver trava_relativa
# Unidades que passaram a entrar quando o filtro virou código CNES: as devolvidas que elas já tinham não são "novas".
UNIDADES_INCORPORADAS = ("UBS 7 - Sobradinho II", "UPA Planaltina")
MIN_ESTOQUE_REL = 0.98  # queda de mais de 2% no estoque de devolvidas (o normal é ~0,04%/dia) = coleta suspeita, o registro não é alterado
TOLERANCIA_MES = 10     # depois deste dia, o mês anterior conta como completo mesmo que a fonte não tenha chegado ao último dia
ANOS_SERIE = 5         # série da análise: de janeiro de (ano - 5) até o mês anterior
SEM_SAIDA_DIAS = 365   # alerta "sem saída há mais de 1 ano"


def _mes_inicial_painel(hoje):
    a, mm = hoje.year, hoje.month - N_PAINEL
    while mm < 1:
        a, mm = a - 1, mm + 12
    return date(a, mm, 1)


def mes_referencia(hoje, ultima):
    """1º dia do mês seguinte ao último mês completo. O mês anterior só é completo quando a última solicitação
    registrada chega ao último dia dele; com a fonte atrasada na virada do mês, ele fica de fora das séries e médias."""
    ini = hoje.replace(day=1)
    fim_ant = ini - timedelta(days=1)
    if _valido(ultima) and _br(ultima) < fim_ant and hoje.day <= TOLERANCIA_MES:
        return fim_ant.replace(day=1)
    return ini


def janela12(hoje):
    """Os 12 meses completos anteriores ao mês de hoje (aqui, o mês de referência), como (ano, mês)."""
    a, mm = hoje.year, hoje.month - 12
    if mm < 1:
        a, mm = a - 1, mm + 12
    out = []
    for _ in range(12):
        out.append((a, mm))
        a, mm = (a, mm + 1) if mm < 12 else (a + 1, 1)
    return out


def coletar(hoje):
    ano12 = (hoje - timedelta(days=45)).year - 1   # primeiro ano necessário para a janela de 12 meses, mesmo com janeiro incompleto
    ini = hoje.year - ANOS_SERIE
    P, PA, U = "descrição_procedimento_solicitacao", "descrição_procedimento_agendamento", m.COD_U
    nomes = m.unidades_norte()     # a região sai do código CNES da unidade solicitante; o nome é ligado depois do filtro
    R = m.filtro_unidades(nomes)
    # a janela de 24 meses do painel pode recuar um mês quando o mês anterior ainda está incompleto (ver mes_referencia);
    # de 1 a 10 de janeiro isso leva o primeiro mês ao ano anterior, e o cubo det_ag precisa cobri-lo
    ano_det = (_mes_inicial_painel(hoje) - timedelta(days=1)).year if hoje.day <= TOLERANCIA_MES else _mes_inicial_painel(hoje).year
    b = m.consultar([
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
         ['Count({<%s,%s,[Calendar.Ano_Solicitacao]={">=%d"}>} distinct codigo_solicitacao)' % (R, FILA_DEV, ano12)]],
        ["dev12", [P, U, "Calendar.Ano_Solicitacao", "Calendar.Mes_Solicitacao"],
         ['Count({<%s,%s,[Calendar.Ano_Solicitacao]={">=%d"}>} distinct codigo_solicitacao)' % (R, DEV, ano12)]],
        ["ag12", [PA, U, "Calendar.Ano_Solicitacao", "Calendar.Mes_Solicitacao"],
         ['Count({<%s,[Calendar.Ano_Solicitacao]={">=%d"}>} distinct codigo_solicitacao)' % (R, ano12)]],
        ["tipo", [P, "tipo_procedimento_solicitacao"], ["Count({<%s>} distinct codigo_solicitacao)" % R]],
        # Quem regula a fila de hoje e onde os agendamentos foram marcados (filtros da aba Filas)
        ["cen", [P, "nome_central_reguladora_solictacao"], ["Count({<%s,%s>} distinct codigo_solicitacao)" % (R, FILA)]],
        # Sem o nome SISREG do executante: uma unidade tem vários nomes lá e as contagens por nome se repetem
        ["exe", [PA, "nome_unidade_executante", "regiao_saude_executante", "Ano_Agendamento", "Mes_Agendamento"],
         ['Count({<%s,[Ano_Agendamento]={">=%d"}>} distinct codigo_solicitacao)' % (R, ano12)]],
        ["ultima", [], ["Date(Max(data_solicitacao))"]],
        ["det_sol", [P, "codigo_solicitacao", "data_solicitacao", "status_solicitacao_solicitacao",
                     "classificacao_risco", "CID_solicitacao", U, "data_devolucao"],
         ["Count({<%s,%s>} codigo_solicitacao)" % (R, FILA_DEV_ESP)]],
        ["det_ag", [PA, "codigo_solicitacao", "Calendar.Data_Solicitacao", "data_agendamento",
                    "status_solicitacao_agendamento", "classificacao_risco", U],
         ['Count({<%s,[Ano_Agendamento]={">=%d"}>} codigo_solicitacao)' % (R, ano_det)]],
    ], layout=True)
    for bloco in b.values():
        if isinstance(bloco, dict) and U in bloco.get("cols", []):
            m.nomear(bloco, nomes)
    return b


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


def central_executante(b, resumo, j12):
    """Grava em cada fila quem regula quem está na fila hoje (cen = [central, n]) e onde os agendamentos dos últimos
    12 meses completos foram marcados (exe = [unidade executante, grupo do Mapa Social, n])."""
    por = {f["p"]: f for f in resumo}
    cen, exe = defaultdict(lambda: defaultdict(int)), defaultdict(lambda: defaultdict(int))
    for p, c, q in b["cen"]["rows"]:
        if p in por and _valido(c):
            cen[p][c] += int(q)
    for p, nome, reg, ano, mes, q in b["exe"]["rows"]:
        if p not in por or not _valido(nome) or mes not in MESES or not ano.isdigit() or (int(ano), MESES.index(mes) + 1) not in j12:
            continue
        exe[p][(nome, reg if _valido(reg) else "Sem grupo")] += int(q)
    for p, f in por.items():
        f["cen"] = sorted(([c, q] for c, q in cen[p].items()), key=lambda x: -x[1])
        f["exe"] = sorted(([n, r, q] for (n, r), q in exe[p].items()), key=lambda x: -x[2])


def montar(b, hoje, ref=None):
    ref = ref or {}
    metas = ref.get("metas", {})
    ultima = b["ultima"]["rows"][0][0]
    ref_mes = mes_referencia(hoje, ultima)
    fim = (ref_mes.year, ref_mes.month - 1) if ref_mes.month > 1 else (ref_mes.year - 1, 12)
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
            elif campo == "out" and k >= (hoje.year, hoje.month):
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
        if pos >= 9 and sum(u12) > 0:   # meses positivos E saldo do ano positivo: um mês muito negativo pode anular os demais
            alertas.append("Fila crescendo: saldo positivo em %d dos últimos 12 meses" % pos)
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
    meta = {"gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M"), "data": hoje.isoformat(),
            "solicitacoes_ate": ultima, "recarga_painel": b.get("_reload", ""),
            "atraso_dias": (hoje - _br(ultima)).days if _valido(ultima) else None}
    if ref_mes != hoje.replace(day=1):   # mês que acabou mas ainda não chegou inteiro da fonte
        meta["mes_incompleto"] = "%s/%s" % (MESES[ref_mes.month - 1], str(ref_mes.year)[2:])

    # Devolvidas por fila × unidade solicitante (abas de devolvidas)
    niveis, esf = ref.get("niveis", []), ref.get("esf", {})
    # Razão de devolvidas: das solicitações feitas pela unidade nos últimos 12 meses completos (pela data da
    # solicitação), quantas estão devolvidas hoje. Cada código conta uma vez; a que foi agendada depois de uma
    # devolução conta como não devolvida (o Mapa Social só guarda a situação atual).
    j12 = set(janela12(ref_mes))
    tot12, dev12 = defaultdict(int), defaultdict(int)
    for chave, alvo in (("sol12", tot12), ("ag12", tot12), ("dev12", dev12)):
        for p, u, ano, mes, q in b[chave]["rows"]:
            if _valido(p) and _valido(u) and mes in MESES and (int(ano), MESES.index(mes) + 1) in j12:
                alvo[(p, u)] += int(q)
    nomes = sorted({u for d in procs.values() for u in d["dev_un"]} | {u for (p, u) in tot12 if p in procs})
    pos_u = {u: k for k, u in enumerate(nomes)}
    d12 = defaultdict(list)
    for (p, u), t in sorted(tot12.items()):
        if p in procs and t:
            d12[ident(p)].append([pos_u[u], min(dev12.get((p, u), 0), t), t])
    central_executante(b, resumo, j12)
    r12 = ["%s/%s" % (MESES[mm - 1], str(a)[2:]) for a, mm in janela12(ref_mes)]
    devolvidas = dict(meta, unidades=[[u, nivel(u, niveis), esf.get(u)] for u in nomes],
                      filas={ident(p): sorted(([pos_u[u], q] for u, q in d["dev_un"].items()), key=lambda x: -x[1])
                             for p, d in procs.items() if d["dev_un"]},
                      janela12=[r12[0], r12[-1]], d12=dict(d12))
    return dict(meta, meses=rot[-N_PAINEL:], filas=resumo), detalhes, devolvidas, meta


ATEND_DIAS = 90   # janela dos atendimentos usados no "tempo até o atendimento"
MIN_UNIDADE = 20  # solicitações na fila, por unidade e cor, para comparar a espera entre unidades


def resultado_agendamento(st):
    """Situação registrada do agendamento: confirmado, falta, sem confirmação do executante, cancelado ou outro."""
    s = (st or "").upper()   # o status bruto traz prefixo (ex.: "AGENDAMENTO / FALTA / EXECUTANTE")
    return ("canc" if "CANCEL" in s else "falta" if "FALTA" in s else "pend" if "PENDENTE" in s
            else "conf" if "CONFIRMADO" in s else "outros")


def estatisticas_cores(b, hoje, ref_mes):
    """Por procedimento e cor: espera de quem está na fila, espera até o atendimento (agendados nos últimos
    90 dias) e entrada/saída mensais médias dos últimos 12 meses completos (base da simulação por cor)."""
    a_, m_ = ref_mes.year, ref_mes.month - 12
    if m_ < 1:
        a_, m_ = a_ - 1, m_ + 12
    ini = date(a_, m_, 1)                     # 12 meses completos: do mesmo mês do ano anterior até o mês passado
    lim_fim = ref_mes                         # primeiro dia depois do último mês completo (exclusivo)
    corte_at = hoje - timedelta(days=ATEND_DIAS)
    acc = defaultdict(lambda: defaultdict(lambda: {"fila": [], "atend": [], "ent": 0, "sai": 0}))
    todos = defaultdict(list)                 # dias de espera de toda a fila (pendentes + reenviadas) por procedimento
    por_un = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))   # procedimento > cor > unidade > dias de espera
    ag12 = defaultdict(lambda: {"n": 0, "conf": 0, "falta": 0, "pend": 0, "canc": 0, "outros": 0})
    vistos, entradas = set(), set()   # entradas: solicitações já contadas na entrada, para nenhuma contar duas vezes
    for p, cod, dt, st, risco, _cid, u, *_ in b["det_sol"]["rows"]:
        if not _valido(p) or (p, cod) in vistos or not _valido(dt):
            continue
        vistos.add((p, cod))
        cor = CORES.get(risco, "sem risco")
        d = _br(dt)
        if st in (PEND, REENV):
            acc[p][cor]["fila"].append((hoje - d).days)
            todos[p].append((hoje - d).days)
            if _valido(u):
                por_un[p][cor][u].append((hoje - d).days)
            if ini <= d < lim_fim:
                entradas.add((p, cod))
                acc[p][cor]["ent"] += 1
    vistos = set()
    for p, cod, dsol, dag, st_ag, risco, *_ in b["det_ag"]["rows"]:
        if not _valido(p) or not _valido(dag) or (p, cod, dag) in vistos:
            continue
        vistos.add((p, cod, dag))
        cor = CORES.get(risco, "sem risco")
        da = _br(dag)
        if ini <= da < lim_fim:
            acc[p][cor]["sai"] += 1           # saída = cada agendamento, como no OUTPUT da série
            ag12[p]["n"] += 1
            ag12[p][resultado_agendamento(st_ag)] += 1
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
                      "atend_mediana": round(statistics.median(a)) if a else None,
                      "atend_n": len(a), "ent": round(v["ent"] / 12, 1), "sai": round(v["sai"] / 12, 1)}
        un, dif = {}, None
        for cor, us in por_un.get(p, {}).items():
            med = sorted(([u, len(x), round(statistics.median(x))] for u, x in us.items() if len(x) >= MIN_UNIDADE), key=lambda z: -z[2])
            if len(med) >= 2:     # só há o que comparar com duas unidades ou mais
                un[cor] = med
                d = med[0][2] - med[-1][2]
                if dif is None or d > dif["d"]:
                    dif = {"cor": cor, "d": d, "max": med[0][2], "umax": med[0][0], "min": med[-1][2], "umin": med[-1][0]}
        r["_un"], r["_dif"], r["_ag"] = un or None, dif, dict(ag12[p]) if p in ag12 else None
        t = todos.get(p, [])
        r["_fila"] = {"media": round(statistics.mean(t)) if t else None,
                      "mediana": round(statistics.median(t)) if t else None, "n": len(t)}
        saida[ident(p)] = r
    return saida


def serie_agendamentos(b, meses):
    """Por fila, mês a mês (os mesmos meses do gráfico, pela data do agendamento): [agendamentos, faltas, sem confirmação
    do executante, confirmados = agendamentos efetivos]. Base do gráfico de absenteísmo da ficha. Um agendamento por código e data, como na estatística de 12 meses."""
    idx = {r: i for i, r in enumerate(meses)}
    n = len(meses)
    serie, vistos = {}, set()
    for p, cod, _dsol, dag, st_ag, *_ in b["det_ag"]["rows"]:
        if not _valido(p) or not _valido(dag) or (p, cod, dag) in vistos:
            continue
        vistos.add((p, cod, dag))
        da = _br(dag)
        i = idx.get("%s/%s" % (MESES[da.month - 1], str(da.year)[2:]))
        if i is None:
            continue
        s = serie.setdefault(ident(p), [[0] * n, [0] * n, [0] * n, [0] * n])
        s[0][i] += 1
        r = resultado_agendamento(st_ag)
        if r == "falta":
            s[1][i] += 1
        elif r == "pend":
            s[2][i] += 1
        elif r == "conf":
            s[3][i] += 1
    return serie


def montar_detalhe(b, hoje):
    """Linhas de cada fila: solicitações (pendentes, reenviadas, devolvidas e em fila de espera) e agendamentos dos últimos 24 meses."""
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
    if "DEVOLVIDA" in st:
        return "devolvida"
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


def mesclar_dia(dia, entradas, saidas, estoque):
    """Junta ao registro do dia (E1, X1) o que mudou desde a publicação anterior (E2, X2), por código SISREG:
    E = (E1 − X2) ∪ (E2 − X1) e X = (X1 − E2) ∪ (X2 − E1). Quem entra e sai (ou sai e volta) no mesmo dia não aparece
    em nenhum dos dois; a saída nova leva o destino novo e a antiga mantém o seu."""
    e1, x1 = {x[2] for x in dia["entradas"]}, {x[2] for x in dia["saidas"]}
    e2, x2 = {x[2] for x in entradas}, {x[2] for x in saidas}
    dia["entradas"] = [x for x in dia["entradas"] if x[2] not in x2] + [x for x in entradas if x[2] not in x1 and x[2] not in e1]
    dia["saidas"] = [x for x in dia["saidas"] if x[2] not in e2] + [x for x in saidas if x[2] not in e1 and x[2] not in x1]
    dia["estoque"] = estoque


def _reg_publicado(reg, hoje_s):
    """O commit de devolvidas_reg.json em main pode falhar (rebase) e o arquivo local daquela rodada se perde, enquanto o
    site publica a versão com os eventos. Para não perder o que já foi registrado, adota do site os dias que faltam
    aqui e, no dia de hoje, o registro mais completo."""
    try:
        with urllib.request.urlopen(PUBLICADO + "devolvidas_reg.json", timeout=60) as r:
            pub = json.loads(r.read().decode("utf-8"))
        assert isinstance(pub.get("dias"), list)
    except Exception as e:
        print("Registro de devolvidas: não li o registro publicado (%s)." % e)
        return
    def traduz(dia):   # índices do registro publicado -> índices do registro local
        pi = lambda i: _indice(reg["p"], pub["p"][i]); ui = lambda i: _indice(reg["u"], pub["u"][i])
        return dict(dia, entradas=[[pi(x[0]), ui(x[1]), *x[2:]] for x in dia["entradas"]],
                    saidas=[[pi(x[0]), ui(x[1]), *x[2:]] for x in dia["saidas"]])
    locais = {d["data"]: d for d in reg["dias"]}
    for d in pub["dias"]:
        l = locais.get(d["data"])
        if l is None:
            reg["dias"].append(traduz(d))
        elif d["data"] == hoje_s and len(d["entradas"]) + len(d["saidas"]) > len(l["entradas"]) + len(l["saidas"]):
            reg["dias"][reg["dias"].index(l)] = traduz(d)
    if pub.get("inicio") and pub["inicio"] < reg["inicio"]:
        reg["inicio"], reg["estoque_inicial"] = pub["inicio"], pub.get("estoque_inicial", reg["estoque_inicial"])


def historico_devolvidas(dev_det, det, hoje, anterior=None):
    """Registro diário da resolução das devolvidas (devolvidas_reg.json, versionado em main; o workflow faz o commit).
    Compara as devolvidas de hoje com as da publicação anterior (detalhe/devolvidas.json no site): códigos que
    ENTRARAM no status devolvida e códigos que SAÍRAM dele, com o destino (reenviada, agendada, pendente, fila de
    espera ou "ausente" = cancelada presumida: o código deixou de constar no Mapa Social, que só lista a solicitação
    em pendente, reenviada, devolvida e fila de espera). O Mapa Social só mostra a situação atual; este registro é a linha de base
    para medir quantas devoluções cada unidade corrige e em quanto tempo. Procedimentos e unidades ficam em listas
    (p, u) e as linhas usam os índices: entradas [pi, ui, código, data da devolução], saídas [..., destino]."""
    arq = PAINEL / "devolvidas_reg.json"
    try:
        reg = json.loads(arq.read_text(encoding="utf-8"))
        assert isinstance(reg.get("dias"), list)
    except (OSError, ValueError, AssertionError) as e:
        print("devolvidas_reg.json ausente ou ilegível (%s); começando o registro hoje." % e)
        # estoque por código único (um código pode ter duas datas de devolução), igual ao "estoque" de cada dia
        reg = {"inicio": hoje.isoformat(), "estoque_inicial": len({r[1] for r in dev_det["r"]}), "p": [], "u": [], "dias": []}
    if anterior is None:
        _reg_publicado(reg, hoje.isoformat())
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
    if len(atual) < MIN_ESTOQUE_REL * len(antes):   # queda brusca do estoque = falha de coleta, não 'canceladas'
        print("::warning::Registro de devolvidas: estoque caiu de %d para %d; coleta suspeita, dia sem comparação." % (len(antes), len(atual)))
        return reg
    status = {}
    for d in det.values():
        for sl in d["s"]:
            status[sl[0]] = sl[2]
        for ag in d["a"]:
            status.setdefault(ag[0], "AGENDADA")
    estreantes = set(UNIDADES_INCORPORADAS) - {r[5] for r in anterior["r"]}
    incorporadas = {cod for cod, (_, u, _) in atual.items() if u in estreantes and cod not in antes}
    if incorporadas:   # entram no estoque, não como devolvidas novas (senão o dia registraria todo o estoque dessas unidades)
        reg["estoque_inicial"] += len(incorporadas)
        reg.setdefault("incorporadas", []).append({"data": hoje.isoformat(), "n": len(incorporadas), "unidades": sorted(estreantes)})
        print("Registro de devolvidas: %d devolvidas de %s incorporadas ao estoque." % (len(incorporadas), ", ".join(sorted(estreantes))))
    entradas = [[_indice(reg["p"], p), _indice(reg["u"], u), cod, dd] for cod, (p, u, dd) in atual.items()
                if cod not in antes and cod not in incorporadas]
    saidas = [[_indice(reg["p"], p), _indice(reg["u"], u), cod, dd, _destino(status.get(cod))] for cod, (p, u, dd) in antes.items() if cod not in atual]
    hoje_s = hoje.isoformat()
    dia = next((d for d in reg["dias"] if d["data"] == hoje_s), None)
    if dia is None:
        dia = {"data": hoje_s, "anterior": anterior.get("data"), "estoque": len(atual), "entradas": [], "saidas": []}
        reg["dias"].append(dia)
    mesclar_dia(dia, entradas, saidas, len(atual))   # nova coleta no mesmo dia: o que mudou se junta ao já registrado
    reg["dias"].sort(key=lambda d: d["data"])
    dest = defaultdict(int)
    for x in dia["saidas"]:
        dest[x[4]] += 1
    print("Registro de devolvidas: %d no estoque, %d entradas e %d saídas hoje (%s); %d dias registrados" % (
        len(atual), len(dia["entradas"]), len(dia["saidas"]), ", ".join("%s %d" % kv for kv in sorted(dest.items())) or "sem saídas", len(reg["dias"])))
    return reg


DESTINOS_FILA = ("agendada", "devolvida", "fila de espera", "ausente")   # ordem das contagens em fila_reg.json


def _fila_atual(det):
    """{código SISREG: id da fila} de quem está na fila hoje (pendentes + reenviadas ao regulador)."""
    return {s[0]: i for i, d in det.items() for s in d["s"] if s[2] in (_curto(PEND), _curto(REENV))}


def historico_fila(det, hoje, anterior=None, publicado=None):
    """Saldo inexplicado (fila_reg.json, versionado em main): compara os códigos que estavam na fila na publicação anterior
    com os de hoje e conta, por mês e por fila, o destino de quem saiu: agendada, devolvida, fila de espera ou ausente (o
    código deixou de constar no Mapa Social sem agendamento). Comparar códigos evita somar séries mensais, que mudam
    para trás a cada recarga. Quem saiu da fila e voltou entre duas coletas não aparece."""
    arq = PAINEL / "fila_reg.json"
    try:
        reg = json.loads(arq.read_text(encoding="utf-8"))
        assert isinstance(reg.get("meses"), dict)
    except (OSError, ValueError, AssertionError) as e:
        print("fila_reg.json ausente ou ilegível (%s); começando o registro." % e)
        reg = {"inicio": None, "ultima": None, "meses": {}}
    if anterior is None:
        try:   # o commit em main pode ter falhado: o que já foi publicado prevalece se estiver mais adiante
            with urllib.request.urlopen(PUBLICADO + "fila_reg.json", timeout=60) as r:
                publicado = json.loads(r.read().decode("utf-8"))
        except Exception as e:
            print("Saldo inexplicado: não li o registro publicado (%s)." % e)
        try:
            with urllib.request.urlopen(PUBLICADO + "detalhe/fila.json", timeout=180) as r:
                anterior = json.loads(r.read().decode("utf-8"))
        except Exception as e:   # primeira coleta ou site fora: hoje só grava a base da próxima comparação
            print("Saldo inexplicado: sem publicação anterior para comparar (%s)." % e)
            return reg
    if publicado and (publicado.get("ultima") or "") > (reg.get("ultima") or ""):
        reg = publicado
    if not anterior.get("r"):
        return reg
    atual = _fila_atual(det)
    antes = {cod: anterior["ids"][k] for k, cod in anterior["r"]}
    if len(atual) < MIN_ESTOQUE_REL * len(antes):   # queda brusca da fila = falha de coleta, não saída de verdade
        print("::warning::Saldo inexplicado: fila caiu de %d para %d; coleta suspeita, dia sem comparação." % (len(antes), len(atual)))
        return reg
    status = {}
    for d in det.values():
        for sl in d["s"]:
            status[sl[0]] = sl[2]
        for ag in d["a"]:
            status.setdefault(ag[0], "AGENDADA")
    mes, total = reg["meses"].setdefault(hoje.isoformat()[:7], {}), [0, 0, 0, 0]
    for cod, i in antes.items():
        if cod not in atual:
            dest = _destino(status.get(cod))
            if dest in DESTINOS_FILA:
                k = DESTINOS_FILA.index(dest)
                mes.setdefault(i, [0, 0, 0, 0])[k] += 1
                total[k] += 1
    reg["inicio"] = reg["inicio"] or anterior.get("data")
    reg["ultima"] = hoje.isoformat()
    print("Saldo inexplicado: saíram %d da fila (%s)" % (sum(total), ", ".join("%s %d" % x for x in zip(DESTINOS_FILA, total))))
    return reg


def acumular_fila(dados, reg):
    """Soma o registro por fila: sx = [agendadas, devolvidas, fila de espera, ausentes]."""
    soma = {}
    for mes in reg["meses"].values():
        for i, v in mes.items():
            s = soma.setdefault(i, [0, 0, 0, 0])
            for k, q in enumerate(v):
                s[k] += q
    for f in dados["filas"]:
        if f["id"] in soma:
            f["sx"] = soma[f["id"]]
    if reg.get("inicio"):
        dados["sx_desde"] = reg["inicio"]


def trava_relativa(dados, anterior=None):
    """Compara a coleta com a publicação anterior; devolve a lista de quedas suspeitas (vazia = pode publicar)."""
    if anterior is None:
        try:
            with urllib.request.urlopen(PUBLICADO + "dados.json", timeout=60) as resp:
                anterior = json.load(resp)
        except Exception as e:   # sem publicação legível não há com que comparar; os mínimos fixos continuam valendo
            print("::warning::Trava da coleta: não li a publicação anterior (%s)." % e)
            return []
    problemas = []
    for rotulo, campo in (("Fila", "fila"), ("Devolvidas", "devolvidas")):
        antes, agora = sum(f[campo] for f in anterior["filas"]), sum(f[campo] for f in dados["filas"])
        if antes and agora < (1 - MAX_QUEDA_TOTAL) * antes:
            problemas.append("%s caiu de %d para %d (%.0f%%)." % (rotulo, antes, agora, (agora / antes - 1) * 100))
    atual = {f["p"]: f["fila"] for f in dados["filas"]}
    for f in anterior["filas"]:
        if f["fila"] >= MIN_FILA_GRANDE and atual.get(f["p"], 0) < f["fila"] / 2:
            problemas.append("%s caiu de %d para %d." % (f["p"], f["fila"], atual.get(f["p"], 0)))
    return problemas


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
    refs = referencias()
    dados, detalhes, devolvidas, meta = montar(b, hoje, refs)
    est = estatisticas_cores(b, hoje, mes_referencia(hoje, dados["solicitacoes_ate"]))
    agm = serie_agendamentos(b, dados["meses"])
    for f in dados["filas"]:
        f["agm"] = agm.get(f["id"])   # [agendamentos, faltas, sem confirmação, confirmados] por mês; None se a fila não tem agendamentos
        cs = dict(est.get(f["id"], {}))
        f["esp_fila"] = cs.pop("_fila", None)   # espera de quem está na fila hoje, fila inteira (média, mediana, n)
        f["esp_un"], f["dif_un"], f["ag12"] = cs.pop("_un", None), cs.pop("_dif", None), cs.pop("_ag", None)
        f["cor_stats"] = cs
        # o protocolo calcula a meta da inserção à execução: compara com a mediana até o atendimento, não com quem ainda espera
        meta_f = refs.get("metas", {}).get(f["p"]) or {}   # só há meta quando uma norma a define (aba Metas_espera)
        pos_a = 1 if f["alertas"] and f["alertas"][0].startswith("Fila crescendo") else 0
        for cor in ("vermelho", "amarelo", "verde", "azul"):
            c_st = cs.get(cor) or {}
            med, n_at, lim = c_st.get("atend_mediana"), c_st.get("atend_n", 0), meta_f.get(cor)
            if med is not None and lim and med > lim:
                fraco = "; só %d %s" % (n_at, "caso" if n_at == 1 else "casos") if n_at < MIN_UNIDADE else ""
                f["alertas"].insert(pos_a, "%s: mediana de %s dias até o atendimento (meta: %d%s)" % (cor.capitalize(), f"{med:,}".replace(",", "."), lim, fraco))
                pos_a += 1
    dados["atend_dias"] = ATEND_DIAS
    if len(dados["filas"]) < MIN_FILAS or not dados["solicitacoes_ate"] or dados["solicitacoes_ate"] == "-":
        print(f"ERRO: coleta incompleta ({len(dados['filas'])} filas, solicitações até {dados['solicitacoes_ate']!r}); nada foi gravado.")
        sys.exit(1)
    sinais_ok = (sum(1 for f in dados["filas"] if f["fila"] > 0) >= MIN_COM_FILA and sum(f["fila"] for f in dados["filas"]) >= MIN_NA_FILA
                 and sum(1 for f in dados["filas"] if f["devolvidas"] > 0) >= MIN_COM_DEV and sum(f["devolvidas"] for f in dados["filas"]) >= MIN_DEVOLVIDAS)
    if not sinais_ok:
        print("ERRO: coleta com fila ou devolvidas muito abaixo do esperado (cubo vazio?); nada foi gravado.")
        sys.exit(1)
    problemas = trava_relativa(dados)
    if problemas and "--liberar-trava" not in sys.argv:
        for x in problemas:
            print("::error::Trava da coleta: " + x)
        print("ERRO: queda brusca em relação à publicação anterior; nada foi gravado. Se a queda for real, rode o workflow com liberar_trava.")
        sys.exit(CODIGO_TRAVA)
    for x in problemas:
        print("::warning::Trava liberada manualmente: " + x)
    det, dev_det = montar_detalhe(b, hoje)
    if not det or not dev_det.get("r"):
        print("ERRO: detalhe sem solicitações ou sem devolvidas; nada foi gravado.")
        sys.exit(1)
    print("[tempo] montagem: %.0f s" % (time.monotonic() - t), flush=True)
    t = time.monotonic()
    hist = historico_atraso(meta, hoje)
    _grava(PAINEL / "atraso.json", hist)
    _grava(PAINEL / "espera.json", historico_espera(dados, hoje))
    _grava(PAINEL / "devolvidas_reg.json", historico_devolvidas(dev_det, det, hoje))
    reg_fila = historico_fila(det, hoje)
    _grava(PAINEL / "fila_reg.json", reg_fila)
    acumular_fila(dados, reg_fila)
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
    atual = _fila_atual(det)
    ids = sorted(set(atual.values()))
    pos = {i: k for k, i in enumerate(ids)}
    _grava(PAINEL / "detalhe" / "fila.json", dict(meta, ids=ids, r=[[pos[i], c] for c, i in atual.items()]))
    print("[tempo] gravação: %.0f s" % (time.monotonic() - t), flush=True)
    _saida_actions(publicar="sim")
    print(f"{dados['gerado_em']}: {len(dados['filas'])} filas, solicitações até {dados['solicitacoes_ate']}, "
          f"{sum(len(d['s']) for d in det.values())} solicitações e {sum(len(d['a']) for d in det.values())} "
          f"agendamentos no detalhe")


if __name__ == "__main__":
    main()

"""Atualiza o painel diário das filas (todas as filas da Região Norte).

Roda todo dia no GitHub Actions (.github/workflows/atualizar.yml), que grava os dados e publica.
Uso local: python coleta/painel.py     (só coleta e grava, sem publicar)

Gera:
  dados.json          resumo de todas as filas (painel)
  analises/<id>.json  detalhes agregados de cada fila, usados pelo botão "Gerar análise"
"""
import hashlib
import json
import os
import sys
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

import mapa_social as m

AQUI = Path(__file__).resolve().parent
PAINEL = Path(os.environ.get("PAINEL_DIR", AQUI.parent))
R = "[regiao_saude_solicitante]={'%s'}" % m.REGIAO
FILA = '[status_solicitacao_solicitacao]={"*PENDENTE*","*REENVIADA*"}'
DEV = '[status_solicitacao_solicitacao]={"*DEVOLVIDA*"}'
MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
CORES = {"Prioridade 1": "vermelho", "Prioridade 2": "amarelo", "Prioridade 3": "verde", "Prioridade 4": "azul"}
N_PAINEL = 24          # meses mostrados no painel
MIN_FILAS = 500        # abaixo disso a coleta é considerada falha (não publica)
META_PADRAO = {"vermelho": 30}   # referência geral quando o protocolo não tem meta cadastrada
ANOS_SERIE = 5         # série da análise: de janeiro de (ano - 5) até o mês anterior


def coletar(hoje):
    ini = hoje.year - ANOS_SERIE
    P, PA, U = "descrição_procedimento_solicitacao", "descrição_procedimento_agendamento", "nome_unidade_solicitante"
    return m.consultar([
        ["fila", [P, "status_solicitacao_solicitacao", "classificacao_risco"],
         ["Count({<%s>} distinct codigo_solicitacao)" % R, "Date(Min({<%s>} data_solicitacao))" % R]],
        ["in_sol", [P, "Calendar.Ano_Solicitacao", "Calendar.Mes_Solicitacao"],
         ['Count({<%s,%s,[Calendar.Ano_Solicitacao]={">=%d"}>} distinct codigo_solicitacao)' % (R, FILA, ini)]],
        ["in_ag", [PA, "Calendar.Ano_Solicitacao", "Calendar.Mes_Solicitacao"],
         ['Count({<%s,[Calendar.Ano_Solicitacao]={">=%d"}>} distinct codigo_solicitacao)' % (R, ini)]],
        ["out", [PA, "Ano_Agendamento", "Mes_Agendamento"],
         ['Count({<%s,[Ano_Agendamento]={">=%d"}>} distinct codigo_solicitacao)' % (R, ini)]],
        ["cid", [P, "CID_solicitacao"], ["Count({<%s,%s>} distinct codigo_solicitacao)" % (R, FILA)]],
        ["dev_un", [P, U], ["Count({<%s,%s>} distinct codigo_solicitacao)" % (R, DEV)]],
        ["un_sol", [P, U], ["Count({<%s>} distinct codigo_solicitacao)" % R]],
        ["un_ag", [PA, U], ["Count({<%s>} distinct codigo_solicitacao)" % R]],
        ["tipo", [P, "tipo_procedimento_solicitacao"], ["Count({<%s>} distinct codigo_solicitacao)" % R]],
        ["ultima", [], ["Date(Max(data_solicitacao))"]],
    ], layout=True)


def _br(s):
    return datetime.strptime(s, "%d/%m/%Y").date()


def ident(p):
    return hashlib.sha1(p.encode("utf-8")).hexdigest()[:12]


def _valido(p):
    return p not in ("-", "")


def metas_ref():
    try:
        return json.loads((PAINEL / "referencias.json").read_text(encoding="utf-8")).get("metas", {})
    except (OSError, ValueError):
        return {}


def montar(b, hoje, metas=None):
    metas = metas or {}
    fim = (hoje.year, hoje.month - 1) if hoje.month > 1 else (hoje.year - 1, 12)
    meses, a, mm = [], hoje.year - ANOS_SERIE, 1
    while (a, mm) <= fim:
        meses.append((a, mm))
        a, mm = (a, mm + 1) if mm < 12 else (a + 1, 1)
    idx = {k: i for i, k in enumerate(meses)}
    n = len(meses)

    procs = defaultdict(lambda: {"in": [0] * n, "out": [0] * n, "cores": {}, "fila": 0, "devolvidas": 0,
                                 "futuros": 0, "tipo": "", "cid": {}, "dev_un": {}, "unidades": set()})
    for p, st, risco, q, antiga in b["fila"]["rows"]:
        if not _valido(p):
            continue
        d, q = procs[p], int(q)
        if "DEVOLVIDA" in st:
            d["devolvidas"] += q
        elif "PENDENTE" in st or "REENVIADA" in st:
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
        esperas = [c["dias"] for c in cores.values() if c["dias"] is not None]
        i = ident(p)
        resumo.append({"id": i, "p": p, "tipo": d["tipo"], "fila": d["fila"], "cores": cores,
                       "devolvidas": d["devolvidas"], "in": ent, "out": sai, "saldo12": sum(u12),
                       "futuros": d["futuros"], "espera_max": max(esperas) if esperas else None, "alertas": alertas})
        detalhes[i] = {"p": p, "tipo": d["tipo"], "meses": rot, "in": d["in"], "out": d["out"], "cores": cores,
                       "cid": d["cid"], "devolvidas_unidade": d["dev_un"], "unidades": sorted(d["unidades"])}
    resumo.sort(key=lambda x: -x["fila"])
    meta = {"gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M"), "data": hoje.isoformat(),
            "solicitacoes_ate": b["ultima"]["rows"][0][0], "recarga_painel": b.get("_reload", "")}
    return dict(meta, meses=rot[-N_PAINEL:], filas=resumo), detalhes, meta


def main():
    hoje = date.today()
    dados, detalhes, meta = montar(coletar(hoje), hoje, metas_ref())
    if len(dados["filas"]) < MIN_FILAS or not dados["solicitacoes_ate"] or dados["solicitacoes_ate"] == "-":
        print(f"ERRO: coleta incompleta ({len(dados['filas'])} filas, solicitações até {dados['solicitacoes_ate']!r}); nada foi gravado.")
        sys.exit(1)
    (PAINEL / "dados.json").write_text(json.dumps(dados, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    pasta = PAINEL / "analises"
    pasta.mkdir(exist_ok=True)
    for velho in pasta.glob("*.json"):   # apaga arquivo a arquivo (rmtree falha com pasta sincronizada pelo OneDrive)
        velho.unlink()
    for i, d in detalhes.items():
        (pasta / f"{i}.json").write_text(json.dumps(dict(meta, **d), ensure_ascii=False, separators=(",", ":")),
                                         encoding="utf-8")
    print(f"{dados['gerado_em']}: {len(dados['filas'])} filas, solicitações até {dados['solicitacoes_ate']}")


if __name__ == "__main__":
    main()

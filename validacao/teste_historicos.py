"""Teste com dados sintéticos das funções de histórico e de trava da coleta (não usa o Qlik nem a rede).

Rodar em painel/:  python validacao/teste_historicos.py      (sai com código != 0 se algo falhar)
Cobre: leitura dos históricos (ausente x corrompido), idempotência de historico_fila (mesma publicação
comparada duas vezes soma uma vez só) e trava_relativa fechando quando a publicação anterior não é lida.
"""
import json
import os
import sys
import tempfile
from datetime import date
from pathlib import Path

TMP = Path(tempfile.mkdtemp(prefix="painel_teste_"))
os.environ["PAINEL_DIR"] = str(TMP)
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "coleta"))
import painel as P  # noqa: E402

falhas = []


def confere(ok, msg):
    print(("ok    " if ok else "FALHA ") + msg)
    if not ok:
        falhas.append(msg)


def sai_com_erro(fn):
    try:
        fn()
    except SystemExit as e:
        return e.code == 1
    return False


# 1) leitura dos históricos
arq = TMP / "x.json"
confere(P._ler_registro(arq, "dias", list) is None, "arquivo ausente começa do zero")
arq.write_text('{"dias": [1]}', encoding="utf-8")
confere(P._ler_registro(arq, "dias", list) == {"dias": [1]}, "arquivo válido é lido")
arq.write_text("<<<<<<< HEAD\n{}", encoding="utf-8")
confere(sai_com_erro(lambda: P._ler_registro(arq, "dias", list)), "arquivo corrompido (marcador de conflito) para a coleta")
arq.write_text('{"meses": []}', encoding="utf-8")
confere(sai_com_erro(lambda: P._ler_registro(arq, "meses", dict)), "arquivo no formato errado para a coleta")
arq.write_text("[1, 2]", encoding="utf-8")
confere(sai_com_erro(lambda: P._ler_registro(arq, "dias", list)), "JSON que não é objeto para a coleta")

# 2) historico_fila: idempotência (100 códigos na fila; a guarda de queda brusca barra perdas de mais de 2%)
PEND = P._curto(P.PEND)
cods = ["c%d" % k for k in range(98)]
det = {"fila1": {"s": [[c, "01/11/2026", PEND] for c in cods] + [["c99", "01/11/2026", "DEVOLVIDA / REGULADOR"]], "a": [["c98", "x", "AGENDADA"]]}}
anterior = {"gerado_em": "22/11/2026 10:00", "data": "2026-11-22", "ids": ["fila1"], "r": [[0, c] for c in cods + ["c98", "c99"]]}
hoje = date(2026, 11, 23)
r1 = P.historico_fila(det, hoje, anterior=anterior)
mes = r1["meses"]["2026-11"]["fila1"]
confere(mes == [1, 1, 0, 0], "primeira comparação conta c98 (agendada) e c99 (devolvida): %s" % mes)
confere(r1.get("base") == "22/11/2026 10:00", "guarda a publicação comparada")
(TMP / "fila_reg.json").write_text(json.dumps(r1), encoding="utf-8")   # o que o workflow versiona
r2 = P.historico_fila(det, hoje, anterior=anterior)
confere(r2["meses"]["2026-11"]["fila1"] == [1, 1, 0, 0], "segunda comparação com a mesma base não soma de novo: %s" % r2["meses"]["2026-11"]["fila1"])
anterior2 = dict(anterior, gerado_em="23/11/2026 12:00", r=[[0, c] for c in cods + ["z1", "z2"]])   # outra publicação: soma normalmente
r3 = P.historico_fila(det, hoje, anterior=anterior2)
confere(r3["meses"]["2026-11"]["fila1"] == [1, 1, 0, 2], "outra publicação soma normalmente (z1 e z2 sumiram do Mapa Social): %s" % r3["meses"]["2026-11"]["fila1"])
(TMP / "fila_reg.json").write_text('{"inicio": null, "ultima": null, "meses": {"2026-11": {}}}', encoding="utf-8")   # arquivo antigo, sem 'base'
r4 = P.historico_fila(det, hoje, anterior=anterior)
confere(r4["meses"]["2026-11"]["fila1"] == [1, 1, 0, 0], "registro antigo, sem 'base', continua funcionando")

# 3) trava fecha quando a publicação anterior não é lida
P.PUBLICADO = "http://127.0.0.1:9/"   # porta fechada
confere(len(P.trava_relativa({"filas": []})) == 1, "sem publicação anterior legível, a trava devolve um problema")
ant = {"filas": [{"p": "X", "fila": 1000, "devolvidas": 100}]}
confere(P.trava_relativa({"filas": [{"p": "X", "fila": 990, "devolvidas": 100}]}, anterior=ant) == [], "queda pequena passa")
confere(len(P.trava_relativa({"filas": [{"p": "X", "fila": 100, "devolvidas": 100}]}, anterior=ant)) >= 1, "queda brusca é barrada")

print("\n%d falha(s)" % len(falhas))
sys.exit(1 if falhas else 0)

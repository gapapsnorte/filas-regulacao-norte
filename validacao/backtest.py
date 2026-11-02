"""Teste histórico das projeções do simulador (rodar em painel/, depois de uma coleta):  python validacao/backtest.py

Pergunta: para prever os 12 meses seguintes de cada fila, o que erra menos, repetir a média dos
últimos 12 meses ou prolongar a reta de tendência?

Método: em cada fila com volume suficiente, escolhe-se um ponto no passado (3 pontos, com 6 meses
de distância entre eles), usam-se só os 24 meses anteriores e compara-se a previsão com a média real dos
12 meses seguintes. Erro = |previsto − real| ÷ real. Os resultados são medianas. O último mês
da série entra só se estiver completo (campo mes_incompleto).
Saída: validacao/resultado_backtest.md (o site cita esses números).
"""
import glob
import json
import statistics as st
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
MIN_MEDIA = 3       # média mensal mínima (entrada e saída) na janela completa para a fila entrar no teste
ORIGENS = (12, 18, 24)  # meses antes do fim da série em que a previsão é feita


def reta(y):
    n, mx, my = len(y), (len(y) - 1) / 2, sum(y) / len(y)
    sxx = sum((i - mx) ** 2 for i in range(n)) or 1
    b = sum((i - mx) * (v - my) for i, v in enumerate(y)) / sxx
    return my - b * mx, b


def prev_media(y):
    return sum(y[-12:]) / 12


def prev_tendencia(y):
    a, b = reta(y)
    return max(0, sum(a + b * (len(y) + k) for k in range(12)) / 12)


def casos():
    for arq in glob.glob(str(RAIZ / "analises" / "*.json")):
        d = json.load(open(arq, encoding="utf8"))
        corte = 1 if d.get("mes_incompleto") else 0
        ent, sai = d["in"], d["out"]
        if corte:
            ent, sai = ent[:-corte], sai[:-corte]
        for o in ORIGENS:
            if len(ent) < o + 24:
                continue
            fim = len(ent) - o
            te, ts = ent[fim - 24:fim], sai[fim - 24:fim]
            re, rs = ent[fim:fim + 12], sai[fim:fim + 12]
            if min(sum(te) / 24, sum(ts) / 24, sum(re) / 12, sum(rs) / 12) < MIN_MEDIA:
                continue
            yield te, ts, sum(re) / 12, sum(rs) / 12


def erro(prev, real):
    return abs(prev - real) / real * 100


def main():
    C = list(casos())
    lin = [("média 12 m", prev_media), ("reta de tendência", prev_tendencia)]
    tab = [(nome, st.median(erro(f(te), re) for te, _, re, _ in C), st.median(erro(f(ts), rs) for _, ts, _, rs in C)) for nome, f in lin]
    pio_e = st.median((re - prev_media(te)) / prev_media(te) * 100 for te, _, re, _ in C if re > prev_media(te))
    pio_s = st.median((prev_media(ts) - rs) / prev_media(ts) * 100 for _, ts, _, rs in C if rs < prev_media(ts))
    txt = ["# Teste histórico das projeções", "",
           f"Casos: {len(C)} (média mensal ≥ {MIN_MEDIA} em entrada e saída; origens {ORIGENS} meses antes do fim da série; "
           "treino de 24 meses; previsão da média dos 12 meses seguintes).", "",
           "| Método | Erro mediano, entrada | Erro mediano, saída |", "|---|---|---|"]
    txt += [f"| {n} | {e:.0f}% | {s:.0f}% |" for n, e, s in tab]
    txt += ["", f"Quando a previsão pela média errou para o lado ruim, o erro mediano foi: entrada {pio_e:.0f}% acima, saída {pio_s:.0f}% abaixo.", ""]
    (RAIZ / "validacao" / "resultado_backtest.md").write_text("\n".join(txt), encoding="utf8")
    print("\n".join(txt))


if __name__ == "__main__":
    main()

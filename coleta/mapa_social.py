"""Leitura do painel público do Mapa Social (Qlik Sense / MPDFT) por um Chrome invisível.

Só puxa campos agregáveis (códigos, datas, status, risco, CID, unidade). Nome, telefone e CNS
nunca são consultados.
"""
import asyncio
import os
import time
from playwright.async_api import async_playwright

URL = "https://paineis-ext.mpdft.mp.br/extensions/mapasauderegulamentacao/mapasauderegulamentacao.html"
APP = "7edb02fa-f1f5-4ea6-b348-049a4ac2d095"
REGIAO = "Região Norte"
# No PC usa o Chrome instalado; no GitHub Actions (CANAL vazio) usa o Chromium do Playwright
CANAL = os.environ.get("MAPA_CHROME_CANAL", "chrome")
PARALELO = int(os.environ.get("MAPA_PARALELO", "6"))   # páginas do Qlik pedidas ao mesmo tempo

_JS = """async ([specs, layout]) => {
  const qlik = await new Promise(r => require(['js/qlik'], r));
  const app = qlik.openApp('%s', {host: location.hostname, prefix: '/', port: location.port, isSecure: true});
  const out = {};
  if (layout) out._reload = (await new Promise(r => app.getAppLayout(r))).qLastReloadTime;
  for (const [name, dims, meas] of specs) {
    const w = dims.length + meas.length;
    const t0 = Date.now();
    const model = await app.createCube({
      qDimensions: dims.map(d => ({qDef: {qFieldDefs: [d]}, qNullSuppression: false})),
      qMeasures: meas.map(m => ({qDef: {qDef: m}})), qSuppressZero: true, qInitialDataFetch: []});
    const lay = await model.getLayout();
    const n = lay.qHyperCube.qSize.qcy, h = Math.floor(10000 / w), tops = [];
    const t1 = Date.now();
    for (let top = 0; top < n; top += h) tops.push(top);
    // páginas de até 10 mil células, pedidas em lotes paralelos (a ordem das linhas é mantida)
    const paginas = [];
    for (let i = 0; i < tops.length; i += %d) {
      paginas.push(...await Promise.all(tops.slice(i, i + %d).map(top =>
        model.getHyperCubeData('/qHyperCubeDef', [{qTop: top, qLeft: 0, qWidth: w, qHeight: Math.min(h, n - top)}]))));
    }
    const rows = [];
    paginas.forEach(pg => pg[0].qMatrix.forEach(r => rows.push(r.map(c => c.qText))));
    await app.destroySessionObject(lay.qInfo.qId);
    console.log('[tempo] ' + name + ': ' + n + ' linhas, ' + tops.length + ' páginas; cálculo ' +
                ((t1 - t0) / 1000).toFixed(1) + ' s, leitura ' + ((Date.now() - t1) / 1000).toFixed(1) + ' s');
    out[name] = {cols: dims.concat(meas), rows};
  }
  return out;
}""" % (APP, PARALELO, PARALELO)


async def _consultar(specs, layout=False):
    async with async_playwright() as p:
        nav = await p.chromium.launch(channel=CANAL or None, headless=True)
        pg = await nav.new_page()
        pg.on("console", lambda msg: msg.text.startswith("[tempo]") and print(msg.text, flush=True))
        # Às vezes a página abre sem a biblioteca do Qlik (ex.: durante a recarga do Mapa Social): espera e recarrega
        for tentativa in range(4):
            try:
                resp = await pg.goto(URL, wait_until="networkidle", timeout=120000)
                await pg.wait_for_function("typeof require === 'function'", timeout=60000)
                break
            except Exception as e:
                try:   # diagnóstico: o que a página devolveu
                    txt = (await pg.inner_text("body"))[:400].replace("\n", " ")
                    print("Mapa Social sem Qlik (tentativa %d): HTTP %s, url %s, título %r, texto %r" % (
                        tentativa + 1, resp.status if resp else "?", pg.url, await pg.title(), txt), flush=True)
                except Exception:
                    print("Mapa Social sem Qlik (tentativa %d): %s" % (tentativa + 1, e), flush=True)
                if tentativa == 3:
                    raise
                await asyncio.sleep(30)
        t = time.monotonic()
        print("[tempo] página do Mapa Social aberta", flush=True)
        res = await pg.evaluate(_JS, [specs, layout])
        print("[tempo] consultas: %.0f s no total" % (time.monotonic() - t), flush=True)
        await nav.close()
        return res


def consultar(specs, layout=False):
    return asyncio.run(_consultar(specs, layout))


def _aspas(txt):
    return txt.replace("'", "''")


def listar_procedimentos():
    """Procedimentos com solicitações da Região Norte, do maior para o menor volume."""
    m = "Count({<[regiao_saude_solicitante]={'%s'}>} distinct codigo_solicitacao)" % REGIAO
    r = consultar([["p", ["descrição_procedimento_solicitacao"], [m]]])["p"]["rows"]
    lista = [(p, int(n)) for p, n in r if p and p != "-"]
    return sorted(lista, key=lambda x: -x[1])


def baixar(procedimento):
    """Bases de solicitações (pendentes, reenviadas, devolvidas) e de agendamentos do procedimento."""
    p = _aspas(procedimento)
    S = "[descrição_procedimento_solicitacao]={'%s'},[regiao_saude_solicitante]={'%s'}" % (p, REGIAO)
    A = "[descrição_procedimento_agendamento]={'%s'},[regiao_saude_solicitante]={'%s'}" % (p, REGIAO)
    specs = [
        ["solicitacoes",
         ["codigo_solicitacao", "data_solicitacao", "status_solicitacao_solicitacao", "classificacao_risco",
          "CID_solicitacao", "nome_unidade_solicitante", "data_devolucao"],
         ["Count({<%s>} codigo_solicitacao)" % S]],
        ["agendamentos",
         ["codigo_solicitacao", "Calendar.Data_Solicitacao", "data_agendamento", "status_solicitacao_agendamento",
          "classificacao_risco", "nome_unidade_solicitante"],
         ["Count({<%s>} codigo_solicitacao)" % A]],
        ["ultima_solicitacao", [], ["Date(Max(data_solicitacao))"]],
    ]
    return consultar(specs, layout=True)

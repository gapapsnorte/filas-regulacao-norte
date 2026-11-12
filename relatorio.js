/* Gera a análise de uma fila (PDF, Word ou Excel) no navegador, a partir de analises/<id>.json.
   Mesma metodologia do programa local (manual GERNO / modelo da Cirurgia Geral de 23/09/2026). */
(function () {
  "use strict";
  const LIBS = {
    pdf: ["https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js",
          "https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js"],
    word: ["https://unpkg.com/docx@8.5.0/build/index.umd.js"],
    excel: ["https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"],
  };
  const MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"];
  const MESES_EXT = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro",
    "outubro", "novembro", "dezembro"];
  const CORES = [["vermelho", "Vermelho", "FF0000"], ["amarelo", "Amarelo", "FFFF00"], ["verde", "Verde", "92D050"],
    ["azul", "Azul", "00B0F0"]];
  const FONTE = "Fonte: Mapa Social do Distrito Federal.";
  const MIN12 = 20;   // mínimo de solicitações em 12 meses para o % de devolvidas entrar no ranking/texto (abaixo disso sai em cinza)
  const CAB_ESQ = "SES-DF  ·  SRSNO  ·  DIRAPS Norte  ·  parceria GERNO e GAPAPS Norte", CAB_DIR = "Análise de Filas de Regulação – Região Norte";
  const carregadas = {};
  let REF = null;

  const fmt = (n, c = 0) => Number(n).toLocaleString("pt-BR", {minimumFractionDigits: c, maximumFractionDigits: c});
  const dataBr = d => `${d.getDate()} ${MESES[d.getMonth()]}. ${d.getFullYear()}`;
  const dd = s => s ? s.split("-").reverse().join("/") : "";
  const cmp = (a, b) => (a < b ? -1 : a > b ? 1 : 0);
  const atrasoUnico = d => d.atraso_dias == null ? "atraso não medido" : d.atraso_dias === 0 ? "sem atraso" : d.atraso_dias === 1 ? "atraso de 1 dia" : `atraso de ${d.atraso_dias} dias`;
  const atrasoTxt = d => { const u = (d.atraso_ultimos || []).filter(x => x != null), mn = Math.min(...u), mx = Math.max(...u);
    return u.length > 1 && mn !== mx ? `${atrasoUnico(d)} hoje (${mn} a ${mx} ${mx === 1 ? "dia" : "dias"} nas últimas ${u.length} medições)` : atrasoUnico(d); };

  function lib(url) {
    if (!carregadas[url]) carregadas[url] = new Promise((ok, erro) => {
      const s = document.createElement("script");
      s.src = url; s.onload = ok; s.onerror = () => erro(new Error("Falha ao carregar " + url));
      document.head.appendChild(s);
    });
    return carregadas[url];
  }

  function nomeCurto(p) {
    const t = p.trim().replace(/^(CONSULTA|EXAME|PROCEDIMENTO) EM /i, "").replace(/\s*-\s*/g, " - ");
    const minus = new Set(["e", "de", "da", "do", "das", "dos", "em", "com", "sem", "para", "a", "o"]);
    return t.toLowerCase().split(/\s+/).map((w, i) => (i && minus.has(w)) ? w : w.charAt(0).toUpperCase() + w.slice(1)).join(" ");
  }

  function mediana(v) {
    if (!v.length) return 0;
    const s = [...v].sort((p, q) => p - q), m = s.length >> 1;
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  }
  // ---- Meses aberrantes e tendência corrigida ----------------------
  // Janela: últimos 24 meses. Theil-Sen iterativo (só para ACHAR os aberrantes): mede o afastamento de cada mês à própria reta, a escala é a
  // oscilação típica (MAD × 1,4826, com piso de √nível por ser contagem) e um mês é aberrante se passar de k = 4
  // escalas acima da reta (no saldo, acima ou abaixo). Os aberrantes ficam de fora do ajuste da reta (mas continuam
  // no gráfico). Sequências de mais de 3 meses seguidos contam como mudança de patamar e não são marcadas.
  // Filas pequenas (< 12 meses com movimento ou < 120 solicitações na janela) não são avaliadas.
  // A reta final e o R² são de mínimos quadrados (como no Excel) calculados SEM os aberrantes.
  // Manter em sincronia com analise.py (aberrantes / analisar_aberrantes).
  const ABERR = {k: 4, janela: 24, minMeses: 12, minTotal: 120, maxSeq: 3};
  // Theil-Sen sobre os meses não marcados: inclinação = mediana das inclinações entre pares; intercepto = mediana de y − a·x
  function tsAjuste(y, flag) {
    const xs = [], ys = [];
    y.forEach((v, i) => { if (!flag[i]) { xs.push(i); ys.push(v); } });
    const m = xs.length, incl = [];
    if (m < 3) return {a: 0, b: mediana(y)};
    for (let i = 0; i < m; i++) for (let j = i + 1; j < m; j++) incl.push((ys[j] - ys[i]) / (xs[j] - xs[i]));
    const a = mediana(incl);
    return {a, b: mediana(ys.map((v, i) => v - a * xs[i]))};
  }
  const elegivelAberr = y => y.filter(v => v !== 0).length >= ABERR.minMeses && y.reduce((s, v) => s + v, 0) >= ABERR.minTotal;
  // Reta de mínimos quadrados e R² clássico sobre os meses não aberrantes
  function mqAjuste(y, flag) {
    const idx = y.map((_, i) => i).filter(i => !flag[i]), m = idx.length;
    if (m < 3) return {a: 0, b: mediana(y), r2: null};
    const mx = idx.reduce((s, i) => s + i, 0) / m, my = idx.reduce((s, i) => s + y[i], 0) / m;
    const sxx = idx.reduce((s, i) => s + (i - mx) ** 2, 0), sxy = idx.reduce((s, i) => s + (i - mx) * (y[i] - my), 0);
    const a = sxx ? sxy / sxx : 0, b = my - a * mx;
    const sst = idx.reduce((s, i) => s + (y[i] - my) ** 2, 0), sse = idx.reduce((s, i) => s + (y[i] - (a * i + b)) ** 2, 0);
    return {a, b, r2: sst ? 1 - sse / sst : null};
  }
  // forcar: meses já marcados à força (saldo herda os aberrantes da saída), sempre fora do ajuste
  function aberrantes(y, {ambos = false, elegivel = true, forcar = null} = {}) {
    const n = y.length, flag = forcar ? forcar.slice() : new Array(n).fill(false);
    let t = tsAjuste(y, flag);
    if (!elegivel || n < ABERR.minMeses) {
      const q0 = mqAjuste(y, flag);
      return {flag, a: q0.a, b: q0.b, ok: false, n: 0, r2: q0.r2};
    }
    for (let it = 0; it < 3; it++) {
      const fit = y.map((_, i) => t.a * i + t.b), res = y.map((v, i) => v - fit[i]);
      const base = res.filter((_, i) => !flag[i]);
      if (base.length < 3) break;
      const mb = mediana(base), mad = 1.4826 * mediana(base.map(v => Math.abs(v - mb)));
      const nivel = fit.filter((_, i) => !flag[i]).map(v => Math.max(v, 1));
      const esc = Math.max(mad, Math.sqrt(Math.max(nivel.reduce((s, v) => s + v, 0) / nivel.length, 1)));
      const novo = res.map((v, i) => v > ABERR.k * esc || (ambos && v < -ABERR.k * esc) || !!(forcar && forcar[i]));
      if (novo.every((v, i) => v === flag[i])) break;
      novo.forEach((v, i) => { flag[i] = v; });
      t = tsAjuste(y, flag);
    }
    for (let i = 0; i < n;) {
      if (!flag[i]) { i++; continue; }
      let j = i; while (j < n && flag[j]) j++;
      if (j - i > ABERR.maxSeq) for (let q = i; q < j; q++) if (!(forcar && forcar[q])) flag[q] = false;
      i = j;
    }
    const q = mqAjuste(y, flag);
    return {flag, a: q.a, b: q.b, ok: true, n: flag.filter(Boolean).length, r2: q.r2};
  }
  // série completa (entrada e saída por mês) → resultado só para os últimos 24 meses; ini = índice do 1º mês avaliado
  function analisarAberrantes(inn, out) {
    const ini = Math.max(0, inn.length - ABERR.janela);
    const i = inn.slice(ini), o = out.slice(ini), s = i.map((v, k) => v - o[k]);
    const eI = elegivelAberr(i), eO = elegivelAberr(o);
    const ro = aberrantes(o, {elegivel: eO});
    // regra de 24/10/2026: mês aberrante na saída é sempre aberrante no saldo
    return {ini, k: ABERR.k, in: aberrantes(i, {elegivel: eI}), out: ro,
      saldo: aberrantes(s, {ambos: true, elegivel: eI || eO, forcar: ro.flag})};
  }
  window._aberrantes = analisarAberrantes;

  // número com sinal explícito, 1 casa, sem "-0,0"
  const sinal = v => { const r = Math.round(v * 10) / 10 + 0; return (r >= 0 ? "+" : "") + fmt(r, 1); };

  function casa(padrao, txt) {
    const re = new RegExp("^" + padrao.split("*").map(s => s.replace(/[.+?^${}()|[\]\\]/g, "\\$&")).join(".*") + "$");
    return re.test(txt);
  }

  // ------------------------------------------------------------------------------ cálculos
  function calcular(d, nomeEscolhido) {
    const notaRef = REF.notas[d.p] || {};
    const r = {p: d.p, nome: nomeEscolhido || notaRef.nome || nomeCurto(d.p), avisos: [], notas: {}};
    r.extracao = new Date(d.data + "T12:00:00");
    r.solicAte = d.solicitacoes_ate;
    r.atraso = atrasoTxt(d);
    const saldo = d.in.map((v, i) => v - d.out[i]);
    r.serie = {meses: d.meses, in: d.in, out: d.out, saldo};
    r.aberr = analisarAberrantes(d.in, d.out);

    r.risco = CORES.map(([k, rot]) => {
      const c = d.cores[k];
      const cs = (d.cor_stats || {})[k] || {};
      return c ? {cor: rot, antiga: dd(c.antiga), n: c.n, dias: c.dias, media: cs.fila_media ?? null, mediana: cs.fila_mediana ?? null, atend: cs.atend_mediana ?? null, natend: cs.atend_n || 0} : {cor: rot, antiga: "", n: 0, dias: null, media: null, mediana: null, atend: cs.atend_mediana ?? null, natend: cs.atend_n || 0};
    });
    if (d.cores["sem risco"]) r.avisos.push(`${d.cores["sem risco"].n} solicitações da fila sem classificação de risco (fora da Figura 2).`);

    // Figura 3 – CIDs
    const cids = Object.entries(d.cid).map(([c, n]) => [c.toUpperCase(), n]);
    let grupos = REF.grupos_cid[d.p];
    if (!grupos) {
      const cat = {};
      cids.forEach(([c, n]) => { const k = c.slice(0, 3); cat[k] = (cat[k] || 0) + n; });
      const top = Object.entries(cat).sort((a, b) => b[1] - a[1]).slice(0, 3).map(x => x[0]);
      grupos = top.map(k => {
        const desc = (cids.find(([c]) => c.startsWith(k)) || [k])[0].split(" - ").slice(1).join(" - ") || k;
        return [k, `${k} - ${desc.length > 35 ? desc.slice(0, 35).replace(/[ ,]+$/, "") + "..." : desc}`];
      });
      r.avisos.push("Sem grupos de CID definidos para este procedimento: a Figura 3.1 usa as 3 categorias mais frequentes (" + top.join(", ") + ").");
    }
    const somaG = {}; grupos.forEach(([, g]) => somaG[g] = 0); somaG["OUTROS CIDS"] = 0;
    cids.forEach(([c, n]) => { const g = grupos.find(([pr]) => c.startsWith(pr)); somaG[g ? g[1] : "OUTROS CIDS"] += n; });
    const totC = Object.values(somaG).reduce((a, b) => a + b, 0) || 1;
    // grupos cadastrados sem nenhuma solicitação nesta fila saem da Figura 3.1 (OUTROS CIDS fica sempre)
    r.cids = Object.entries(somaG).filter(([g, n]) => n || g === "OUTROS CIDS").map(([g, n]) => ({grupo: g, n, pct: Math.round(n / totC * 1000) / 10}));
    const grupoDe = c => { const g = grupos.find(([pr]) => c.startsWith(pr)); return g ? g[1] : "OUTROS CIDS"; };
    r.cidsDetalhe = cids.sort((a, b) => b[1] - a[1]).map(([c, n]) => [c, n, grupoDe(c)]);
    // Tabela 1: 15 categorias (3 caracteres) mais frequentes; subcategorias registradas como x.1, x.2...
    const cats = {};
    r.cidsDetalhe.forEach(([c, n, g]) => {
      const k = c.slice(0, 3), cod = c.split(" - ")[0];
      const cat = cats[k] || (cats[k] = {k, n: 0, g, desc: null, subs: []});
      cat.n += n; cat.subs.push({cod, txt: c, n});
      if (cod === k) cat.desc = c;
    });
    const ord = Object.values(cats).sort((a, b) => b.n - a.n);
    let acum = 0;
    r.topCids = [];
    // as 15 categorias mais frequentes, agrupadas em sequência por grupo da Figura 3.1 (grupos do maior para o menor; OUTROS CIDS por último);
    // dentro do grupo, por frequência. O % acumulado segue a ordem em que a tabela é lida.
    const top = ord.slice(0, 15), somaGrupo = {};
    top.forEach(c => somaGrupo[c.g] = (somaGrupo[c.g] || 0) + c.n);
    top.sort((a, b) => (a.g === "OUTROS CIDS") - (b.g === "OUTROS CIDS") || somaGrupo[b.g] - somaGrupo[a.g] || cmp(a.g, b.g) || b.n - a.n);
    top.forEach((cat, i) => {
      acum += cat.n;
      const desc = cat.desc || `${cat.k} - ${(cat.subs[0].txt.split(" - ").slice(1).join(" - ") || "")} (categoria)`;
      r.topCids.push({num: `${i + 1}`, c: desc, n: cat.n, pct: cat.n / totC * 100, acum: acum / totC * 100, g: cat.g, cat: true});
      if (cat.subs.length > 1) cat.subs.sort((a, b) => b.n - a.n).forEach((s, j) =>
        r.topCids.push({num: `${i + 1}.${j + 1}`, c: s.cod === cat.k ? `${s.cod} - sem subcategoria` : s.txt, n: s.n, pct: s.n / totC * 100, acum: null, g: "", cat: false}));
    });
    r.restoCids = {qt: Math.max(ord.length - 15, 0), n: totC - acum};
    r.totFila = totC;

    // Figura 4 e Tabela 1 – devolvidas
    const nivel = u => { for (const [p, n, dest] of REF.niveis) if (casa(p, u)) return [n, dest]; return ["OUTROS", false]; };
    const dev = Object.entries(d.devolvidas_unidade);
    const totD = dev.reduce((s, [, n]) => s + n, 0);
    const outros = dev.filter(([u]) => nivel(u)[0] === "OUTROS").map(([u]) => u);
    if (outros.length) r.avisos.push("Unidades sem nível de atenção: " + outros.join("; "));
    const destacadas = dev.filter(([u, n]) => nivel(u)[1] && n / Math.max(totD, 1) >= 0.05).map(([u]) => u);
    const fatias = {}; ["APS", "DIRASE", ...destacadas, "HRS", "HRPL", "UPA", ...(outros.length ? ["OUTROS"] : [])].forEach(k => fatias[k] = 0);
    dev.forEach(([u, n]) => { fatias[destacadas.includes(u) ? u : nivel(u)[0]] += n; });
    r.niveis = Object.entries(fatias).map(([k, n]) => ({nivel: k, n, pct: Math.round(n / Math.max(totD, 1) * 100)}));
    r.notas.destacadas = destacadas;
    r.totDev = totD;
    r.devUnidade = dev.map(([u, n]) => ({u, nivel: destacadas.includes(u) ? u : nivel(u)[0], n})).sort((a, b) => b.n - a.n);

    const aps = dev.filter(([u]) => nivel(u)[0] === "APS").sort((a, b) => cmp(a[0], b[0]));
    const totA = aps.reduce((s, [, n]) => s + n, 0);
    const q12 = u => (d.dev12 && d.dev12.u[u]) || null;
    r.ubs = aps.map(([u, n]) => {
      const e = REF.esf[u], q = q12(u);
      return {u, n, pct: n / Math.max(totA, 1) * 100, esf: e ?? null, taxa: e ? n / e : null, d12: q ? q[0] : null, t12: q ? q[1] : null, pct12: q && q[1] ? q[0] / q[1] * 100 : null};
    }).sort((a, b) => (b.taxa ?? -1) - (a.taxa ?? -1) || b.n - a.n);  // ranking: quem mais devolve por eSF primeiro
    r.janela12 = d.dev12 ? d.dev12.janela : null;
    r.ag12 = d.ag12 || null;
    // total da APS: todas as UBS da referência (também as sem devolvidas), como no programa local
    const apsQ = d.dev12 ? Object.entries(d.dev12.u).filter(([u]) => REF.esf[u] != null) : [];
    r.d12Aps = apsQ.reduce((t, [, q]) => t + q[0], 0); r.t12Aps = apsQ.reduce((t, [, q]) => t + q[1], 0);
    // UBS sem devolvidas mas com solicitações no período também têm % (0%): entram só na nota
    r.notas.semDev12 = d.dev12 ? Object.entries(d.dev12.u).filter(([u, q]) => REF.esf[u] != null && q[1] >= MIN12 && !aps.some(([x]) => x === u)).map(([u, q]) => `${u} (0 de ${fmt(q[1])})`) : [];
    const semEsf = r.ubs.filter(x => x.esf == null).map(x => x.u);
    if (semEsf.length) r.avisos.push("Unidades da APS sem nº de eSF na referência: " + semEsf.join("; "));
    r.totAps = totA;
    r.esfTotal = Object.values(REF.esf).reduce((a, b) => a + b, 0);
    const vistas = new Set(d.unidades);
    r.notas.ausentes = Object.keys(REF.esf).filter(u => !vistas.has(u));
    r.notas.semDev = Object.keys(REF.esf).filter(u => vistas.has(u) && !aps.some(([x]) => x === u));
    r.notas.azul = notaRef.azul || "";
    r.meta = REF.metas && REF.metas[d.p] || null;
    r.risco.forEach(x => { const m = r.meta && r.meta[x.cor.toLowerCase()]; x.meta = m || null; x.acima = !!(m && x.atend != null && x.atend > m); });   // o protocolo calcula da inserção à execução
    const azul = r.risco[3].n;
    if (azul && r.notas.azul) r.avisos.push(`Há ${azul} solicitações AZUL na fila, embora a nota do procedimento diga que o protocolo não prevê azul.`);
    if (!r.notas.azul) r.avisos.push("Sem nota sobre a classificação de risco deste procedimento: confira no protocolo vigente se há classificação AZUL.");
    // Figura 1.1: resultado dos agendamentos por mês (agm = [agendamentos, faltas, sem confirmação, confirmados], pela data do agendamento)
    if (d.agm && d.agm_meses && d.agm[0].reduce((x, y) => x + y, 0) >= MIN12 && d.agm[3]) r.abs = {meses: d.agm_meses, ag: d.agm[0], falta: d.agm[1], pend: d.agm[2], conf: d.agm[3], ent: d.in.slice(-d.agm[0].length)};
    r.textos = textos(r);
    return r;
  }

  // resultado dos agendamentos: a saída efetiva é faixa porque parte deles nunca foi confirmada pelo executante
  function agTxt(a) {
    if (!a || !a.n) return "";
    const pc = x => fmt(x / a.n * 100);
    return `Dos ${fmt(a.n)} agendamentos dos últimos 12 meses, ${pc(a.conf)}% foram confirmados, ${pc(a.falta)}% foram faltas, ${pc(a.pend)}% seguem sem confirmação do executante e ${pc(a.canc)}% foram cancelados; a saída efetiva ficou entre ${fmt(a.conf / 12, 1)} e ${fmt((a.conf + a.pend) / 12, 1)} atendimentos/mês. `;
  }
  function textos(r) {
    const s = r.serie, n = s.in.length, t = {};
    const media = a => a.reduce((x, y) => x + y, 0) / Math.max(a.length, 1);
    // Médias: últimos 12 meses contra os 12 anteriores, e média dos últimos 24 meses (a projeção do painel usa a média de 12 meses)
    const u12 = a => a.slice(-12), a12 = a => a.slice(-24, -12), us = s.saldo.slice(-12);
    const ab = r.aberr, nj = n - ab.ini;
    const varPct = (x, y) => y ? (x / y - 1) * 100 : null;
    const comp = (nome, a) => {
      const m1 = media(u12(a)), m0 = media(a12(a)), pc = n >= 24 ? varPct(m1, m0) : null;
      return `${nome}: ${fmt(m1, 1)}/mês nos últimos 12 meses` + (n >= 24 ? ` contra ${fmt(m0, 1)}/mês nos 12 anteriores` + (pc != null ? ` (${pc >= 0 ? "+" : ""}${fmt(pc, 1)}%)` : "") : "");
    };
    const retaTxt = (nome, t) => `${nome} ${sinal(t.a)}/mês` + (t.r2 != null ? ` (R² ${fmt(t.r2, 2)})` : " (R² não calculável)");
    t.fig1 = `A média mensal dos últimos 24 meses (${s.meses[n - 24 < 0 ? 0 : n - 24]} a ${s.meses[n - 1]}) foi de ${fmt(media(s.in.slice(-24)), 1)} solicitações na entrada (INPUT), ${fmt(media(s.out.slice(-24)), 1)} na saída (OUTPUT) e ${sinal(media(s.saldo.slice(-24)))} no saldo regulatório. ` +
      `Comparando as médias anuais, ${comp("entrada", s.in)}; ${comp("saída", s.out)}. ` +
      `Nos últimos 12 meses, o saldo regulatório foi positivo em ${us.filter(x => x > 0).length} meses, somando ${fmt(us.reduce((a, b) => a + b, 0))} solicitações a mais do que as agendadas no período. ` + agTxt(r.ag12) +
      `Média e tendência respondem a perguntas diferentes: a média dos últimos 12 meses mostra o nível atual e é a base das projeções do painel; a linha de tendência dos últimos ${nj} meses (mínimos quadrados, sem os meses aberrantes) mostra a direção e não é usada para projetar: ${retaTxt("INPUT", ab.in)}, ${retaTxt("OUTPUT", ab.out)}, ${retaTxt("SALDO", ab.saldo)}. Quanto menor o R², menos a reta explica a oscilação mensal.`;
    const ext = [];
    for (const [k, a, chave] of [["INPUT", s.in, "in"], ["OUTPUT", s.out, "out"], ["SALDO REGULATÓRIO", s.saldo, "saldo"]]) {
      ab[chave].flag.forEach((x, i) => { if (x) ext.push(`${k} de ${s.meses[ab.ini + i]} (${fmt(a[ab.ini + i])})`); });
    }
    if (ext.length) t.fig1 += " Meses aberrantes (mais de 4 vezes a oscilação típica em relação à tendência; mantidos no gráfico e descontados da linha de tendência), a explicar (mutirão, suspensão de agenda etc.): " + ext.join("; ") + ".";
    {   // esforço de intervenção: média de 12 meses constante (mesma conta do plano do painel); cenário pior = entrada +18%, saída −25%
      const fila = r.risco.reduce((a, x) => a + x.n, 0);
      if (fila > 0 && n >= 12) {
        const ent = media(u12(s.in)), sai = media(u12(s.out)), up = v => fmt(Math.ceil(v - 1e-9) + 0);
        const est = Math.max(0, ent - sai), zer = Math.max(0, ent - sai + fila / 12);
        const entP = ent * 1.18, saiP = sai * 0.75, estP = Math.max(0, entP - saiP), zerP = Math.max(0, entP - saiP + fila / 12);
        t.fig1 += ` Esforço de intervenção (média dos últimos 12 meses mantida constante; é uma conta, não uma previsão): ` +
          (est > 0.05 ? `para a fila parar de crescer seriam necessários cerca de ${up(est)} agendamentos a mais por mês` + (sai > 0 ? ` (${fmt(est / sai * 100, 0)}% da saída atual)` : "") : `a saída atual já basta para a fila não crescer`) +
          `; para zerar as ${fmt(fila)} solicitações em 12 meses, cerca de ${up(zer)} a mais por mês durante o ano, e depois manter a oferta igual à entrada (${fmt(ent, 1)}/mês). ` +
          `Num cenário pior que o normal (entrada 18% maior e saída 25% menor, os erros típicos do teste com o passado das filas): ${up(estP)} a mais por mês para estabilizar e ${up(zerP)} para zerar em 12 meses.`;
      }
    }
    if (r.abs) {
      const A = r.abs, S = a => a.reduce((x, y) => x + y, 0), tt = S(A.ag), pc = x => fmt(x / tt * 100, 1);
      const m12 = a => S(a.slice(-12)), t12 = m12(A.ag);
      const pf = A.ag.map((q, i) => q >= MIN12 ? A.falta[i] / q * 100 : -1), im = pf.indexOf(Math.max(...pf));
      const P = projecaoAbs(A), f1 = v => fmt(Math.abs(v), 1);
      const efeito = P.d1 > 0 ? `faltariam entre ${f1(P.d1)} e ${f1(P.d2)} atendimentos por mês para dar conta do que entra` : P.d2 < 0 ? `os atendimentos superariam a entrada em ${f1(P.d2)} a ${f1(P.d1)} por mês` : `os atendimentos ficariam entre ${f1(P.d1)} abaixo e ${f1(P.d2)} acima da entrada, por mês`;
      t.fig1b = `Nos ${A.meses.length} meses de ${A.meses[0]} a ${A.meses[A.meses.length - 1]}, dos ${fmt(tt)} agendamentos, ${pc(S(A.falta))}% foram faltas (absenteísmo) e ${pc(S(A.pend))}% seguem sem confirmação do executante; ${pc(S(A.conf))}% foram efetivos (confirmados).` +
        (t12 >= MIN12 ? ` Nos últimos 12 meses (${fmt(t12)} agendamentos): ${fmt(m12(A.falta) / t12 * 100, 1)}% de faltas, ${fmt(m12(A.pend) / t12 * 100, 1)}% sem confirmação e ${fmt(m12(A.conf) / t12 * 100, 1)}% efetivos.` : "") +
        (im >= 0 && pf[im] >= 0 ? ` O mês com mais faltas foi ${A.meses[im]} (${fmt(pf[im], 1)}% de ${fmt(A.ag[im])} agendamentos).` : "") +
        ` Projeção para os próximos 12 meses, mantendo a média dos últimos 12: saída efetiva entre ${fmt(P.pMin, 1)} (só confirmados) e ${fmt(P.pMax, 1)} (confirmados e sem confirmação) por mês, contra entrada média de ${fmt(P.eMed, 1)}; nessa conta ${efeito}. A conta compara a entrada com atendimentos, enquanto o saldo regulatório da Figura 1 a compara com agendamentos (as faltas ocupam vagas e tiram a solicitação da fila, mas não são atendimentos).`;
    }
    const tot = r.risco.reduce((a, x) => a + x.n, 0);
    t.fig2 = `A fila tem ${fmt(tot)} solicitações pendentes ou reenviadas.` + (tot ? " Por prioridade, " + r.risco.filter(x => x.n).map(x =>
      `${x.cor.toLowerCase()}: ${fmt(x.n)} (${fmt(x.n / tot * 100, 1)}%), a mais antiga de ${x.antiga}` + (x.media != null ? `, com espera média de ${fmt(x.media)} dias e mediana de ${fmt(x.mediana)} dias` : "")).join("; ") + "." : "");
    if (r.meta && tot) {
      const acima = r.risco.filter(x => x.acima), ok = r.risco.filter(x => x.meta && x.atend != null && !x.acima);
      t.fig2 += acima.length ? " Em relação às metas do protocolo, calculadas da inserção à execução, a mediana até o atendimento ultrapassa a meta em " + acima.map(x => `${x.cor.toLowerCase()} (${fmt(x.atend)} dias, ${fmt(x.natend)} ${x.natend === 1 ? "caso" : "casos"}${x.natend < MIN12 ? ", pouco confiável" : ""}; meta de ${x.meta})`).join(", ") + "." : "";
      t.fig2 += ok.length ? ` Dentro da meta: ${ok.map(x => x.cor.toLowerCase()).join(", ")}.` : "";
    }
    t.fig3 = "Os principais grupos de CID são " + [...r.cids].sort((a, b) => b.n - a.n).filter(x => x.n).map(x => `${x.grupo} (${fmt(x.n)} solicitações; ${fmt(x.pct)}%)`).join(", ") + ".";
    const top = [...r.ubs].sort((a, b) => b.n - a.n).slice(0, 3);
    const tx = r.ubs.filter(x => x.taxa != null).sort((a, b) => b.taxa - a.taxa).slice(0, 3);
    t.fig4 = r.totDev ? `Das ${fmt(r.totDev)} solicitações devolvidas, ${fmt(r.niveis[0].n)} (${r.niveis[0].pct}%) vieram da APS.` +
      (top.length ? " Em números absolutos, concentram mais devoluções " + top.map(x => `${x.u} (${x.n})`).join(", ") +
        "; proporcionalmente ao número de eSF, " + tx.map(x => `${x.u} (${fmt(x.taxa, 1)} por eSF; ${fmt(x.esf)} eSF)`).join(", ") + "." : "")
      : "Não há solicitações devolvidas para este procedimento.";
    const ubs12 = r.ubs.filter(x => x.pct12 != null && x.t12 >= MIN12).sort((a, b) => b.pct12 - a.pct12).slice(0, 3);
    if (ubs12.length) t.fig4 += ` Em relação ao que cada UBS solicitou nos últimos 12 meses, o maior percentual de devolvidas está em ` + ubs12.map(x => `${x.u} (${fmt(x.pct12, 1)}%, ${fmt(x.d12)} de ${fmt(x.t12)})`).join(", ") + (r.t12Aps ? `; na APS como um todo, ${fmt(r.d12Aps / r.t12Aps * 100, 1)}%.` : ".");
    return t;
  }

  function notas(r) {
    const e = dataBr(r.extracao), s = r.serie, n = s.meses.length;
    const [m0, a0] = s.meses[0].split("/"), [m1, a1] = s.meses[n - 1].split("/");
    const n2 = [FONTE, `Nota: Dados extraídos em ${e}. Considera as solicitações pendentes e reenviadas; "mais antiga" é a data da solicitação mais antiga de cada prioridade; a média e a mediana são do tempo, em dias, que as solicitações dessa prioridade já esperaram até a data da extração. O Mapa Social registra solicitações até ${r.solicAte} (${r.atraso} em relação ao SISREG), por isso os totais podem diferir do SISREG.`];
    if (r.meta) n2.push("Metas de tempo de espera do protocolo vigente (mediana), comparadas com a mediana até o atendimento dos agendados nos últimos 90 dias: " + ["vermelho", "amarelo", "verde", "azul"].filter(c => r.meta[c]).map(c => `${c} até ${r.meta[c]} dias`).join(", ") + `. ${(r.meta.fonte || "").trim().replace(/\.$/, "")}.`);
    if (!r.meta) n2.push("Não há meta de tempo de espera por prioridade definida em normativa da SES-DF para este procedimento (levantamento de 14/10/2026 nos protocolos de acesso e notas técnicas vigentes).");
    if (r.notas.azul) n2.push("*" + r.notas.azul);
    let nt = `Nota: Percentuais calculados em relação às ${fmt(r.totAps)} solicitações devolvidas provenientes da Atenção Primária.`;
    if (r.notas.ausentes.length) nt += " O Mapa Social não apresenta solicitações deste procedimento da(s) " + r.notas.ausentes.join(", ") + ".";
    if (r.notas.semDev.length) nt += " UBS sem solicitações devolvidas: " + r.notas.semDev.join(", ") + ".";
    nt += ` UBS ordenadas da que mais devolve por eSF para a que menos devolve. Total de eSF da Região Norte: ${r.esfTotal} (referência: RH APS Norte, junho/2026).`;
    if (r.janela12) nt += ` "% ainda devolvidas hoje (12 m)" = das solicitações deste procedimento feitas pela UBS de ${r.janela12[0]} a ${r.janela12[1]} (pela data da solicitação), quantas estão devolvidas hoje; cada solicitação conta uma vez e a que foi reenviada e agendada conta como não devolvida, por isso o percentual é um piso. Janela de 12 meses para refletir a prática atual da unidade; percentuais de UBS com menos de ${MIN12} solicitações do procedimento no período aparecem em cinza. A coluna "Solicitações por eSF", em amarelo, é o critério de ordenação.` +
      (r.notas.semDev12.length ? ` UBS com ${MIN12} ou mais solicitações no período e nenhuma devolvida: ${r.notas.semDev12.join(", ")}.` : "");
    return {
      titulo: `Análise da fila de ${/^CONSULTA/i.test(r.p) ? "Consulta em " : ""}${r.nome} - Região Norte`,
      f1: `Figura 1 – Comportamento da fila de ${r.nome}`,
      n1: [FONTE, `Nota: Dados extraídos em ${e}, referentes ao período de ${MESES_EXT[MESES.indexOf(m0)]} de 20${a0} a ${MESES_EXT[MESES.indexOf(m1)]} de 20${a1} (a série começa em janeiro de cinco anos antes do ano corrente, como no modelo de análise). ` +
        `Média e tendência respondem a perguntas diferentes: a média diz quanto entra ou sai, em geral, por mês; a tendência diz para onde a curva está indo. Por exemplo, duas filas podem ter a mesma média de 100 solicitações por mês, uma estável e outra subindo 10 por mês. As médias mostradas no alto da figura são dos últimos 24 meses. A linha horizontal pontilhada de cada curva é a média dos últimos 12 meses, base das projeções do painel: no teste com o passado das filas da região, prever pela média errou menos (mediana de 18% na entrada e 26% na saída) do que prolongar a tendência (21% e 36%). A média olha para trás: se uma subida continuar de fato, ela fica abaixo do que vai acontecer (a fila do exemplo que sobe 10 por mês teria média de 100 em um ano e de 220 no seguinte). Por isso a tendência segue na figura como aviso: quando está bem inclinada para cima, a projeção é um piso, não o número certo. A média não tem R² por ser uma linha sem inclinação. As linhas de tendência (mínimos quadrados) cobrem os últimos ${n - r.aberr.ini} meses, mostram só a direção da curva e são calculadas sem os meses aberrantes (identificados pelo método de Theil-Sen: mais de 4 vezes a oscilação típica em relação à tendência; no saldo, acima ou abaixo, e sempre que a saída do mês for aberrante); o R² ao lado de cada linha considera só os meses não aberrantes, ` +
        "que permanecem no gráfico em tracejado e com círculo vazio. Quando o valor ultrapassa a escala, o ponto é desenhado no limite do eixo, com o valor real escrito ao lado (nesses meses a altura não é fiel). Meses anteriores à janela não foram avaliados."],
      f1b: `Figura 1.1 – Resultado dos agendamentos, entrada e projeção da saída efetiva em ${r.nome}`,
      n1b: [FONTE, `Nota: Dados extraídos em ${e}. Barras: agendamentos de cada mês (pela data do agendamento; um por código e data), por resultado; linha: solicitações (entrada) do mês. Efetivos = confirmados pelo executante (a pessoa foi atendida). Sem confirmação = o executante não informou se a pessoa foi atendida (a chave do agendamento não foi fechada); nos meses mais recentes o registro ainda pode chegar. Cancelados e outras situações completam o total.`, `Projeção: média dos últimos 12 meses mantida constante, a que errou menos no teste com o passado das filas; não prevê mudança de oferta nem de absenteísmo. A faixa da saída efetiva vai de só os confirmados até confirmados mais sem confirmação. O resumo da parte de baixo refere-se aos últimos 12 meses completos.`],
      f2: `Figura 2 – Tempo de espera por prioridade em ${r.nome}`, n2,
      f3: `Figura 3.1 – Distribuição de solicitações pendentes em ${r.nome} por Código Internacional de Doenças (CID)`,
      n3: [FONTE, `Nota: Dados extraídos em ${e}. Rótulos apresentam número absoluto e percentual em relação ao total de ${fmt(r.cids.reduce((t, x) => t + x.n, 0))} solicitações pendentes e reenviadas.`],
      f4: `Figura 4 – Proporção de solicitações devolvidas de ${r.nome} na Região Norte`,
      n4: [FONTE, `Nota: Dados extraídos em ${e}. Rótulos apresentam número absoluto e percentual em relação ao total de ${fmt(r.totDev)} solicitações devolvidas.` +
        (r.notas.destacadas.length ? " DIRASE inclui Policlínicas, CEO, NUFIRF e CAPS; unidades com 5% ou mais das devolvidas aparecem separadas." : "")],
      tc: `Figura 3.2 – Detalhamento das categorias de CID mais frequentes nas solicitações pendentes em ${r.nome}`,
      ntc: [FONTE, `Nota: Dados extraídos em ${e}. Percentuais em relação às ${fmt(r.totFila)} solicitações pendentes e reenviadas.` +
        " As 15 categorias do CID (3 caracteres) mais frequentes aparecem em negrito; abaixo de cada uma, recuadas, as subcategorias registradas." +
        (r.restoCids.qt ? ` As demais ${fmt(r.restoCids.qt)} categorias somam ${fmt(r.restoCids.n)} solicitações (${fmt(r.restoCids.n / Math.max(r.totFila, 1) * 100, 1)}%).` : "")],
      t1: `Tabela 1 – Distribuição das solicitações devolvidas de ${r.nome} na APS, por UBS`,
      nt1: [FONTE, nt],
      rascunho: "[Rascunho para revisão] ",
    };
  }

  // ------------------------------------------------------------------------------ gráficos (canvas)
  function canvas(w, h) {
    const c = document.createElement("canvas"); c.width = w; c.height = h;
    const g = c.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, w, h);
    g.font = "22px Calibri, Arial, sans-serif"; g.textBaseline = "middle";
    return [c, g];
  }
  function passo(f) { const p = 10 ** Math.floor(Math.log10(f || 1)); const m = f / p; return (m <= 1 ? 1 : m <= 2 ? 2 : m <= 5 ? 5 : 10) * p; }

  // Figura 1.1: barras empilhadas por mês (agendamentos por resultado), entrada em linha, projeção de 12 meses da saída efetiva
  // e, embaixo, o resumo dos últimos 12 meses (barra a 100% e faixa da saída efetiva). Devolve {url, w, h}.
  const COR_AG = [["Efetivos (confirmados)", "#15803d"], ["Sem confirmação do executante", "#2563eb"], ["Faltas (absenteísmo)", "#d97706"], ["Cancelados e outros", "#9ca3af"]];
  function projecaoAbs(A) {
    const m12 = a => a.slice(-12).reduce((x, y) => x + y, 0) / Math.max(1, Math.min(12, a.length));
    const pMin = m12(A.conf), pMax = m12(A.conf) + m12(A.pend), eMed = m12(A.ent);
    return {pMin, pMax, eMed, d1: eMed - pMax, d2: eMed - pMin};
  }
  function graficoAbs(r) {
    const A = r.abs, a12 = r.ag12 && r.ag12.n >= MIN12 ? r.ag12 : null, n = A.ag.length, H12 = 12, slots = n + H12;
    const W = 1400, H1 = 640, H = a12 ? 1010 : H1, L = 80, Rm = 20, T = 100, B = 120;
    const [c, g] = canvas(W, H), P = projecaoAbs(A);
    const ou = A.ag.map((q, i) => q - A.falta[i] - A.pend[i] - A.conf[i]);
    const top = Math.max(1, ...A.ag, ...A.ent, P.pMax, P.eMed), st = passo(top / 5), ymax = Math.ceil(top / st) * st;
    const Xc = i => L + (i + .5) / slots * (W - L - Rm), bw = (W - L - Rm) / slots * .68, Y = v => T + (ymax - v) / ymax * (H1 - T - B);
    const hat = (() => { const [pc, pg] = canvas(14, 14); pg.fillStyle = "#fff"; pg.fillRect(0, 0, 14, 14); pg.strokeStyle = "#15803d"; pg.lineWidth = 3; [-14, 0, 14].forEach(o => { pg.beginPath(); pg.moveTo(o, 14); pg.lineTo(o + 14, 0); pg.stroke(); }); return g.createPattern(pc, "repeat"); })();
    g.font = "20px Calibri, Arial, sans-serif"; g.textAlign = "left";
    let lx = L;
    const item = (t, desenha, ly) => { desenha(lx, ly); g.fillStyle = "#404040"; g.fillText(t, lx + 36, ly); lx += g.measureText(t).width + 80; };
    COR_AG.forEach(([t, cor]) => item(t, (x, y) => { g.fillStyle = cor; g.fillRect(x, y - 8, 26, 16); }, 26));
    lx = L;
    item("Entrada (solicitações)", (x, y) => { g.strokeStyle = "#000"; g.lineWidth = 3; g.beginPath(); g.moveTo(x, y); g.lineTo(x + 26, y); g.stroke(); }, 58);
    item("Projeção da saída efetiva", (x, y) => { g.fillStyle = hat; g.fillRect(x, y - 8, 26, 16); g.strokeStyle = "#15803d"; g.lineWidth = 1.5; g.strokeRect(x, y - 8, 26, 16); }, 58);
    g.strokeStyle = "#D9D9D9"; g.lineWidth = 1.2; g.fillStyle = "#595959"; g.textAlign = "right";
    for (let v = 0; v <= ymax + 1e-9; v += st) { g.beginPath(); g.moveTo(L, Y(v)); g.lineTo(W - Rm, Y(v)); g.stroke(); g.fillText(fmt(v), L - 10, Y(v)); }
    const [ml, al] = A.meses[n - 1].split("/");
    let mi = MESES.indexOf(ml), ai = +al;
    const rot = A.meses.slice();
    for (let j = 0; j < H12; j++) { mi++; if (mi > 11) { mi = 0; ai++; } rot.push(`${MESES[mi]}/${String(ai).padStart(2, "0")}`); }
    for (let i = 0; i < slots; i += 3) { g.save(); g.translate(Xc(i) + 6, H1 - B + 14); g.rotate(-Math.PI / 4); g.fillStyle = i >= n ? "#15803d" : "#595959"; g.fillText(rot[i], 0, 0); g.restore(); }
    for (let i = 0; i < n; i++) {
      let base = 0;
      [A.conf[i], A.pend[i], A.falta[i], ou[i]].forEach((v, k) => { if (v > 0) { g.fillStyle = COR_AG[k][1]; g.fillRect(Xc(i) - bw / 2, Y(base + v), bw, Y(base) - Y(base + v)); base += v; } });
    }
    const xs = Xc(n - .5);
    g.strokeStyle = "#7F7F7F"; g.lineWidth = 1.5; g.setLineDash([6, 5]); g.beginPath(); g.moveTo(xs, T - 12); g.lineTo(xs, H1 - B); g.stroke(); g.setLineDash([]);
    g.textAlign = "left"; g.fillStyle = "#15803d"; g.font = "bold 20px Calibri, Arial, sans-serif"; g.fillText("projeção: média de 12 meses, mantida", xs + 8, T - 14);
    for (let j = 0; j < H12; j++) {
      const x = Xc(n + j) - bw / 2;
      g.fillStyle = hat; g.fillRect(x, Y(P.pMin), bw, Y(0) - Y(P.pMin)); g.strokeStyle = "#15803d"; g.lineWidth = 1.5; g.strokeRect(x, Y(P.pMin), bw, Y(0) - Y(P.pMin));
      g.setLineDash([5, 4]); g.strokeRect(x, Y(P.pMax), bw, Y(P.pMin) - Y(P.pMax)); g.setLineDash([]);
    }
    g.strokeStyle = "#000"; g.lineWidth = 3; g.beginPath(); A.ent.forEach((v, i) => { if (i) g.lineTo(Xc(i), Y(v)); else g.moveTo(Xc(i), Y(v)); }); g.stroke();
    A.ent.forEach((v, i) => { g.beginPath(); g.arc(Xc(i), Y(v), 4.5, 0, 7); g.fillStyle = "#fff"; g.fill(); g.lineWidth = 2; g.stroke(); });
    g.lineWidth = 3; g.setLineDash([3, 7]); g.beginPath(); g.moveTo(Xc(n - 1), Y(P.eMed)); g.lineTo(Xc(slots - 1), Y(P.eMed)); g.stroke(); g.setLineDash([]);
    g.textAlign = "right"; g.fillStyle = "#000"; g.font = "bold 20px Calibri, Arial, sans-serif"; g.fillText(`entrada ${fmt(P.eMed, 1)}/mês`, W - Rm, Y(P.eMed) - 14);
    if (a12) {
      const y0 = H1 + 20, wb = 620, tot = a12.n, am = tot / 12, mn = a12.conf / 12, mx = (a12.conf + a12.pend) / 12;
      g.font = "bold 22px Calibri, Arial, sans-serif"; g.textAlign = "left"; g.fillStyle = "#262626";
      g.fillText(`Resultado dos ${fmt(tot)} agendamentos (12 meses)`, L, y0 + 10);
      const partes = [a12.conf, a12.pend, a12.falta, a12.canc + a12.outros];
      let x = L;
      partes.forEach((v, k) => {
        const w = v / tot * wb;
        if (v > 0) { g.fillStyle = COR_AG[k][1]; g.fillRect(x, y0 + 30, w, 50); if (w >= 50) { g.fillStyle = "#fff"; g.font = "bold 22px Calibri, Arial, sans-serif"; g.textAlign = "center"; g.fillText(fmt(v / tot * 100) + "%", x + w / 2, y0 + 56); } }
        x += w;
      });
      g.textAlign = "left"; g.font = "21px Calibri, Arial, sans-serif";
      partes.forEach((v, k) => { const yy = y0 + 112 + k * 34; g.fillStyle = COR_AG[k][1]; g.fillRect(L, yy - 9, 18, 18); g.fillStyle = "#262626"; g.fillText(`${COR_AG[k][0]}: ${fmt(v / tot * 100)}% (${fmt(v)})`, L + 30, yy); });
      const x2 = 760, X2 = v => x2 + v / am * wb;
      g.font = "bold 22px Calibri, Arial, sans-serif"; g.fillStyle = "#262626"; g.fillText("Saída efetiva (atendimentos/mês)", x2, y0 + 10);
      g.fillStyle = "#E5E7EB"; g.fillRect(x2, y0 + 56, wb, 40); g.fillStyle = "#15803d"; g.fillRect(x2, y0 + 56, X2(mn) - x2, 40);
      g.fillStyle = hat; g.fillRect(X2(mn), y0 + 56, X2(mx) - X2(mn), 40); g.strokeStyle = "#15803d"; g.lineWidth = 1.5; g.strokeRect(X2(mn), y0 + 56, X2(mx) - X2(mn), 40);
      g.fillStyle = "#15803d"; g.textAlign = "center"; g.font = "bold 21px Calibri, Arial, sans-serif";
      g.fillText(`${fmt(mn, 1)} a ${fmt(mx, 1)}`, Math.min(Math.max((X2(mn) + X2(mx)) / 2, x2 + 60), x2 + wb - 60), y0 + 40);
      g.font = "20px Calibri, Arial, sans-serif"; g.fillStyle = "#595959"; g.textAlign = "left"; g.fillText("0", x2, y0 + 120);
      g.textAlign = "right"; g.fillText(`${fmt(am, 1)} agendamentos/mês`, x2 + wb, y0 + 120);
      g.textAlign = "left"; g.fillStyle = "#262626"; g.fillText("Mínimo: só os confirmados. Máximo: confirmados + sem confirmação.", x2, y0 + 160);
    }
    return {url: c.toDataURL("image/png"), w: W, h: H};
  }

  function graficoSerie(r) {
    const W = 2200, H = 960, L = 80, Rm = 250, T = 80, B = 170;
    const [c, g] = canvas(W, H), s = r.serie, n = s.in.length, ab = r.aberr;
    const ser = [["INPUT", s.in, "#000000", ab.in], ["OUTPUT", s.out, "#00B050", ab.out], ["SALDO REGULATÓRIO", s.saldo, "#FF0000", ab.saldo]];
    // aberrantes (só na janela) não entram na escala; as retas de tendência entram
    const fl = ser.map(([, , , t]) => Array.from({length: n}, (_, i) => i >= ab.ini && t.flag[i - ab.ini]));
    const reta = (t, i) => t.a * (i - ab.ini) + t.b;
    const todos = [];
    ser.forEach(([, a, , t], j) => { a.forEach((v, i) => { if (!fl[j][i]) todos.push(v); }); todos.push(reta(t, ab.ini), reta(t, n - 1)); });
    const st = passo((Math.max(...todos) - Math.min(0, ...todos)) / 6);
    const ymin = Math.floor(Math.min(0, ...todos) / st) * st, ymax = Math.ceil(Math.max(1, ...todos) / st) * st;
    const X = i => L + i * (W - L - Rm) / (n - 1), Y = v => T + (ymax - v) / (ymax - ymin) * (H - T - B);
    const Yc = v => Y(Math.min(ymax, Math.max(ymin, v)));
    g.strokeStyle = "#D9D9D9"; g.lineWidth = 1.2; g.fillStyle = "#595959"; g.textAlign = "right";
    for (let v = ymin; v <= ymax + 1e-9; v += st) {
      g.beginPath(); g.moveTo(L, Y(v)); g.lineTo(W - Rm, Y(v)); g.stroke(); g.fillText(fmt(v), L - 12, Y(v));
    }
    g.textAlign = "right";
    for (let i = 0; i < n; i += 2) { g.save(); g.translate(X(i) + 6, H - B + 18); g.rotate(-Math.PI / 4); g.fillText(s.meses[i], 0, 0); g.restore(); }
    ser.forEach(([, a, cor], j) => {
      g.strokeStyle = cor; g.fillStyle = cor; g.lineWidth = 4;
      for (let i = 1; i < n; i++) {
        g.setLineDash(fl[j][i] || fl[j][i - 1] ? [12, 9] : []);
        g.beginPath(); g.moveTo(X(i - 1), Yc(a[i - 1])); g.lineTo(X(i), Yc(a[i])); g.stroke();
      }
      g.setLineDash([]);
      a.forEach((v, i) => {
        g.beginPath(); g.arc(X(i), Yc(v), fl[j][i] ? 8 : 5.5, 0, 7);
        if (fl[j][i]) { g.fillStyle = "#fff"; g.fill(); g.lineWidth = 3; g.stroke(); g.fillStyle = cor; g.lineWidth = 4; } else g.fill();
        if (fl[j][i] && (v > ymax || v < ymin)) {
          g.font = "bold 20px Calibri, Arial, sans-serif"; g.textAlign = "center"; g.fillStyle = cor;
          g.fillText((v > ymax ? "▲ " : "▼ ") + fmt(v), X(i), Yc(v) + (v > ymax ? 30 + (i % 3) * 22 : -16 - (i % 3) * 22)); g.font = "22px Calibri, Arial, sans-serif";
        }
      });
    });
    g.textAlign = "left"; g.font = "20px Calibri, Arial, sans-serif";
    const ty = [];
    ser.forEach(([, , cor, t]) => {
      g.strokeStyle = cor; g.lineWidth = 1.6; g.beginPath();
      g.moveTo(X(ab.ini), Y(reta(t, ab.ini))); g.lineTo(X(n - 1), Y(reta(t, n - 1))); g.stroke();
      let y = Y(reta(t, n - 1)) - 14; while (ty.some(z => Math.abs(z - y) < 24)) y += 24; ty.push(y);
      g.fillStyle = cor; g.fillText(sinal(t.a) + "/mês" + (t.r2 != null ? " · R² " + fmt(t.r2, 2) : ""), X(n - 1) + 10, y);
    });
    // média dos últimos 12 meses (a que a projeção usa): linha horizontal pontilhada sobre os 12 meses que ela resume
    ser.forEach(([, a, cor]) => {
      const m12 = a.slice(-12).reduce((x, y) => x + y, 0) / Math.max(1, Math.min(12, a.length)), y = Yc(m12);
      g.strokeStyle = cor; g.lineWidth = 2.2; g.setLineDash([3, 7]); g.beginPath(); g.moveTo(X(Math.max(0, n - 12)), y); g.lineTo(X(n - 1), y); g.stroke(); g.setLineDash([]);
    });
    g.font = "20px Calibri, Arial, sans-serif";
    // legenda das médias mensais no alto do gráfico
    const media = a => a.reduce((x, y) => x + y, 0) / Math.max(a.length, 1);
    const j24 = Math.max(0, n - 24), m24 = a => media(a.slice(j24));
    const mm = [["INPUT médio", m24(s.in), "#000000"], ["OUTPUT médio", m24(s.out), "#00B050"],
      ["SALDO REGULATÓRIO médio", m24(s.saldo), "#FF0000"]];
    g.font = "bold 22px Calibri, Arial, sans-serif";
    const tit = `Média mensal dos últimos 24 meses (${s.meses[j24]} a ${s.meses[n - 1]}):`;
    const m12v = [media(s.in.slice(-12)), media(s.out.slice(-12)), media(s.saldo.slice(-12))];
    const partes = mm.map(([t, v], i) => `${t}: ${t.startsWith("SALDO") ? sinal(v) : fmt(v, 1)} (12 m: ${t.startsWith("SALDO") ? sinal(m12v[i]) : fmt(m12v[i], 1)})`);
    const lt = g.measureText(tit).width + 24 + partes.reduce((a, p) => a + g.measureText(p).width + 90, 0);
    let lx = (W - lt) / 2; const ly = 36;
    g.fillStyle = "#404040"; g.fillText(tit, lx, ly); lx += g.measureText(tit).width + 24;
    partes.forEach((p, i) => {
      g.strokeStyle = mm[i][2]; g.lineWidth = 4; g.beginPath(); g.moveTo(lx, ly); g.lineTo(lx + 36, ly); g.stroke();
      g.fillStyle = "#404040"; g.fillText(p, lx + 46, ly); lx += g.measureText(p).width + 90;
    });
    g.font = "20px Calibri, Arial, sans-serif";
    const leg = [["INPUT", "#000", true], ["OUTPUT", "#00B050", true], ["SALDO REGULATÓRIO", "#FF0000", true],
      ["Linear (INPUT)", "#000", false], ["Linear (OUTPUT)", "#00B050", false], ["Linear (SALDO REGULATÓRIO)", "#FF0000", false],
      ["Média dos últimos 12 meses", "#595959", "m12"],
      ["Mês aberrante", "#7F7F7F", "ab"]];
    g.font = "20px Calibri, Arial, sans-serif";
    const larg = leg.map(([t]) => g.measureText(t).width + 70), totL = larg.reduce((a, b) => a + b, 0);
    let x = (W - totL) / 2; const y = H - 28;
    leg.forEach(([t, cor, grossa], i) => {
      g.strokeStyle = cor; g.lineWidth = grossa === "m12" ? 2.2 : grossa ? 4 : 1.6; g.setLineDash(grossa === "ab" ? [8, 6] : grossa === "m12" ? [3, 7] : []); g.beginPath(); g.moveTo(x, y); g.lineTo(x + 44, y); g.stroke(); g.setLineDash([]);
      if (grossa === "ab") { g.fillStyle = "#fff"; g.lineWidth = 3; g.beginPath(); g.arc(x + 22, y, 8, 0, 7); g.fill(); g.stroke(); }
      else if (grossa && grossa !== "m12") { g.fillStyle = cor; g.beginPath(); g.arc(x + 22, y, 5.5, 0, 7); g.fill(); }
      g.fillStyle = "#404040"; g.fillText(t, x + 52, y); x += larg[i];
    });
    return c.toDataURL("image/png");
  }

  function pizza(valores, rotulos, cores, rotulo) {
    const W = 1280, H = 920, cx = 640, cy = 470, R = 300;
    const [c, g] = canvas(W, H), tot = valores.reduce((a, b) => a + b, 0) || 1;
    let ang = -Math.PI / 2, peq = 0;
    valores.forEach((v, i) => {
      if (!v) return;
      const a = v / tot * Math.PI * 2;
      g.fillStyle = cores[i]; g.beginPath(); g.moveTo(cx, cy); g.arc(cx, cy, R, ang, ang + a); g.closePath(); g.fill();
      g.strokeStyle = "#fff"; g.lineWidth = 4; g.stroke();
      const m = ang + a / 2, dist = v / tot > 0.06 ? 0.72 : 1.12 + 0.15 * (peq++ % 3);  // fatias pequenas: rótulo fora, em até 3 alturas (cabe no topo da imagem)
      g.fillStyle = "#000"; g.font = "bold 26px Calibri, Arial, sans-serif"; g.textAlign = "center";
      g.fillText(rotulo(v, tot), cx + Math.cos(m) * R * dist, cy + Math.sin(m) * R * dist);
      ang += a;
    });
    g.font = "22px Calibri, Arial, sans-serif"; g.textAlign = "left";
    const larg = rotulos.map(t => g.measureText(t).width + 56);
    const linhas = []; let atual = [], w = 0;
    larg.forEach((l, i) => { if (w + l > W - 60 && atual.length) { linhas.push(atual); atual = []; w = 0; } atual.push(i); w += l; });
    linhas.push(atual);
    linhas.forEach((ls, k) => {
      let x = (W - ls.reduce((s, i) => s + larg[i], 0)) / 2; const y = H - 40 - (linhas.length - 1 - k) * 36;
      ls.forEach(i => { g.fillStyle = cores[i]; g.fillRect(x, y - 10, 34, 20); g.fillStyle = "#404040"; g.fillText(rotulos[i], x + 42, y); x += larg[i]; });
    });
    return c.toDataURL("image/png");
  }
  const PAL_CID = ["#4472C4", "#ED7D31", "#FF0000", "#7F7F7F", "#70AD47", "#7030A0", "#00B0F0", "#9E480E", "#264478"];
  const coresCids = r => [...PAL_CID.slice(0, r.cids.length - 1), "#FFC000"];
  const corGrupo = r => Object.fromEntries(r.cids.map((x, i) => [x.grupo, coresCids(r)[i]]));
  // versão clara da cor do grupo (78% de branco) para o fundo das linhas de categoria da Figura 3.2
  const tom = hex => { const n = parseInt(hex.slice(1), 16), f = c => Math.round(255 - (255 - c) * 0.22); return [f(n >> 16), f((n >> 8) & 255), f(n & 255)]; };
  const tomHex = hex => tom(hex).map(v => v.toString(16).padStart(2, "0")).join("").toUpperCase();
  const graficoCids = r => pizza(r.cids.map(x => x.n), r.cids.map(x => x.grupo), coresCids(r), (v, t) => `${fmt(v)} (${Math.round(v / t * 100)}%)`);
  function graficoNiveis(r) {
    const base = {APS: "#70AD47", DIRASE: "#2F5597", HRS: "#FF0000", HRPL: "#C00000", UPA: "#FFC000", OUTROS: "#7F7F7F"};
    const extra = ["#7030A0", "#00B0F0", "#A5A5A5", "#264478"]; let k = 0;
    return pizza(r.niveis.map(x => x.n), r.niveis.map(x => x.nivel), r.niveis.map(x => base[x.nivel] || extra[k++ % 4]),
      (v, t) => `${v}; ${Math.round(v / t * 100)}%`);
  }

  // ------------------------------------------------------------------------------ PDF
  // chrome = cabeçalho e rodapé institucionais em todas as páginas (false gera o documento sem eles)
  async function gerarPdf(r, chrome = true) {
    for (const u of LIBS.pdf) await lib(u);
    const {jsPDF} = window.jspdf, N = notas(r);
    const doc = new jsPDF({orientation: "landscape", unit: "mm", format: "a4"});
    let y = chrome ? 22 : 18, W = 297, M = 20;
    const LIM = () => (W > 250 ? 200 : 285) - (chrome ? 6 : 0);
    const texto = (t, tam = 12, opt = {}) => {
      doc.setFont("times", opt.negrito ? "bold" : "normal"); doc.setFontSize(tam);
      let ls = doc.splitTextToSize(t, W - 2 * M);
      const h = tam * 0.42;
      if (y + ls.length * h > LIM()) { novaPag(); doc.setFont("times", opt.negrito ? "bold" : "normal"); doc.setFontSize(tam); ls = doc.splitTextToSize(t, W - 2 * M); }   // refaz a quebra: a página nova é retrato (largura menor)
      ls.forEach(l => {
        if (opt.realce) { doc.setFillColor(255, 255, 0); doc.rect(opt.centro ? (W - doc.getTextWidth(l)) / 2 : M, y - h * 0.8, doc.getTextWidth(l), h + 0.6, "F"); }
        doc.text(l, opt.centro ? W / 2 : M, y, {align: opt.centro ? "center" : "left"}); y += h;
      });
      y += opt.depois ?? 1;
    };
    const novaPag = () => { doc.addPage("a4", "portrait"); W = 210; M = 25; y = chrome ? 25 : 22; };
    const imagem = (url, larg, alt) => { if (y + alt > LIM()) novaPag(); doc.addImage(url, "PNG", (W - larg) / 2, y, larg, alt, undefined, "FAST"); y += alt + 3; };   // sem "FAST" o PNG vai descomprimido (PDF de ~18 MB)
    const rodape = ls => { ls.forEach((l, i) => texto(l, 9, {depois: i === ls.length - 1 ? 5 : 0})); };
    const garantir = h => { if (y + h > LIM()) novaPag(); };  // mantém título junto da figura/tabela
    const rascunho = t => texto(N.rascunho + t, 11, {realce: true, depois: 7});

    texto(N.titulo, 12, {negrito: true, centro: true, depois: 6});
    texto(N.f1, 12, {depois: 2});
    imagem(graficoSerie(r), 250, 109.1);
    rodape(N.n1); rascunho(r.textos.fig1);
    if (r.abs) {   // Figura 1.1 (subitem: o modelo não é renumerado)
      if (W > 250) novaPag();
      const gi = graficoAbs(r), alt = 160 * gi.h / gi.w;
      garantir(alt + 12); texto(N.f1b, 12, {depois: 2});
      imagem(gi.url, 160, alt);
      rodape(N.n1b); rascunho(r.textos.fig1b);
    }

    if (W > 250) novaPag(); else y += 3;   // se o texto da Figura 1 já passou para a página retrato, a Figura 2 segue nela
    garantir(70); texto(N.f2, 12, {depois: 3});
    const corFill = {Vermelho: [255, 0, 0], Amarelo: [255, 255, 0], Verde: [146, 208, 80], Azul: [0, 176, 240]};
    doc.autoTable({
      startY: y, margin: {left: M, right: M}, theme: "grid", styles: {font: "helvetica", fontSize: 10, halign: "center", textColor: 0, lineColor: 120, lineWidth: 0.2},
      head: [["Classificação", "Mais antiga", "Solicitações", "Espera média (dias)", "Espera mediana (dias)"]],
      headStyles: {fillColor: 255, textColor: 0, fontStyle: "bold"},
      body: r.risco.map(x => [x.cor, x.antiga, x.n ? fmt(x.n) : "", x.media != null && x.n ? fmt(x.media) : "", x.mediana != null && x.n ? fmt(x.mediana) : ""]),
      foot: [["TOTAL", "", fmt(r.risco.reduce((a, x) => a + x.n, 0)), "", ""]],
      footStyles: {fillColor: 255, textColor: 0, fontStyle: "bold"}, showFoot: "lastPage",
      didParseCell: h => { if (h.section === "body") h.cell.styles.fillColor = corFill[h.row.raw[0]]; },
    });
    y = doc.lastAutoTable.finalY + 4;
    rodape(N.n2); rascunho(r.textos.fig2);
    garantir(115); texto(N.f3, 12, {depois: 2});
    imagem(graficoCids(r), 140, 100.6);
    rodape(N.n3);
    garantir(45); texto(N.tc, 12, {depois: 3});
    const cg31 = corGrupo(r);
    doc.autoTable({
      startY: y, margin: {left: M, right: M}, theme: "grid",
      styles: {font: "times", fontSize: 8.5, textColor: 0, lineColor: 120, lineWidth: 0.2, cellPadding: 0.8},
      headStyles: {fillColor: 255, textColor: 0, fontStyle: "bold", halign: "center", fontSize: 9.5},
      columnStyles: {0: {halign: "left", cellWidth: 72}, 1: {halign: "right"}, 2: {halign: "right"}, 3: {halign: "right"}, 4: {halign: "left"}},
      head: [["CID", "Solicitações", "%", "% acum.", "Grupo na Figura 3.1"]],
      body: r.topCids.map(x => [x.cat ? x.c : "      " + x.c, fmt(x.n), fmt(x.pct, 1) + "%", x.acum != null ? fmt(x.acum, 1) + "%" : "", x.g]),
      didParseCell: h => { const x = r.topCids[h.row.index]; if (h.section === "body" && x.cat) { h.cell.styles.fontStyle = "bold"; h.cell.styles.fillColor = cg31[x.g] ? tom(cg31[x.g]) : 242; } },
    });
    y = doc.lastAutoTable.finalY + 4;
    rodape(N.ntc); rascunho(r.textos.fig3);
    garantir(115); texto(N.f4, 12, {depois: 2});
    imagem(graficoNiveis(r), 140, 100.6);
    rodape(N.n4);
    garantir(45); texto(N.t1, 12, {depois: 3});
    doc.autoTable({
      startY: y, margin: {left: M, right: M}, theme: "grid",
      styles: {font: "times", fontSize: 9, textColor: 0, lineColor: 120, lineWidth: 0.2, cellPadding: 0.8},
      headStyles: {fillColor: 255, textColor: 0, fontStyle: "bold", halign: "center", fontSize: 10},
      columnStyles: {0: {halign: "left", cellWidth: 40}, 1: {halign: "right"}, 2: {halign: "right"}, 3: {halign: "right"}, 4: {halign: "right"}, 5: {halign: "right"}, 6: {halign: "right"}},
      head: [["UNIDADE", "Número de solicitações", "Percentual da APS", "Quantidade de eSF", "Solicitações por eSF", "Solicitações (12 m)", "% ainda devolvidas hoje (12 m)"]],
      body: r.ubs.map(x => [x.u, fmt(x.n), fmt(x.pct, 2) + "%", x.esf != null ? fmt(x.esf) : "", x.taxa != null ? fmt(x.taxa, 2) : "", x.t12 != null ? fmt(x.t12) : "", x.pct12 != null ? fmt(x.pct12, 1) + "%" : ""]),
      foot: [["TOTAL", fmt(r.totAps), "100,00%", fmt(r.esfTotal), fmt(r.totAps / r.esfTotal, 2), r.t12Aps ? fmt(r.t12Aps) : "", r.t12Aps ? fmt(r.d12Aps / r.t12Aps * 100, 1) + "%" : ""]],
      footStyles: {fillColor: 255, textColor: 0, fontStyle: "normal", halign: "right"}, showFoot: "lastPage",
      didParseCell: h => { if (h.column.index === 4 && h.section !== "head") h.cell.styles.fillColor = [255, 255, 204];   // critério de ordenação em amarelo claro
        if (h.section === "body" && h.column.index === 6 && (r.ubs[h.row.index].t12 || 0) < MIN12) h.cell.styles.textColor = 140; },
    });
    y = doc.lastAutoTable.finalY + 4;
    rodape(N.nt1); rascunho(r.textos.fig4);
    if (chrome) {
      const np = doc.getNumberOfPages(), ext = r.extracao.toLocaleDateString("pt-BR");
      for (let i = 1; i <= np; i++) {
        doc.setPage(i); const pw = doc.internal.pageSize.getWidth(), ph = doc.internal.pageSize.getHeight(), mm = pw > 250 ? 20 : 25;
        doc.setFont("times", "normal"); doc.setFontSize(9); doc.setTextColor(60);
        doc.text(CAB_ESQ, mm, 10);
        doc.text(CAB_DIR, pw - mm, 10, {align: "right"});
        doc.setDrawColor(120); doc.setLineWidth(0.3); doc.line(mm, 12.5, pw - mm, 12.5);
        doc.line(mm, ph - 11, pw - mm, ph - 11);
        doc.setFontSize(8.5);
        doc.text(`${N.titulo}  ·  dados extraídos em ${ext}  ·  versão piloto`, mm, ph - 7);
        doc.text(`Página ${i} de {total}`, pw - mm, ph - 7, {align: "right"});
      }
      doc.setTextColor(0);
      if (typeof doc.putTotalPages === "function") doc.putTotalPages("{total}");
    }
    doc.save(N.titulo + ".pdf");
  }

  // ------------------------------------------------------------------------------ Word
  async function gerarWord(r) {
    for (const u of LIBS.word) await lib(u);
    const D = window.docx, N = notas(r);
    const bytes = url => Uint8Array.from(atob(url.split(",")[1]), ch => ch.charCodeAt(0));
    const run = (t, o = {}) => new D.TextRun({text: t, font: "Times New Roman", size: (o.tam || 12) * 2, bold: o.negrito, highlight: o.realce ? "yellow" : undefined});
    const par = (t, o = {}) => new D.Paragraph({children: Array.isArray(t) ? t : [run(t, o)], alignment: o.centro ? D.AlignmentType.CENTER : o.just ? D.AlignmentType.JUSTIFIED : undefined,
      spacing: {after: (o.depois ?? 0) * 20, before: (o.antes ?? 0) * 20}});
    const img = (url, w, h) => new D.Paragraph({alignment: D.AlignmentType.CENTER, children: [new D.ImageRun({data: bytes(url), transformation: {width: w, height: h}})]});
    const rodape = ls => ls.map((l, i) => par(l, {tam: 9, depois: i === ls.length - 1 ? 12 : 0}));
    const rascunho = t => par([run(N.rascunho, {tam: 11, negrito: true, realce: true}), run(t, {tam: 11, realce: true})], {just: true, depois: 12});
    const cel = (t, o = {}) => new D.TableCell({children: [new D.Paragraph({alignment: o.al ?? D.AlignmentType.CENTER,
      children: [new D.TextRun({text: t, font: o.fonte || "Calibri", size: (o.tam || 11) * 2, bold: o.negrito, color: o.cor})]})],
      shading: o.fill ? {type: D.ShadingType.CLEAR, color: "auto", fill: o.fill} : undefined, width: o.w ? {size: o.w, type: D.WidthType.DXA} : undefined});
    const fill = Object.fromEntries(CORES.map(([, rot, f]) => [rot, f]));
    const tRisco = new D.Table({alignment: D.AlignmentType.CENTER, rows: [
      new D.TableRow({children: ["Classificação", "Mais antiga", "Solicitações", "Espera média (dias)", "Espera mediana (dias)"].map(t => cel(t, {negrito: true}))}),
      ...r.risco.map(x => new D.TableRow({children: [x.cor, x.antiga, x.n ? fmt(x.n) : "", x.media != null && x.n ? fmt(x.media) : "", x.mediana != null && x.n ? fmt(x.mediana) : ""].map(t => cel(t, {fill: fill[x.cor]}))})),
      new D.TableRow({children: [cel("TOTAL", {negrito: true, al: D.AlignmentType.LEFT}), cel(""), cel(fmt(r.risco.reduce((a, x) => a + x.n, 0)), {negrito: true}), cel(""), cel("")]}),
    ], columnWidths: [1700, 1600, 1600, 1900, 1900]});
    const R_ = D.AlignmentType.RIGHT, L_ = D.AlignmentType.LEFT, o9 = {fonte: "Times New Roman", tam: 9};
    const tUbs = new D.Table({alignment: D.AlignmentType.CENTER, columnWidths: [2350, 1150, 1150, 1050, 1150, 1100, 1100], rows: [
      new D.TableRow({tableHeader: true, children: ["UNIDADE", "Número de solicitações", "Percentual da APS", "Quantidade de eSF", "Solicitações por eSF", "Solicitações (12 m)", "% ainda devolvidas hoje (12 m)"].map(t => cel(t, {fonte: "Times New Roman", tam: 9, negrito: true}))}),
      ...r.ubs.map(x => new D.TableRow({children: [cel(x.u, {...o9, al: L_}), cel(fmt(x.n), {...o9, al: R_}), cel(fmt(x.pct, 2) + "%", {...o9, al: R_}),
        cel(x.esf != null ? fmt(x.esf) : "", {...o9, al: R_}), cel(x.taxa != null ? fmt(x.taxa, 2) : "", {...o9, al: R_, fill: "FFFFCC"}), cel(x.t12 != null ? fmt(x.t12) : "", {...o9, al: R_}), cel(x.pct12 != null ? fmt(x.pct12, 1) + "%" : "", {...o9, al: R_, cor: (x.t12 || 0) < MIN12 ? "8C8C8C" : undefined})]})),
      new D.TableRow({children: [cel("TOTAL", {...o9, al: L_}), cel(fmt(r.totAps), {...o9, al: R_}), cel("100,00%", {...o9, al: R_}), cel(fmt(r.esfTotal), {...o9, al: R_}), cel(fmt(r.totAps / r.esfTotal, 2), {...o9, al: R_, fill: "FFFFCC"}),
        cel(r.t12Aps ? fmt(r.t12Aps) : "", {...o9, al: R_}), cel(r.t12Aps ? fmt(r.d12Aps / r.t12Aps * 100, 1) + "%" : "", {...o9, al: R_})]}),
    ]});
    const cgw = corGrupo(r);   // cor de cada grupo da Figura 3.1, para o fundo das linhas de categoria da Figura 3.2
    const tCid = new D.Table({alignment: D.AlignmentType.CENTER, columnWidths: [3700, 1400, 800, 1000, 2200], rows: [
      new D.TableRow({tableHeader: true, children: ["CID", "Solicitações", "%", "% acum.", "Grupo na Figura 3.1"].map(t => cel(t, {fonte: "Times New Roman", tam: 10, negrito: true}))}),
      ...r.topCids.map(x => { const o = {...o9, negrito: x.cat, fill: x.cat ? (cgw[x.g] ? tomHex(cgw[x.g]) : "F2F2F2") : undefined};
        return new D.TableRow({children: [cel(x.cat ? x.c : "      " + x.c, {...o, al: L_}), cel(fmt(x.n), {...o, al: R_}),
          cel(fmt(x.pct, 1) + "%", {...o, al: R_}), cel(x.acum != null ? fmt(x.acum, 1) + "%" : "", {...o, al: R_}), cel(x.g, {...o, al: L_})]}); }),
    ]});
    // cabeçalho e rodapé institucionais; w = largura útil da página em DXA, para o texto à direita
    const ext = r.extracao.toLocaleDateString("pt-BR"), tab = () => (typeof D.Tab === "function" ? new D.Tab() : "\t");
    const linha = (lado, cor = "808080") => ({[lado]: {style: D.BorderStyle.SINGLE, size: 4, color: cor, space: 2}});
    const cab = w => new D.Header({children: [new D.Paragraph({tabStops: [{type: D.TabStopType.RIGHT, position: w}], border: linha("bottom"),
      children: [new D.TextRun({text: CAB_ESQ, font: "Times New Roman", size: 18}), new D.TextRun({font: "Times New Roman", size: 18, children: [tab(), CAB_DIR]})]})]});
    const rod = w => new D.Footer({children: [new D.Paragraph({tabStops: [{type: D.TabStopType.RIGHT, position: w}], border: linha("top"),
      children: [new D.TextRun({text: `${N.titulo}  ·  dados extraídos em ${ext}  ·  versão piloto`, font: "Times New Roman", size: 17}),
        new D.TextRun({font: "Times New Roman", size: 17, children: [tab(), "Página ", D.PageNumber.CURRENT, " de ", D.PageNumber.TOTAL_PAGES]})]})]});
    const gi = r.abs ? graficoAbs(r) : null;
    const doc = new D.Document({sections: [
      {properties: {page: {size: {orientation: D.PageOrientation.LANDSCAPE, width: 11906, height: 16838}, margin: {top: 1134, bottom: 1134, left: 1134, right: 1134, header: 500, footer: 500}}},
        headers: {default: cab(14450)}, footers: {default: rod(14450)},
        children: [par(N.titulo, {negrito: true, centro: true, depois: 18}), par(N.f1, {depois: 4}), img(graficoSerie(r), 940, 410), ...rodape(N.n1), rascunho(r.textos.fig1)]},
      {properties: {page: {size: {orientation: D.PageOrientation.PORTRAIT, width: 11906, height: 16838}, margin: {top: 1247, bottom: 1247, left: 1417, right: 1417, header: 500, footer: 500}}},
        headers: {default: cab(8950)}, footers: {default: rod(8950)},
        children: [...(r.abs ? [par(N.f1b, {depois: 4}), img(gi.url, 600, 600 * gi.h / gi.w), ...rodape(N.n1b), rascunho(r.textos.fig1b)] : []), par(N.f2, {depois: 6}), tRisco, ...rodape(N.n2), rascunho(r.textos.fig2),
          par(N.f3, {depois: 4}), img(graficoCids(r), 540, 388), ...rodape(N.n3), new D.Paragraph({pageBreakBefore: true, spacing: {after: 120}, children: [run(N.tc)]}), tCid, ...rodape(N.ntc), rascunho(r.textos.fig3),
          par(N.f4, {depois: 4}), img(graficoNiveis(r), 540, 388), ...rodape(N.n4),
          par(N.t1, {antes: 6, depois: 6}), tUbs, ...rodape(N.nt1), rascunho(r.textos.fig4)]},
    ]});
    baixar(await D.Packer.toBlob(doc), N.titulo + ".docx");
  }

  // ------------------------------------------------------------------------------ Excel
  async function gerarExcel(r) {
    for (const u of LIBS.excel) await lib(u);
    const X = window.XLSX, wb = X.utils.book_new(), N = notas(r);
    const aba = (nome, linhas, larg) => { const ws = X.utils.aoa_to_sheet(linhas); ws["!cols"] = larg.map(w => ({wch: w})); X.utils.book_append_sheet(wb, ws, nome); };
    aba("Leia-me", [["Informação", "Valor"], ["Procedimento", r.p], ["Região", "Região Norte"], ["Data da extração", r.extracao.toLocaleDateString("pt-BR")],
      ["Solicitações registradas até", r.solicAte], ["Fonte", "Mapa Social do Distrito Federal (MPDFT) – dados agregados, sem identificação de usuários"],
      ["Gerado pelo painel", location.href.split("#")[0]]], [30, 90]);
    aba("Saldo Regulatório", [["Mês/Ano", "INPUT", "OUTPUT", "SALDO REGULATÓRIO", "Mês aberrante em"], ...r.serie.meses.map((m, i) => [m, r.serie.in[i], r.serie.out[i], r.serie.saldo[i],
      [["INPUT", "in"], ["OUTPUT", "out"], ["SALDO", "saldo"]].filter(([, k]) => i >= r.aberr.ini && r.aberr[k].flag[i - r.aberr.ini]).map(([t]) => t).join("; ")])], [10, 10, 10, 20, 24]);
    if (r.abs) aba("Agendamentos por mês", [["Mês/Ano", "Agendamentos", "Efetivos (confirmados)", "Sem confirmação do executante", "Faltas", "Cancelados e outros", "% faltas", "Entrada (solicitações)"],
      ...r.abs.meses.map((m, i) => [m, r.abs.ag[i], r.abs.conf[i], r.abs.pend[i], r.abs.falta[i], r.abs.ag[i] - r.abs.conf[i] - r.abs.pend[i] - r.abs.falta[i], r.abs.ag[i] ? Math.round(r.abs.falta[i] / r.abs.ag[i] * 1000) / 10 : "", r.abs.ent[i]])], [10, 14, 20, 28, 10, 20, 10, 20]);
    if (r.ag12) aba("Agendamentos 12 meses", [["Situação do agendamento", "Agendamentos", "%"],
      ...[["Confirmado", "conf"], ["Falta", "falta"], ["Sem confirmação do executante", "pend"], ["Cancelado", "canc"], ["Outra situação", "outros"]].map(([t, k]) => [t, r.ag12[k], Math.round(r.ag12[k] / r.ag12.n * 1000) / 10]),
      ["TOTAL", r.ag12.n, 100]], [34, 14, 8]);
    aba("Tempo de espera por risco", [["Classificação", "Mais antiga", "Solicitações", "Espera média (dias)", "Espera mediana (dias)", "Espera da mais antiga (dias)", "Mediana até o atendimento (dias)", "Atendimentos (90 dias)", "Meta do protocolo (dias, mediana)", "Mediana acima da meta?"],
      ...r.risco.map(x => [x.cor, dataXl(x.antiga), x.n, x.n ? x.media ?? "" : "", x.n ? x.mediana ?? "" : "", x.dias ?? "", x.atend ?? "", x.natend || "", x.meta ?? "", x.meta && x.atend != null ? (x.acima ? "Sim" : "Não") : ""]),
      ["TOTAL", "", r.risco.reduce((a, x) => a + x.n, 0), "", "", "", "", ""]], [14, 12, 12, 22, 20, 22, 22, 14]);
    aba("Figura 3.1 CIDs grupos", [["Grupo", "Solicitações", "%"], ...r.cids.map(x => [x.grupo, x.n, x.pct])], [45, 12, 8]);
    aba("Figura 3.2 CIDs detalhe", [["Nível", "CID", "Solicitações", "%", "% acumulado", "Grupo na Figura 3.1"],
      ...r.topCids.map(x => [x.cat ? "Categoria" : "Subcategoria", x.c, x.n, Math.round(x.pct * 10) / 10, x.acum != null ? Math.round(x.acum * 10) / 10 : "", x.g])], [13, 80, 12, 8, 12, 36]);
    aba("CIDs (todos)", [["CID", "Solicitações pendentes/reenviadas", "Grupo na Figura 3.1"], ...r.cidsDetalhe], [80, 30, 40]);
    aba("Devolvidas por nível", [["Nível", "Solicitações", "%"], ...r.niveis.map(x => [x.nivel, x.n, x.pct])], [24, 12, 8]);
    aba("Devolvidas por unidade", [["Unidade solicitante", "Nível / fatia", "Solicitações devolvidas"], ...r.devUnidade.map(x => [x.u, x.nivel, x.n])], [32, 22, 22]);
    aba("Devolvidas APS por UBS", [["UNIDADE", "Número de solicitações", "Percentual da APS", "Quantidade de eSF", "Solicitações por eSF", "Solicitações (12 m)", "Devolvidas (12 m)", "% ainda devolvidas hoje (12 m)"],
      ...r.ubs.map(x => [x.u, x.n, Math.round(x.pct * 100) / 100, x.esf ?? "", x.taxa != null ? Math.round(x.taxa * 100) / 100 : "", x.t12 ?? "", x.d12 ?? "", x.pct12 != null ? Math.round(x.pct12 * 10) / 10 : ""]),
      ["TOTAL", r.totAps, 100, r.esfTotal, Math.round(r.totAps / r.esfTotal * 100) / 100, r.t12Aps || "", r.d12Aps || "", r.t12Aps ? Math.round(r.d12Aps / r.t12Aps * 1000) / 10 : ""]], [26, 22, 18, 18, 20, 18, 16, 18]);
    aba("Notas", [[N.titulo], [""], ...[N.n1, N.n2, N.n3, N.ntc, N.n4, N.nt1].flat().map(t => [t])], [160]);
    aba("Avisos", [["Avisos para revisão"], ...(r.avisos.length ? r.avisos : ["Nenhum"]).map(a => [a])], [140]);
    X.writeFile(wb, `Bases e tabelas - ${r.nome} (${r.extracao.toLocaleDateString("pt-BR").replace(/\//g, "-")}).xlsx`);
  }

  function baixar(blob, nome) {
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = nome;
    document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
  }

  // analises/<id>.json não traz a média/mediana de espera por prioridade (cor_stats): vem do dados.json
  async function carregarAnalise(id, cache) {
    const d = await (await fetch(`analises/${id}.json`, cache ? {cache} : undefined)).json();
    try { const dj = await (await fetch("dados.json", cache ? {cache} : undefined)).json(); const fj = dj.filas.find(x => x.id === id) || {}; d.cor_stats = fj.cor_stats || null; d.ag12 = fj.ag12 || null; d.agm = fj.agm || null; d.agm_meses = dj.meses || null; } catch { d.cor_stats = null; d.ag12 = null; d.agm = null; d.agm_meses = null; }
    // % de devolvidas nos últimos 12 meses por unidade (Tabela 1): devolvidas.json, d12[id] = [[índice da unidade, devolvidas, solicitações], ...]
    try { const dv = await (await fetch("devolvidas.json", cache ? {cache} : undefined)).json(); const u = {};
      (dv.d12 && dv.d12[id] || []).forEach(([ui, dd_, tt]) => { u[dv.unidades[ui][0]] = [dd_, tt]; }); d.dev12 = {janela: dv.janela12 || null, u}; } catch { d.dev12 = null; }
    return d;
  }
  window.gerarAnalise = async function (id, formato, nome) {
    if (!REF) REF = await (await fetch("referencias.json", {cache: "no-store"})).json();
    const d = await carregarAnalise(id, "no-store");
    const r = calcular(d, nome);
    await ({pdf: gerarPdf, word: gerarWord, excel: gerarExcel})[formato](r);
    return r;
  };
  // ------------------------------------------------------------------------------ dados detalhados (linhas do Mapa Social)
  // Datas vão como número de série do Excel (Date do JS desloca um dia por causa do fuso histórico de Brasília)
  const dataNum = s => { const [d, m, a] = (s || "").split("/").map(Number); return a ? (Date.UTC(a, m - 1, d) - Date.UTC(1899, 11, 30)) / 864e5 : 0; };
  const dataXl = s => dataNum(s) ? {t: "n", v: dataNum(s), z: "dd/mm/yyyy"} : (s || "");
  const nivelDe = u => { for (const [p, n] of REF.niveis) if (casa(p, u)) return n; return "OUTROS"; };
  const COR_TXT = {vermelho: "Vermelho", amarelo: "Amarelo", verde: "Verde", azul: "Azul"};
  function folha(X, linhas, larg) {
    const ws = X.utils.aoa_to_sheet(linhas);
    ws["!cols"] = larg.map(w => ({wch: w}));
    if (linhas.length > 1) ws["!autofilter"] = {ref: X.utils.encode_range({s: {r: 0, c: 0}, e: {r: linhas.length - 1, c: larg.length - 1}})};
    return ws;
  }
  function leiaMe(d, extra) {
    return [["Informação", "Valor"], ["Fonte", "Mapa Social do Distrito Federal (MPDFT) – painel público de regulação"],
      ["Região", "Região Norte (unidade solicitante)"], ["Data da extração", dd(d.data)],
      ["Solicitações registradas até", `${d.solicitacoes_ate} (${atrasoTxt(d)} em relação ao SISREG)`],
      ...extra, ["Identificação", "Sem nome, telefone ou CNS. O código da solicitação é o código SISREG, público no Mapa Social."],
      ["Gerado pelo painel", location.href.split("#")[0]]];
  }
  window.baixarDetalhe = async function (id, nome) {
    for (const u of LIBS.excel) await lib(u);
    if (!REF) REF = await (await fetch("referencias.json", {cache: "no-store"})).json();
    const r = await fetch(`detalhe/${id}.json`, {cache: "no-store"});
    if (!r.ok) throw new Error("dados detalhados ainda não publicados para esta fila");
    const d = await r.json(), X = window.XLSX, wb = X.utils.book_new();
    const p = d.p || (await (await fetch(`analises/${id}.json`)).json()).p;
    nome = nome || (REF.notas[p] && REF.notas[p].nome) || nomeCurto(p);
    const S = d.s.slice().sort((a, b) => cmp(a[2], b[2]) || dataNum(a[1]) - dataNum(b[1]));
    X.utils.book_append_sheet(wb, folha(X, [["Código SISREG da solicitação", "Data da solicitação", "Status", "Classificação de risco",
      "CID", "Unidade solicitante", "Nível de atenção", "Data da devolução", "Dias desde a solicitação"],
      ...S.map(([c, ds, st, cor, cid, u, dv]) => [c, dataXl(ds), st, COR_TXT[cor] || "", cid, u, u ? nivelDe(u) : "", dataXl(dv),
        /PENDENTE|REENVIADA/.test(st) && ds ? Math.round(dataNum(dd(d.data)) - dataNum(ds)) : ""])],
      [16, 12, 24, 12, 60, 34, 12, 12, 12]), "Solicitações");
    const A = d.a.slice().sort((a, b) => dataNum(b[2]) - dataNum(a[2]));
    X.utils.book_append_sheet(wb, folha(X, [["Código SISREG da solicitação", "Data da solicitação", "Data do agendamento", "Status do agendamento",
      "Classificação de risco", "Unidade solicitante"], ...A.map(([c, ds, da, st, cor, u]) => [c, dataXl(ds), dataXl(da), st, COR_TXT[cor] || "", u])],
      [16, 12, 12, 36, 12, 34]), "Agendamentos");
    X.utils.book_append_sheet(wb, folha(X, leiaMe(d, [["Procedimento", p],
      ["Aba Solicitações", "Pendentes e reenviadas ao regulador (fila), devolvidas e em fila de espera (status PENDENTE / FILA DE ESPERA), uma linha por solicitação."],
      ["Aba Agendamentos", `Agendamentos com data a partir de ${dd(d.agendamentos_desde)}, inclusive os marcados à frente, com cancelados e faltas (mesma regra do OUTPUT).`]]), [30, 110]), "Leia-me");
    X.writeFile(wb, `Dados detalhados - ${nome} (${dd(d.data).replace(/\//g, "-")}).xlsx`);
    return {s: S.length, a: A.length};
  };
  let DEV_DET = null;
  // filtro: {unidades?: [...], procs?: [...], tipo?: "", nivel?: "", titulo}; porId: mapa id -> fila do painel (para o tipo)
  window.baixarDevolvidas = async function (filtro, porId) {
    for (const u of LIBS.excel) await lib(u);
    if (!REF) REF = await (await fetch("referencias.json", {cache: "no-store"})).json();
    if (!DEV_DET) { const r = await fetch("detalhe/devolvidas.json", {cache: "no-store"}); if (!r.ok) throw new Error("devolvidas detalhadas ainda não publicadas"); DEV_DET = await r.json(); }
    const tipoDe = {}; Object.values(porId || {}).forEach(f => tipoDe[f.p] = f.tipo);
    const us = filtro.unidades && new Set(filtro.unidades), ps = filtro.procs && new Set(filtro.procs);
    const L = DEV_DET.r.filter(([pi, , , , , u]) => {
      const p = DEV_DET.p[pi];
      return (!us || us.has(u)) && (!ps || ps.has(p)) && (!filtro.tipo || tipoDe[p] === filtro.tipo) && (!filtro.nivel || nivelDe(u) === filtro.nivel);
    }).sort((a, b) => cmp(a[5], b[5]) || cmp(DEV_DET.p[a[0]], DEV_DET.p[b[0]]) || dataNum(a[2]) - dataNum(b[2]));
    if (!L.length) throw new Error("nenhuma devolvida com esse filtro");
    const X = window.XLSX, wb = X.utils.book_new();
    X.utils.book_append_sheet(wb, folha(X, [["Unidade solicitante", "Nível de atenção", "Procedimento", "Tipo", "Código SISREG da solicitação",
      "Data da solicitação", "Data da devolução", "Classificação de risco", "CID"],
      ...L.map(([pi, c, ds, cor, cid, u, dv]) => [u, nivelDe(u), DEV_DET.p[pi], tipoDe[DEV_DET.p[pi]] || "", c, dataXl(ds), dataXl(dv), COR_TXT[cor] || "", cid])],
      [34, 12, 55, 10, 16, 12, 12, 12, 60]), "Devolvidas");
    X.utils.book_append_sheet(wb, folha(X, leiaMe(DEV_DET, [["Conteúdo", "Solicitações devolvidas pelo regulador (status DEVOLVIDA / REGULADOR), uma linha por solicitação, sem recorte de ano."],
      ["Filtro", [filtro.unidades && filtro.unidades.length <= 5 ? "Unidade(s): " + filtro.unidades.join("; ") : filtro.unidades ? `${filtro.unidades.length} unidades` : "",
        filtro.procs ? "Procedimento: " + filtro.procs.join("; ") : "", filtro.tipo ? "Tipo: " + filtro.tipo : "", filtro.nivel ? "Nível: " + filtro.nivel : ""].filter(Boolean).join(" · ") || "nenhum"]]), [30, 110]), "Leia-me");
    X.writeFile(wb, `${filtro.titulo} (${dd(DEV_DET.data).replace(/\//g, "-")}).xlsx`.replace(/[\\/:*?"<>|]/g, "-"));
    return L.length;
  };

  fetch("referencias.json", {cache: "no-store"}).then(x => x.json()).then(j => { REF = REF || j; if (window._aoCarregarRef) window._aoCarregarRef(); }).catch(() => {});
  window._analiseMeta = p => REF && REF.metas && REF.metas[p] || null;
  window._analiseNomeCurto = p => (REF && REF.notas[p] && REF.notas[p].nome) || nomeCurto(p);
})();

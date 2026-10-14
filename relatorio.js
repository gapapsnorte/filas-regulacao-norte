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
  const carregadas = {};
  let REF = null;

  const fmt = (n, c = 0) => Number(n).toLocaleString("pt-BR", {minimumFractionDigits: c, maximumFractionDigits: c});
  const dataBr = d => `${d.getDate()} ${MESES[d.getMonth()]}. ${d.getFullYear()}`;
  const dd = s => s ? s.split("-").reverse().join("/") : "";
  const cmp = (a, b) => (a < b ? -1 : a > b ? 1 : 0);
  const atrasoTxt = d => d.atraso_dias != null ? `atraso de ${d.atraso_dias} dias` : "atraso de 6 a 7 dias";

  function lib(url) {
    if (!carregadas[url]) carregadas[url] = new Promise((ok, erro) => {
      const s = document.createElement("script");
      s.src = url; s.onload = ok; s.onerror = () => erro(new Error("Falha ao carregar " + url));
      document.head.appendChild(s);
    });
    return carregadas[url];
  }

  function nomeCurto(p) {
    let t = p.trim().replace(/^(CONSULTA|EXAME|PROCEDIMENTO) EM /i, "").replace(/\s*-\s*/g, " - ");
    const minus = new Set(["e", "de", "da", "do", "das", "dos", "em", "com", "sem", "para", "a", "o"]);
    return t.toLowerCase().split(/\s+/).map((w, i) => (i && minus.has(w)) ? w : w.charAt(0).toUpperCase() + w.slice(1)).join(" ");
  }

  function tendencia(y) {
    const n = y.length, x = [...y.keys()], mx = (n - 1) / 2, my = y.reduce((a, b) => a + b, 0) / n;
    let sxy = 0, sxx = 0;
    x.forEach(i => { sxy += (i - mx) * (y[i] - my); sxx += (i - mx) ** 2; });
    const a = sxx ? sxy / sxx : 0, b = my - a * mx;
    let ssr = 0, sst = 0;
    y.forEach((v, i) => { ssr += (v - (a * i + b)) ** 2; sst += (v - my) ** 2; });
    return {a, b, r2: sst ? 1 - ssr / sst : 0};
  }

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

    r.risco = CORES.map(([k, rot]) => {
      const c = d.cores[k];
      return c ? {cor: rot, antiga: dd(c.antiga), n: c.n, dias: c.dias} : {cor: rot, antiga: "", n: 0, dias: null};
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
    r.cids = Object.entries(somaG).map(([g, n]) => ({grupo: g, n, pct: Math.round(n / totC * 1000) / 10}));
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
    ord.slice(0, 15).forEach((cat, i) => {
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
    r.ubs = aps.map(([u, n]) => {
      const e = REF.esf[u];
      return {u, n, pct: n / Math.max(totA, 1) * 100, esf: e ?? null, taxa: e ? n / e : null};
    }).sort((a, b) => (b.taxa ?? -1) - (a.taxa ?? -1) || b.n - a.n);  // ranking: quem mais devolve por eSF primeiro
    const semEsf = r.ubs.filter(x => x.esf == null).map(x => x.u);
    if (semEsf.length) r.avisos.push("Unidades da APS sem nº de eSF na referência: " + semEsf.join("; "));
    r.totAps = totA;
    r.esfTotal = Object.values(REF.esf).reduce((a, b) => a + b, 0);
    const vistas = new Set(d.unidades);
    r.notas.ausentes = Object.keys(REF.esf).filter(u => !vistas.has(u));
    r.notas.semDev = Object.keys(REF.esf).filter(u => vistas.has(u) && !aps.some(([x]) => x === u));
    r.notas.azul = notaRef.azul || "";
    r.meta = REF.metas && REF.metas[d.p] || null;
    r.risco.forEach(x => { const m = r.meta && r.meta[x.cor.toLowerCase()]; x.meta = m || null; x.acima = !!(m && x.dias != null && x.dias > m); });
    const azul = r.risco[3].n;
    if (azul && r.notas.azul) r.avisos.push(`Há ${azul} solicitações AZUL na fila, embora a nota do procedimento diga que o protocolo não prevê azul.`);
    if (!r.notas.azul) r.avisos.push("Sem nota sobre a classificação de risco deste procedimento: confira no protocolo vigente se há classificação AZUL.");
    r.textos = textos(r);
    return r;
  }

  function textos(r) {
    const s = r.serie, n = s.in.length, t = {};
    const media = a => a.reduce((x, y) => x + y, 0) / Math.max(a.length, 1);
    const p12 = s.in.slice(0, 12), u12 = s.in.slice(-12), us = s.saldo.slice(-12);
    const v = media(p12) ? (media(u12) / media(p12) - 1) * 100 : 0;
    const ai = tendencia(s.in).a, ao = tendencia(s.out).a;
    t.fig1 = `Entre ${s.meses[0]} e ${s.meses[n - 1]}, a média mensal de entrada (INPUT) passou de ${fmt(media(p12), 1)} solicitações nos primeiros 12 meses para ${fmt(media(u12), 1)} nos últimos 12 meses (${v >= 0 ? "+" : ""}${fmt(v, 1)}%). ` +
      `Nos últimos 12 meses, o saldo regulatório foi positivo em ${us.filter(x => x > 0).length} meses, somando ${fmt(us.reduce((a, b) => a + b, 0))} solicitações a mais do que as agendadas no período. ` +
      `A tendência linear indica variação média de ${fmt(ai * 12, 1)} solicitações/mês por ano no INPUT e de ${fmt(ao * 12, 1)} no OUTPUT.`;
    const ext = [];
    for (const [k, a] of [["INPUT", s.in], ["OUTPUT", s.out]]) {
      const m = media(a), dp = Math.sqrt(a.reduce((x, y) => x + (y - m) ** 2, 0) / Math.max(a.length - 1, 1));
      a.forEach((y, i) => {
        if (y > m + 2 * dp) ext.push(`${k} de ${s.meses[i]} (${y}) muito acima da média (${fmt(m)})`);
        else if (y < m - 2 * dp) ext.push(`${k} de ${s.meses[i]} (${y}) muito abaixo da média (${fmt(m)})`);
      });
    }
    if (ext.length) t.fig1 += " Meses fora do padrão, a explicar (mutirão, suspensão de agenda etc.): " + ext.join("; ") + ".";
    const tot = r.risco.reduce((a, x) => a + x.n, 0);
    t.fig2 = `A fila tem ${fmt(tot)} solicitações pendentes ou reenviadas.` + (tot ? " Por prioridade, " + r.risco.filter(x => x.n).map(x =>
      `${x.cor.toLowerCase()}: ${fmt(x.n)} (${fmt(x.n / tot * 100, 1)}%), a mais antiga de ${x.antiga}, com ${fmt(x.dias)} dias de espera`).join("; ") + "." : "");
    if (r.meta && tot) {
      const acima = r.risco.filter(x => x.n && x.acima), ok = r.risco.filter(x => x.n && x.meta && !x.acima);
      t.fig2 += acima.length ? " Em relação às metas do protocolo, a espera máxima ultrapassa a meta em " + acima.map(x => `${x.cor.toLowerCase()} (${fmt(x.dias)} dias; meta de ${x.meta})`).join(", ") + "." : "";
      t.fig2 += ok.length ? ` Dentro da meta: ${ok.map(x => x.cor.toLowerCase()).join(", ")}.` : "";
    }
    t.fig3 = "Os principais grupos de CID são " + [...r.cids].sort((a, b) => b.n - a.n).filter(x => x.n).map(x => `${x.grupo} (${fmt(x.pct)}%)`).join(", ") + ".";
    const top = [...r.ubs].sort((a, b) => b.n - a.n).slice(0, 3);
    const tx = r.ubs.filter(x => x.taxa != null).sort((a, b) => b.taxa - a.taxa).slice(0, 3);
    t.fig4 = r.totDev ? `Das ${fmt(r.totDev)} solicitações devolvidas, ${fmt(r.niveis[0].n)} (${r.niveis[0].pct}%) vieram da APS.` +
      (top.length ? " Em números absolutos, concentram mais devoluções " + top.map(x => `${x.u} (${x.n})`).join(", ") +
        "; proporcionalmente ao número de eSF, " + tx.map(x => `${x.u} (${fmt(x.taxa, 1)} por eSF)`).join(", ") + "." : "")
      : "Não há solicitações devolvidas para este procedimento.";
    return t;
  }

  function notas(r) {
    const e = dataBr(r.extracao), s = r.serie, n = s.meses.length;
    const [m0, a0] = s.meses[0].split("/"), [m1, a1] = s.meses[n - 1].split("/");
    const n2 = [FONTE, `Nota: Dados extraídos em ${e}. Considera as solicitações pendentes e reenviadas; o tempo de espera vai da solicitação mais antiga de cada prioridade até a data da extração. O Mapa Social registra solicitações até ${r.solicAte} (${r.atraso} em relação ao SISREG), por isso os totais podem diferir do SISREG.`];
    if (r.meta) n2.push("Metas de tempo de espera do protocolo vigente (mediana): " + ["vermelho", "amarelo", "verde", "azul"].filter(c => r.meta[c]).map(c => `${c} até ${r.meta[c]} dias`).join(", ") + `. ${(r.meta.fonte || "").trim().replace(/\.$/, "")}.`);
    if (r.notas.azul) n2.push("*" + r.notas.azul);
    let nt = `Nota: Percentuais calculados em relação às ${fmt(r.totAps)} solicitações devolvidas provenientes da Atenção Primária.`;
    if (r.notas.ausentes.length) nt += " O Mapa Social não apresenta solicitações deste procedimento da(s) " + r.notas.ausentes.join(", ") + ".";
    if (r.notas.semDev.length) nt += " UBS sem solicitações devolvidas: " + r.notas.semDev.join(", ") + ".";
    nt += ` UBS ordenadas da que mais devolve por eSF para a que menos devolve. Total de eSF da Região Norte: ${r.esfTotal}.`;
    return {
      titulo: `Análise da fila de ${/^CONSULTA/i.test(r.p) ? "Consulta em " : ""}${r.nome} - Região Norte`,
      f1: `Figura 1 – Comportamento da fila de ${r.nome}`,
      n1: [FONTE, `Nota: Dados extraídos em ${e}, referentes ao período de ${MESES_EXT[MESES.indexOf(m0)]} de 20${a0} a ${MESES_EXT[MESES.indexOf(m1)]} de 20${a1}.`],
      f2: `Figura 2 – Tempo de espera por prioridade em ${r.nome}`, n2,
      f3: `Figura 3.1 – Distribuição de solicitações pendentes em ${r.nome} por Código Internacional de Doenças (CID)`,
      n3: [FONTE, `Nota: Dados extraídos em ${e}.`],
      f4: `Figura 4 – Proporção de solicitações devolvidas de ${r.nome} na Região Norte`,
      n4: [FONTE, `Nota: Dados extraídos em ${e}. Rótulos apresentam número absoluto e percentual em relação ao total de ${fmt(r.totDev)} solicitações devolvidas.` +
        (r.notas.destacadas.length ? " DIRASE inclui Policlínicas, CEO, NUFIRF e CAPS; unidades com 5% ou mais das devolvidas aparecem separadas." : "")],
      tc: `Figura 3.2 – Detalhamento das categorias de CID mais frequentes nas solicitações pendentes em ${r.nome}`,
      ntc: [FONTE, `Nota: Dados extraídos em ${e}. Percentuais em relação às ${fmt(r.totFila)} solicitações pendentes e reenviadas.` +
        " As 15 categorias do CID (3 caracteres) mais frequentes aparecem em negrito; abaixo de cada uma, recuadas, as subcategorias registradas." +
        (r.restoCids.qt ? ` As demais ${fmt(r.restoCids.qt)} categorias somam ${fmt(r.restoCids.n)} solicitações (${fmt(r.restoCids.n / Math.max(r.totFila, 1) * 100, 1)}%).` : "") +
""],
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

  function graficoSerie(r) {
    const W = 2200, H = 920, L = 80, Rm = 150, T = 30, B = 170;
    const [c, g] = canvas(W, H), s = r.serie, n = s.in.length;
    const todos = [...s.in, ...s.out, ...s.saldo];
    const st = passo((Math.max(...todos) - Math.min(0, ...todos)) / 6);
    const ymin = Math.floor(Math.min(0, ...todos) / st) * st, ymax = Math.ceil(Math.max(1, ...todos) / st) * st;
    const X = i => L + i * (W - L - Rm) / (n - 1), Y = v => T + (ymax - v) / (ymax - ymin) * (H - T - B);
    g.strokeStyle = "#D9D9D9"; g.lineWidth = 1.2; g.fillStyle = "#595959"; g.textAlign = "right";
    for (let v = ymin; v <= ymax + 1e-9; v += st) {
      g.beginPath(); g.moveTo(L, Y(v)); g.lineTo(W - Rm, Y(v)); g.stroke(); g.fillText(fmt(v), L - 12, Y(v));
    }
    g.textAlign = "right";
    for (let i = 0; i < n; i += 2) { g.save(); g.translate(X(i) + 6, H - B + 18); g.rotate(-Math.PI / 4); g.fillText(s.meses[i], 0, 0); g.restore(); }
    const ser = [["INPUT", s.in, "#000000"], ["OUTPUT", s.out, "#00B050"], ["SALDO REGULATÓRIO", s.saldo, "#FF0000"]];
    for (const [, a, cor] of ser) {
      g.strokeStyle = cor; g.fillStyle = cor; g.lineWidth = 4; g.beginPath();
      a.forEach((v, i) => i ? g.lineTo(X(i), Y(v)) : g.moveTo(X(i), Y(v))); g.stroke();
      a.forEach((v, i) => { g.beginPath(); g.arc(X(i), Y(v), 5.5, 0, 7); g.fill(); });
    }
    g.textAlign = "left"; g.font = "20px Calibri, Arial, sans-serif";
    const r2y = [];
    for (const [, a, cor] of ser) {
      const t = tendencia(a); g.strokeStyle = cor; g.lineWidth = 1.6; g.beginPath();
      g.moveTo(X(0), Y(t.b)); g.lineTo(X(n - 1), Y(t.a * (n - 1) + t.b)); g.stroke();
      let y = Y(t.a * (n - 1) + t.b) - 14; while (r2y.some(z => Math.abs(z - y) < 24)) y += 24; r2y.push(y);
      g.fillStyle = "#000"; g.fillText("R² = " + t.r2.toFixed(4).replace(".", ","), X(n - 1) + 10, y);
    }
    const leg = [["INPUT", "#000", true], ["OUTPUT", "#00B050", true], ["SALDO REGULATÓRIO", "#FF0000", true],
      ["Linear (INPUT)", "#000", false], ["Linear (OUTPUT)", "#00B050", false], ["Linear (SALDO REGULATÓRIO)", "#FF0000", false]];
    g.font = "20px Calibri, Arial, sans-serif";
    const larg = leg.map(([t]) => g.measureText(t).width + 70), totL = larg.reduce((a, b) => a + b, 0);
    let x = (W - totL) / 2; const y = H - 28;
    leg.forEach(([t, cor, grossa], i) => {
      g.strokeStyle = cor; g.lineWidth = grossa ? 4 : 1.6; g.beginPath(); g.moveTo(x, y); g.lineTo(x + 44, y); g.stroke();
      if (grossa) { g.fillStyle = cor; g.beginPath(); g.arc(x + 22, y, 5.5, 0, 7); g.fill(); }
      g.fillStyle = "#404040"; g.fillText(t, x + 52, y); x += larg[i];
    });
    return c.toDataURL("image/png");
  }

  function pizza(valores, rotulos, cores, rotulo) {
    const W = 1280, H = 920, cx = 640, cy = 410, R = 340;
    const [c, g] = canvas(W, H), tot = valores.reduce((a, b) => a + b, 0) || 1;
    let ang = -Math.PI / 2, peq = 0;
    valores.forEach((v, i) => {
      if (!v) return;
      const a = v / tot * Math.PI * 2;
      g.fillStyle = cores[i]; g.beginPath(); g.moveTo(cx, cy); g.arc(cx, cy, R, ang, ang + a); g.closePath(); g.fill();
      g.strokeStyle = "#fff"; g.lineWidth = 4; g.stroke();
      const m = ang + a / 2, dist = v / tot > 0.06 ? 0.72 : 1.1 + 0.12 * (peq++ % 2);
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
  const graficoCids = r => pizza(r.cids.map(x => x.n), r.cids.map(x => x.grupo),
    [...["#4472C4", "#ED7D31", "#FF0000", "#7F7F7F", "#70AD47", "#7030A0", "#00B0F0", "#9E480E", "#264478"].slice(0, r.cids.length - 1), "#FFC000"], (v, t) => `${Math.round(v / t * 100)}%`);
  function graficoNiveis(r) {
    const base = {APS: "#70AD47", DIRASE: "#2F5597", HRS: "#FF0000", HRPL: "#C00000", UPA: "#FFC000", OUTROS: "#7F7F7F"};
    const extra = ["#7030A0", "#00B0F0", "#A5A5A5", "#264478"]; let k = 0;
    return pizza(r.niveis.map(x => x.n), r.niveis.map(x => x.nivel), r.niveis.map(x => base[x.nivel] || extra[k++ % 4]),
      (v, t) => `${v}; ${Math.round(v / t * 100)}%`);
  }

  // ------------------------------------------------------------------------------ PDF
  async function gerarPdf(r) {
    for (const u of LIBS.pdf) await lib(u);
    const {jsPDF} = window.jspdf, N = notas(r);
    const doc = new jsPDF({orientation: "landscape", unit: "mm", format: "a4"});
    let y = 18, W = 297, M = 20;
    const texto = (t, tam = 12, opt = {}) => {
      doc.setFont("times", opt.negrito ? "bold" : "normal"); doc.setFontSize(tam);
      const ls = doc.splitTextToSize(t, W - 2 * M), h = tam * 0.42;
      if (y + ls.length * h > (W > 250 ? 200 : 285)) { novaPag(); }
      ls.forEach(l => {
        if (opt.realce) { doc.setFillColor(255, 255, 0); doc.rect(opt.centro ? (W - doc.getTextWidth(l)) / 2 : M, y - h * 0.8, doc.getTextWidth(l), h + 0.6, "F"); }
        doc.text(l, opt.centro ? W / 2 : M, y, {align: opt.centro ? "center" : "left"}); y += h;
      });
      y += opt.depois ?? 1;
    };
    const novaPag = () => { doc.addPage("a4", "portrait"); W = 210; M = 25; y = 22; };
    const imagem = (url, larg, alt) => { if (y + alt > (W > 250 ? 200 : 285)) novaPag(); doc.addImage(url, "PNG", (W - larg) / 2, y, larg, alt); y += alt + 3; };
    const rodape = ls => { ls.forEach((l, i) => texto(l, 10, {depois: i === ls.length - 1 ? 5 : 0})); };
    const garantir = h => { if (y + h > (W > 250 ? 200 : 285)) novaPag(); };  // mantém título junto da figura/tabela
    const rascunho = t => texto(N.rascunho + t, 11, {realce: true, depois: 7});

    texto(N.titulo, 12, {negrito: true, centro: true, depois: 6});
    texto(N.f1, 12, {depois: 2});
    imagem(graficoSerie(r), 250, 104.5);
    rodape(N.n1); rascunho(r.textos.fig1);

    novaPag();
    texto(N.f2, 12, {depois: 3});
    const corFill = {Vermelho: [255, 0, 0], Amarelo: [255, 255, 0], Verde: [146, 208, 80], Azul: [0, 176, 240]};
    doc.autoTable({
      startY: y, margin: {left: M, right: M}, theme: "grid", styles: {font: "helvetica", fontSize: 10, halign: "center", textColor: 0, lineColor: 120, lineWidth: 0.2},
      head: [["Atualizado", r.extracao.toLocaleDateString("pt-BR"), "", ""], ["Classificação", "Mais antiga", "Solicitações", "Tempo de espera em DIAS"]],
      headStyles: {fillColor: 255, textColor: 0, fontStyle: "bold"},
      body: r.risco.map(x => [x.cor, x.antiga, x.n ? fmt(x.n) : "", x.dias != null ? fmt(x.dias) : ""]),
      foot: [["TOTAL", "", fmt(r.risco.reduce((a, x) => a + x.n, 0)), ""]],
      footStyles: {fillColor: 255, textColor: 0, fontStyle: "bold"}, showFoot: "lastPage",
      didParseCell: h => { if (h.section === "body") h.cell.styles.fillColor = corFill[h.row.raw[0]]; },
    });
    y = doc.lastAutoTable.finalY + 4;
    rodape(N.n2); rascunho(r.textos.fig2);
    garantir(115); texto(N.f3, 12, {depois: 2});
    imagem(graficoCids(r), 140, 100.6);
    rodape(N.n3);
    garantir(45); texto(N.tc, 12, {depois: 3});
    doc.autoTable({
      startY: y, margin: {left: M, right: M}, theme: "grid",
      styles: {font: "times", fontSize: 8.5, textColor: 0, lineColor: 120, lineWidth: 0.2, cellPadding: 0.8},
      headStyles: {fillColor: 255, textColor: 0, fontStyle: "bold", halign: "center", fontSize: 9.5},
      columnStyles: {0: {halign: "left", cellWidth: 72}, 1: {halign: "right"}, 2: {halign: "right"}, 3: {halign: "right"}, 4: {halign: "left"}},
      head: [["CID", "Solicitações", "%", "% acum.", "Grupo na Figura 3.1"]],
      body: r.topCids.map(x => [x.cat ? x.c : "      " + x.c, fmt(x.n), fmt(x.pct, 1) + "%", x.acum != null ? fmt(x.acum, 1) + "%" : "", x.g]),
      didParseCell: h => { if (h.section === "body" && r.topCids[h.row.index].cat) { h.cell.styles.fontStyle = "bold"; h.cell.styles.fillColor = 242; } },
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
      columnStyles: {0: {halign: "left"}, 1: {halign: "right"}, 2: {halign: "right"}, 3: {halign: "right"}, 4: {halign: "right"}},
      head: [["UNIDADE", "Número de solicitações", "Percentual da APS", "Quantidade de eSF", "Solicitações por eSF"]],
      body: r.ubs.map(x => [x.u, fmt(x.n), fmt(x.pct, 2) + "%", x.esf != null ? fmt(x.esf) : "", x.taxa != null ? fmt(x.taxa, 2) : ""]),
      foot: [["TOTAL", fmt(r.totAps), "100,00%", fmt(r.esfTotal), fmt(r.totAps / r.esfTotal, 2)]],
      footStyles: {fillColor: 255, textColor: 0, fontStyle: "normal", halign: "right"}, showFoot: "lastPage",
    });
    y = doc.lastAutoTable.finalY + 4;
    rodape(N.nt1); rascunho(r.textos.fig4);
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
    const rodape = ls => ls.map((l, i) => par(l, {tam: 10, depois: i === ls.length - 1 ? 12 : 0}));
    const rascunho = t => par([run(N.rascunho, {tam: 11, negrito: true, realce: true}), run(t, {tam: 11, realce: true})], {just: true, depois: 12});
    const cel = (t, o = {}) => new D.TableCell({children: [new D.Paragraph({alignment: o.al ?? D.AlignmentType.CENTER,
      children: [new D.TextRun({text: t, font: o.fonte || "Calibri", size: (o.tam || 11) * 2, bold: o.negrito})]})],
      shading: o.fill ? {type: D.ShadingType.CLEAR, color: "auto", fill: o.fill} : undefined, width: o.w ? {size: o.w, type: D.WidthType.DXA} : undefined});
    const fill = Object.fromEntries(CORES.map(([, rot, f]) => [rot, f]));
    const tRisco = new D.Table({alignment: D.AlignmentType.CENTER, rows: [
      new D.TableRow({children: [cel("Atualizado", {negrito: true, al: D.AlignmentType.LEFT}), cel(r.extracao.toLocaleDateString("pt-BR"), {negrito: true}), cel(""), cel("")]}),
      new D.TableRow({children: ["Classificação", "Mais antiga", "Solicitações", "Tempo de espera em DIAS"].map(t => cel(t, {negrito: true}))}),
      ...r.risco.map(x => new D.TableRow({children: [x.cor, x.antiga, x.n ? fmt(x.n) : "", x.dias != null ? fmt(x.dias) : ""].map(t => cel(t, {fill: fill[x.cor]}))})),
      new D.TableRow({children: [cel("TOTAL", {negrito: true, al: D.AlignmentType.LEFT}), cel(""), cel(fmt(r.risco.reduce((a, x) => a + x.n, 0)), {negrito: true}), cel("")]}),
    ], columnWidths: [2040, 2040, 2040, 2600]});
    const R_ = D.AlignmentType.RIGHT, L_ = D.AlignmentType.LEFT, o9 = {fonte: "Times New Roman", tam: 9};
    const tUbs = new D.Table({alignment: D.AlignmentType.CENTER, columnWidths: [2950, 1600, 1600, 1500, 1600], rows: [
      new D.TableRow({tableHeader: true, children: ["UNIDADE", "Número de solicitações", "Percentual da APS", "Quantidade de eSF", "Solicitações por eSF"].map(t => cel(t, {fonte: "Times New Roman", tam: 10, negrito: true}))}),
      ...r.ubs.map(x => new D.TableRow({children: [cel(x.u, {...o9, al: L_}), cel(fmt(x.n), {...o9, al: R_}), cel(fmt(x.pct, 2) + "%", {...o9, al: R_}),
        cel(x.esf != null ? fmt(x.esf) : "", {...o9, al: R_}), cel(x.taxa != null ? fmt(x.taxa, 2) : "", {...o9, al: R_})]})),
      new D.TableRow({children: [cel("TOTAL", {...o9, al: L_}), cel(fmt(r.totAps), {...o9, al: R_}), cel("100,00%", {...o9, al: R_}), cel(fmt(r.esfTotal), {...o9, al: R_}), cel(fmt(r.totAps / r.esfTotal, 2), {...o9, al: R_})]}),
    ]});
    const tCid = new D.Table({alignment: D.AlignmentType.CENTER, columnWidths: [3700, 1400, 800, 1000, 2200], rows: [
      new D.TableRow({tableHeader: true, children: ["CID", "Solicitações", "%", "% acum.", "Grupo na Figura 3.1"].map(t => cel(t, {fonte: "Times New Roman", tam: 10, negrito: true}))}),
      ...r.topCids.map(x => { const o = {...o9, negrito: x.cat, fill: x.cat ? "F2F2F2" : undefined};
        return new D.TableRow({children: [cel(x.cat ? x.c : "      " + x.c, {...o, al: L_}), cel(fmt(x.n), {...o, al: R_}),
          cel(fmt(x.pct, 1) + "%", {...o, al: R_}), cel(x.acum != null ? fmt(x.acum, 1) + "%" : "", {...o, al: R_}), cel(x.g, {...o, al: L_})]}); }),
    ]});
    const doc = new D.Document({sections: [
      {properties: {page: {size: {orientation: D.PageOrientation.LANDSCAPE, width: 16838, height: 11906}, margin: {top: 1000, bottom: 1000, left: 1134, right: 1134}}},
        children: [par(N.titulo, {negrito: true, centro: true, depois: 18}), par(N.f1, {depois: 4}), img(graficoSerie(r), 940, 393), ...rodape(N.n1), rascunho(r.textos.fig1)]},
      {properties: {page: {size: {orientation: D.PageOrientation.PORTRAIT, width: 11906, height: 16838}, margin: {top: 1134, bottom: 1134, left: 1417, right: 1417}}},
        children: [par(N.f2, {depois: 6}), tRisco, ...rodape(N.n2), rascunho(r.textos.fig2),
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
    aba("Saldo Regulatório", [["Mês/Ano", "INPUT", "OUTPUT", "SALDO REGULATÓRIO"], ...r.serie.meses.map((m, i) => [m, r.serie.in[i], r.serie.out[i], r.serie.saldo[i]])], [10, 10, 10, 20]);
    aba("Tempo de espera por risco", [["Classificação", "Mais antiga", "Solicitações", "Tempo de espera em DIAS", "Meta do protocolo (dias)", "Acima da meta?"],
      ...r.risco.map(x => [x.cor, x.antiga, x.n, x.dias ?? "", x.meta ?? "", x.meta && x.dias != null ? (x.acima ? "Sim" : "Não") : ""]),
      ["TOTAL", "", r.risco.reduce((a, x) => a + x.n, 0), "", "", ""]], [14, 12, 12, 22, 22, 14]);
    aba("Figura 3.1 CIDs grupos", [["Grupo", "Solicitações", "%"], ...r.cids.map(x => [x.grupo, x.n, x.pct])], [45, 12, 8]);
    aba("Figura 3.2 CIDs detalhe", [["Nível", "CID", "Solicitações", "%", "% acumulado", "Grupo na Figura 3.1"],
      ...r.topCids.map(x => [x.cat ? "Categoria" : "Subcategoria", x.c, x.n, Math.round(x.pct * 10) / 10, x.acum != null ? Math.round(x.acum * 10) / 10 : "", x.g])], [13, 80, 12, 8, 12, 36]);
    aba("CIDs (todos)", [["CID", "Solicitações pendentes/reenviadas", "Grupo na Figura 3.1"], ...r.cidsDetalhe], [80, 30, 40]);
    aba("Devolvidas por nível", [["Nível", "Solicitações", "%"], ...r.niveis.map(x => [x.nivel, x.n, x.pct])], [24, 12, 8]);
    aba("Devolvidas por unidade", [["Unidade solicitante", "Nível / fatia", "Solicitações devolvidas"], ...r.devUnidade.map(x => [x.u, x.nivel, x.n])], [32, 22, 22]);
    aba("Devolvidas APS por UBS", [["UNIDADE", "Número de solicitações", "Percentual da APS", "Quantidade de eSF", "Solicitações por eSF"],
      ...r.ubs.map(x => [x.u, x.n, Math.round(x.pct * 100) / 100, x.esf ?? "", x.taxa != null ? Math.round(x.taxa * 100) / 100 : ""]),
      ["TOTAL", r.totAps, 100, r.esfTotal, Math.round(r.totAps / r.esfTotal * 100) / 100]], [26, 22, 18, 18, 20]);
    aba("Notas", [[N.titulo], [""], ...[N.n1, N.n2, N.n3, N.ntc, N.n4, N.nt1].flat().map(t => [t])], [160]);
    aba("Avisos", [["Avisos para revisão"], ...(r.avisos.length ? r.avisos : ["Nenhum"]).map(a => [a])], [140]);
    X.writeFile(wb, `Bases e tabelas - ${r.nome} (${r.extracao.toLocaleDateString("pt-BR").replace(/\//g, "-")}).xlsx`);
  }

  function baixar(blob, nome) {
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = nome;
    document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
  }

  window.gerarAnalise = async function (id, formato, nome) {
    if (!REF) REF = await (await fetch("referencias.json", {cache: "no-store"})).json();
    const d = await (await fetch(`analises/${id}.json`, {cache: "no-store"})).json();
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
      ["Aba Solicitações", "Pendentes e reenviadas ao regulador (fila) e devolvidas, uma linha por solicitação. Não inclui fila de espera."],
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

  fetch("referencias.json", {cache: "no-store"}).then(x => x.json()).then(j => { REF = REF || j; }).catch(() => {});
  window._analiseMeta = p => REF && REF.metas && REF.metas[p] || null;
  window._analiseNomeCurto = p => (REF && REF.notas[p] && REF.notas[p].nome) || nomeCurto(p);
  window._analiseCalcular = async (id) => { if (!REF) REF = await (await fetch("referencias.json")).json(); return calcular(await (await fetch(`analises/${id}.json`)).json()); };
})();

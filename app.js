// Use sempre a URL pública de Web App do Apps Script (sem /a/macros/<dominio>/)
// para permitir chamadas a partir do GitHub Pages sem dependência de sessão Google.
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwZ9itEkqLp9QWPn5NK1olS9j9FLrGrIdVpnXIszbDL7Wv_pWL4zxBrMRlUz1MqcHq-pw/exec";

const STORAGE_KEYS = {
  benchmarks: "viab_benchmarks_v1",
  scenarios: "viab_scenarios_v1",
  terrenos: "viab_terrenos_v3",
};

const PROJECT_TYPES = [
  { id: "loteamento", label: "Loteamento", icon: "🏘️", enabled: true },
  { id: "incorporacao", label: "Incorporação", icon: "🏗️", enabled: false },
  { id: "bts", label: "BTS / Locação", icon: "🏭", enabled: false },
];

const BENCHMARK_TEMPLATE = {
  loteamento: {
    urban: {
      areaLoteavelPct: 88.7,
      areaLotesPct: 50.0,
      areaLiquidaVendaPct: 100.0,
      lotesPossiveis: 500,
    },
    financial: {
      margemFinalPct: 27.0,
      margemOperacionalPct: 32.0,
      custoObrasPct: 35.0,
      precoVendaM2: 750.0,
    },
  },
  horizontal: {
    urban: {
      areaLoteavelPct: 0,
      areaLotesPct: 0,
      areaLiquidaVendaPct: 0,
      lotesPossiveis: 0,
    },
    financial: {
      margemFinalPct: 0,
      margemOperacionalPct: 0,
      custoObrasPct: 0,
      precoVendaM2: 0,
    },
  },
  vertical: {
    urban: {
      areaLoteavelPct: 0,
      areaLotesPct: 0,
      areaLiquidaVendaPct: 0,
      lotesPossiveis: 0,
    },
    financial: {
      margemFinalPct: 0,
      margemOperacionalPct: 0,
      custoObrasPct: 0,
      precoVendaM2: 0,
    },
  },
};

const state = {
  view: "home",
  phase: null,
  projectType: null,
  showBenchmarkModal: false,
  benchmarkMessage: "",
  sheetMessage: "",
  studySheetMessage: "",
  showStudyPickerModal: false,
  sheetStudies: [],
  showFeedbackModal: false,
  feedbackType: "sugestao",
  feedbackText: "",
  feedbackMessage: "",
  study: {
    studyId: "",
    nomeEstudo: "",
    cidade: "",
    urban: {
      areaTotal: 0,
      percApp: 0,
      percRem: 0,
      percNaoEd: 0,
      percInst: 5,
      percPubl: 15,
      percViario: 30,
    },
    product: {
      nLotes: 0,
      areaMedia: 0,
      precoM2: 0,
    },
    costs: {
      infraMode: "pct",
      infraM2: 0,
      infraPct: 30,
      projetoMode: "R$",
      projetoR: 0,
      projetoPct: 0,
      licenciamentoMode: "R$",
      licenciamentoR: 0,
      licenciamentoPct: 0,
      registroMode: "R$",
      registroR: 0,
      registroPct: 0,
      manutPosPct: 2,
      marketingPct: 2.5,
      corretagemPct: 5,
      adminPct: 0.75,
      impostosPct: 4,
      permFisicaPct: 0,
      permFinPct: 0,
      permFinExcImpostos: true,
      permFinExcCorretagem: true,
      permFinExcMarketing: false,
      permFinExcAdmin: true,
      contingenciasPct: 1.5,
      terrenoM2: 0,
    },
  },
  calc: null,
  savedScenarios: loadLocal(STORAGE_KEYS.scenarios, []),
  benchmarks: loadLocal(STORAGE_KEYS.benchmarks, BENCHMARK_TEMPLATE),
  terrenos: loadLocal("viab_terrenos", []),
  terrenoForm: { nome: "", cidade: "", estado: "", projeto: "", etapa: "", areaGleba: 0, areaApp: 0, fotoBase64: "", fotoNome: "", kmlBase64: "", kmlNome: "" },
  terrenoMessage: "",
  showTerrenoPickerModal: false,
};

function loadLocal(key, fallback) {
  try {
    const v = localStorage.getItem(key);
    return v ? JSON.parse(v) : fallback;
  } catch {
    return fallback;
  }
}

function saveLocal(key, value) {
  localStorage.setItem(key, JSON.stringify(value));
}

function clone(v) {
  return JSON.parse(JSON.stringify(v));
}

function fmt(v, d = 2) {
  return Number(v || 0).toLocaleString("pt-BR", {
    minimumFractionDigits: d,
    maximumFractionDigits: d,
  });
}

function rs(v) {
  return `R$ ${fmt(v, 2)}`;
}

function pc(v) {
  return `${fmt(v, 2)}%`;
}

function num(v) {
  const n = parseFloat(v);
  return Number.isFinite(n) ? n : 0;
}

function pctOf(value, total) {
  return total ? (value / total) * 100 : 0;
}

function deltaPct(base, value) {
  return base ? ((value / base) - 1) * 100 : 0;
}

function safeId(txt) {
  return (txt || "").replace(/[^\w\-]+/g, "_");
}

let _activePath = null;
let _activeRaw = "";

function fmtBR(v) {
  const n = Number(v || 0);
  return n.toLocaleString("pt-BR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function getNestedValue(path) {
  const keys = path.split(".");
  let ref = state.study;
  try {
    for (const k of keys) ref = ref[k];
    return ref;
  } catch {
    return 0;
  }
}

function getDefaultStudy() {
  return {
    studyId: "",
    nomeEstudo: "",
    cidade: "",
    urban: {
      areaTotal: 0,
      percApp: 0,
      percRem: 0,
      percNaoEd: 0,
      percInst: 5,
      percPubl: 15,
      percViario: 30,
    },
    product: {
      nLotes: 0,
      areaMedia: 0,
      precoM2: 0,
    },
    costs: {
      infraMode: "pct",
      infraM2: 0,
      infraPct: 30,
      projetoMode: "R$",
      projetoR: 0,
      projetoPct: 0,
      licenciamentoMode: "R$",
      licenciamentoR: 0,
      licenciamentoPct: 0,
      registroMode: "R$",
      registroR: 0,
      registroPct: 0,
      manutPosPct: 2,
      marketingPct: 2.5,
      corretagemPct: 5,
      adminPct: 0.75,
      impostosPct: 4,
      permFisicaPct: 0,
      permFinPct: 0,
      permFinExcImpostos: true,
      permFinExcCorretagem: true,
      permFinExcMarketing: false,
      permFinExcAdmin: true,
      contingenciasPct: 1.5,
      terrenoM2: 0,
    },
  };
}

function normalizeLoadedStudy(payload) {
  const base = getDefaultStudy();
  const loaded = payload && payload.study ? payload.study : {};
  const loadedCosts = { ...(loaded.costs || {}) };
  const seemsOldInfraDefault = loadedCosts.infraMode === "m2"
    && num(loadedCosts.infraM2) === 30
    && num(loadedCosts.infraPct) === 0;
  if (!("infraMode" in loadedCosts) || seemsOldInfraDefault) {
    loadedCosts.infraMode = "pct";
    loadedCosts.infraPct = 30;
    loadedCosts.infraM2 = 0;
  }

  return {
    ...base,
    ...loaded,
    urban: { ...base.urban, ...(loaded.urban || {}) },
    product: { ...base.product, ...(loaded.product || {}) },
    costs: { ...base.costs, ...loadedCosts },
  };
}

function kpiWithBm(title, value, sub, actualNum, bmNum, higherIsBetter = true) {
  if (!bmNum) return kpi(title, value, sub);
  const isOk = higherIsBetter ? actualNum >= bmNum : actualNum <= bmNum;
  return `
    <div class="kpi ${isOk ? "bm-ok" : "bm-bad"}">
      <div class="k-title">${title}</div>
      <div class="k-value">${value}</div>
      <div class="k-sub">${sub}</div>
      <div class="k-bm">Benchmark: ${pc(bmNum)}</div>
    </div>
  `;
}

function compute(study) {
  const u = study.urban;
  const p = study.product;
  const c = study.costs;

  const areaTotal = num(u.areaTotal);
  const areaApp = areaTotal * num(u.percApp) / 100;
  const areaRem = areaTotal * num(u.percRem) / 100;
  const areaNaoEd = areaTotal * num(u.percNaoEd) / 100;
  const areaLoteavel = Math.max(0, areaTotal - areaApp - areaRem - areaNaoEd);

  const areaInst = areaLoteavel * num(u.percInst) / 100;
  const areaPubl = areaLoteavel * num(u.percPubl) / 100;
  const areaViario = areaLoteavel * num(u.percViario) / 100;
  const pctLotesVend = Math.max(0, 100 - num(u.percInst) - num(u.percPubl) - num(u.percViario));
  const areaLotesVend = Math.max(0, areaLoteavel - areaInst - areaPubl - areaViario);

  const nLotes = Math.round(num(p.nLotes));
  const areaMedia = num(p.areaMedia);
  const precoM2 = num(p.precoM2);

  const areaTotalLotes = nLotes * areaMedia;
  const areaPermutaFis = areaTotalLotes * num(c.permFisicaPct) / 100;
  const areaLiquidaVenda = Math.max(0, areaTotalLotes - areaPermutaFis);

  const ticketMedio = areaMedia * precoM2;
  const vgvPotencial = areaTotalLotes * precoM2;
  const vgvBruto = areaLiquidaVenda * precoM2;

  const permutaFisicaR = areaPermutaFis * precoM2;

  // Deduções comerciais antes da receita líquida
  const impostosR = vgvBruto * num(c.impostosPct) / 100;
  const corretagemR = vgvBruto * num(c.corretagemPct) / 100;
  const marketingR = vgvBruto * num(c.marketingPct) / 100;

  const receitaLiquida = vgvBruto - impostosR - corretagemR - marketingR;

  const terrenoR = num(c.terrenoM2) * areaTotal;
  const infraR = c.infraMode === "pct"
    ? vgvBruto * num(c.infraPct) / 100
    : num(c.infraM2) * areaLoteavel;
  const projetoFinalR = c.projetoMode === "pct"
    ? vgvBruto * num(c.projetoPct) / 100
    : num(c.projetoR);
  const licenciamentoFinalR = c.licenciamentoMode === "pct"
    ? vgvBruto * num(c.licenciamentoPct) / 100
    : num(c.licenciamentoR);
  const registroR = c.registroMode === "pct"
    ? vgvBruto * num(c.registroPct) / 100
    : num(c.registroR);
  const manutPosR = infraR * num(c.manutPosPct) / 100;

  const custoObrasTotal = infraR + projetoFinalR + licenciamentoFinalR + registroR + manutPosR;
  const contingenciasR = custoObrasTotal * num(c.contingenciasPct) / 100;
  const resultadoOperacional = receitaLiquida - terrenoR - custoObrasTotal - contingenciasR;

  // Após resultado operacional: admin, permuta financeira (base ajustável)
  const adminR = vgvBruto * num(c.adminPct) / 100;
  const permFinBrutoR = vgvBruto * num(c.permFinPct) / 100;
  let permFinReducaoR = 0;
  if (c.permFinExcImpostos)   permFinReducaoR += impostosR;
  if (c.permFinExcCorretagem) permFinReducaoR += corretagemR;
  if (c.permFinExcMarketing)  permFinReducaoR += marketingR;
  if (c.permFinExcAdmin)      permFinReducaoR += adminR;
  const permFinR = Math.max(0, permFinBrutoR - permFinReducaoR);
  const resultadoFinal = resultadoOperacional - permFinR - adminR;

  const custoTotal = terrenoR + custoObrasTotal + impostosR + adminR + corretagemR + marketingR + permFinR + contingenciasR;
  const custoM2Lotes = areaTotalLotes ? custoObrasTotal / areaTotalLotes : 0;
  const relPrecoCusto = custoM2Lotes ? precoM2 / custoM2Lotes : 0;
  const lotesPossiveis = areaMedia ? Math.floor(areaLotesVend / areaMedia) : 0;
  const areaMediaCalculada = nLotes ? areaLotesVend / nLotes : 0;

  return {
    areaTotal,
    areaApp,
    areaRem,
    areaNaoEd,
    areaLoteavel,
    areaInst,
    areaPubl,
    areaViario,
    pctLotesVend,
    areaLotesVend,
    nLotes,
    lotesPossiveis,
    areaMedia,
    areaMediaCalculada,
    areaTotalLotes,
    areaPermutaFis,
    areaLiquidaVenda,
    precoM2,
    ticketMedio,
    vgvPotencial,
    vgvBruto,
    permutaFisicaR,
    impostosR,
    corretagemR,
    marketingR,
    receitaLiquida,
    terrenoR,
    infraR,
    projetoFinalR,
    licenciamentoFinalR,
    registroR,
    manutPosR,
    custoObrasTotal,
    resultadoOperacional,
    permFinR,
    contingenciasR,
    adminR,
    resultadoFinal,
    custoTotal,
    custoM2Lotes,
    relPrecoCusto,
    margemLiquidaPct: pctOf(receitaLiquida, vgvBruto),
    margemOperacionalPct: pctOf(resultadoOperacional, vgvBruto),
    margemFinalPct: pctOf(resultadoFinal, vgvBruto),
    custoObrasPct: pctOf(custoObrasTotal, vgvBruto),
    custoTotalPct: pctOf(custoTotal, vgvBruto),
    areaLoteavelPct: pctOf(areaLoteavel, areaTotal),
    areaLotesSobreLoteavelPct: pctLotesVend,
    areaLotesSobreTotalPct: pctOf(areaLotesVend, areaTotal),
    areaLiquidaVendaPct: pctOf(areaTotalLotes, areaLotesVend),
    vgvPotencialPctSobreBruto: vgvBruto ? (vgvPotencial / vgvBruto) * 100 : 100,
    permutaFisicaPctSobreBruto: vgvBruto ? (permutaFisicaR / vgvBruto) * 100 : 0,
  };
}

function setNested(path, value) {
  const keys = path.split(".");
  let ref = state.study;
  for (let i = 0; i < keys.length - 1; i++) ref = ref[keys[i]];
  ref[keys[keys.length - 1]] = value;
  rerender();
}

function setView(view, extra = {}) {
  Object.assign(state, extra);
  state.view = view;
  rerender();
}

function inputField(label, path, value, opts = {}) {
  const { suffix = "", prefix = "", full = false, text = false, integer = false } = opts;
  const isPercentField = suffix.includes("%");
  const displayVal = text
    ? (_activePath === path ? _activeRaw : (value ?? ""))
    : (_activePath === path ? _activeRaw : (integer ? fmt(value, 0) : (isPercentField ? fmt(value, 2) : fmtBR(value))));

  return `
    <div class="field">
      <label>${label}</label>
      <div class="input-wrap ${full ? "full" : ""}">
        ${prefix ? `<span class="affix left">${prefix}</span>` : ""}
        <input class="inp ${text ? "text" : ""}"
          type="text"
          ${!text ? 'inputmode="decimal"' : ''}
          ${text ? 'data-text="true"' : ''}
          ${integer ? 'data-integer="true"' : ''}
          value="${displayVal}"
          data-path="${path}" />
        ${suffix ? `<span class="affix">${suffix}</span>` : ""}
      </div>
    </div>
  `;
}

// modeField: campo com mini-toggle para escolher modo de entrada
// modes = [{id, label}]  fieldsByMode = { modeId: { path, value, prefix?, suffix? } }
function modeField(label, modePath, modes, fieldsByMode) {
  const currentMode = getNestedValue(modePath) || modes[0].id;
  const f = fieldsByMode[currentMode];
  const toggleHtml = modes.map(m =>
    `<button type="button" class="mode-btn${currentMode === m.id ? " active" : ""}"
      onclick="setNested('${modePath}','${m.id}')">${m.label}</button>`
  ).join("");
  const isPercentField = (f.suffix || "").includes("%");
  const displayVal = _activePath === f.path
    ? _activeRaw
    : (isPercentField ? fmt(f.value, 2) : fmtBR(f.value));
  return `
    <div class="field">
      <label class="label-with-mode">
        <span>${label}</span>
        <span class="mode-toggle">${toggleHtml}</span>
      </label>
      <div class="input-wrap full">
        ${f.prefix ? `<span class="affix left">${f.prefix}</span>` : ""}
        <input class="inp" type="text" inputmode="decimal"
          value="${displayVal}" data-path="${f.path}" />
        ${f.suffix ? `<span class="affix">${f.suffix}</span>` : ""}
      </div>
    </div>
  `;
}

function calcDisplay(label, value, sub = "") {
  return `
    <div class="field">
      <label>${label}</label>
      <div class="calc-display">
        <span class="calc-value">${value}</span>
        ${sub ? `<span class="calc-sub">${sub}</span>` : ""}
      </div>
    </div>
  `;
}

function benchmarkInput(label, type, group, key, value) {
  return `
    <div class="field">
      <label>${label}</label>
      <div class="input-wrap full">
        <input class="inp" type="number" value="${value}" data-benchmark="${type}.${group}.${key}" />
      </div>
    </div>
  `;
}

function homeView() {
  return `
    <div class="view">
      <div class="page-header">
        <div class="header-title">
          <small>Sistema de análise</small>
          <h1>Viabilidade Imobiliária</h1>
          <p>Selecione a fase do estudo para iniciar a análise.</p>
        </div>
      </div>

      <div class="container">
        <div class="home-section-label">Fases do estudo</div>
        <div class="card-grid-3" style="margin-bottom:28px">
          <div class="phase-card available" onclick="setView('projectType', { phase: 'estudo_preliminar' })">
            <div class="icon">🏗️</div>
            <h3>Estudo Preliminar</h3>
            <p>Análise inicial de viabilidade com dados básicos do terreno e indicadores econômico-financeiros.</p>
            <span class="badge ok">DISPONÍVEL</span>
          </div>

          <div class="phase-card">
            <div class="icon">📐</div>
            <h3 style="color:#999">Pré-Projeto</h3>
            <p style="color:#b1b1b1">Estrutura preparada para evolução futura.</p>
            <span class="badge wait">EM BREVE</span>
          </div>
        </div>

        <div class="home-section-label">Configurações</div>
        <div class="card-grid-3">
          <div class="phase-card available" onclick="openBenchmarks()">
            <div class="icon">📊</div>
            <h3>Benchmark</h3>
            <p>Cadastre benchmarks urbanísticos e financeiros para comparação nos estudos.</p>
            <span class="badge ok">EDITÁVEL</span>
          </div>
          <div class="phase-card available" onclick="setView('terrenos')">
            <div class="icon">📍</div>
            <h3>Terrenos</h3>
            <p>Registre terrenos por área temática e visualize no mapa com busca rápida.</p>
            <span class="badge ok">DISPONÍVEL</span>
          </div>
        </div>
      </div>

      ${state.showBenchmarkModal ? benchmarkModal() : ""}
    </div>
  `;
}

function projectTypeView() {
  return `
    <div class="view">
      <div class="page-header">
        <div class="top-breadcrumb">
          <button class="nav-btn" onclick="setView('home')">← Voltar</button>
          <span class="crumb">Estudo preliminar</span>
          <span class="sep">›</span>
          <span class="crumb">Tipo de projeto</span>
        </div>
        <div class="btn-row">
          <button class="btn orange" onclick="openBenchmarks()">Benchmark</button>
        </div>
      </div>

      <div class="container">
        <div class="card-grid-3">
          ${PROJECT_TYPES.map((item) => `
            <div class="phase-card ${item.enabled ? "available" : ""}" ${item.enabled ? `onclick="startProject('${item.id}')"` : ""} style="${item.enabled ? "" : "opacity:.58;cursor:not-allowed"}">
              <div class="icon">${item.icon}</div>
              <h3>${item.label}</h3>
              <p>${item.id === "loteamento" ? "Parcelamento do solo urbano em lotes para venda individual." : "Estrutura já preparada, mas sem cálculo completo nesta etapa."}</p>
              <span class="badge ${item.enabled ? "ok" : "wait"}">${item.enabled ? "DISPONÍVEL" : "EM BREVE"}</span>
            </div>
          `).join("")}
        </div>
      </div>

      ${state.showBenchmarkModal ? benchmarkModal() : ""}
    </div>
  `;
}

function loteamentoView() {
  state.calc = compute(state.study);
  const c = state.calc;
  const bm = state.benchmarks.loteamento || BENCHMARK_TEMPLATE.loteamento;

  return `
    <div class="view">
      <div class="page-header no-print">
        <div class="top-breadcrumb">
          <button class="nav-btn" onclick="setView('home')">← Home</button>
          <span class="crumb">Estudo preliminar</span>
          <span class="sep">›</span>
          <span class="crumb">Loteamento</span>
        </div>
        <div class="btn-row">
          <button class="btn gray" onclick="newStudy()">Novo estudo</button>
          <button class="btn blue" onclick="openStudyPicker()">Buscar estudos salvos</button>
          <button class="btn blue" onclick="openTerrenoPicker()">📍 Selecionar terreno</button>
        </div>
      </div>

      <div class="container print-area">
        <div class="field">
          <label>Nome do estudo</label>
          <div class="input-wrap full" style="max-width:780px">
            <input class="inp text" type="text" value="${_activePath === "nomeEstudo" ? _activeRaw : (state.study.nomeEstudo || "")}" data-path="nomeEstudo" data-text="true" />
          </div>
        </div>

        <div class="field">
          <label>Cidade / referência</label>
          <div class="input-wrap full" style="max-width:540px">
            <input class="inp text" type="text" value="${_activePath === "cidade" ? _activeRaw : (state.study.cidade || "")}" data-path="cidade" data-text="true" />
          </div>
        </div>

        <div class="card-grid-2">
          <div>
            <div class="section">
              <div class="section-head head-primary">1. Dados da área (gleba)</div>
              <div class="section-body">
                ${inputField("Área total da gleba", "urban.areaTotal", state.study.urban.areaTotal, { suffix: "m²", full: true })}
                <div class="form-row">
                  <div>${inputField("APP (área de preservação)", "urban.percApp", state.study.urban.percApp, { suffix: "%" })}</div>
                  <div class="value-inline">${fmt(c.areaApp)} m²</div>
                </div>
                <div class="form-row">
                  <div>${inputField("Área remanescente", "urban.percRem", state.study.urban.percRem, { suffix: "%" })}</div>
                  <div class="value-inline">${fmt(c.areaRem)} m²</div>
                </div>
                <div class="form-row">
                  <div>${inputField("Área não edificante", "urban.percNaoEd", state.study.urban.percNaoEd, { suffix: "%" })}</div>
                  <div class="value-inline">${fmt(c.areaNaoEd)} m²</div>
                </div>
                <div class="highlight" style="background:var(--primary)">
                  <div><div class="title">Área loteável</div></div>
                  <div style="text-align:right">
                    <div class="val">${fmt(c.areaLoteavel)}</div>
                    <div class="sub">m² · ${pc(c.areaLoteavelPct)}</div>
                  </div>
                </div>
              </div>
            </div>

            <div class="section">
              <div class="section-head head-blue">2. Distribuição da área loteável</div>
              <div class="section-body">
                <div class="form-row">
                  <div>${inputField("Áreas institucionais", "urban.percInst", state.study.urban.percInst, { suffix: "%" })}</div>
                  <div class="value-inline">${fmt(c.areaInst)} m²</div>
                </div>
                <div class="form-row">
                  <div>${inputField("Áreas públicas (praças / verdes)", "urban.percPubl", state.study.urban.percPubl, { suffix: "%" })}</div>
                  <div class="value-inline">${fmt(c.areaPubl)} m²</div>
                </div>
                <div class="form-row">
                  <div>${inputField("Sistema viário", "urban.percViario", state.study.urban.percViario, { suffix: "%" })}</div>
                  <div class="value-inline">${fmt(c.areaViario)} m²</div>
                </div>
                <div class="highlight" style="background:var(--green)">
                  <div><div class="title">Área dos lotes (vendável)</div></div>
                  <div style="text-align:right">
                    <div class="val">${fmt(c.areaLotesVend)}</div>
                    <div class="sub">m² · ${pc(c.areaLotesSobreLoteavelPct)}</div>
                  </div>
                </div>
              </div>
            </div>
          </div>

          <div>
            <div class="section">
              <div class="section-head head-orange">Indicadores urbanísticos</div>
              <div class="section-body">
                <div class="kpi-grid">
                  ${kpi("Área loteável", fmt(c.areaLoteavel), pc(c.areaLoteavelPct))}
                  ${kpiWithBm("Área lotes vendáveis", fmt(c.areaLotesVend), pc(c.areaLotesSobreLoteavelPct), c.areaLotesSobreLoteavelPct, bm.urban.areaLotesPct, true)}
                  ${kpi("Área líquida de venda", fmt(c.areaLiquidaVenda), pc(c.areaLiquidaVendaPct))}
                </div>
              </div>
            </div>

            <div class="section">
              <div class="section-head head-blue">Distribuição — área loteável</div>
              <div class="section-body">
                ${bar("Institucional", c.areaInst, c.areaLoteavel, "#4c93b3")}
                ${bar("Público / verdes", c.areaPubl, c.areaLoteavel, "#5fac79")}
                ${bar("Sistema viário", c.areaViario, c.areaLoteavel, "#d18b45")}
                ${bar("Lotes (vendável)", c.areaLotesVend, c.areaLoteavel, "var(--green)")}
              </div>
            </div>
          </div>
        </div>

        <div class="section" style="margin-top:18px">
          <div class="section-head head-primary">3. Parâmetros do produto e preço</div>
          <div class="section-body">
            <div style="display:grid;grid-template-columns:.65fr .9fr 1fr .65fr .8fr 1.2fr;gap:14px">
              ${inputField("Número de lotes", "product.nLotes", state.study.product.nLotes, { full: true, integer: true })}
              ${inputField("Área média do lote", "product.areaMedia", state.study.product.areaMedia, { suffix: "m²", full: true })}
              ${inputField("Preço por m²", "product.precoM2", state.study.product.precoM2, { prefix: "R$", suffix: "/m²", full: true })}
              ${inputField("Permuta física", "costs.permFisicaPct", state.study.costs.permFisicaPct, { suffix: "%", full: true })}
              ${inputField("Permuta financeira", "costs.permFinPct", state.study.costs.permFinPct, { suffix: "%", full: true })}
              ${inputField("Terreno", "costs.terrenoM2", state.study.costs.terrenoM2, { prefix: "R$", suffix: "/m² gleba", full: true })}
            </div>
            <div style="display:grid;grid-template-columns:.65fr .9fr 1fr .65fr .8fr 1.2fr;gap:14px;margin-top:12px;padding-top:12px;border-top:1px solid var(--border)">
              ${calcDisplay("Preço médio do lote", rs(c.ticketMedio))}
              ${calcDisplay("Custo aquisição terreno", rs(c.terrenoR))}
              ${calcDisplay("Área entregue — perm. física", `${fmt(c.areaPermutaFis)} m²`)}
              ${calcDisplay("Valor bruto perm. financeira", rs(c.permFinR))}
              <div></div>
              <div></div>
            </div>
          </div>
        </div>

        <div class="section" style="margin-top:12px">
          <div class="section-head head-orange">4. Estrutura de custos</div>
          <div class="section-body">
            <div style="display:grid;grid-template-columns:1.2fr 1fr 1fr 1fr;gap:14px;margin-bottom:14px">
              ${modeField("Infraestrutura", "costs.infraMode",
                [{id:"pct", label:"% VGV"}, {id:"m2", label:"R$/m²lot."}],
                {
                  m2:  { path:"costs.infraM2",  value: state.study.costs.infraM2,  prefix:"R$", suffix:"/m²lot." },
                  pct: { path:"costs.infraPct",  value: state.study.costs.infraPct,  suffix:"%VGV" }
                }
              )}
              ${modeField("Projetos", "costs.projetoMode",
                [{id:"R$", label:"R$"}, {id:"pct", label:"% VGV"}],
                {
                  "R$": { path:"costs.projetoR",   value: state.study.costs.projetoR,   prefix:"R$" },
                  pct:  { path:"costs.projetoPct",  value: state.study.costs.projetoPct,  suffix:"%VGV" }
                }
              )}
              ${modeField("Licenciamento e Custos Ambientais", "costs.licenciamentoMode",
                [{id:"R$", label:"R$"}, {id:"pct", label:"% VGV"}],
                {
                  "R$": { path:"costs.licenciamentoR",   value: state.study.costs.licenciamentoR,   prefix:"R$" },
                  pct:  { path:"costs.licenciamentoPct",  value: state.study.costs.licenciamentoPct,  suffix:"%VGV" }
                }
              )}
              ${modeField("Registro", "costs.registroMode",
                [{id:"R$", label:"R$"}, {id:"pct", label:"% VGV"}],
                {
                  "R$": { path:"costs.registroR",   value: state.study.costs.registroR,   prefix:"R$" },
                  pct:  { path:"costs.registroPct", value: state.study.costs.registroPct, suffix:"%VGV" }
                }
              )}
            </div>
            <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:14px">
              ${inputField("Manutenção pós-obra", "costs.manutPosPct", state.study.costs.manutPosPct, { suffix: "%infra", full: true })}
              ${inputField("Contingências", "costs.contingenciasPct", state.study.costs.contingenciasPct, { suffix: "%obras", full: true })}
              ${inputField("Marketing", "costs.marketingPct", state.study.costs.marketingPct, { suffix: "%VGV", full: true })}
              ${inputField("Corretagem", "costs.corretagemPct", state.study.costs.corretagemPct, { suffix: "%VGV", full: true })}
              ${inputField("Administração e gestão", "costs.adminPct", state.study.costs.adminPct, { suffix: "%VGV", full: true })}
              ${inputField("Impostos", "costs.impostosPct", state.study.costs.impostosPct, { suffix: "%VGV", full: true })}
            </div>
          </div>
        </div>

        <div class="card-grid-2" style="margin-top:18px">
          <div class="proforma-wrap">
            <div class="proforma-title">PROFORMA LOTEAMENTO</div>
            <div class="proforma-name">${state.study.nomeEstudo || "Sem nome"}</div>
            <table class="proforma-table">
              <thead>
                <tr>
                  <th style="width:52%;text-align:left"></th>
                  <th style="width:20%">R$</th>
                  <th style="width:18%">R$/m² venda líquida</th>
                  <th style="width:10%;white-space:nowrap">% VGV</th>
                </tr>
              </thead>
              <tbody>
                ${proformaRow(`VGV Potencial (${fmt(c.areaTotalLotes, 0)} m²)`, c.vgvPotencial, c.vgvPotencialPctSobreBruto, "row-pre-header")}
                ${c.permutaFisicaR > 0 ? proformaRow(`(-) Permuta física (${fmt(c.areaPermutaFis, 0)} m²)`, -c.permutaFisicaR, c.permutaFisicaPctSobreBruto, "row-pre-deduct") : ""}
                ${proformaRow("Receita bruta (VGV)", c.vgvBruto, 100, "row-sec")}
                ${proformaRow("(-) Impostos", -c.impostosR, pctOf(c.impostosR, c.vgvBruto), "row-expense")}
                ${proformaRow("(-) Corretagem", -c.corretagemR, pctOf(c.corretagemR, c.vgvBruto), "row-expense")}
                ${proformaRow("(-) Marketing e vendas", -c.marketingR, pctOf(c.marketingR, c.vgvBruto), "row-expense")}
                ${proformaRow("= Receita líquida", c.receitaLiquida, c.margemLiquidaPct, "row-sub")}
                ${proformaRow("(-) Pagamento do terreno", -c.terrenoR, pctOf(c.terrenoR, c.vgvBruto), "row-expense")}
                ${proformaRow("(-) Infraestrutura", -c.infraR, pctOf(c.infraR, c.vgvBruto), "row-expense")}
                ${proformaRow("(-) Projetos", -c.projetoFinalR, pctOf(c.projetoFinalR, c.vgvBruto), "row-expense")}
                ${proformaRow("(-) Licenciamento e Custos Ambientais", -c.licenciamentoFinalR, pctOf(c.licenciamentoFinalR, c.vgvBruto), "row-expense")}
                ${proformaRow("(-) Registro", -c.registroR, pctOf(c.registroR, c.vgvBruto), "row-expense")}
                ${proformaRow("(-) Manutenção pós-obra", -c.manutPosR, pctOf(c.manutPosR, c.vgvBruto), "row-expense")}
                ${proformaRow("(-) Contingências", -c.contingenciasR, pctOf(c.contingenciasR, c.vgvBruto), "row-expense")}
                ${proformaRow("= Resultado operacional", c.resultadoOperacional, c.margemOperacionalPct, "row-sub")}
                ${proformaRow("(-) Permuta financeira", -c.permFinR, pctOf(c.permFinR, c.vgvBruto), "row-expense")}
                ${proformaRow("(-) Administração e gestão de carteira", -c.adminR, pctOf(c.adminR, c.vgvBruto), "row-expense")}
                ${proformaRow("= Resultado final", c.resultadoFinal, c.margemFinalPct, "row-result row-result-final")}
              </tbody>
            </table>
          </div>

          <div class="section">
            <div class="section-head head-blue">Indicadores financeiros</div>
            <div class="section-body">
              <div class="kpi-grid">
                ${kpi("Receita líquida", rs(c.receitaLiquida), pc(c.margemLiquidaPct))}
                ${kpiWithBm("Resultado operacional", rs(c.resultadoOperacional), pc(c.margemOperacionalPct), c.margemOperacionalPct, bm.financial.margemOperacionalPct, true)}
                ${kpiWithBm("Resultado final", rs(c.resultadoFinal), pc(c.margemFinalPct), c.margemFinalPct, bm.financial.margemFinalPct, true)}
                ${kpiWithBm("Custo obras / VGV", pc(c.custoObrasPct), rs(c.custoObrasTotal), c.custoObrasPct, bm.financial.custoObrasPct, false)}
                ${kpi("Preço médio por lote", rs(c.ticketMedio), `${fmt(c.nLotes, 0)} lotes`)}
                ${kpi("Relação preço/custo m²", `${fmt(c.relPrecoCusto)}x`, `Preço: R$ ${fmt(c.precoM2)}/m²`)}
              </div>
              ${state.study.costs.permFinPct > 0 ? `
                <div class="perm-fin-opts" style="margin-top:12px">
                  <p class="hint">Base da permuta financeira — excluir do cálculo:</p>
                  <div class="checkbox-group">
                    <label class="checkbox-label"><input type="checkbox" ${state.study.costs.permFinExcImpostos ? "checked" : ""} onchange="setNested('costs.permFinExcImpostos',this.checked)"> Impostos</label>
                    <label class="checkbox-label"><input type="checkbox" ${state.study.costs.permFinExcCorretagem ? "checked" : ""} onchange="setNested('costs.permFinExcCorretagem',this.checked)"> Corretagem</label>
                    <label class="checkbox-label"><input type="checkbox" ${state.study.costs.permFinExcMarketing ? "checked" : ""} onchange="setNested('costs.permFinExcMarketing',this.checked)"> Marketing / publicidade</label>
                    <label class="checkbox-label"><input type="checkbox" ${state.study.costs.permFinExcAdmin ? "checked" : ""} onchange="setNested('costs.permFinExcAdmin',this.checked)"> Gestão de carteira (admin)</label>
                  </div>
                </div>
              ` : ""}
              <div class="spacer"></div>
              <div class="footer-actions no-print">
                <div class="btn-row">
                  <button class="btn primary" onclick="exportPDF()">Exportar PDF</button>
                  <button class="btn green" onclick="exportExcel()">Exportar Excel</button>
                  <button class="btn orange" onclick="sendStudyToSheet()">Registrar na Google Sheet</button>
                  <button class="btn blue" onclick="syncBenchmarksToSheet()">Salvar benchmarks</button>
                </div>
                ${state.sheetMessage
                  ? `<div class="${state.sheetMessage.startsWith("Erro") || state.sheetMessage.startsWith("Falha") ? "error" : "notice"}">${state.sheetMessage}</div>`
                  : `<div class="muted">Exporte ou envie o estudo para o Google Sheets.</div>`}
              </div>
            </div>
          </div>
        </div>

        <div class="section" style="margin-top:18px">
          <div class="section-head head-green">Comparação de cenários</div>
          <div class="section-body">
            <div class="footer-actions no-print" style="margin-bottom:12px">
              <div class="btn-row">
                <button class="btn gray" onclick="saveScenario()">Guardar cenário</button>
                <button class="btn danger" onclick="clearScenarios()">Limpar comparação</button>
              </div>
              <div class="muted">Salve um cenário, altere valores e salve de novo para comparar.</div>
            </div>
            ${comparisonView()}
          </div>
        </div>
      </div>

      ${state.showBenchmarkModal ? benchmarkModal() : ""}
      ${state.showStudyPickerModal ? studyPickerModal() : ""}
      ${state.showTerrenoPickerModal ? terrenoPickerModal() : ""}
    </div>
  `;
}

function kpi(title, value, sub) {
  return `
    <div class="kpi">
      <div class="k-title">${title}</div>
      <div class="k-value">${value}</div>
      <div class="k-sub">${sub}</div>
    </div>
  `;
}

function bar(label, value, total, color) {
  const p = pctOf(value, total);
  return `
    <div class="bar-group">
      <div class="bar-meta">
        <span>${label}</span>
        <strong>${pc(p)}</strong>
      </div>
      <div class="bar-track">
        <div class="bar-fill" style="width:${Math.min(100, p)}%;background:${color}"></div>
      </div>
    </div>
  `;
}

function proformaRow(label, value, pct, cls) {
  const perArea = state.calc ? (state.calc.areaLiquidaVenda ? value / state.calc.areaLiquidaVenda : 0) : 0;
  return `
    <tr class="${cls}">
      <td>${label}</td>
      <td class="num col-main">${value < 0 ? `(${fmt(Math.abs(value))})` : fmt(value)}</td>
      <td class="num col-secondary">${perArea < 0 ? `(${fmt(Math.abs(perArea), 2)})` : fmt(perArea, 2)}</td>
      <td class="num">${pc(pct)}</td>
    </tr>
  `;
}

function summaryRows(calc) {
  return [
    { label: "VGV", value: calc.vgvBruto, kind: "receita" },
    { label: "Receita líquida", value: calc.receitaLiquida, kind: "receita" },
    { label: "Terreno", value: calc.terrenoR, kind: "despesa" },
    { label: "Custo obras total", value: calc.custoObrasTotal, kind: "despesa" },
    { label: "Resultado operacional", value: calc.resultadoOperacional, kind: "resultado" },
    { label: "Impostos", value: calc.impostosR, kind: "despesa" },
    { label: "Administração e gestão", value: calc.adminR, kind: "despesa" },
    { label: "Resultado final", value: calc.resultadoFinal, kind: "resultado" },
    { label: "Margem final %", value: calc.margemFinalPct, kind: "resultado_pct" },
  ];
}

function comparisonView() {
  const items = state.savedScenarios || [];
  const a = items.length >= 1 ? items[items.length - 2] || items[0] : null;
  const b = items.length >= 2 ? items[items.length - 1] : null;

  return `
    <div class="compare-grid">
      <div class="compare-box">
        <h4>Cenário A</h4>
        ${scenarioBox(a)}
      </div>
      <div class="compare-box">
        <h4>Cenário B</h4>
        ${scenarioBox(b)}
      </div>
      <div class="compare-box">
        <h4>Variação percentual</h4>
        ${variationBox(a, b)}
      </div>
    </div>
  `;
}

function scenarioBox(item) {
  if (!item) return `<div class="muted">Ainda não há cenário salvo.</div>`;
  return `
    <div class="muted"><strong>${item.study.nomeEstudo || "Sem nome"}</strong><br>${new Date(item.savedAt).toLocaleString("pt-BR")}</div>
    <div class="spacer"></div>
    <table class="compare-table">
      <tbody>
        ${summaryRows(item.calc).map((row, idx) => `
          <tr class="metric-${row.kind}">
            <td>${row.label}</td>
            <td>${idx === 8 ? pc(row.value) : rs(row.value)}</td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
}

function variationBox(a, b) {
  if (!a || !b) return `<div class="muted">Salve dois cenários para ativar a comparação.</div>`;
  const rowsA = summaryRows(a.calc);
  const rowsB = summaryRows(b.calc);

  return `
    <div class="muted" style="visibility:hidden">&nbsp;<br>&nbsp;</div>
    <div class="spacer"></div>
    <table class="compare-table">
      <tbody>
        ${rowsA.map((row, idx) => `
          <tr class="metric-${row.kind}">
            <td>${row.label}</td>
            <td class="${deltaPct(row.value, rowsB[idx].value) >= 0 ? "var-pos" : "var-neg"}">${pc(deltaPct(row.value, rowsB[idx].value))}</td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
}

function benchmarkModal() {
  const b = state.benchmarks;
  return `
    <div class="modal-overlay" onclick="closeBenchmarks(event)">
      <div class="modal" onclick="event.stopPropagation()">
        <h3>Benchmarks por tipo de projeto</h3>
        <p>Esses valores ficam registrados no navegador e também podem ser enviados para a Google Sheet.</p>

        ${state.benchmarkMessage ? `<div class="notice">${state.benchmarkMessage}</div><div class="spacer"></div>` : ""}

        <div class="benchmark-grid">
          ${benchmarkCard("loteamento", "Loteamento", b.loteamento)}
          ${benchmarkCard("horizontal", "Incorporação Horizontal", b.horizontal)}
          ${benchmarkCard("vertical", "Incorporação Vertical", b.vertical)}
        </div>

        <div class="spacer"></div>

        <div class="footer-actions">
          <div class="btn-row">
            <button class="btn green" onclick="saveBenchmarksLocal()">Salvar benchmark</button>
            <button class="btn orange" onclick="syncBenchmarksToSheet()">Salvar benchmark na planilha</button>
            <button class="btn blue" onclick="loadBenchmarksFromSheet()">Buscar benchmark da planilha</button>
          </div>
          <button class="btn gray" onclick="state.showBenchmarkModal=false; state.benchmarkMessage=''; rerender()">Fechar</button>
        </div>
      </div>
    </div>
  `;
}

function benchmarkCard(type, title, data) {
  return `
    <div class="mini-card">
      <h4>${title}</h4>
      ${benchmarkInput("Área lotes vendáveis (%)", type, "urban", "areaLotesPct", data.urban.areaLotesPct)}
      ${benchmarkInput("Margem final / VGV (%)", type, "financial", "margemFinalPct", data.financial.margemFinalPct)}
      ${benchmarkInput("Margem operacional (%)", type, "financial", "margemOperacionalPct", data.financial.margemOperacionalPct)}
      ${benchmarkInput("Custo obras / VGV (%)", type, "financial", "custoObrasPct", data.financial.custoObrasPct)}
    </div>
  `;
}

function buildPayload() {
  state.calc = compute(state.study);
  return {
    action: "saveStudy",
    payload: {
      timestamp: new Date().toISOString(),
      phase: state.phase,
      projectType: state.projectType,
      study: clone(state.study),
      results: clone(state.calc),
      benchmarks: clone(state.benchmarks[state.projectType || "loteamento"] || {}),
    },
  };
}

async function postToAppsScript(action, payload) {
  if (!APPS_SCRIPT_URL || APPS_SCRIPT_URL.includes("COLE_AQUI")) {
    throw new Error("A URL do Apps Script ainda não foi configurada.");
  }
  let response;
  try {
    response = await fetch(APPS_SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ action, payload }),
    });
  } catch (error) {
    throw new Error("Falha de rede ao salvar no Apps Script. Verifique a URL publicada e a implantação do Web App.");
  }

  if (!response.ok) {
    throw new Error(`Falha HTTP ${response.status} ao salvar no Apps Script.`);
  }

  const raw = await response.text();
  try {
    const parsed = JSON.parse(raw);
    if (!parsed.ok) {
      throw new Error(parsed.message || "Apps Script retornou erro ao salvar.");
    }
    return parsed;
  } catch {
    throw new Error(`Resposta inválida do Apps Script: ${raw.slice(0, 160)}`);
  }
}

async function getFromAppsScript(action, params = {}) {
  if (!APPS_SCRIPT_URL || APPS_SCRIPT_URL.includes("COLE_AQUI")) {
    throw new Error("A URL do Apps Script ainda não foi configurada.");
  }

  const query = new URLSearchParams({ action, ...params }).toString();
  const url = `${APPS_SCRIPT_URL}?${query}`;

  let response;
  try {
    response = await fetch(url);
  } catch (error) {
    throw new Error("Falha de rede ao consultar o Apps Script. Verifique se a URL é pública e se o deploy está ativo.");
  }

  if (!response.ok) {
    throw new Error(`Falha HTTP ${response.status}`);
  }

  const raw = await response.text();

  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch {
    throw new Error(`Resposta inválida do Apps Script: ${raw.slice(0, 160)}`);
  }

  if (!parsed.ok) {
    throw new Error(parsed.message || "Apps Script retornou erro.");
  }

  return parsed;
}

function openBenchmarks() {
  state.showBenchmarkModal = true;
  state.benchmarkMessage = "";
  rerender();
}

function closeBenchmarks(e) {
  if (e.target.classList.contains("modal-overlay")) {
    state.showBenchmarkModal = false;
    state.benchmarkMessage = "";
    rerender();
  }
}

function saveBenchmarksLocal() {
  saveLocal(STORAGE_KEYS.benchmarks, state.benchmarks);
  state.benchmarkMessage = "Benchmarks salvos com sucesso no navegador.";
  rerender();
}

async function syncBenchmarksToSheet() {
  try {
    saveLocal(STORAGE_KEYS.benchmarks, state.benchmarks);
    await postToAppsScript("saveBenchmarks", state.benchmarks);
    const msg = "Benchmarks enviados para a planilha. Aguarde 1 a 2 segundos e confira a aba benchmarks.";
    state.benchmarkMessage = msg;
    state.sheetMessage = msg;
  } catch (err) {
    const msg = `Falha ao salvar benchmarks na planilha: ${err.message}`;
    state.benchmarkMessage = msg;
    state.sheetMessage = msg;
  }
  rerender();
}

async function loadBenchmarksFromSheet() {
  try {
    const res = await getFromAppsScript("getBenchmarks");
    if (res && res.data) {
      state.benchmarks = res.data;
      saveLocal(STORAGE_KEYS.benchmarks, state.benchmarks);
      state.benchmarkMessage = "Benchmarks carregados da planilha com sucesso.";
    } else {
      state.benchmarkMessage = "Nenhum benchmark encontrado na planilha.";
    }
  } catch (err) {
    state.benchmarkMessage = `Falha ao buscar benchmarks: ${err.message}`;
  }
  rerender();
}

async function openStudyPicker() {
  state.showStudyPickerModal = true;
  state.sheetStudies = [];
  state.studySheetMessage = "Carregando estudos da planilha...";
  rerender();

  try {
    const res = await getFromAppsScript("listStudies");
    state.sheetStudies = Array.isArray(res.data) ? res.data : [];
    state.studySheetMessage = state.sheetStudies.length
      ? "Selecione um estudo para preencher os campos automaticamente."
      : "Nenhum estudo encontrado na aba viabilidade.";
  } catch (err) {
    state.studySheetMessage = `Erro ao carregar estudos: ${err.message}`;
  }

  rerender();
}

async function applyStudyFromSheet(index) {
  const selected = state.sheetStudies[index];
  if (!selected || !selected.id) return;

  try {
    const res = await getFromAppsScript("getStudy", { id: selected.id });
    const payload = res.data;

    if (!payload || !payload.study) {
      throw new Error("O estudo retornado não possui estrutura válida.");
    }

    state.study = normalizeLoadedStudy(payload);
    state.phase = payload.phase || "estudo_preliminar";
    state.projectType = payload.projectType || "loteamento";
    state.view = "loteamento";
    state.sheetMessage = `Estudo "${selected.nomeEstudo}" carregado da planilha.`;
    state.showStudyPickerModal = false;
    state.studySheetMessage = "";
  } catch (err) {
    state.studySheetMessage = `Erro ao carregar estudo: ${err.message}`;
  }

  rerender();
}

function studyPickerModal() {
  return `
    <div class="modal-overlay" onclick="closeStudyPicker(event)">
      <div class="modal study-modal" onclick="event.stopPropagation()">
        <h3>Estudos salvos na planilha</h3>
        <p>${state.studySheetMessage || "Selecione um estudo para preencher o formulário."}</p>
        <div class="study-list">
          ${state.sheetStudies.map((s, idx) => `
            <button class="study-item" onclick="applyStudyFromSheet(${idx})">
              <strong>${s.nomeEstudo || "Sem nome"}</strong>
              <span>${s.cidade || "Sem cidade"} · ${s.studyId || "Sem ID"} · ${s.timestamp ? new Date(s.timestamp).toLocaleString("pt-BR") : "-"}</span>
            </button>
          `).join("") || `<div class="muted">Sem estudos para listar.</div>`}
        </div>
        <div class="spacer"></div>
        <div class="footer-actions">
          <button class="btn gray" onclick="closeStudyPicker()">Fechar</button>
        </div>
      </div>
    </div>
  `;
}

function closeStudyPicker(e) {
  if (!e || (e.target && e.target.classList && e.target.classList.contains("modal-overlay"))) {
    state.showStudyPickerModal = false;
    state.studySheetMessage = "";
    rerender();
  }
}

// ─── Feedback ────────────────────────────────────────────────────────────────

function openFeedback() {
  state.showFeedbackModal = true;
  state.feedbackMessage = "";
  rerender();
}

function closeFeedback(e) {
  if (!e || (e.target && e.target.classList && e.target.classList.contains("modal-overlay"))) {
    state.showFeedbackModal = false;
    state.feedbackMessage = "";
    rerender();
  }
}

async function submitFeedback() {
  if (!state.feedbackText.trim()) {
    state.feedbackMessage = "Por favor, descreva seu feedback antes de enviar.";
    rerender();
    return;
  }
  try {
    await postToAppsScript("saveFeedback", {
      timestamp: new Date().toISOString(),
      tipo: state.feedbackType,
      texto: state.feedbackText.trim(),
      studyId: state.study.studyId || "",
      projectType: state.projectType || "",
    });
    state.feedbackMessage = "Feedback enviado com sucesso. Obrigado!";
    state.feedbackText = "";
  } catch (err) {
    state.feedbackMessage = `Erro ao enviar feedback: ${err.message}`;
  }
  rerender();
}

function feedbackModal() {
  const tipos = [
    { id: "sugestao", label: "Sugestão" },
    { id: "erro",     label: "Erro" },
    { id: "duvida",   label: "Dúvida" },
  ];
  return `
    <div class="modal-overlay" onclick="closeFeedback(event)">
      <div class="modal feedback-modal" onclick="event.stopPropagation()">
        <h3>Enviar feedback</h3>
        <p>Nos conte sua sugestão, reporte um erro ou tire uma dúvida sobre o sistema.</p>
        ${state.feedbackMessage ? `<div class="notice ${state.feedbackMessage.includes("sucesso") ? "notice-ok" : "notice-err"}">${state.feedbackMessage}</div><div class="spacer"></div>` : ""}
        <div class="feedback-types">
          ${tipos.map(t => `
            <label class="type-option${state.feedbackType === t.id ? " active" : ""}">
              <input type="radio" name="feedbackType" value="${t.id}"
                ${state.feedbackType === t.id ? "checked" : ""}
                onchange="state.feedbackType='${t.id}';rerender()">
              ${t.label}
            </label>
          `).join("")}
        </div>
        <div class="spacer"></div>
        <textarea class="feedback-text" rows="5" placeholder="Descreva aqui..."
          oninput="state.feedbackText=this.value">${state.feedbackText}</textarea>
        <div class="spacer"></div>
        <div class="footer-actions">
          <div class="btn-row">
            <button class="btn orange" onclick="submitFeedback()">Enviar</button>
            <button class="btn gray" onclick="closeFeedback()">Fechar</button>
          </div>
        </div>
      </div>
    </div>
  `;
}

// ─── Google Sheets: envio de estudo ──────────────────────────────────────────

async function sendStudyToSheet() {
  try {
    const payload = buildPayload().payload;
    await postToAppsScript("saveStudy", payload);
    state.sheetMessage = "Estudo enviado para a Google Sheet. Aguarde 1 a 2 segundos e confira a aba viabilidade.";
  } catch (err) {
    state.sheetMessage = `Erro ao enviar estudo: ${err.message}`;
  }
  rerender();
}

function saveScenario() {
  const calc = compute(state.study);
  const arr = loadLocal(STORAGE_KEYS.scenarios, []);
  arr.push({
    savedAt: new Date().toISOString(),
    study: clone(state.study),
    calc,
  });
  state.savedScenarios = arr.slice(-6);
  saveLocal(STORAGE_KEYS.scenarios, state.savedScenarios);
  state.sheetMessage = "Cenário guardado com sucesso para comparação.";
  rerender();
}

function clearScenarios() {
  state.savedScenarios = [];
  saveLocal(STORAGE_KEYS.scenarios, []);
  state.sheetMessage = "Comparação limpa.";
  rerender();
}

function exportPDF() {
  window.print();
}

async function exportExcel() {
  const c = compute(state.study);
  const wb = new ExcelJS.Workbook();

  wb.creator = "Viabilidade Imobiliária";
  wb.lastModifiedBy = "Viabilidade Imobiliária";
  wb.created = new Date();
  wb.modified = new Date();
  wb.company = "UP";
  wb.subject = "Estudo de viabilidade";
  wb.title = state.study.nomeEstudo || "Viabilidade Loteamento";

  const nomeEstudo = state.study.nomeEstudo || "Sem nome";
  const cidade = state.study.cidade || "-";
  const fase = state.phase || "-";
  const tipoProjeto = state.projectType || "-";

  const moneyFmt = 'R$ #,##0.00';
  const numFmt = '#,##0.00';
  const pctFmt = '0.00%';

  function applyTitle(cell, text, fill = "064B59") {
    cell.value = text;
    cell.font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: fill }
    };
    cell.border = {
      top: { style: "thin", color: { argb: "FFFFFFFF" } },
      left: { style: "thin", color: { argb: "FFFFFFFF" } },
      bottom: { style: "thin", color: { argb: "FFFFFFFF" } },
      right: { style: "thin", color: { argb: "FFFFFFFF" } }
    };
  }

  function applySectionHeader(cell, text, fill = "0B5C6B") {
    cell.value = text;
    cell.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
    cell.alignment = { vertical: "middle", horizontal: "left" };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: fill }
    };
    cell.border = {
      top: { style: "thin", color: { argb: "FFFFFFFF" } },
      left: { style: "thin", color: { argb: "FFFFFFFF" } },
      bottom: { style: "thin", color: { argb: "FFFFFFFF" } },
      right: { style: "thin", color: { argb: "FFFFFFFF" } }
    };
  }

  function applyLabel(cell) {
    cell.font = { bold: true, size: 10, color: { argb: "1F2937" } };
    cell.alignment = { vertical: "middle", horizontal: "left" };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "F3F4F6" }
    };
    cell.border = {
      top: { style: "thin", color: { argb: "D1D5DB" } },
      left: { style: "thin", color: { argb: "D1D5DB" } },
      bottom: { style: "thin", color: { argb: "D1D5DB" } },
      right: { style: "thin", color: { argb: "D1D5DB" } }
    };
  }

  function applyValue(cell, format = null, bold = false) {
    cell.font = { bold, size: 10, color: { argb: "111827" } };
    cell.alignment = { vertical: "middle", horizontal: "right" };
    cell.border = {
      top: { style: "thin", color: { argb: "D1D5DB" } },
      left: { style: "thin", color: { argb: "D1D5DB" } },
      bottom: { style: "thin", color: { argb: "D1D5DB" } },
      right: { style: "thin", color: { argb: "D1D5DB" } }
    };
    if (format) cell.numFmt = format;
  }

  function applyTextValue(cell, bold = false) {
    cell.font = { bold, size: 10, color: { argb: "111827" } };
    cell.alignment = { vertical: "middle", horizontal: "left" };
    cell.border = {
      top: { style: "thin", color: { argb: "D1D5DB" } },
      left: { style: "thin", color: { argb: "D1D5DB" } },
      bottom: { style: "thin", color: { argb: "D1D5DB" } },
      right: { style: "thin", color: { argb: "D1D5DB" } }
    };
  }

  function applyKpiBox(ws, startRow, startCol, title, value, subtitle, isMoney = false, isPercent = false) {
    const titleCell = ws.getCell(startRow, startCol);
    const valueCell = ws.getCell(startRow + 1, startCol);
    const subCell = ws.getCell(startRow + 2, startCol);

    ws.mergeCells(startRow, startCol, startRow, startCol + 2);
    ws.mergeCells(startRow + 1, startCol, startRow + 1, startCol + 2);
    ws.mergeCells(startRow + 2, startCol, startRow + 2, startCol + 2);

    titleCell.value = title;
    titleCell.font = { bold: true, size: 10, color: { argb: "4B5563" } };
    titleCell.alignment = { vertical: "middle", horizontal: "left" };
    titleCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "EFF6FF" }
    };

    valueCell.value = value;
    valueCell.font = { bold: true, size: 16, color: { argb: "1F4B8F" } };
    valueCell.alignment = { vertical: "middle", horizontal: "left" };

    if (typeof value === "number") {
      if (isMoney) valueCell.numFmt = moneyFmt;
      if (isPercent) valueCell.numFmt = pctFmt;
      if (!isMoney && !isPercent) valueCell.numFmt = numFmt;
    }

    subCell.value = subtitle;
    subCell.font = { size: 10, color: { argb: "6B7280" } };
    subCell.alignment = { vertical: "middle", horizontal: "left" };

    [titleCell, valueCell, subCell].forEach((cell) => {
      cell.border = {
        top: { style: "thin", color: { argb: "BFDBFE" } },
        left: { style: "thin", color: { argb: "BFDBFE" } },
        bottom: { style: "thin", color: { argb: "BFDBFE" } },
        right: { style: "thin", color: { argb: "BFDBFE" } }
      };
    });
  }

  // =========================
  // ABA 1 - RESUMO EXECUTIVO
  // =========================
  const wsResumo = wb.addWorksheet("Resumo Executivo", {
    views: [{ state: "frozen", ySplit: 5 }]
  });

  wsResumo.columns = [
    { width: 24 },
    { width: 18 },
    { width: 18 },
    { width: 18 },
    { width: 18 },
    { width: 18 },
    { width: 18 },
    { width: 18 },
    { width: 18 },
    { width: 18 },
    { width: 18 },
    { width: 18 }
  ];

  wsResumo.mergeCells("A1:L1");
  applyTitle(wsResumo.getCell("A1"), "PROFORMA LOTEAMENTO", "075985");
  wsResumo.getRow(1).height = 26;

  wsResumo.mergeCells("A2:L2");
  wsResumo.getCell("A2").value = nomeEstudo;
  wsResumo.getCell("A2").font = { bold: true, size: 14, color: { argb: "1F2937" } };
  wsResumo.getCell("A2").alignment = { horizontal: "center", vertical: "middle" };
  wsResumo.getRow(2).height = 22;

  applySectionHeader(wsResumo.getCell("A4"), "IDENTIFICACAO", "0B5C6B");
  wsResumo.mergeCells("A4:F4");

  const identificacao = [
    ["Cidade / Referência", cidade],
    ["Fase", fase],
    ["Tipo de projeto", tipoProjeto],
    ["Study ID", state.study.studyId || ""]
  ];

  let row = 5;
  identificacao.forEach(([label, value]) => {
    applyLabel(wsResumo.getCell(`A${row}`));
    wsResumo.getCell(`A${row}`).value = label;
    wsResumo.mergeCells(`B${row}:F${row}`);
    applyTextValue(wsResumo.getCell(`B${row}`));
    wsResumo.getCell(`B${row}`).value = value;
    row++;
  });

  applySectionHeader(wsResumo.getCell("H4"), "INDICADORES URBANISTICOS", "C8752A");
  wsResumo.mergeCells("H4:L4");

  const urb = [
    ["Área loteável", c.areaLoteavel, numFmt],
    ["Área lotes vendáveis", c.areaLotesVend, numFmt],
    ["Área total dos lotes", c.areaTotalLotes, numFmt],
    ["Área líquida de venda", c.areaLiquidaVenda, numFmt],
    ["Lotes possíveis", c.lotesPossiveis, '0'],
    ["Área média calculada", c.areaMediaCalculada, numFmt]
  ];

  row = 5;
  urb.forEach(([label, value, fmtLocal]) => {
    applyLabel(wsResumo.getCell(`H${row}`));
    wsResumo.getCell(`H${row}`).value = label;
    wsResumo.mergeCells(`I${row}:L${row}`);
    applyValue(wsResumo.getCell(`I${row}`), fmtLocal);
    wsResumo.getCell(`I${row}`).value = value;
    row++;
  });

  applySectionHeader(wsResumo.getCell("A11"), "INDICADORES FINANCEIROS", "0B5C6B");
  wsResumo.mergeCells("A11:L11");

  applyKpiBox(wsResumo, 12, 1, "Receita líquida", c.receitaLiquida, `${fmt(c.margemLiquidaPct)}%`, true, false);
  applyKpiBox(wsResumo, 12, 4, "Resultado operacional", c.resultadoOperacional, `${fmt(c.margemOperacionalPct)}%`, true, false);
  applyKpiBox(wsResumo, 12, 7, "Resultado final", c.resultadoFinal, `${fmt(c.margemFinalPct)}%`, true, false);
  applyKpiBox(wsResumo, 12, 10, "Margem final / VGV", c.margemFinalPct / 100, "rentabilidade final", false, true);

  applyKpiBox(wsResumo, 16, 1, "Custo total", c.custoTotal, `${fmt(c.custoTotalPct)}%`, true, false);
  applyKpiBox(wsResumo, 16, 4, "Custo obras total", c.custoObrasTotal, `${fmt(c.custoObrasPct)}%`, true, false);
  applyKpiBox(wsResumo, 16, 7, "Preço médio por lote", c.ticketMedio, `${fmt(c.nLotes, 0)} lotes`, true, false);
  applyKpiBox(wsResumo, 16, 10, "Relação preço/custo m²", c.relPrecoCusto, `Preço: R$ ${fmt(c.precoM2)}/m²`, false, false);

  // =========================
  // ABA 2 - PROFORMA
  // =========================
  const wsProforma = wb.addWorksheet("Proforma");

  wsProforma.columns = [
    { width: 42 },
    { width: 18 },
    { width: 20 },
    { width: 14 }
  ];

  wsProforma.mergeCells("A1:D1");
  applyTitle(wsProforma.getCell("A1"), "PROFORMA LOTEAMENTO", "075985");

  wsProforma.mergeCells("A2:D2");
  wsProforma.getCell("A2").value = nomeEstudo;
  wsProforma.getCell("A2").font = { bold: true, size: 13, color: { argb: "1F2937" } };
  wsProforma.getCell("A2").alignment = { horizontal: "center", vertical: "middle" };

  const headerRow = wsProforma.getRow(4);
  ["Linha", "R$", "R$/m² venda líquida", "% VGV"].forEach((txt, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = txt;
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "0B5C6B" }
    };
    cell.alignment = { horizontal: i === 0 ? "left" : "center", vertical: "middle" };
    cell.border = {
      top: { style: "thin", color: { argb: "FFFFFF" } },
      left: { style: "thin", color: { argb: "FFFFFF" } },
      bottom: { style: "thin", color: { argb: "FFFFFF" } },
      right: { style: "thin", color: { argb: "FFFFFF" } }
    };
  });

  const proformaRows = [
    { label: `VGV Potencial (${fmt(c.areaTotalLotes, 0)} m\u00b2)`, v: c.vgvPotencial, pct: c.vgvPotencialPctSobreBruto },
    ...(c.permutaFisicaR > 0 ? [{ label: "(-) Permuta fisica", v: -c.permutaFisicaR, pct: c.permutaFisicaPctSobreBruto }] : []),
    { label: "Receita bruta (VGV)", v: c.vgvBruto, pct: 100 },
    { label: "(-) Impostos", v: -c.impostosR, pct: pctOf(c.impostosR, c.vgvBruto) },
    { label: "(-) Corretagem", v: -c.corretagemR, pct: pctOf(c.corretagemR, c.vgvBruto) },
    { label: "(-) Marketing e vendas", v: -c.marketingR, pct: pctOf(c.marketingR, c.vgvBruto) },
    { label: "= Receita liquida", v: c.receitaLiquida, pct: c.margemLiquidaPct },
    { label: "(-) Pagamento do terreno", v: -c.terrenoR, pct: pctOf(c.terrenoR, c.vgvBruto) },
    { label: "(-) Infraestrutura", v: -c.infraR, pct: pctOf(c.infraR, c.vgvBruto) },
    { label: "(-) Projetos", v: -c.projetoFinalR, pct: pctOf(c.projetoFinalR, c.vgvBruto) },
    { label: "(-) Licenciamento e Custos Ambientais", v: -c.licenciamentoFinalR, pct: pctOf(c.licenciamentoFinalR, c.vgvBruto) },
    { label: "(-) Registro", v: -c.registroR, pct: pctOf(c.registroR, c.vgvBruto) },
    { label: "(-) Manutencao pos-obra", v: -c.manutPosR, pct: pctOf(c.manutPosR, c.vgvBruto) },
    { label: "(-) Contingencias", v: -c.contingenciasR, pct: pctOf(c.contingenciasR, c.vgvBruto) },
    { label: "= Resultado operacional", v: c.resultadoOperacional, pct: c.margemOperacionalPct },
    { label: "(-) Permuta financeira", v: -c.permFinR, pct: pctOf(c.permFinR, c.vgvBruto) },
    { label: "(-) Administracao e gestao de carteira", v: -c.adminR, pct: pctOf(c.adminR, c.vgvBruto) },
    { label: "= Resultado final", v: c.resultadoFinal, pct: c.margemFinalPct },
  ];

  let proformaStart = 5;
  proformaRows.forEach((row, i) => {
    const r = wsProforma.getRow(proformaStart + i);
    r.getCell(1).value = row.label;
    r.getCell(2).value = row.v;
    r.getCell(3).value = c.areaLiquidaVenda ? row.v / c.areaLiquidaVenda : 0;
    r.getCell(4).value = (row.pct || 0) / 100;

    r.getCell(2).numFmt = moneyFmt;
    r.getCell(3).numFmt = moneyFmt;
    r.getCell(4).numFmt = pctFmt;

    let fill = "FFFFFF";
    let fontColor = "111827";
    let bold = false;

    const lbl = row.label;
    if (lbl.startsWith("=")) { fill = "DDF4D7"; bold = true; fontColor = "166534"; }
    else if (lbl.startsWith("(-) Permuta") || lbl.startsWith("(-) Admin") || lbl.startsWith("(-) Impostos") || lbl.startsWith("(-) Corretagem") || lbl.startsWith("(-) Marketing")) { fill = "FCE7E7"; }
    else if (lbl.startsWith("(-) Pagamento") || lbl.startsWith("(-) Infra") || lbl.startsWith("(-) Proj") || lbl.startsWith("(-) Licen") || lbl.startsWith("(-) Regist") || lbl.startsWith("(-) Manut") || lbl.startsWith("(-) Cont")) { fill = "F9FAFB"; }
    else if (lbl.includes("VGV Potencial") || lbl.includes("Receita bruta")) { fill = "DDEFF2"; bold = true; }
    else if (lbl.includes("Permuta fisica")) { fill = "FCE7E7"; }

    for (let col = 1; col <= 4; col++) {
      const cell = r.getCell(col);
      cell.font = { bold, color: { argb: fontColor } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: fill }
      };
      cell.alignment = {
        vertical: "middle",
        horizontal: col === 1 ? "left" : "right"
      };
      cell.border = {
        top: { style: "thin", color: { argb: "D1D5DB" } },
        left: { style: "thin", color: { argb: "D1D5DB" } },
        bottom: { style: "thin", color: { argb: "D1D5DB" } },
        right: { style: "thin", color: { argb: "D1D5DB" } }
      };
    }
  });

  // =========================
  // ABA 3 - PREMISSAS
  // =========================
  const wsPremissas = wb.addWorksheet("Premissas");

  wsPremissas.columns = [
    { width: 32 },
    { width: 18 },
    { width: 16 }
  ];

  wsPremissas.mergeCells("A1:C1");
  applyTitle(wsPremissas.getCell("A1"), "PREMISSAS DO ESTUDO", "0B5C6B");

  const premissas = [
    ["Nome do estudo", nomeEstudo, ""],
    ["Cidade / referencia", cidade, ""],
    ["Area total da gleba", state.study.urban.areaTotal, "m2"],
    ["APP", state.study.urban.percApp, "%"],
    ["Area remanescente", state.study.urban.percRem, "%"],
    ["Area nao edificante", state.study.urban.percNaoEd, "%"],
    ["Areas institucionais", state.study.urban.percInst, "%"],
    ["Areas publicas", state.study.urban.percPubl, "%"],
    ["Sistema viario", state.study.urban.percViario, "%"],
    ["Numero de lotes", state.study.product.nLotes, "un"],
    ["Area media do lote", state.study.product.areaMedia, "m2"],
    ["Preco por m2", state.study.product.precoM2, "R$/m2"],
    ["Permuta fisica", state.study.costs.permFisicaPct, "%"],
    ["Permuta financeira", state.study.costs.permFinPct, "%"],
    ["Terreno", state.study.costs.terrenoM2, "R$/m2"],
    ["Infraestrutura", state.study.costs.infraMode === "pct" ? state.study.costs.infraPct : state.study.costs.infraM2, state.study.costs.infraMode === "pct" ? "% VGV" : "R$/m2lot."],
    ["Modo infra", state.study.costs.infraMode, ""],
    ["Projetos", state.study.costs.projetoR, "R$"],
    ["Modo projetos", state.study.costs.projetoMode, ""],
    ["Licenciamento e Custos Ambientais", state.study.costs.licenciamentoR, "R$"],
    ["Modo licenciamento", state.study.costs.licenciamentoMode, ""],
    ["Registro", state.study.costs.registroR, "R$"],
    ["Manutencao", state.study.costs.manutPosPct, "%"],
    ["Marketing", state.study.costs.marketingPct, "%"],
    ["Corretagem", state.study.costs.corretagemPct, "%"],
    ["Administracao", state.study.costs.adminPct, "%"],
    ["Impostos vendas", state.study.costs.impostosPct, "%"],
    ["Contingencias", state.study.costs.contingenciasPct, "%"]
  ];

  let premRow = 3;
  premissas.forEach(([label, value, unidade]) => {
    applyLabel(wsPremissas.getCell(`A${premRow}`));
    wsPremissas.getCell(`A${premRow}`).value = label;

    const valCell = wsPremissas.getCell(`B${premRow}`);
    valCell.value = value;

    if (typeof value === "number") {
      if (String(unidade).includes("R$")) {
        applyValue(valCell, moneyFmt);
      } else {
        applyValue(valCell, numFmt);
      }
    } else {
      applyTextValue(valCell);
    }

    applyTextValue(wsPremissas.getCell(`C${premRow}`));
    wsPremissas.getCell(`C${premRow}`).value = unidade;

    premRow++;
  });

  // =========================
  // DOWNLOAD
  // =========================
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob(
    [buffer],
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
  );

  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `${safeId(nomeEstudo || "viabilidade_loteamento")}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(link.href);
}

function newStudy() {
  state.study = {
    studyId: "EST-001",
    nomeEstudo: "",
    cidade: "",
    urban: {
      areaTotal: 0,
      percApp: 0,
      percRem: 0,
      percNaoEd: 0,
      percInst: 5,
      percPubl: 15,
      percViario: 30,
    },
    product: {
      nLotes: 0,
      areaMedia: 0,
      precoM2: 0,
    },
    costs: {
      infraMode: "pct",
      infraM2: 0,
      infraPct: 30,
      projetoMode: "R$",
      projetoR: 0,
      projetoPct: 0,
      licenciamentoMode: "R$",
      licenciamentoR: 0,
      licenciamentoPct: 0,
      registroMode: "R$",
      registroR: 0,
      registroPct: 0,
      manutPosPct: 2,
      marketingPct: 2.5,
      corretagemPct: 5,
      adminPct: 0.75,
      impostosPct: 4,
      permFisicaPct: 0,
      permFinPct: 0,
      permFinExcImpostos: true,
      permFinExcCorretagem: true,
      permFinExcMarketing: false,
      permFinExcAdmin: true,
      contingenciasPct: 1.5,
      terrenoM2: 0,
    },
  };
  state.sheetMessage = "Novo estudo iniciado.";
  rerender();
}

function startProject(type) {
  state.projectType = type;
  if (!state.phase) {
    state.phase = "estudo_preliminar";
  }
  state.view = type === "loteamento" ? "loteamento" : "projectType";
  rerender();
}

function globalOverlays() {
  return `
    <button class="fab-feedback" onclick="openFeedback()" title="Enviar feedback">💬</button>
    ${state.showFeedbackModal ? feedbackModal() : ""}
  `;
}

function render() {
  const view = (() => {
    if (state.view === "home") return homeView();
    if (state.view === "projectType") return projectTypeView();
    if (state.view === "loteamento") return loteamentoView();
    if (state.view === "terrenos") return terrenosView();
    return homeView();
  })();
  return view + globalOverlays();
}

function attachEvents() {
  document.querySelectorAll("[data-path]").forEach((el) => {
    el.addEventListener("focus", (e) => {
      _activePath = e.target.getAttribute("data-path");
      _activeRaw = e.target.value;
    });

    el.addEventListener("blur", (e) => {
      const path = e.target.getAttribute("data-path");
      const isText = e.target.getAttribute("data-text") === "true";
      const isInteger = e.target.getAttribute("data-integer") === "true";
      if (!isText) {
        const val = getNestedValue(path);
        e.target.value = isInteger ? fmt(val, 0) : fmtBR(val);
      }
      _activePath = null;
      _activeRaw = "";
    });

    el.addEventListener("input", (e) => {
      const path = e.target.getAttribute("data-path");
      const isText = e.target.getAttribute("data-text") === "true";
      const isInteger = e.target.getAttribute("data-integer") === "true";
      _activePath = path;
      _activeRaw = e.target.value;

      if (isText) {
        setNested(path, e.target.value);
      } else {
        const raw = e.target.value.replace(/\./g, "").replace(",", ".");
        const parsed = num(raw);
        setNested(path, isInteger ? Math.trunc(parsed) : parsed);
      }
    });
  });

  document.querySelectorAll("[data-benchmark]").forEach((el) => {
    el.addEventListener("input", (e) => {
      const path = e.target.getAttribute("data-benchmark").split(".");
      let ref = state.benchmarks;
      for (let i = 0; i < path.length - 1; i++) ref = ref[path[i]];
      ref[path[path.length - 1]] = num(e.target.value);
    });
  });
}

function rerender() {
  const root = document.getElementById("root");
  const scrollY = window.scrollY;
  const savedPath = _activePath;
  const savedCursor = (() => {
    const el = document.activeElement;
    if (!el || !savedPath) return null;
    try {
      return el.selectionEnd;
    } catch {
      return null;
    }
  })();

  root.innerHTML = render();
  attachEvents();

  // Init Leaflet maps for terrain cards with KML
  if (state.view === "terrenos") {
    requestAnimationFrame(() => {
      state.terrenos.forEach(t => { if (t.kmlBase64) initTerrainMap(t); });
    });
  }

  window.scrollTo(0, scrollY);
  if (savedPath) {
    const el = document.querySelector(`[data-path="${savedPath}"]`);
    if (el) {
      el.focus();
      if (savedCursor !== null) {
        try {
          el.setSelectionRange(savedCursor, savedCursor);
        } catch {}
      }
    }
  }
}

// ─── Terrenos ─────────────────────────────────────────────────────────────────

const UF_LIST = ["AC","AL","AP","AM","BA","CE","DF","ES","GO","MA","MT","MS","MG","PA","PB","PR","PE","PI","RJ","RN","RS","RO","RR","SC","SP","SE","TO"];
const TERRENO_TEMAS = [
  { group: "UP", items: ["Regularização", "Urbitá"] },
  { group: "Novos Negócios", items: ["Vespasiano", "Alto Paraíso"] },
];

function generateTerrenoId() {
  return "TER-" + Date.now().toString(36).toUpperCase() + Math.random().toString(36).slice(2, 5).toUpperCase();
}

function setTerrenoField(key, value) {
  state.terrenoForm[key] = value;
  rerender();
}

function handleKmlUpload(input) {
  const file = input.files[0];
  if (!file) return;
  state.terrenoForm.kmlNome = file.name;
  rerender();
}

function handleTerrenoBoardUpload(input) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    state.terrenoForm.quadroImagem = String(reader.result || "");
    state.terrenoForm.quadroImagemNome = file.name;
    rerender();
  };
  reader.readAsDataURL(file);
}

function handleKmlUpload(input) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    state.terrenoForm.kmlBase64 = e.target.result;
    state.terrenoForm.kmlNome = file.name;
    rerender();
  };
  reader.readAsDataURL(file);
}

function initTerrainMap(terreno) {
  const el = document.getElementById("map-" + terreno.id);
  if (!el || el._mapInit) return;
  el._mapInit = true;
  const map = L.map(el, { zoomControl: true });
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    attribution: "© OpenStreetMap contributors",
    maxZoom: 19,
  }).addTo(map);
  const kmlLayer = omnivore.kml.parse(atob(terreno.kmlBase64.split(",")[1]));
  kmlLayer.addTo(map);
  kmlLayer.on("ready", () => {
    const bounds = kmlLayer.getBounds();
    if (bounds.isValid()) map.fitBounds(bounds);
  });
}

async function saveTerrenoLocal() {
  const f = state.terrenoForm;
  if (!f.nome.trim()) {
    state.terrenoMessage = "Informe o nome do terreno antes de salvar.";
    rerender();
    return;
  }
  // Colunas enviadas ao Google Sheets (aba "terrenos"):
  // id, createdAt, nome, cidade, estado, projeto, etapa, areaGleba, areaApp, kmlNome
  // fotoBase64 e kmlBase64 ficam apenas no localStorage
  const t = {
    id: generateTerrenoId(),
    createdAt: new Date().toISOString(),
    nome: f.nome.trim(),
    cidade: f.cidade.trim(),
    estado: f.estado,
    projeto: f.projeto.trim(),
    etapa: f.etapa,
    areaGleba: f.areaGleba,
    areaApp: f.areaApp,
    fotoBase64: f.fotoBase64,
    fotoNome: f.fotoNome,
    kmlBase64: f.kmlBase64,
    kmlNome: f.kmlNome,
  };
  state.terrenos.push(t);
  saveLocal("viab_terrenos", state.terrenos);

  // Envia para Sheets sem foto/kml (campos grandes demais para célula)
  try {
    const { fotoBase64: _foto, kmlBase64: _kml, ...semBinarios } = t;
    await postToAppsScript("saveTerrain", semBinarios);
    state.terrenoMessage = "Terreno salvo com sucesso.";
  } catch (err) {
    state.terrenoMessage = "Salvo localmente. Erro ao enviar para a planilha: " + err.message;
  }

  state.terrenoForm = { nome: "", cidade: "", estado: "", projeto: "", etapa: "", areaGleba: 0, areaApp: 0, fotoBase64: "", fotoNome: "", kmlBase64: "", kmlNome: "" };
  rerender();
}

function removeTerrenoLocal(id) {
  state.terrenos = state.terrenos.filter((t) => t.id !== id);
  saveLocal(STORAGE_KEYS.terrenos, state.terrenos);
  rerender();
}

function setTerrenoTema(tema) {
  state.terrenoTema = tema;
  state.terrenoSearch = "";
  state.terrenoSelecionadoId = null;
  state.terrenoMessage = "";
  rerender();
}

function setTerrenoSearch(v) {
  state.terrenoSearch = v || "";
  rerender();
}

function selectTerrenoMapa(id) {
  state.terrenoSelecionadoId = id;
  rerender();
}

function terrenoMapUrl(terreno) {
  const q = terreno
    ? `${terreno.nome} ${terreno.cidade || ""} ${terreno.estado || ""}`.trim()
    : "Brasil";
  return `https://www.google.com/maps?q=${encodeURIComponent(q)}&output=embed`;
}

function terrenoCadastroForm() {
  const f = state.terrenoForm;
  const cards = state.terrenos.map(t => `
    <div class="terreno-card">
      <div class="terreno-card-img">
        ${t.fotoBase64
          ? `<img src="${t.fotoBase64}" alt="${t.nome}" />`
          : `<div class="terreno-card-no-img">📍</div>`}
      </div>
      <div class="terreno-card-body">
        <div class="terreno-card-title">${t.nome}</div>
        <div class="terreno-card-sub">${t.cidade}${t.estado ? " · " + t.estado : ""}${t.projeto ? " · " + t.projeto : ""}</div>
        <div class="terreno-card-meta">
          <span>Gleba: ${fmt(t.areaGleba, 0)} m²</span>
          <span>APP: ${fmt(t.areaApp, 0)} m²</span>
          ${t.etapa ? `<span>Etapa ${t.etapa}</span>` : ""}
        </div>
        ${t.kmlBase64 ? `<div id="map-${t.id}" class="terreno-card-map"></div>` : ""}
      </div>
      <button class="terreno-card-del" onclick="removeTerrenoLocal('${t.id}')" title="Remover">✕</button>
    </div>
  `).join("");

  return `
    <div class="section">
      <div class="section-head head-blue">Cadastro único de terreno</div>
      <div class="section-body">
        <div class="field">
          <label>Tema</label>
          <div class="input-wrap full">
            <select class="inp" onchange="setTerrenoField('tema',this.value)" style="cursor:pointer">
              <option value="">Selecione…</option>
              ${TERRENO_TEMAS.flatMap(group => group.items).map(item => `<option value="${item}"${f.tema === item ? " selected" : ""}>${item}</option>`).join("")}
            </select>
          </div>
        </div>
        <div class="field">
          <label>Nome</label>
          <div class="input-wrap full">
            <input class="inp text" type="text" value="${f.nome}"
              oninput="setTerrenoField('nome',this.value)" placeholder="Ex: Gleba Santa Clara" />
          </div>
        </div>
        <div style="display:grid;grid-template-columns:1fr .6fr;gap:14px">
          <div class="field">
            <label>Cidade</label>
            <div class="input-wrap full">
              <input class="inp text" type="text" value="${f.cidade}" oninput="setTerrenoField('cidade',this.value)" />
            </div>
          </div>
          <div class="field">
            <label>Estado (UF)</label>
            <div class="input-wrap full">
              <select class="inp" onchange="setTerrenoField('estado',this.value)" style="cursor:pointer">
                <option value="">—</option>
                ${UF_LIST.map(uf => `<option value="${uf}"${f.estado === uf ? " selected" : ""}>${uf}</option>`).join("")}
              </select>
            </div>
          </div>
        </div>
        <div class="field">
          <label>Área da gleba</label>
          <div class="input-wrap full">
            <input class="inp" type="text" inputmode="decimal"
              value="${f.areaGleba || ""}"
              oninput="setTerrenoField('areaGleba',parseFloat(this.value.replace(/\\./g,'').replace(',','.'))||0)" />
            <span class="affix">m²</span>
          </div>
        </div>
        <div class="field">
          <label>Arquivo KML</label>
          <input type="file" accept=".kml,.kmz,application/vnd.google-earth.kml+xml,application/vnd.google-earth.kmz" class="terreno-file-input"
            onchange="handleKmlUpload(this)" />
          ${f.kmlNome ? `<div class="muted" style="margin-top:4px">Arquivo selecionado: ${f.kmlNome}</div>` : ""}
        </div>
        <div class="field">
          <label>Quadro visual (imagem + notas)</label>
          <div class="terreno-board">
            <div class="terreno-board-media">
              <input id="terreno-board-upload" type="file" accept="image/*" class="terreno-board-upload-input" onchange="handleTerrenoBoardUpload(this)" />
              <label for="terreno-board-upload" class="terreno-board-dropzone">
                ${f.quadroImagem
                  ? `<img src="${f.quadroImagem}" alt="Prévia do quadro do terreno" class="terreno-board-image" />`
                  : `<div><strong>Adicionar imagem do terreno</strong><span>Arraste ou clique para carregar uma referência visual (Notion/Trello style).</span></div>`}
              </label>
              ${f.quadroImagem ? `<div class="btn-row" style="margin-top:8px"><button class="btn gray" type="button" onclick="clearTerrenoBoardImage()">Remover imagem</button></div>` : ""}
            </div>
            <textarea class="feedback-text terreno-board-notes" placeholder="Anotações do quadro visual, checklist, ideias de produto, etc." oninput="setTerrenoField('quadroNotas',this.value)">${f.quadroNotas || ""}</textarea>
          </div>
        </div>
        ${state.terrenoMessage ? `<div class="${state.terrenoMessage.startsWith("Salvo") || state.terrenoMessage.includes("sucesso") ? "notice" : "error"}" style="margin-bottom:10px">${state.terrenoMessage}</div>` : ""}
        <div class="btn-row">
          <button class="btn green" onclick="saveTerrenoLocal()">Salvar terreno</button>
        </div>
      </div>
    </div>
  `;
}

function terrenosView() {
  if (!state.terrenoTema) {
    return `
      <div class="view">
        <div class="page-header">
          <div class="top-breadcrumb">
            <button class="nav-btn" onclick="setView('home')">← Home</button>
            <span class="crumb">Terrenos</span>
          </div>
        </div>
        <div class="container">
          <div class="card-grid-2 bottom">
            <div class="section">
              <div class="section-head head-primary">Terrenos</div>
              <div class="section-body">
                <p class="muted" style="margin-bottom:14px">Selecione um tema para abrir a visualização em mapa.</p>
                ${TERRENO_TEMAS.map(group => `
                  <div class="terreno-group">
                    <h4>${group.group}</h4>
                    <div class="btn-row">
                      ${group.items.map(item => `<button class="btn blue" onclick="setTerrenoTema('${item}')">${item}</button>`).join("")}
                    </div>
                  </div>
                `).join("")}
              </div>
            </div>
            ${terrenoCadastroForm()}
          </div>
        </div>
      </div>
    `;
  }

  const terrenosTema = state.terrenos.filter((t) => t.tema === state.terrenoTema);
  const filtered = terrenosTema.filter((t) => {
    const q = state.terrenoSearch.trim().toLowerCase();
    if (!q) return true;
    const txt = `${t.nome} ${t.cidade || ""} ${t.estado || ""}`.toLowerCase();
    return txt.includes(q);
  });
  const selected = filtered.find((t) => t.id === state.terrenoSelecionadoId) || filtered[0] || null;
  if (!state.terrenoSelecionadoId && selected) {
    state.terrenoSelecionadoId = selected.id;
  }
  return `
    <div class="view">
      <div class="page-header">
        <div class="top-breadcrumb">
          <button class="nav-btn" onclick="setTerrenoTema(null)">← Terrenos</button>
          <span class="crumb">${state.terrenoTema}</span>
        </div>
      </div>

      <div class="container">
        <div class="terrenos-layout">
          <div class="section">
            <div class="section-head head-primary">Áreas cadastradas (${filtered.length})</div>
            <div class="section-body">
              <div class="field">
                <label>Pesquisar</label>
                <div class="input-wrap full">
                  <input class="inp text" type="text" value="${state.terrenoSearch}" oninput="setTerrenoSearch(this.value)"
                    placeholder="Buscar terreno por nome/cidade..." />
                </div>
              </div>
              <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px">
                <div class="field">
                  <label>Projeto</label>
                  <div class="input-wrap full">
                    <input class="inp text" type="text" value="${f.projeto}"
                      oninput="setTerrenoField('projeto',this.value)" placeholder="Ex: Residencial Bela Vista" />
                  </div>
                </div>
                <div class="field">
                  <label>Etapa</label>
                  <div class="input-wrap full">
                    <input class="inp" type="text" inputmode="numeric" value="${f.etapa}"
                      oninput="setTerrenoField('etapa',this.value.replace(/\\D/g,''))" placeholder="Ex: 1" />
                  </div>
                </div>
              </div>
              <div style="display:grid;grid-template-columns:1fr .6fr;gap:14px">
                <div class="field">
                  <label>Cidade</label>
                  <div class="input-wrap full">
                    <input class="inp text" type="text" value="${f.cidade}"
                      oninput="setTerrenoField('cidade',this.value)" />
                  </div>
                </div>
                <div class="field">
                  <label>Estado (UF)</label>
                  <div class="input-wrap full">
                    <select class="inp" onchange="setTerrenoField('estado',this.value)" style="cursor:pointer">
                      <option value="">—</option>
                      ${UF_LIST.map(uf => `<option value="${uf}"${f.estado === uf ? " selected" : ""}>${uf}</option>`).join("")}
                    </select>
                  </div>
                </div>
              </div>
              <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px">
                <div class="field">
                  <label>Área da gleba</label>
                  <div class="input-wrap full">
                    <input class="inp" type="text" inputmode="decimal"
                      value="${f.areaGleba || ""}"
                      oninput="setTerrenoField('areaGleba',parseFloat(this.value.replace(/\\./g,'').replace(',','.'))||0)" />
                    <span class="affix">m²</span>
                  </div>
                </div>
                <div class="field">
                  <label>APP (Preservação Ambiental)</label>
                  <div class="input-wrap full">
                    <input class="inp" type="text" inputmode="decimal"
                      value="${f.areaApp || ""}"
                      oninput="setTerrenoField('areaApp',parseFloat(this.value.replace(/\\./g,'').replace(',','.'))||0)" />
                    <span class="affix">m²</span>
                  </div>
                </div>
              </div>
              <div class="field">
                <label>Foto da gleba</label>
                <input type="file" accept="image/*" class="terreno-file-input"
                  onchange="handleFotoUpload(this)" />
                ${f.fotoBase64 ? `<div class="terreno-preview"><img class="terreno-thumb-large" src="${f.fotoBase64}" alt="${f.fotoNome}" /></div>` : ""}
              </div>
              <div class="field">
                <label>Arquivo KML do terreno</label>
                <input type="file" accept=".kml" class="terreno-file-input"
                  onchange="handleKmlUpload(this)" />
                ${f.kmlNome ? `<div class="muted" style="font-size:12px;margin-top:4px">📎 ${f.kmlNome}</div>` : ""}
              </div>
              ${state.terrenoMessage ? `<div class="${state.terrenoMessage.startsWith("Salvo") || state.terrenoMessage.includes("sucesso") ? "notice" : "error"}" style="margin-bottom:10px">${state.terrenoMessage}</div>` : ""}
              <div class="btn-row">
                <button class="btn green" onclick="saveTerrenoLocal()">Salvar terreno</button>
              </div>
            </div>
          </div>

          <div>
            <div class="section">
              <div class="section-head head-orange">Terrenos cadastrados (${state.terrenos.length})</div>
              <div class="section-body">
                ${state.terrenos.length === 0
                  ? `<div class="muted">Nenhum terreno cadastrado ainda.</div>`
                  : `<div class="terreno-cards-grid">${cards}</div>`
                }
              </div>
            </div>
          </div>
        </div>
      </div>

    </div>
  `;
}

function openTerrenoPicker() {
  if (state.terrenos.length === 0) {
    state.sheetMessage = "Nenhum terreno cadastrado. Acesse 'Terrenos' na home para adicionar.";
    rerender();
    return;
  }
  state.showTerrenoPickerModal = true;
  rerender();
}

function closeTerrenoPicker(e) {
  if (!e || (e.target && e.target.classList && e.target.classList.contains("modal-overlay"))) {
    state.showTerrenoPickerModal = false;
    rerender();
  }
}

function selectTerrenoForStudy(id) {
  const t = state.terrenos.find((x) => x.id === id);
  if (!t) return;
  if (t.cidade) state.study.cidade = t.cidade;
  if (t.areaGleba) state.study.urban.areaTotal = t.areaGleba;
  state.sheetMessage = `Terreno "${t.nome}" aplicado ao estudo.`;
  state.showTerrenoPickerModal = false;
  rerender();
}

function terrenoPickerModal() {
  return `
    <div class="modal-overlay" onclick="closeTerrenoPicker(event)">
      <div class="modal study-modal" onclick="event.stopPropagation()">
        <h3>Selecionar terreno</h3>
        <p>Escolha um terreno para pré-preencher os dados do estudo.</p>
        <div class="study-list">
          ${state.terrenos.map(t => `
            <button class="study-item" onclick="selectTerrenoForStudy('${t.id}')">
              <strong>${t.nome}</strong>
              <span>${t.cidade}${t.estado ? " · " + t.estado : ""} · Gleba: ${fmt(t.areaGleba)} m²${t.tema ? " · Tema: " + t.tema : ""}</span>
            </button>
          `).join("")}
        </div>
        <div class="spacer"></div>
        <div class="footer-actions">
          <button class="btn gray" onclick="closeTerrenoPicker()">Fechar</button>
        </div>
      </div>
    </div>
  `;
}

function bootApp() {
  const root = document.getElementById("root");
  if (!root) {
    console.error("Elemento #root não encontrado.");
    return;
  }

  window.state = state;
  window.rerender = rerender;
  window.setView = setView;
  window.openBenchmarks = openBenchmarks;
  window.closeBenchmarks = closeBenchmarks;
  window.saveBenchmarksLocal = saveBenchmarksLocal;
  window.syncBenchmarksToSheet = syncBenchmarksToSheet;
  window.loadBenchmarksFromSheet = loadBenchmarksFromSheet;
  window.startProject = startProject;
  window.newStudy = newStudy;
  window.saveScenario = saveScenario;
  window.clearScenarios = clearScenarios;
  window.exportPDF = exportPDF;
  window.exportExcel = exportExcel;
  window.sendStudyToSheet = sendStudyToSheet;
  window.openStudyPicker = openStudyPicker;
  window.closeStudyPicker = closeStudyPicker;
  window.applyStudyFromSheet = applyStudyFromSheet;
  window.openFeedback = openFeedback;
  window.closeFeedback = closeFeedback;
  window.submitFeedback = submitFeedback;
  window.setTerrenoField = setTerrenoField;
  window.handleKmlUpload = handleKmlUpload;
  window.handleTerrenoBoardUpload = handleTerrenoBoardUpload;
  window.clearTerrenoBoardImage = clearTerrenoBoardImage;
  window.saveTerrenoLocal = saveTerrenoLocal;
  window.removeTerrenoLocal = removeTerrenoLocal;
  window.setTerrenoTema = setTerrenoTema;
  window.setTerrenoSearch = setTerrenoSearch;
  window.selectTerrenoMapa = selectTerrenoMapa;
  window.openTerrenoPicker = openTerrenoPicker;
  window.closeTerrenoPicker = closeTerrenoPicker;
  window.selectTerrenoForStudy = selectTerrenoForStudy;

  try {
    rerender();
  } catch (error) {
    console.error("Erro ao renderizar a aplicação:", error);
    root.innerHTML = `
      <div style="padding:24px;font-family:Arial,sans-serif">
        <h2>Erro ao carregar a aplicação</h2>
        <p>Abra o console do navegador para ver os detalhes.</p>
        <pre style="white-space:pre-wrap;background:#f5f5f5;padding:12px;border-radius:8px;">${String(error && error.message ? error.message : error)}</pre>
      </div>
    `;
  }
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", bootApp);
} else {
  bootApp();
}

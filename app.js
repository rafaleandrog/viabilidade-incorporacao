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
  feedbackNome: "",
  feedbackText: "",
  feedbackMessage: "",
  study: {
    studyId: "",
    nomeEstudo: "",
    cidade: "",
    terrenoId: "",
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
      registroR: 0,
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
  terrenos: loadLocal(STORAGE_KEYS.terrenos, []),
  terrenoForm: { tema: "", nome: "", cidade: "", estado: "", projeto: "", etapa: "", areaGleba: 0, areaApp: 0, fotoBase64: "", fotoNome: "", quadroImagem: "", quadroImagemNome: "", quadroNotas: "", apeloNotas: {}, apeloDetalhes: {}, lote: "", zona: "", setor: "", codilog: "", caBas: "N/I", caMax: "N/I", gabarito: "N.A.", cotaParte: "N/I", incentivo: "N/I", valorRef: 0, viabilidadePct: 0, coordinates: [] },
  terrenoMessage: "",
  terrenoTema: null,
  terrenoSearch: "",
  terrenoSelecionadoId: null,
  showTerrenoPickerModal: false,
  appsScriptActionCache: {},
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

function getCachedAction(cacheKey) {
  return state.appsScriptActionCache[cacheKey] || "";
}

function setCachedAction(cacheKey, action) {
  state.appsScriptActionCache[cacheKey] = action;
}

async function runActionWithCache(cacheKey, actions, executor) {
  const cached = getCachedAction(cacheKey);
  const ordered = cached ? [cached, ...actions.filter((a) => a !== cached)] : [...actions];
  let lastError = null;
  for (const action of ordered) {
    try {
      const result = await executor(action);
      setCachedAction(cacheKey, action);
      return result;
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError || new Error("Ação não suportada pelo Apps Script.");
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
  return (txt || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Za-z0-9_-]+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");
}

let _activePath = null;
let _activeRaw = "";
let _deckMap = null;
let _mapIs3D = false;
let _mapSavedVS = null;
let _mapPavimentos = {};

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
    terrenoId: "",
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
      registroR: 0,
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
  const registroR = num(c.registroR);
  const manutPosR = infraR * num(c.manutPosPct) / 100;

  const custoObrasTotal = infraR + projetoFinalR + licenciamentoFinalR + registroR + manutPosR;
  const contingenciasR = custoObrasTotal * num(c.contingenciasPct) / 100;
  const resultadoOperacional = receitaLiquida - terrenoR - custoObrasTotal - contingenciasR;

  // Após resultado operacional: admin, permuta financeira (base ajustável)
  const adminR = vgvBruto * num(c.adminPct) / 100;
  const permFinBrutoR = vgvBruto * num(c.permFinPct) / 100;
  let basePermFinPctReducao = 0;
  if (c.permFinExcImpostos)   basePermFinPctReducao += num(c.impostosPct);
  if (c.permFinExcCorretagem) basePermFinPctReducao += num(c.corretagemPct);
  if (c.permFinExcMarketing)  basePermFinPctReducao += num(c.marketingPct);
  if (c.permFinExcAdmin)      basePermFinPctReducao += num(c.adminPct);
  const basePermFinFator = Math.max(0, 1 - (basePermFinPctReducao / 100));
  const permFinR = permFinBrutoR * basePermFinFator;
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
    permFinBrutoR,
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
  if (view === "terrenos") {
    void loadTerrenosFromSheet();
  }
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
  const terrenoVinculado = state.terrenos.find((t) => t.id === state.study.terrenoId);

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
        </div>
      </div>

      <div class="container print-area">
        <div class="field">
          <label>Terreno (planilha)</label>
          <div class="btn-row" style="align-items:center;gap:8px;flex-wrap:wrap">
            <div class="input-wrap full" style="max-width:720px">
              <select class="inp" onchange="handleStudyTerrenoChange(this.value)" style="cursor:pointer">
                <option value="">— Selecione um terreno —</option>
                ${state.terrenos.map((t) => `<option value="${t.id}"${state.study.terrenoId === t.id ? " selected" : ""}>${t.nome}${t.cidade ? ` · ${t.cidade}` : ""}</option>`).join("")}
              </select>
            </div>
            <button class="btn blue" type="button" onclick="refreshTerrenosForStudy()">↻ Atualizar</button>
          </div>
          ${terrenoVinculado ? `<div class="terreno-tag" style="margin-top:6px">📍 Terreno vinculado: <strong>${terrenoVinculado.nome}</strong>${terrenoVinculado.cidade ? ` · ${terrenoVinculado.cidade}` : ""}</div>` : ""}
        </div>
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
            ${(() => {
              const ratio = c.areaLotesVend > 0 ? c.areaTotalLotes / c.areaLotesVend : 0;
              const pct   = Math.min(ratio, 1) * 100;
              const over  = ratio > 1;
              return `
                <div style="margin-top:12px;padding-top:12px;border-top:1px solid var(--border)">
                  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:5px">
                    <span style="font-size:11px;font-weight:600;color:var(--muted)">Área total lotes vs. área vendável</span>
                    <span style="font-size:11px;font-weight:700;color:${over ? '#dc2626' : 'var(--primary)'}">${fmt(c.areaTotalLotes,0)} m² / ${fmt(c.areaLotesVend,0)} m² (${fmt(ratio * 100, 1)}%)</span>
                  </div>
                  <div style="height:8px;background:#e5e7eb;border-radius:6px;overflow:hidden">
                    <div style="height:100%;width:${pct}%;background:${over ? '#dc2626' : 'var(--primary)'};border-radius:6px;transition:width .3s"></div>
                  </div>
                  ${over ? `<div style="margin-top:6px;font-size:11px;color:#dc2626;font-weight:600">⚠ Área total dos lotes (${fmt(c.areaTotalLotes,0)} m²) excede a área vendável (${fmt(c.areaLotesVend,0)} m²). Reduza o número ou a área média dos lotes.</div>` : ""}
                </div>
              `;
            })()}
            <div style="display:grid;grid-template-columns:.65fr .9fr 1fr .65fr .8fr 1.2fr;gap:14px;margin-top:12px">
              ${calcDisplay("Preço médio do lote", rs(c.ticketMedio))}
              ${calcDisplay("Custo aquisição terreno", rs(c.terrenoR))}
              ${calcDisplay("Área entregue — perm. física", `${fmt(c.areaPermutaFis)} m²`)}
              ${calcDisplay("Valor líquido perm. financeira", rs(c.permFinR))}
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
                [{id:"m2", label:"R$/m²lot."}, {id:"pct", label:"% VGV"}],
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
              ${inputField("Registro", "costs.registroR", state.study.costs.registroR, { prefix: "R$", full: true })}
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
    <div class="compare-head muted"><strong>${item.study.nomeEstudo || "Sem nome"}</strong><br>${new Date(item.savedAt).toLocaleString("pt-BR")}</div>
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
    <div class="compare-head muted"><strong>Base: Cenário A</strong><br>Comparado com Cenário B</div>
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
  const b = normalizeBenchmarks(state.benchmarks);
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
  const safe = {
    urban: {
      ...BENCHMARK_TEMPLATE[type].urban,
      ...((data && data.urban) || {}),
    },
    financial: {
      ...BENCHMARK_TEMPLATE[type].financial,
      ...((data && data.financial) || {}),
    },
  };
  return `
    <div class="mini-card">
      <h4>${title}</h4>
      ${benchmarkInput("Área lotes vendáveis (%)", type, "urban", "areaLotesPct", safe.urban.areaLotesPct)}
      ${benchmarkInput("Margem final / VGV (%)", type, "financial", "margemFinalPct", safe.financial.margemFinalPct)}
      ${benchmarkInput("Margem operacional (%)", type, "financial", "margemOperacionalPct", safe.financial.margemOperacionalPct)}
      ${benchmarkInput("Custo obras / VGV (%)", type, "financial", "custoObrasPct", safe.financial.custoObrasPct)}
    </div>
  `;
}

function normalizeBenchmarks(raw) {
  const base = clone(BENCHMARK_TEMPLATE);
  const incoming = raw && typeof raw === "object" ? raw : {};
  for (const type of Object.keys(base)) {
    const src = incoming[type] || {};
    base[type] = {
      urban: { ...base[type].urban, ...(src.urban || {}) },
      financial: { ...base[type].financial, ...(src.financial || {}) },
    };
  }
  return base;
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

async function postToAppsScriptWithFallback(actions, payload) {
  const actionList = Array.isArray(actions) ? actions : [actions];
  let lastError = null;

  for (const action of actionList) {
    try {
      return await postToAppsScript(action, payload);
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError || new Error("Falha ao enviar dados para o Apps Script.");
}

function normalizeTerrenoFromSheet(raw) {
  if (!raw || typeof raw !== "object") return null;
  const nome = String(raw.nome || raw.name || "").trim();
  if (!nome) return null;
  const id = String(raw.id || raw.terrenoId || `terreno_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`);
  return {
    id,
    timestamp: raw.timestamp || raw.createdAt || new Date().toISOString(),
    nome,
    cidade: String(raw.cidade || raw.city || "").trim(),
    estado: String(raw.estado || raw.uf || "").trim(),
    projeto: String(raw.projeto || raw.project || "").trim(),
    etapa: String(raw.etapa || raw.stage || "").trim(),
    areaGleba: num(raw.areaGleba || raw.gleba || 0),
    areaApp: num(raw.areaApp || raw.app || 0),
    fotoBase64: "",
    fotoNome: "",
    quadroImagem: "",
    quadroNotas: String(raw.quadroNotas || raw.notas || "").trim(),
    apeloNotas: {},
    apeloDetalhes: {},
  };
}

async function loadTerrenosFromSheet() {
  const actions = ["listTerrenos", "listTerrains", "getTerrenos", "getTerrains"];
  try {
    const res = await runActionWithCache("get:terrenos", actions, (action) => getFromAppsScript(action));
    const rows = Array.isArray(res.data) ? res.data : [];
    const normalized = rows.map(normalizeTerrenoFromSheet).filter(Boolean);
    state.terrenos = normalized;
    saveLocal(STORAGE_KEYS.terrenos, state.terrenos);
    state.terrenoSheetMessage = normalized.length
      ? `Terrenos carregados da planilha (${normalized.length}).`
      : "Nenhum terreno encontrado na planilha.";
    rerender();
    return normalized;
  } catch (lastError) {
    state.terrenoSheetMessage = `Erro ao buscar terrenos na planilha: ${lastError ? lastError.message : "ação não suportada no Apps Script."}`;
    rerender();
    throw lastError || new Error("Não foi possível carregar terrenos da planilha.");
  }
}

function flattenBenchmarkRows(rows) {
  const out = {};
  rows.forEach((row) => {
    const tipo = String(row.tipo || row.type || "").toLowerCase();
    if (!tipo || !BENCHMARK_TEMPLATE[tipo]) return;
    out[tipo] = out[tipo] || { urban: {}, financial: {} };
    if (row.areaLotesPct != null) out[tipo].urban.areaLotesPct = num(row.areaLotesPct);
    if (row.margemFinalPct != null) out[tipo].financial.margemFinalPct = num(row.margemFinalPct);
    if (row.margemOperacionalPct != null) out[tipo].financial.margemOperacionalPct = num(row.margemOperacionalPct);
    if (row.custoObrasPct != null) out[tipo].financial.custoObrasPct = num(row.custoObrasPct);
  });
  return out;
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
  state.benchmarks = normalizeBenchmarks(state.benchmarks);
  saveLocal(STORAGE_KEYS.benchmarks, state.benchmarks);
  state.benchmarkMessage = "Benchmarks salvos com sucesso no navegador.";
  rerender();
}

async function syncBenchmarksToSheet() {
  try {
    state.benchmarks = normalizeBenchmarks(state.benchmarks);
    saveLocal(STORAGE_KEYS.benchmarks, state.benchmarks);
    await runActionWithCache(
      "post:saveBenchmarks",
      ["saveBenchmarks", "saveBenchmark"],
      (action) => postToAppsScript(action, state.benchmarks),
    );
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
    const res = await runActionWithCache(
      "get:benchmarks",
      ["getBenchmarks", "listBenchmarks", "getBenchmark"],
      (action) => getFromAppsScript(action),
    );
    if (res && res.data) {
      state.benchmarks = Array.isArray(res.data)
        ? normalizeBenchmarks(flattenBenchmarkRows(res.data))
        : normalizeBenchmarks(res.data);
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
      nome: state.feedbackNome.trim(),
      texto: state.feedbackText.trim(),
      studyId: state.study.studyId || "",
      projectType: state.projectType || "",
    });
    state.feedbackMessage = "Feedback enviado com sucesso. Obrigado!";
    state.feedbackNome = "";
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
        <input class="inp text" type="text" placeholder="Seu nome (opcional)"
          value="${state.feedbackNome}" oninput="state.feedbackNome=this.value" style="margin-bottom:8px" />
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
  wb.title = safeId(state.study.nomeEstudo || "Viabilidade_Loteamento");

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
    ...(c.permutaFisicaR > 0 ? [{ label: `(-) Permuta fisica (${fmt(c.areaPermutaFis, 0)} m2)`, v: -c.permutaFisicaR, pct: c.permutaFisicaPctSobreBruto }] : []),
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
    ["Projetos", state.study.costs.projetoMode === "pct" ? state.study.costs.projetoPct : state.study.costs.projetoR, state.study.costs.projetoMode === "pct" ? "% VGV" : "R$"],
    ["Modo projetos", state.study.costs.projetoMode, ""],
    ["Licenciamento e Custos Ambientais", state.study.costs.licenciamentoMode === "pct" ? state.study.costs.licenciamentoPct : state.study.costs.licenciamentoR, state.study.costs.licenciamentoMode === "pct" ? "% VGV" : "R$"],
    ["Modo licenciamento", state.study.costs.licenciamentoMode, ""],
    ["Registro", state.study.costs.registroMode === "pct" ? state.study.costs.registroPct : state.study.costs.registroR, state.study.costs.registroMode === "pct" ? "% VGV" : "R$"],
    ["Modo registro", state.study.costs.registroMode, ""],
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
    terrenoId: "",
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
      registroR: 0,
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

  document.querySelectorAll("[data-terreno-path]").forEach((el) => {
    const isNum    = el.getAttribute("data-terreno-num")    === "true";
    const isDigits = el.getAttribute("data-terreno-digits") === "true";
    el.addEventListener("focus", (e) => {
      _activePath = "terreno:" + e.target.getAttribute("data-terreno-path");
      _activeRaw  = e.target.value;
    });
    el.addEventListener("blur", (e) => {
      if (isNum) {
        const key = e.target.getAttribute("data-terreno-path");
        const v = state.terrenoForm[key];
        e.target.value = v ? fmtBR(v) : "";
      }
      _activePath = null;
      _activeRaw  = "";
    });
    el.addEventListener("input", (e) => {
      const key = e.target.getAttribute("data-terreno-path");
      _activePath = "terreno:" + key;
      _activeRaw  = e.target.value;
      if (isNum) {
        state.terrenoForm[key] = parseFloat(e.target.value.replace(/\./g, "").replace(",", ".")) || 0;
      } else if (isDigits) {
        const cleaned = e.target.value.replace(/\D/g, "");
        state.terrenoForm[key] = cleaned;
        if (e.target.value !== cleaned) { e.target.value = cleaned; _activeRaw = cleaned; }
      } else {
        state.terrenoForm[key] = e.target.value;
      }
      rerender();
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

  // Initialize deck.gl map when on terrenos tema view
  if (state.view === "terrenos" && state.terrenoTema) {
    setTimeout(initDeckMap, 0);
  } else {
    destroyDeckMap();
  }
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

  try {
    root.innerHTML = render();
    attachEvents();
  } catch (error) {
    console.error("Erro ao renderizar:", error);
    root.innerHTML = `<div style="padding:24px;font-family:Arial,sans-serif">
      <h2>Erro ao renderizar</h2>
      <pre style="background:#f5f5f5;padding:12px;border-radius:8px;white-space:pre-wrap">${String(error && error.message ? error.message : error)}</pre>
      <button onclick="location.reload()" style="margin-top:12px;padding:8px 16px;cursor:pointer">Recarregar</button>
    </div>`;
    return;
  }

  window.scrollTo(0, scrollY);
  if (savedPath) {
    let el = document.querySelector(`[data-path="${savedPath}"]`);
    if (!el && savedPath.startsWith("terreno:")) {
      el = document.querySelector(`[data-terreno-path="${savedPath.slice(8)}"]`);
    }
    if (!el && savedPath === "_search") {
      el = document.getElementById("terreno-search-inp");
    }
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

const APELO_FATORES = [
  "Localização", "Acesso viário", "Proximidade centros",
  "Infraestrutura entorno", "Vetor de crescimento", "Topografia",
  "Condição jurídica", "Potencial de valorização",
  "Concorrência", "Demanda estrutural",
];

function generateTerrenoId() {
  return "TER-" + Date.now().toString(36).toUpperCase() + Math.random().toString(36).slice(2, 5).toUpperCase();
}

function setTerrenoField(key, value) {
  state.terrenoForm[key] = value;
  rerender();
}

function handleTerrenoBoardUpload(input) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    state.terrenoForm.fotoBase64 = String(reader.result || "");
    state.terrenoForm.fotoNome = file.name;
    rerender();
  };
  reader.readAsDataURL(file);
}

function clearTerrenoBoardImage() {
  state.terrenoForm.quadroImagem = "";
  state.terrenoForm.quadroImagemNome = "";
  rerender();
}

function handleFotoUpload(input) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    state.terrenoForm.fotoBase64 = String(e.target.result || "");
    state.terrenoForm.fotoNome = file.name;
    rerender();
  };
  reader.readAsDataURL(file);
}

function setTerrenoApelo(fator, nota) {
  state.terrenoForm.apeloNotas[fator] = nota;
  rerender();
}

function setTerrenoApeloDetalhe(fator, texto) {
  state.terrenoForm.apeloDetalhes[fator] = texto;
}

function renderStars(fator, nota) {
  return Array.from({ length: 5 }, (_, i) =>
    `<span class="apelo-star${i < nota ? " filled" : ""}" onclick="setTerrenoApelo('${fator}',${i + 1})">${i < nota ? "★" : "☆"}</span>`
  ).join("");
}

async function saveTerrenoLocal() {
  const f = state.terrenoForm;
  if (!f.nome.trim()) {
    state.terrenoMessage = "Informe o nome do terreno antes de salvar.";
    rerender();
    return;
  }
  const t = {
    id: generateTerrenoId(),
    timestamp: new Date().toISOString(),
    nome: f.nome.trim(),
    cidade: f.cidade.trim(),
    estado: f.estado,
    projeto: f.projeto.trim(),
    etapa: f.etapa,
    areaGleba: f.areaGleba,
    areaApp: f.areaApp,
    fotoBase64: f.fotoBase64,
    fotoNome: f.fotoNome,
    quadroImagem: f.quadroImagem,
    quadroNotas: f.quadroNotas,
    apeloNotas: { ...f.apeloNotas },
    apeloDetalhes: { ...f.apeloDetalhes },
    lote: f.lote || "",
    zona: f.zona || "",
    setor: f.setor || "",
    codilog: f.codilog || "",
    caBas: f.caBas || "N/I",
    caMax: f.caMax || "N/I",
    gabarito: f.gabarito || "N.A.",
    cotaParte: f.cotaParte || "N/I",
    incentivo: f.incentivo || "N/I",
    valorRef: f.valorRef || 0,
    viabilidadePct: f.viabilidadePct || 0,
    coordinates: f.coordinates ? [...f.coordinates] : [],
  };
  if (!t.projeto && state.terrenoTema) {
    t.projeto = state.terrenoTema;
  }
  // Envia para Sheets e só persiste local após confirmação de sucesso.
  try {
    const terrenoPayload = {
      id: t.id,
      timestamp: t.timestamp,
      createdAt: t.timestamp,
      nome: t.nome,
      cidade: t.cidade,
      uf: t.estado,
      estado: t.estado,
      projeto: t.projeto,
      etapa: t.etapa,
      areaGleba: t.areaGleba,
      areaApp: t.areaApp,
      kmlNomeArquivo: "",
      kmlNome: "",
      quadroImagemUrl: "",
      imagemUrl: "",
      quadroNotas: t.quadroNotas,
      notas: t.quadroNotas,
      apeloNotas: { ...t.apeloNotas },
      apeloDetalhes: { ...t.apeloDetalhes },
      apeloComercial: APELO_FATORES.map((fator) => ({
        fator,
        nota: num(t.apeloNotas[fator] || 0),
        detalhe: String(t.apeloDetalhes[fator] || ""),
      })),
      lote: t.lote || "",
      zona: t.zona || "",
      setor: t.setor || "",
      codilog: t.codilog || "",
      caBas: t.caBas || "N/I",
      caMax: t.caMax || "N/I",
      gabarito: t.gabarito || "N.A.",
      cotaParte: t.cotaParte || "N/I",
      incentivo: t.incentivo || "N/I",
      valorRef: num(t.valorRef || 0),
      viabilidadePct: num(t.viabilidadePct || 0),
      status: "ativo",
    };
    await postToAppsScriptWithFallback(["saveTerreno", "saveTerrain"], terrenoPayload);
    state.terrenos.push(t);
    saveLocal(STORAGE_KEYS.terrenos, state.terrenos);
    state.terrenoMessage = "Terreno salvo na planilha com sucesso.";
  } catch (err) {
    state.terrenoMessage = "Erro ao salvar terreno na planilha: " + err.message;
    rerender();
    return;
  }

  state.terrenoForm = { tema: "", nome: "", cidade: "", estado: "", projeto: "", etapa: "", areaGleba: 0, areaApp: 0, fotoBase64: "", fotoNome: "", quadroImagem: "", quadroImagemNome: "", quadroNotas: "", apeloNotas: {}, apeloDetalhes: {}, lote: "", zona: "", setor: "", codilog: "", caBas: "N/I", caMax: "N/I", gabarito: "N.A.", cotaParte: "N/I", incentivo: "N/I", valorRef: 0, viabilidadePct: 0, coordinates: [] };
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
  _activePath = "_search";
  _activeRaw = v || "";
  state.terrenoSearch = v || "";
  rerender();
}

function selectTerrenoMapa(id) {
  state.terrenoSelecionadoId = id;
  rerender();
}

function terrenoCadastroForm() {
  const f = state.terrenoForm;
  const cards = state.terrenos.map(t => `
    <div class="terreno-card${t.id === state.study.terrenoId ? " terreno-card-selected" : ""}">
      ${t.id === state.study.terrenoId ? `<div class="terreno-card-badge">✔ Em uso no estudo</div>` : ""}
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
      </div>
      <button class="terreno-card-del" onclick="removeTerrenoLocal('${t.id}')" title="Remover">✕</button>
    </div>
  `).join("");

  return `
    <div class="section">
      <div class="section-head head-blue">Cadastro único de terreno</div>
      <div class="section-body">
        <div class="field">
          <label>Projeto</label>
          <div class="input-wrap full">
            <select class="inp" onchange="setTerrenoField('projeto',this.value)" style="cursor:pointer">
              <option value="">Selecione…</option>
              ${TERRENO_TEMAS.flatMap(group => group.items).map(item => `<option value="${item}"${f.projeto === item ? " selected" : ""}>${item}</option>`).join("")}
            </select>
          </div>
        </div>
        <div class="field">
          <label>Etapa</label>
          <div class="input-wrap full">
            <input class="inp text" type="text" inputmode="numeric"
              value="${_activePath === 'terreno:etapa' ? _activeRaw : (f.etapa || '')}"
              data-terreno-path="etapa" data-terreno-digits="true" placeholder="Ex: 1" />
          </div>
        </div>
        <div class="field">
          <label>Nome</label>
          <div class="input-wrap full">
            <input class="inp text" type="text"
              value="${_activePath === 'terreno:nome' ? _activeRaw : f.nome}"
              data-terreno-path="nome" placeholder="Ex: Gleba Santa Clara" />
          </div>
        </div>
        <div style="display:grid;grid-template-columns:1fr .6fr;gap:14px">
          <div class="field">
            <label>Cidade</label>
            <div class="input-wrap full">
              <input class="inp text" type="text"
                value="${_activePath === 'terreno:cidade' ? _activeRaw : f.cidade}"
                data-terreno-path="cidade" />
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
              value="${_activePath === 'terreno:areaGleba' ? _activeRaw : (f.areaGleba ? fmtBR(f.areaGleba) : '')}"
              data-terreno-path="areaGleba" data-terreno-num="true" />
            <span class="affix">m²</span>
          </div>
        </div>
        <div class="section-head head-primary" style="margin-top:18px;border-radius:8px 8px 0 0">4. Legislação Urbanística</div>
        <div style="padding:14px 0 4px;display:grid;grid-template-columns:1fr 1fr;gap:12px">
          <div class="field"><label>Lote nº</label>
            <div class="input-wrap full"><input class="inp text" type="text" value="${f.lote}" data-terreno-path="lote" placeholder="Ex: 5" /></div>
          </div>
          <div class="field"><label>Zona</label>
            <div class="input-wrap full"><input class="inp text" type="text" value="${f.zona}" data-terreno-path="zona" placeholder="Ex: ZR3" /></div>
          </div>
        </div>
        <div class="field"><label>Setor-Quadra-Lote-Dígito</label>
          <div class="input-wrap full"><input class="inp text" type="text" value="${f.setor}" data-terreno-path="setor" placeholder="Ex: SIQN QI 9 LT 7" /></div>
        </div>
        <div class="field"><label>Codilog</label>
          <div class="input-wrap full"><input class="inp text" type="text" value="${f.codilog}" data-terreno-path="codilog" placeholder="Ex: 000000" /></div>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
          <div class="field"><label>CA Básico</label>
            <div class="input-wrap full"><input class="inp text" type="text" value="${f.caBas}" data-terreno-path="caBas" placeholder="Ex: 2.0" /></div>
          </div>
          <div class="field"><label>CA Máximo</label>
            <div class="input-wrap full"><input class="inp text" type="text" value="${f.caMax}" data-terreno-path="caMax" placeholder="Ex: 4.0" /></div>
          </div>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
          <div class="field"><label>Gabarito de altura</label>
            <div class="input-wrap full"><input class="inp text" type="text" value="${f.gabarito}" data-terreno-path="gabarito" placeholder="Ex: N.A." /></div>
          </div>
          <div class="field"><label>Cota parte (m²)</label>
            <div class="input-wrap full"><input class="inp text" type="text" value="${f.cotaParte}" data-terreno-path="cotaParte" placeholder="Ex: 500" /></div>
          </div>
        </div>
        <div class="field"><label>Incentivos</label>
          <div class="input-wrap full"><input class="inp text" type="text" value="${f.incentivo}" data-terreno-path="incentivo" placeholder="Ex: EHIS, ZEIS" /></div>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
          <div class="field"><label>VGV de referência (R$)</label>
            <div class="input-wrap full"><input class="inp" type="text" inputmode="decimal"
              value="${_activePath === 'terreno:valorRef' ? _activeRaw : (f.valorRef ? fmtBR(f.valorRef) : '')}"
              data-terreno-path="valorRef" data-terreno-num="true"
              placeholder="0,00" /></div>
          </div>
          <div class="field"><label>Viabilidade (%)</label>
            <div class="input-wrap full"><input class="inp" type="text" inputmode="decimal"
              value="${_activePath === 'terreno:viabilidadePct' ? _activeRaw : (f.viabilidadePct ? fmtBR(f.viabilidadePct) : '')}"
              data-terreno-path="viabilidadePct" data-terreno-num="true"
              placeholder="0,00" /><span class="affix">%</span></div>
          </div>
        </div>
        <div class="section-head head-orange" style="margin-top:18px;border-radius:8px 8px 0 0">5. Apelo Comercial do Imóvel</div>
        <p class="muted" style="font-size:12px;margin:6px 0 10px">Avalie de 1 a 5 estrelas cada fator. A análise considera inserção urbana, infraestrutura local, potencial construtivo e dinâmica econômica da cidade.</p>
        <table class="apelo-table">
          <thead>
            <tr><th>Fator</th><th>Avaliação</th><th>Detalhamento</th></tr>
          </thead>
          <tbody>
            ${APELO_FATORES.map((fator, i) => `
              <tr class="${i % 2 === 0 ? "apelo-row-a" : "apelo-row-b"}">
                <td class="apelo-fator">${fator}</td>
                <td class="apelo-avaliacao">
                  ${renderStars(fator, f.apeloNotas[fator] || 0)}
                  <span class="apelo-nota">(${f.apeloNotas[fator] || 0}/5)</span>
                </td>
                <td><input class="inp text apelo-detalhe" type="text"
                  value="${(f.apeloDetalhes[fator] || "").replace(/"/g, "&quot;")}"
                  oninput="setTerrenoApeloDetalhe('${fator}',this.value)"
                  placeholder="Descreva..." /></td>
              </tr>
            `).join("")}
          </tbody>
        </table>
        ${state.terrenoMessage ? `<div class="${state.terrenoMessage.startsWith("Salvo") || state.terrenoMessage.includes("sucesso") ? "notice" : "error"}" style="margin-bottom:10px;margin-top:14px">${state.terrenoMessage}</div>` : ""}
        <div class="btn-row" style="margin-top:14px">
          <button class="btn green" onclick="saveTerrenoLocal()">Salvar terreno</button>
        </div>
      </div>
    </div>
  `;
}

function terrenosView() {
  const f = state.terrenoForm;
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
          ${state.terrenoSheetMessage ? `<div class="${state.terrenoSheetMessage.startsWith("Erro") ? "error" : "notice"}" style="margin-bottom:12px">${state.terrenoSheetMessage}</div>` : ""}
          <div class="card-grid-2 bottom">
            <div class="section">
              <div class="section-head head-primary">Terrenos</div>
              <div class="section-body">
                <p class="muted" style="margin-bottom:14px">Selecione um projeto para abrir a visualização em mapa.</p>
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

  const terrenosTema = state.terrenos.filter((t) => t.projeto === state.terrenoTema);
  const q = (state.terrenoSearch || "").trim().toLowerCase();
  const filtered = q
    ? terrenosTema.filter(t => `${t.nome} ${t.cidade || ""}`.toLowerCase().includes(q))
    : terrenosTema;

  if (!state.terrenoSelecionadoId && filtered.length > 0) {
    state.terrenoSelecionadoId = filtered[0].id;
  }
  const sel = filtered.find(t => t.id === state.terrenoSelecionadoId) || filtered[0] || null;

  // ── Panel body ───────────────────────────────────────────────────────────────
  const pav = (_mapPavimentos[sel && sel.id] || 4);
  const area = sel ? (sel.areaGleba || 0) : 0;
  const caMaxF = sel ? (parseFloat((sel.caMax || "4").replace(",", ".")) || 4) : 4;
  const esc = Math.min(0.70, Math.sqrt(caMaxF / pav));
  const lamina = area * esc * esc;
  const vgv = sel ? (sel.valorRef || 0) : 0;
  const viab = sel ? (sel.viabilidadePct || 0) : 0;

  function mVbox(lbl, val, green) {
    return `<div class="mapa-campo"><label class="mapa-lbl">${lbl}</label><div class="mapa-vbox${green ? " green" : ""}">${val}</div></div>`;
  }
  function mVbox2(lbl1, v1, lbl2, v2) {
    return `<div class="mapa-g2">${mVbox(lbl1, v1)}${mVbox(lbl2, v2)}</div>`;
  }

  let painelBody = "";
  if (!sel) {
    painelBody = `
      <div style="text-align:center;padding:60px 20px;color:#9ca3af">
        <svg width="56" height="56" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" style="margin:0 auto 14px;opacity:.35;display:block">
          <path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/>
        </svg>
        <p style="margin:0 0 16px;font-size:14px">Nenhum terreno cadastrado neste projeto.</p>
        <button class="btn blue" onclick="setTerrenoTema(null)">← Voltar e cadastrar</button>
      </div>`;
  } else {
    const listItems = filtered.length > 1 ? `
      <div class="mapa-lista">
        ${filtered.map(t => `
          <button class="mapa-lista-item${t.id === sel.id ? " active" : ""}" onclick="mapaSelectTerreno('${t.id}')">
            <span class="mapa-lista-nome">${t.nome}</span>
            <span class="mapa-lista-sub">${t.cidade || ""}${t.etapa ? " · Et. " + t.etapa : ""}</span>
          </button>`).join("")}
      </div>` : "";

    painelBody = `
      ${listItems}
      <div class="mapa-pb-inner">
        <div class="mapa-sec">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#2d5f52" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
          <h3>Legislação</h3>
        </div>
        ${sel.lote ? mVbox("LOTE Nº", sel.lote) : ""}
        ${mVbox2("ÁREA", `<span style="font-weight:800;color:#166534">${fmt(area, 2)}</span> <span style="font-size:11px;color:#16a34a">m²</span>`, "ZONA", sel.zona || "N/I")}
        ${mVbox("SETOR-QUADRA-LOTE-DÍGITO", sel.setor || "N/I")}
        ${mVbox("CODILOG", sel.codilog || "N/I")}
        <div class="mapa-campo"><label class="mapa-lbl">COEFICIENTE DE APROVEITAMENTO</label>
          <div class="mapa-g2">
            <div><div style="font-size:10px;color:#6b7280;font-weight:600;margin-bottom:4px">Básico</div><div class="mapa-vbox">${sel.caBas || "N/I"}</div></div>
            <div><div style="font-size:10px;color:#6b7280;font-weight:600;margin-bottom:4px">Máximo</div><div class="mapa-vbox">${sel.caMax || "N/I"}</div></div>
          </div>
        </div>
        ${mVbox("GABARITO DE ALTURA", sel.gabarito || "N.A.")}
        ${mVbox("COTA PARTE", sel.cotaParte || "N/I")}
        ${mVbox("INCENTIVOS", sel.incentivo || "N/I")}

        <div class="mapa-sec" style="margin-top:20px">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#2d5f52" stroke-width="2"><rect x="2" y="7" width="20" height="15" rx="2"/><polyline points="17 2 12 7 7 2"/></svg>
          <h3>Estudo de Volumetria</h3>
        </div>
        <div class="mapa-vol-box">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px">
            <span class="mapa-lbl" style="margin:0">Pavimentos</span>
            <span id="mapa-val-pav" style="font-size:22px;font-weight:800;color:#166534">${pav}</span>
          </div>
          <input type="range" min="2" max="11" step="1" value="${pav}"
            data-terreno-id="${sel.id}" data-area="${area}" data-camaxf="${caMaxF}"
            oninput="mapaUpdatePav(this.value,'${sel.id}')"
            style="width:100%;accent-color:#2d5f52;cursor:pointer">
          <div style="display:flex;justify-content:space-between;font-size:10px;color:#6b7280;margin-top:3px">
            <span>2 pav.</span><span>11 pav.</span>
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-top:12px">
            <div class="mapa-si-item"><div class="v">${pav}</div><div class="l">Pavimentos</div></div>
            <div class="mapa-si-item"><div class="v">${pav * 3}m</div><div class="l">Altura</div></div>
            <div class="mapa-si-item"><div class="v">${fmt(lamina, 0)}m²</div><div class="l">Lâmina</div></div>
          </div>
        </div>

        <div class="mapa-card-dark">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px">
            <div>
              <div style="color:rgba(255,255,255,.7);font-size:11px;font-weight:600;margin-bottom:3px">ROI</div>
              <div style="color:#fff;font-size:16px;font-weight:700">${fmt(viab * 1.2, 1)}%</div>
            </div>
            <div style="text-align:right">
              <div style="color:rgba(255,255,255,.7);font-size:11px;font-weight:600;margin-bottom:3px">VGV</div>
              <div style="color:#fff;font-size:16px;font-weight:700">R$ ${fmt(vgv, 0)}</div>
            </div>
          </div>
          <div style="border-top:1px solid rgba(255,255,255,.2);padding-top:14px">
            <div style="color:rgba(255,255,255,.9);font-size:12px;font-weight:600;margin-bottom:10px">Indicadores</div>
            <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px">
              <div><div style="color:rgba(255,255,255,.6);font-size:10px;margin-bottom:3px">Viabilidade</div><div style="color:#fff;font-size:14px;font-weight:700">${fmt(viab, 1)}%</div></div>
              <div><div style="color:rgba(255,255,255,.6);font-size:10px;margin-bottom:3px">Margem</div><div style="color:#fff;font-size:14px;font-weight:700">${fmt(viab * 0.8, 1)}%</div></div>
              <div><div style="color:rgba(255,255,255,.6);font-size:10px;margin-bottom:3px">Lucro</div><div style="color:#10b981;font-size:14px;font-weight:700">${fmt(viab * 1.5, 1)}%</div></div>
            </div>
          </div>
        </div>
      </div>`;
  }

  return `
    <div class="mapa-fullscreen">
      <div class="mapa-painel" id="mapa-painel">
        <div class="mapa-painel-header">
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="rgba(255,255,255,.65)" stroke-width="2"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/></svg>
            <span style="color:rgba(255,255,255,.7);font-size:11px;font-weight:600">Estudo de Viabilidade</span>
          </div>
          <div style="display:flex;align-items:center;justify-content:space-between">
            <h2 style="margin:0;font-size:17px;font-weight:700;color:#fff">
              ${sel ? sel.nome : state.terrenoTema}
              ${filtered.length > 1 ? `<span class="badge" style="font-size:11px;padding:2px 8px;border-radius:12px;background:rgba(255,255,255,.25);margin-left:8px">${filtered.length}</span>` : ""}
            </h2>
            <button onclick="setTerrenoTema(null)"
              style="background:rgba(255,255,255,.15);border:none;color:#fff;width:30px;height:30px;border-radius:6px;cursor:pointer;font-size:18px;line-height:1;flex-shrink:0">
              &times;
            </button>
          </div>
          <div style="margin-top:10px">
            <input id="terreno-search-inp" class="inp text mapa-search" type="text" value="${state.terrenoSearch}"
              oninput="setTerrenoSearch(this.value)"
              placeholder="Buscar terreno..." />
          </div>
        </div>
        <div class="mapa-painel-body">
          ${painelBody}
        </div>
      </div>

      <div id="terreno-project-map" class="mapa-canvas"></div>

      <button class="mapa-fab" id="mapa-btn3d" onclick="toggleMapa3D()">
        ${_mapIs3D ? "Vista 2D" : "Vista 3D"}
      </button>
    </div>
  `;
}

// ─── Mapa deck.gl ──────────────────────────────────────────────────────────────

function mapaSelectTerreno(id) {
  state.terrenoSelecionadoId = id;
  rerender();
}

function mapaUpdatePav(val, id) {
  const n = parseInt(val, 10);
  _mapPavimentos[id] = n;
  const pavEl = document.getElementById("mapa-val-pav");
  if (pavEl) pavEl.textContent = n;
  const area = parseFloat(document.querySelector("input[type=range][data-terreno-id]")?.dataset.area || 0);
  const caMaxF = parseFloat(document.querySelector("input[type=range][data-terreno-id]")?.dataset.camaxf || 4);
  const esc = Math.min(0.70, Math.sqrt(caMaxF / n));
  const lamina = area * esc * esc;
  const items = document.querySelectorAll(".mapa-si-item .v");
  if (items[0]) items[0].textContent = n;
  if (items[1]) items[1].textContent = (n * 3) + "m";
  if (items[2]) items[2].textContent = fmt(lamina, 0) + "m²";
  renderMapLayers();
}

function toggleMapa3D() {
  _mapIs3D = !_mapIs3D;
  const btn = document.getElementById("mapa-btn3d");
  if (btn) btn.textContent = _mapIs3D ? "Vista 2D" : "Vista 3D";
  renderMapLayers();
}

function destroyDeckMap() {
  if (_deckMap) {
    try { _deckMap.finalize(); } catch {}
    _deckMap = null;
  }
}

function makeMapLayers() {
  if (typeof deck === "undefined") return [];
  const tema = state.terrenoTema;
  if (!tema) return [];
  const terrenos = state.terrenos.filter(t => t.projeto === tema);
  const features = terrenos
    .filter(t => t.coordinates && t.coordinates.length > 2)
    .map(t => ({ id: t.id, nome: t.nome, polygon: t.coordinates }));

  const selId = state.terrenoSelecionadoId;
  const pav = _mapPavimentos[selId] || 4;

  const layers = [new deck.PolygonLayer({
    id: "terrenos-layer",
    data: features,
    extruded: false,
    getPolygon: d => d.polygon,
    getFillColor: d => d.id === selId ? [134, 239, 172, 200] : [45, 95, 82, 180],
    getLineColor: d => d.id === selId ? [22, 163, 74, 255] : [45, 95, 82, 255],
    getLineWidth: d => d.id === selId ? 3 : 1.5,
    lineWidthMinPixels: 1,
    pickable: true,
    autoHighlight: true,
    highlightColor: [134, 239, 172, 120],
    updateTriggers: { getFillColor: [selId], getLineColor: [selId], getLineWidth: [selId] },
    onClick: info => { if (info.object) mapaSelectTerreno(info.object.id); },
  })];

  if (_mapIs3D && selId) {
    const selFeat = features.find(f => f.id === selId);
    if (selFeat) {
      const selTerreno = state.terrenos.find(t => t.id === selId);
      const caMaxF = selTerreno ? (parseFloat((selTerreno.caMax || "4").replace(",", ".")) || 4) : 4;
      const esc = Math.min(0.70, Math.sqrt(caMaxF / pav));
      const cx = selFeat.polygon.reduce((s, c) => s + c[0], 0) / selFeat.polygon.length;
      const cy = selFeat.polygon.reduce((s, c) => s + c[1], 0) / selFeat.polygon.length;
      const scaled = selFeat.polygon.map(c => [cx + (c[0] - cx) * esc, cy + (c[1] - cy) * esc]);
      layers.push(new deck.PolygonLayer({
        id: "edificio-layer",
        data: [{ polygon: scaled, elevation: pav * 30 }],
        extruded: true,
        getPolygon: d => d.polygon,
        getElevation: d => d.elevation,
        getFillColor: [255, 80, 0, 230],
        getLineColor: [255, 120, 20, 255],
        getLineWidth: 1,
        lineWidthMinPixels: 1,
        pickable: false,
      }));
    }
  }
  return layers;
}

function renderMapLayers() {
  if (_deckMap) {
    try { _deckMap.setProps({ layers: makeMapLayers() }); } catch {}
  }
}

function initDeckMap() {
  const container = document.getElementById("terreno-project-map");
  if (!container) { destroyDeckMap(); return; }
  if (typeof deck === "undefined") return;

  destroyDeckMap();

  const tema = state.terrenoTema;
  const terrenos = tema ? state.terrenos.filter(t => t.projeto === tema && t.coordinates && t.coordinates.length > 2) : [];
  let centerLon = -47.8220, centerLat = -15.6581, zoom = 13;
  if (terrenos.length > 0) {
    const allLons = terrenos.flatMap(t => t.coordinates.map(c => c[0]));
    const allLats = terrenos.flatMap(t => t.coordinates.map(c => c[1]));
    centerLon = (Math.min(...allLons) + Math.max(...allLons)) / 2;
    centerLat = (Math.min(...allLats) + Math.max(...allLats)) / 2;
    zoom = 15;
  }

  const vs = _mapSavedVS || {};
  _deckMap = new deck.DeckGL({
    container: "terreno-project-map",
    mapStyle: "https://basemaps.cartocdn.com/gl/positron-gl-style/style.json",
    initialViewState: {
      longitude: vs.longitude || centerLon,
      latitude:  vs.latitude  || centerLat,
      zoom:      vs.zoom      || zoom,
      pitch:     vs.pitch     || 0,
      bearing:   vs.bearing   || 0,
    },
    controller: { dragRotate: true, touchRotate: true },
    layers: makeMapLayers(),
    onViewStateChange: e => { _mapSavedVS = e.viewState; },
    getCursor: s => s.isHovering ? "pointer" : "default",
    getTooltip: o => o.object && o.object.nome
      ? { html: `<div style="background:#1f2937;color:#fff;padding:6px 12px;border-radius:6px;font-size:13px;font-weight:600">${o.object.nome}</div>`, style: { background: "transparent", border: "none" } }
      : null,
  });
}

async function openTerrenoPicker() {
  state.sheetMessage = "Carregando terrenos da planilha...";
  rerender();
  try {
    await loadTerrenosFromSheet();
  } catch {
    // Mensagem já tratada em loadTerrenosFromSheet
  }
  if (state.terrenos.length === 0) {
    state.sheetMessage = "Nenhum terreno encontrado na planilha.";
    rerender();
    return;
  }
  state.sheetMessage = "";
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
  state.study.terrenoId = t.id;
  state.sheetMessage = `Terreno "${t.nome}" aplicado ao estudo.`;
  state.showTerrenoPickerModal = false;
  rerender();
}

function handleStudyTerrenoChange(id) {
  if (!id) {
    state.study.terrenoId = "";
    rerender();
    return;
  }
  selectTerrenoForStudy(id);
}

async function refreshTerrenosForStudy() {
  state.sheetMessage = "Atualizando terrenos da planilha...";
  rerender();
  try {
    await loadTerrenosFromSheet();
    state.sheetMessage = "Terrenos atualizados com sucesso.";
  } catch {
    state.sheetMessage = state.terrenoSheetMessage || "Falha ao atualizar terrenos da planilha.";
  }
  rerender();
}

function getTerrenoMapQuery(terreno) {
  if (!terreno) return "Brasil";
  return [terreno.nome, terreno.cidade, terreno.estado].filter(Boolean).join(", ");
}

function getTerrenoMapEmbedUrl(terreno) {
  return `https://maps.google.com/maps?q=${encodeURIComponent(getTerrenoMapQuery(terreno))}&z=14&output=embed`;
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
              <span>${t.cidade}${t.estado ? " · " + t.estado : ""} · Gleba: ${fmt(t.areaGleba)} m²${t.projeto ? " · Projeto: " + t.projeto : ""}</span>
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
  window.handleTerrenoBoardUpload = handleTerrenoBoardUpload;
  window.clearTerrenoBoardImage = clearTerrenoBoardImage;
  window.handleFotoUpload = handleFotoUpload;
  window.setTerrenoApelo = setTerrenoApelo;
  window.setTerrenoApeloDetalhe = setTerrenoApeloDetalhe;
  window.saveTerrenoLocal = saveTerrenoLocal;
  window.removeTerrenoLocal = removeTerrenoLocal;
  window.setTerrenoTema = setTerrenoTema;
  window.setTerrenoSearch = setTerrenoSearch;
  window.selectTerrenoMapa = selectTerrenoMapa;
  window.mapaSelectTerreno = mapaSelectTerreno;
  window.mapaUpdatePav = mapaUpdatePav;
  window.toggleMapa3D = toggleMapa3D;
  window.openTerrenoPicker = openTerrenoPicker;
  window.refreshTerrenosForStudy = refreshTerrenosForStudy;
  window.closeTerrenoPicker = closeTerrenoPicker;
  window.selectTerrenoForStudy = selectTerrenoForStudy;
  window.handleStudyTerrenoChange = handleStudyTerrenoChange;

  try {
    state.benchmarks = normalizeBenchmarks(state.benchmarks);
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

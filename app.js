const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwZ9itEkqLp9QWPn5NK1olS9j9FLrGrIdVpnXIszbDL7Wv_pWL4zxBrMRlUz1MqcHq-pw/exec";

const STORAGE_KEYS = {
  benchmarks: "viab_benchmarks_v1",
  scenarios: "viab_scenarios_v1",
};

const PROJECT_TYPES = [
  { id: "loteamento", label: "Loteamento", icon: "🏘️", enabled: true },
  { id: "horizontal", label: "Incorporação Horizontal", icon: "🏡", enabled: false },
  { id: "vertical", label: "Incorporação Vertical", icon: "🏢", enabled: false },
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
  study: {
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
      infraM2: 0,
      projetoR: 0,
      registroR: 0,
      manutPosPct: 2,
      marketingPct: 0,
      corretagemPct: 0,
      adminPct: 0,
      impostosPct: 0,
      permFisicaPct: 0,
      permFinPct: 0,
      contingenciasPct: 0,
      terrenoM2: 0,
      houseMes: 0,
      houseCorretores: 6,
      houseMeses: 6,
    },
  },
  calc: null,
  savedScenarios: loadLocal(STORAGE_KEYS.scenarios, []),
  benchmarks: loadLocal(STORAGE_KEYS.benchmarks, BENCHMARK_TEMPLATE),
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
  if (n === 0) return "";
  return n.toLocaleString("pt-BR", {
    minimumFractionDigits: 0,
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
      infraM2: 0,
      projetoR: 0,
      registroR: 0,
      manutPosPct: 2,
      marketingPct: 0,
      corretagemPct: 0,
      adminPct: 0,
      impostosPct: 0,
      permFisicaPct: 0,
      permFinPct: 0,
      contingenciasPct: 0,
      terrenoM2: 0,
      houseMes: 0,
      houseCorretores: 6,
      houseMeses: 6,
    },
  };
}

function normalizeLoadedStudy(payload) {
  const base = getDefaultStudy();
  const loaded = payload && payload.study ? payload.study : {};

  return {
    ...base,
    ...loaded,
    urban: { ...base.urban, ...(loaded.urban || {}) },
    product: { ...base.product, ...(loaded.product || {}) },
    costs: { ...base.costs, ...(loaded.costs || {}) },
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

  const impostosR = vgvBruto * num(c.impostosPct) / 100;
  const corretagemR = vgvBruto * num(c.corretagemPct) / 100;
  const marketingR = vgvBruto * num(c.marketingPct) / 100;
  const houseR = num(c.houseMes) * num(c.houseCorretores) * num(c.houseMeses);

  const receitaLiquida = vgvBruto - impostosR - corretagemR - marketingR - houseR;

  const terrenoR = num(c.terrenoM2) * areaTotal;
  const infraR = num(c.infraM2) * areaLoteavel;
  const projetoR = num(c.projetoR);
  const registroR = num(c.registroR);
  const manutPosR = infraR * num(c.manutPosPct) / 100;

  const custoObrasTotal = infraR + projetoR + registroR + manutPosR;
  const resultadoOperacional = receitaLiquida - terrenoR - custoObrasTotal;

  const permFinR = vgvBruto * num(c.permFinPct) / 100;
  const contingenciasR = custoObrasTotal * num(c.contingenciasPct) / 100;
  const adminR = vgvBruto * num(c.adminPct) / 100;
  const resultadoFinal = resultadoOperacional - permFinR - contingenciasR - adminR;

  const custoTotal = terrenoR + custoObrasTotal + impostosR + adminR + corretagemR + marketingR + houseR + permFinR + contingenciasR;
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
    houseR,
    receitaLiquida,
    terrenoR,
    infraR,
    projetoR,
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
  const { suffix = "", prefix = "", full = false, text = false } = opts;
  const isPercentField = suffix.includes("%");
  const displayVal = text
    ? (_activePath === path ? _activeRaw : (value ?? ""))
    : (_activePath === path ? _activeRaw : (isPercentField ? fmt(value, 2) : fmtBR(value)));

  return `
    <div class="field">
      <label>${label}</label>
      <div class="input-wrap ${full ? "full" : ""}">
        ${prefix ? `<span class="affix left">${prefix}</span>` : ""}
        <input class="inp ${text ? "text" : ""}"
          type="text"
          ${!text ? 'inputmode="decimal"' : ''}
          ${text ? 'data-text="true"' : ''}
          value="${displayVal}"
          data-path="${path}" />
        ${suffix ? `<span class="affix">${suffix}</span>` : ""}
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
            <p>Cadastre benchmarks urbanísticos e financeiros para Loteamento, Horizontal e Vertical.</p>
            <span class="badge ok">EDITÁVEL</span>
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
                  ${kpi("Área total dos lotes", fmt(c.areaTotalLotes), `${fmt(c.nLotes, 0)} lotes`)}
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
              ${inputField("Número de lotes", "product.nLotes", state.study.product.nLotes, { full: true })}
              ${inputField("Área média do lote", "product.areaMedia", state.study.product.areaMedia, { suffix: "m²", full: true })}
              ${inputField("Preço por m²", "product.precoM2", state.study.product.precoM2, { prefix: "R$", suffix: "/m²", full: true })}
              ${inputField("Permuta física", "costs.permFisicaPct", state.study.costs.permFisicaPct, { suffix: "%", full: true })}
              ${inputField("Permuta financeira", "costs.permFinPct", state.study.costs.permFinPct, { suffix: "%", full: true })}
              ${inputField("Terreno", "costs.terrenoM2", state.study.costs.terrenoM2, { prefix: "R$", suffix: "/m²total", full: true })}
            </div>
          </div>
        </div>

        <div class="section" style="margin-top:12px">
          <div class="section-head head-orange">4. Estrutura de custos</div>
          <div class="section-body">
            <div style="display:grid;grid-template-columns:1.25fr 1fr 1fr .65fr .65fr .65fr;gap:14px;margin-bottom:14px">
              ${inputField("Infraestrutura", "costs.infraM2", state.study.costs.infraM2, { prefix: "R$", suffix: "/m²lot.", full: true })}
              ${inputField("Projeto e licenciamento", "costs.projetoR", state.study.costs.projetoR, { prefix: "R$", full: true })}
              ${inputField("Registro", "costs.registroR", state.study.costs.registroR, { prefix: "R$", full: true })}
              ${inputField("Manutenção", "costs.manutPosPct", state.study.costs.manutPosPct, { suffix: "%infra", full: true })}
              ${inputField("Marketing", "costs.marketingPct", state.study.costs.marketingPct, { suffix: "%VGV", full: true })}
              ${inputField("Corretagem", "costs.corretagemPct", state.study.costs.corretagemPct, { suffix: "%VGV", full: true })}
            </div>
            <div style="display:grid;grid-template-columns:.65fr .65fr 1fr .65fr .65fr .65fr;gap:14px">
              ${inputField("Administração", "costs.adminPct", state.study.costs.adminPct, { suffix: "%VGV", full: true })}
              ${inputField("Impostos vendas", "costs.impostosPct", state.study.costs.impostosPct, { suffix: "%VGV", full: true })}
              ${inputField("House comercial", "costs.houseMes", state.study.costs.houseMes, { prefix: "R$", suffix: "/mês", full: true })}
              ${inputField("Corretores", "costs.houseCorretores", state.study.costs.houseCorretores, { full: true })}
              ${inputField("Meses", "costs.houseMeses", state.study.costs.houseMeses, { full: true })}
              ${inputField("Contingências", "costs.contingenciasPct", state.study.costs.contingenciasPct, { suffix: "%obras", full: true })}
            </div>
          </div>
        </div>

        <div class="card-grid-2" style="margin-top:18px">
          <div class="section">
            <div class="section-head head-primary">Proforma financeiro</div>
            <div class="section-body">
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
                    ${proformaRow("(-) Impostos sobre vendas", -c.impostosR, pctOf(c.impostosR, c.vgvBruto), "row-expense")}
                    ${proformaRow("(-) Corretagem", -c.corretagemR, pctOf(c.corretagemR, c.vgvBruto), "row-expense")}
                    ${proformaRow("(-) Marketing e vendas", -c.marketingR, pctOf(c.marketingR, c.vgvBruto), "row-expense")}
                    ${proformaRow("(-) House comercial", -c.houseR, pctOf(c.houseR, c.vgvBruto), "row-expense")}
                    ${proformaRow("= Receita líquida", c.receitaLiquida, c.margemLiquidaPct, "row-sub")}
                    ${proformaRow("(-) Pagamento do terreno", -c.terrenoR, pctOf(c.terrenoR, c.vgvBruto), "row-expense")}
                    ${proformaRow("(-) Infraestrutura", -c.infraR, pctOf(c.infraR, c.vgvBruto), "row-expense")}
                    ${proformaRow("(-) Projeto e licenciamento", -c.projetoR, pctOf(c.projetoR, c.vgvBruto), "row-expense")}
                    ${proformaRow("(-) Registro", -c.registroR, pctOf(c.registroR, c.vgvBruto), "row-expense")}
                    ${proformaRow("(-) Manutenção pós-obra", -c.manutPosR, pctOf(c.manutPosR, c.vgvBruto), "row-expense")}
                    ${proformaRow("= Resultado operacional", c.resultadoOperacional, c.margemOperacionalPct, "row-sub")}
                    ${proformaRow("(-) Permuta financeira", -c.permFinR, pctOf(c.permFinR, c.vgvBruto), "row-expense")}
                    ${proformaRow("(-) Contingências", -c.contingenciasR, pctOf(c.contingenciasR, c.vgvBruto), "row-expense")}
                    ${proformaRow("(-) Administração e gestão", -c.adminR, pctOf(c.adminR, c.vgvBruto), "row-expense")}
                    ${proformaRow("= Resultado final", c.resultadoFinal, c.margemFinalPct, "row-result row-result-final")}
                  </tbody>
                </table>
              </div>
            </div>
          </div>

          <div class="section">
            <div class="section-head head-blue">Indicadores financeiros</div>
            <div class="section-body">
              <div class="kpi-grid">
                ${kpi("Receita líquida", rs(c.receitaLiquida), pc(c.margemLiquidaPct))}
                ${kpiWithBm("Resultado operacional", rs(c.resultadoOperacional), pc(c.margemOperacionalPct), c.margemOperacionalPct, bm.financial.margemOperacionalPct, true)}
                ${kpiWithBm("Resultado final", rs(c.resultadoFinal), pc(c.margemFinalPct), c.margemFinalPct, bm.financial.margemFinalPct, true)}
                ${kpiWithBm("Margem final / VGV", pc(c.margemFinalPct), "rentabilidade final", c.margemFinalPct, bm.financial.margemFinalPct, true)}
                ${kpi("Custo total", rs(c.custoTotal), pc(c.custoTotalPct))}
                ${kpiWithBm("Custo obras / VGV", pc(c.custoObrasPct), rs(c.custoObrasTotal), c.custoObrasPct, bm.financial.custoObrasPct, false)}
                ${kpi("Preço médio por lote", rs(c.ticketMedio), `${fmt(c.nLotes, 0)} lotes`)}
                ${kpi("Relação preço/custo m²", `${fmt(c.relPrecoCusto)}x`, `Preço: R$ ${fmt(c.precoM2)}/m²`)}
              </div>
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

  await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    mode: "no-cors",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action, payload }),
  });

  return {
    ok: true,
    message: "Requisição enviada ao Apps Script."
  };
}

async function getFromAppsScript(action, params = {}) {
  if (!APPS_SCRIPT_URL || APPS_SCRIPT_URL.includes("COLE_AQUI")) {
    throw new Error("A URL do Apps Script ainda não foi configurada.");
  }

  const query = new URLSearchParams({ action, ...params }).toString();
  const url = `${APPS_SCRIPT_URL}?${query}`;

  const response = await fetch(url);

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
  if (!selected || !selected.studyId) return;

  try {
    const res = await getFromAppsScript("getStudy", { studyId: selected.rowNumber });
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

function exportExcel() {
  const c = compute(state.study);
  const wb = XLSX.utils.book_new();

  const resumo = [
    ["Campo", "Valor"],
    ["Nome do estudo", state.study.nomeEstudo],
    ["Cidade", state.study.cidade],
    ["Fase", state.phase],
    ["Tipo de projeto", state.projectType],
    ["Área total", c.areaTotal],
    ["Área loteável", c.areaLoteavel],
    ["Área lotes vendáveis", c.areaLotesVend],
    ["Número de lotes", c.nLotes],
    ["Área média do lote", c.areaMedia],
    ["Preço venda R$/m²", c.precoM2],
    ["VGV", c.vgvBruto],
    ["Receita líquida", c.receitaLiquida],
    ["Resultado operacional", c.resultadoOperacional],
    ["Resultado final", c.resultadoFinal],
  ];

  const proformaRows = [
    ["Receita bruta (VGV)", c.vgvBruto, 100],
    ["Corretagem", -c.corretagemR, pctOf(c.corretagemR, c.vgvBruto)],
    ["Marketing e vendas", -c.marketingR, pctOf(c.marketingR, c.vgvBruto)],
    ["Pagamento fixo - house", -c.houseR, pctOf(c.houseR, c.vgvBruto)],
    ["Permuta financeira", -c.permFinR, pctOf(c.permFinR, c.vgvBruto)],
    ["Receita líquida", c.receitaLiquida, c.margemLiquidaPct],
    ["Pagamento do terreno", -c.terrenoR, pctOf(c.terrenoR, c.vgvBruto)],
    ["Custo obras total", -c.custoObrasTotal, c.custoObrasPct],
    ["Infraestrutura", -c.infraR, pctOf(c.infraR, c.vgvBruto)],
    ["Projeto e licenciamento", -c.projetoR, pctOf(c.projetoR, c.vgvBruto)],
    ["Registro", -c.registroR, pctOf(c.registroR, c.vgvBruto)],
    ["Manutenção pós-obra", -c.manutPosR, pctOf(c.manutPosR, c.vgvBruto)],
    ["Resultado operacional", c.resultadoOperacional, c.margemOperacionalPct],
    ["Impostos sobre vendas", -c.impostosR, pctOf(c.impostosR, c.vgvBruto)],
    ["Administração e gestão", -c.adminR, pctOf(c.adminR, c.vgvBruto)],
    ["Resultado final", c.resultadoFinal, c.margemFinalPct],
  ];

  const proforma = [
    ["Linha", "R$", "R$/m² venda líquida", "% VGV"],
    ...proformaRows.map((row) => [
      row[0],
      row[1],
      c.areaLiquidaVenda ? row[1] / c.areaLiquidaVenda : 0,
      row[2],
    ]),
  ];

  const wsResumo = XLSX.utils.aoa_to_sheet(resumo);
  wsResumo["!cols"] = [{ wch: 34 }, { wch: 24 }];
  for (let i = 1; i < resumo.length; i += 1) {
    const cell = wsResumo[`B${i + 1}`];
    if (cell && typeof resumo[i][1] === "number") {
      cell.z = "#,##0.00";
    }
  }

  const wsProforma = XLSX.utils.aoa_to_sheet(proforma);
  wsProforma["!cols"] = [{ wch: 34 }, { wch: 18 }, { wch: 20 }, { wch: 11 }];
  for (let i = 1; i < proforma.length; i += 1) {
    const valueCell = wsProforma[`B${i + 1}`];
    const perAreaCell = wsProforma[`C${i + 1}`];
    const pctCell = wsProforma[`D${i + 1}`];
    if (valueCell) valueCell.z = '"R$" #,##0.00';
    if (perAreaCell) perAreaCell.z = '"R$" #,##0.00';
    if (pctCell) {
      if (typeof pctCell.v === "number") pctCell.v = pctCell.v / 100;
      pctCell.z = "0.00%";
    }
  }

  XLSX.utils.book_append_sheet(wb, wsResumo, "Resumo");
  XLSX.utils.book_append_sheet(wb, wsProforma, "Proforma");

  const nome = safeId(state.study.nomeEstudo || "viabilidade_loteamento");
  XLSX.writeFile(wb, `${nome}.xlsx`);
}

function newStudy() {
  state.study = getDefaultStudy();
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

function render() {
  if (state.view === "home") return homeView();
  if (state.view === "projectType") return projectTypeView();
  if (state.view === "loteamento") return loteamentoView();
  return homeView();
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
      if (!isText) {
        const val = getNestedValue(path);
        e.target.value = fmtBR(val);
      }
      _activePath = null;
      _activeRaw = "";
    });

    el.addEventListener("input", (e) => {
      const path = e.target.getAttribute("data-path");
      const isText = e.target.getAttribute("data-text") === "true";
      _activePath = path;
      _activeRaw = e.target.value;

      if (isText) {
        setNested(path, e.target.value);
      } else {
        const raw = e.target.value.replace(/\./g, "").replace(",", ".");
        setNested(path, num(raw));
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

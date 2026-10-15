/* ===================== Constantes ===================== */

const SITES = [
  { id: 1, name: "VILLE" },
  { id: 2, name: "PUNA" },
  { id: 3, name: "TRVO" },
  { id: 4, name: "PAO" },
  { id: 5, name: "ARUE" },
  { id: 6, name: "MOOREA" },
];

// Numéros de site attendus par le système à l'IMPORT (colonnes siteN/capaN du
// fichier stock) : NE PAS CHANGER, c'est l'ordre du fichier stock lui-même.
// Numéros de site attendus à l'EXPORT (colonne "Site" du fichier envoyé) :
// différents de l'import, fournis par l'utilisateur.
const EXPORT_SITE_ID = {
  1: 1, // VILLE
  2: 2, // PUNA
  3: 5, // TRVO
  4: 6, // PAO
  5: 7, // ARUE
  6: 8, // MOOREA
};

function exportSiteId(internalSiteId) {
  return EXPORT_SITE_ID[internalSiteId] ?? internalSiteId;
}

const DEFAULT_COLMAP = {
  code: "code",
  libelle: "libelle",
  codebarre: "codebarre",
  pv: "prix_vente",
  entrepotStock: "site0",
  sitePrefix: "site",
  capaPrefix: "capa",
};

const LS_KEYS = {
  colMap: "capa_colmap_v2",
  stock: "capa_stock_v1",
  stockMeta: "capa_stock_meta_v1",
  session: "capa_session_v1",
  prepa: "capa_prepa_v1",
  mail: "capa_mail_v1",
};

/* ===================== État ===================== */

const state = {
  colMap: loadJSON(LS_KEYS.colMap, DEFAULT_COLMAP),
  stock: null, // { byCode: Map, byBarcode: Map, list: [] }
  stockMeta: loadJSON(LS_KEYS.stockMeta, null), // { fileName, count, importedAt }
  session: loadJSON(LS_KEYS.session, []), // [{site, code, capa, libelle}]
  prepa: loadJSON(LS_KEYS.prepa, []), // [{code, qty, libelle}]
  editingPrepaIdx: null,
  mail: loadJSON(LS_KEYS.mail, { clientId: "", to: "" }),
  sending: false,
  currentArticle: null,
  scannerActive: false,
  editingSessionIdx: null,
};

/* ===================== Utilitaires ===================== */

function loadJSON(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch (e) {
    return fallback;
  }
}

function saveJSON(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
    return true;
  } catch (e) {
    console.warn("Stockage local plein ou indisponible", e);
    return false;
  }
}

// Normalise un code (interne ou code-barre) en chaîne de chiffres, en gérant
// les valeurs venues d'Excel en notation scientifique (ex: 8,00084E+12).
function normalizeCode(val) {
  if (val === null || val === undefined || val === "") return "";
  if (typeof val === "number") {
    if (!isFinite(val)) return "";
    return Math.round(val).toString();
  }
  let s = String(val).trim();
  if (s === "") return "";
  const normalized = s.replace(",", ".");
  if (/^-?\d+(\.\d+)?e[+-]?\d+$/i.test(normalized)) {
    const num = Number(normalized);
    if (isFinite(num)) return Math.round(num).toString();
  }
  const digitsOnly = s.replace(/[^\d]/g, "");
  return digitsOnly !== "" ? digitsOnly.replace(/^0+(?=\d)/, "") || "0" : s;
}

function toNumberOrZero(val) {
  if (val === null || val === undefined || val === "") return 0;
  if (typeof val === "number") return isFinite(val) ? val : 0;
  const n = Number(String(val).replace(",", "."));
  return isFinite(n) ? n : 0;
}

function siteName(id) {
  const s = SITES.find((x) => x.id === Number(id));
  return s ? s.name : "?";
}

function formatPV(val) {
  if (!val) return "0";
  return Number(val).toLocaleString("fr-FR", { maximumFractionDigits: 0 });
}

let toastTimer = null;
function showToast(msg, ms = 2200) {
  const el = document.getElementById("toast");
  el.textContent = msg;
  el.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove("show"), ms);
}

/* ===================== Import du fichier stock ===================== */

async function handleStockFile(file) {
  try {
    showToast("Lecture du fichier…", 1500);
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type: "array", cellText: false, cellDates: false });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });
    if (!rows.length) {
      showToast("Fichier vide ou illisible");
      return;
    }
    buildStockIndex(rows);
    state.stockMeta = { fileName: file.name, count: rows.length, importedAt: new Date().toISOString() };
    saveJSON(LS_KEYS.stockMeta, state.stockMeta);
    persistStock(rows);
    renderStockStatus();
    showToast(`${rows.length} articles importés`);
    goToScreen("scan");
  } catch (err) {
    console.error(err);
    showToast("Erreur pendant l'import du fichier");
  }
}

function persistStock(rows) {
  const ok = saveJSON(LS_KEYS.stock, rows);
  if (!ok) {
    showToast("Fichier importé mais trop volumineux pour être mémorisé — à réimporter au prochain lancement", 3500);
  }
}

function buildStockIndex(rows) {
  const cm = state.colMap;
  const byCode = new Map();
  const byBarcode = new Map();
  const list = [];

  for (const row of rows) {
    const code = normalizeCode(row[cm.code]);
    const barcode = normalizeCode(row[cm.codebarre]);
    const libelle = String(row[cm.libelle] ?? "").trim();
    if (!code && !barcode) continue;

    const item = {
      code,
      barcode,
      libelle,
      pv: toNumberOrZero(row[cm.pv]),
      entrepotStock: toNumberOrZero(row[cm.entrepotStock]),
      bySite: {},
      capaBySite: {},
    };

    for (const s of SITES) {
      item.bySite[s.id] = toNumberOrZero(row[`${cm.sitePrefix}${s.id}`]);
      item.capaBySite[s.id] = toNumberOrZero(row[`${cm.capaPrefix}${s.id}`]);
    }

    if (code) byCode.set(code, item);
    if (barcode) byBarcode.set(barcode, item);
    list.push(item);
  }

  state.stock = { byCode, byBarcode, list };
}

function tryRestoreStockFromCache() {
  const rows = loadJSON(LS_KEYS.stock, null);
  if (rows && Array.isArray(rows) && rows.length) {
    buildStockIndex(rows);
    return true;
  }
  return false;
}

function renderStockStatus() {
  const el = document.getElementById("stock-status");
  if (!state.stockMeta || !state.stock) {
    el.innerHTML = `<div class="empty-state"><div class="emoji">📦</div><p>Aucun fichier stock importé pour l'instant.<br>Importe un fichier CSV ou XLS pour voir les désignations, le PV et le stock pendant le scan.</p></div>`;
    return;
  }
  const d = new Date(state.stockMeta.importedAt);
  el.innerHTML = `
    <div class="chip-status ok">✓ Fichier chargé</div>
    <div style="margin-top:10px; font-size:13.5px; color:var(--ink-soft); line-height:1.6;">
      <strong style="color:var(--ink)">${escapeHtml(state.stockMeta.fileName)}</strong><br>
      ${state.stockMeta.count} articles · importé le ${d.toLocaleDateString("fr-FR")} à ${d.toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" })}
    </div>`;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

/* ===================== Recherche / Article courant ===================== */

function searchArticles(query) {
  const q = query.trim().toLowerCase();
  if (!state.stock || q.length < 2) return [];
  const results = [];
  for (const item of state.stock.list) {
    if (
      (item.code && item.code.includes(q)) ||
      (item.barcode && item.barcode.includes(q)) ||
      (item.libelle && item.libelle.toLowerCase().includes(q))
    ) {
      results.push(item);
      if (results.length >= 20) break;
    }
  }
  return results;
}

function findExact(code) {
  if (!state.stock) return null;
  const norm = normalizeCode(code);
  return state.stock.byBarcode.get(norm) || state.stock.byCode.get(norm) || null;
}

function selectArticle(item) {
  state.currentArticle = item;
  renderArticleCard();
}

function selectManualCode(rawCode) {
  const norm = normalizeCode(rawCode);
  if (!norm) return;
  const found = findExact(norm);
  if (found) {
    selectArticle(found);
  } else {
    state.currentArticle = { code: norm, barcode: "", libelle: "", unknown: true, pv: 0, entrepotStock: 0, bySite: {}, capaBySite: {} };
    renderArticleCard();
  }
}

// Valeur de capa affichée pour un site donné : celle déjà ajoutée en session
// si présente, sinon la valeur d'origine du fichier stock.
function displayedCapaFor(item, siteId) {
  const sessionEntry = state.session.find((e) => e.site === siteId && e.code === item.code);
  if (sessionEntry) return sessionEntry.capa;
  return item.capaBySite[siteId] ?? 0;
}

function renderArticleCard() {
  const wrap = document.getElementById("article-card-wrap");
  const item = state.currentArticle;
  if (!item) {
    wrap.innerHTML = "";
    return;
  }

  const siteRows = SITES.map((s) => {
    const stock = item.bySite[s.id] ?? 0;
    const capa = displayedCapaFor(item, s.id);
    const isModified = capa !== (item.capaBySite[s.id] ?? 0);
    return `
      <div class="site-edit-row${isModified ? " modified" : ""}" data-site="${s.id}">
        <div class="ser-name">${s.name}</div>
        <div class="ser-stock">stock ${stock}</div>
        <div class="stepper stepper-sm">
          <button type="button" class="ser-minus" data-site="${s.id}">−</button>
          <input type="number" inputmode="numeric" class="ser-input" data-site="${s.id}" data-orig="${item.capaBySite[s.id] ?? 0}" value="${capa}">
          <button type="button" class="ser-plus" data-site="${s.id}">+</button>
        </div>
      </div>`;
  }).join("");

  const inPrepa = state.prepa.find((p) => p.code === item.code);

  wrap.innerHTML = `
    <div class="card article-card">
      <div class="ac-head">
      ${item.unknown ? `<div class="chip-status off" style="align-self:flex-start">⚠️ Article absent du fichier stock</div>` : ""}
      <div class="ac-titlebar">
      <div>
        <div class="article-title">${item.libelle ? escapeHtml(item.libelle) : "(sans désignation)"}</div>
        <div class="article-code">Code ${escapeHtml(item.code || "—")}${item.barcode ? " · CB " + escapeHtml(item.barcode) : ""}</div>
      </div>
      <div class="stat-grid stat-grid-2">
        <div class="stat"><div class="v">${formatPV(item.pv)}</div><div class="l">PV</div></div>
        <div class="stat"><div class="v">${item.entrepotStock}</div><div class="l">Stock entrepôt</div></div>
      </div>
      </div>
      </div>
      <div class="ac-body">
      <div class="ac-capa">
      <div>
        <p class="section-title" style="margin-bottom:8px;">Capa par site</p>
        <div class="site-editor">${siteRows}</div>
      </div>
      <button class="btn btn-primary" id="add-entry-btn">Enregistrer les capa modifiées</button>
      </div>

      <div class="ac-prepa">
      <div class="prepa-block">
        <p class="section-title" style="margin-bottom:8px;">Prépa</p>
        ${inPrepa ? `<div class="chip-status off" style="margin-bottom:10px;">Déjà en prépa : ${inPrepa.qty}</div>` : ""}
        <div class="stepper" style="margin-bottom:10px;">
          <button type="button" id="prepa-minus">−</button>
          <input type="number" inputmode="numeric" id="prepa-input" placeholder="Quantité" value="">
          <button type="button" id="prepa-plus">+</button>
        </div>
        <button class="btn btn-secondary${inPrepa ? " is-blocked" : ""}" id="add-prepa-btn">Ajouter à la prépa</button>
      </div>
      </div>
      </div>

      <button class="btn btn-ghost" id="close-article-btn">Fermer l'article</button>
    </div>
  `;

  wrap.querySelectorAll(".ser-minus").forEach((btn) => {
    btn.onclick = () => stepSiteCapa(btn.dataset.site, -1);
  });
  wrap.querySelectorAll(".ser-plus").forEach((btn) => {
    btn.onclick = () => stepSiteCapa(btn.dataset.site, 1);
  });
  wrap.querySelectorAll(".ser-input").forEach((inp) => {
    inp.addEventListener("input", () => markRowModified(inp));
  });
  document.getElementById("add-entry-btn").onclick = saveModifiedCapas;
  document.getElementById("prepa-minus").onclick = () => stepPrepaInput(-1);
  document.getElementById("prepa-plus").onclick = () => stepPrepaInput(1);
  document.getElementById("add-prepa-btn").onclick = addToPrepa;
  document.getElementById("close-article-btn").onclick = clearArticleSelection;
}

function stepPrepaInput(delta) {
  const inp = document.getElementById("prepa-input");
  inp.value = Math.max(0, (parseInt(inp.value, 10) || 0) + delta);
}

function addToPrepa() {
  const item = state.currentArticle;
  if (!item || !item.code) {
    showToast("Aucun article sélectionné");
    return;
  }
  const existing = state.prepa.find((p) => p.code === item.code);
  if (existing) {
    showToast(`Erreur : article déjà en prépa (qté ${existing.qty}). Modifie-le depuis l'onglet Session.`, 3800);
    return;
  }
  const qty = parseInt(document.getElementById("prepa-input").value, 10);
  if (isNaN(qty) || qty < 1) {
    showToast("Renseigne une quantité valide (1 minimum)");
    return;
  }
  state.prepa.push({ code: item.code, qty, libelle: item.libelle });
  saveJSON(LS_KEYS.prepa, state.prepa);
  renderPrepaTab();
  renderArticleCard();
  showToast(`Ajouté à la prépa · ${item.libelle || item.code}`);
}

function markRowModified(inputEl) {
  const row = inputEl.closest(".site-edit-row");
  const orig = Number(inputEl.dataset.orig);
  const val = parseInt(inputEl.value, 10);
  row.classList.toggle("modified", !isNaN(val) && val !== orig);
}

function stepSiteCapa(siteId, delta) {
  const inp = document.querySelector(`.ser-input[data-site="${siteId}"]`);
  const val = Math.max(0, (parseInt(inp.value, 10) || 0) + delta);
  inp.value = val;
  markRowModified(inp);
}

function saveModifiedCapas() {
  const item = state.currentArticle;
  if (!item || !item.code) {
    showToast("Aucun article sélectionné");
    return;
  }

  let changedCount = 0;
  document.querySelectorAll(".ser-input").forEach((inp) => {
    const siteId = Number(inp.dataset.site);
    const orig = Number(inp.dataset.orig);
    const val = parseInt(inp.value, 10);
    if (isNaN(val) || val < 0) return;
    if (val === orig) {
      // revenu à la valeur d'origine : on retire une éventuelle entrée de session
      const idx = state.session.findIndex((e) => e.site === siteId && e.code === item.code);
      if (idx >= 0) {
        state.session.splice(idx, 1);
        changedCount++;
      }
      return;
    }
    const idx = state.session.findIndex((e) => e.site === siteId && e.code === item.code);
    const entry = { site: siteId, code: item.code, capa: val, libelle: item.libelle };
    if (idx >= 0) state.session[idx] = entry;
    else state.session.push(entry);
    changedCount++;
  });

  saveJSON(LS_KEYS.session, state.session);
  renderSessionTab();

  if (changedCount === 0) {
    showToast("Aucune modification par rapport au fichier");
  } else {
    showToast(`${changedCount} capa mise(s) à jour · ${item.libelle || item.code}`);
    renderArticleCard();
  }
}

function clearArticleSelection() {
  state.currentArticle = null;
  document.getElementById("article-card-wrap").innerHTML = "";
  const searchInput = document.getElementById("search-input");
  if (searchInput) searchInput.value = "";
  document.getElementById("suggestions").innerHTML = "";
}

/* ===================== Session ===================== */

function updateSessionBadge() {
  const total = state.session.length + state.prepa.length;
  const badge = document.getElementById("tab-session-badge");
  badge.textContent = total || "";
  badge.style.display = total ? "inline-flex" : "none";
}

function renderSessionTab() {
  const list = document.getElementById("session-list");
  const countEl = document.getElementById("session-count");
  countEl.textContent = state.session.length;
  updateSessionBadge();

  if (!state.session.length) {
    list.innerHTML = `<div class="empty-state"><div class="emoji">🧾</div><p>Aucune capa modifiée pour l'instant.<br>Scanne ou recherche un article pour commencer.</p></div>`;
    return;
  }

  list.innerHTML = state.session
    .map((e, i) => {
      if (state.editingSessionIdx === i) {
        return `
        <div class="session-item editing">
          <div class="si-main">
            <div class="si-name">${e.libelle ? escapeHtml(e.libelle) : "Article " + escapeHtml(e.code)}</div>
            <div class="si-meta">${siteName(e.site)} · code ${escapeHtml(e.code)}</div>
          </div>
          <div class="stepper stepper-sm">
            <button type="button" class="si-edit-minus">−</button>
            <input type="number" inputmode="numeric" id="si-edit-input" value="${e.capa}">
            <button type="button" class="si-edit-plus">+</button>
          </div>
          <button class="si-save" data-idx="${i}" aria-label="Valider">✓</button>
        </div>`;
      }
      return `
      <div class="session-item">
        <div class="si-main">
          <div class="si-name">${e.libelle ? escapeHtml(e.libelle) : "Article " + escapeHtml(e.code)}</div>
          <div class="si-meta">${siteName(e.site)} · code ${escapeHtml(e.code)}</div>
        </div>
        <div class="si-capa">${e.capa}</div>
        <button class="si-edit" data-idx="${i}" aria-label="Modifier">✎</button>
        <button class="si-del" data-idx="${i}" aria-label="Supprimer">✕</button>
      </div>`;
    })
    .join("");

  list.querySelectorAll(".si-del").forEach((btn) => {
    btn.onclick = () => {
      const idx = Number(btn.dataset.idx);
      state.session.splice(idx, 1);
      state.editingSessionIdx = null;
      saveJSON(LS_KEYS.session, state.session);
      renderSessionTab();
    };
  });

  list.querySelectorAll(".si-edit").forEach((btn) => {
    btn.onclick = () => {
      state.editingSessionIdx = Number(btn.dataset.idx);
      renderSessionTab();
      const inp = document.getElementById("si-edit-input");
      if (inp) inp.focus();
    };
  });

  const minusBtn = list.querySelector(".si-edit-minus");
  const plusBtn = list.querySelector(".si-edit-plus");
  if (minusBtn) minusBtn.onclick = () => stepEditInput(-1);
  if (plusBtn) plusBtn.onclick = () => stepEditInput(1);

  list.querySelectorAll(".si-save").forEach((btn) => {
    btn.onclick = () => {
      const idx = Number(btn.dataset.idx);
      const inp = document.getElementById("si-edit-input");
      const val = parseInt(inp.value, 10);
      if (isNaN(val) || val < 0) {
        showToast("Quantité invalide");
        return;
      }
      state.session[idx].capa = val;
      state.editingSessionIdx = null;
      saveJSON(LS_KEYS.session, state.session);
      renderSessionTab();
      showToast("Quantité mise à jour");
    };
  });
}

function renderPrepaTab() {
  const list = document.getElementById("prepa-list");
  document.getElementById("prepa-count").textContent = state.prepa.length;
  updateSessionBadge();

  if (!state.prepa.length) {
    list.innerHTML = `<div class="empty-state"><div class="emoji">📋</div><p>Aucun article en prépa pour l'instant.</p></div>`;
    return;
  }

  list.innerHTML = state.prepa
    .map((p, i) => {
      const name = p.libelle ? escapeHtml(p.libelle) : "Article " + escapeHtml(p.code);
      if (state.editingPrepaIdx === i) {
        return `
        <div class="session-item editing">
          <div class="si-main">
            <div class="si-name">${name}</div>
            <div class="si-meta">code ${escapeHtml(p.code)}</div>
          </div>
          <div class="stepper stepper-sm">
            <button type="button" class="pr-edit-minus">−</button>
            <input type="number" inputmode="numeric" id="prepa-edit-input" value="${p.qty}">
            <button type="button" class="pr-edit-plus">+</button>
          </div>
          <button class="si-save pr-save" data-idx="${i}" aria-label="Valider">✓</button>
        </div>`;
      }
      return `
      <div class="session-item">
        <div class="si-main">
          <div class="si-name">${name}</div>
          <div class="si-meta">code ${escapeHtml(p.code)}</div>
        </div>
        <div class="si-capa si-qty">${p.qty}</div>
        <button class="si-edit pr-edit" data-idx="${i}" aria-label="Modifier">✎</button>
        <button class="si-del pr-del" data-idx="${i}" aria-label="Supprimer">✕</button>
      </div>`;
    })
    .join("");

  list.querySelectorAll(".pr-del").forEach((btn) => {
    btn.onclick = () => {
      state.prepa.splice(Number(btn.dataset.idx), 1);
      state.editingPrepaIdx = null;
      saveJSON(LS_KEYS.prepa, state.prepa);
      renderPrepaTab();
    };
  });

  list.querySelectorAll(".pr-edit").forEach((btn) => {
    btn.onclick = () => {
      state.editingPrepaIdx = Number(btn.dataset.idx);
      renderPrepaTab();
      const inp = document.getElementById("prepa-edit-input");
      if (inp) inp.focus();
    };
  });

  const stepEdit = (delta) => {
    const inp = document.getElementById("prepa-edit-input");
    inp.value = Math.max(0, (parseInt(inp.value, 10) || 0) + delta);
  };
  const minusBtn = list.querySelector(".pr-edit-minus");
  const plusBtn = list.querySelector(".pr-edit-plus");
  if (minusBtn) minusBtn.onclick = () => stepEdit(-1);
  if (plusBtn) plusBtn.onclick = () => stepEdit(1);

  list.querySelectorAll(".pr-save").forEach((btn) => {
    btn.onclick = () => {
      const val = parseInt(document.getElementById("prepa-edit-input").value, 10);
      if (isNaN(val) || val < 1) {
        showToast("Quantité invalide (1 minimum)");
        return;
      }
      state.prepa[Number(btn.dataset.idx)].qty = val;
      state.editingPrepaIdx = null;
      saveJSON(LS_KEYS.prepa, state.prepa);
      renderPrepaTab();
      showToast("Quantité mise à jour");
    };
  });
}

function clearPrepa() {
  if (!state.prepa.length) return;
  if (!confirm("Vider toute la prépa en cours ? Cette action est irréversible.")) return;
  state.prepa = [];
  state.editingPrepaIdx = null;
  saveJSON(LS_KEYS.prepa, state.prepa);
  renderPrepaTab();
}

function stepEditInput(delta) {
  const inp = document.getElementById("si-edit-input");
  const val = Math.max(0, (parseInt(inp.value, 10) || 0) + delta);
  inp.value = val;
}

function clearSession() {
  if (!state.session.length) return;
  if (!confirm("Vider toutes les capa en cours ? Cette action est irréversible.")) return;
  state.session = [];
  state.editingSessionIdx = null;
  saveJSON(LS_KEYS.session, state.session);
  renderSessionTab();
}

/* ===================== Export ===================== */

function buildExportRows() {
  // Colonnes strictement numériques : site (numéro d'export) / code article / capa
  // Uniquement les capa réellement modifiées par rapport au fichier stock.
  const header = ["Site", "Code article", "CAPA"];
  const dataRows = state.session.map((e) => [exportSiteId(e.site), Number(e.code), Number(e.capa)]);
  return [header, ...dataRows];
}

function buildPrepaExportRows() {
  // Colonnes numériques : code article / quantité
  const header = ["Code article", "Quantite"];
  const dataRows = state.prepa.map((p) => [Number(p.code), Number(p.qty)]);
  return [header, ...dataRows];
}

function buildWorkbook(kind) {
  const rows = kind === "prepa" ? buildPrepaExportRows() : buildExportRows();
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, kind === "prepa" ? "PREPA" : "CAPA");
  return wb;
}

function timestampForFile() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`;
}

/* ===================== Envoi Gmail ===================== */

let gmailToken = null; // { token, exp }

function parseRecipients(str) {
  return String(str || "")
    .split(/[,;\s]+/)
    .map((x) => x.trim())
    .filter(Boolean);
}

function isValidEmail(e) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e);
}

function mailConfigured() {
  return !!(state.mail.clientId && parseRecipients(state.mail.to).length);
}

// À appeler directement depuis le clic (le popup Google exige un geste utilisateur).
function getGmailToken() {
  return new Promise((resolve, reject) => {
    if (gmailToken && Date.now() < gmailToken.exp - 60000) return resolve(gmailToken.token);
    if (!window.google || !google.accounts || !google.accounts.oauth2) {
      return reject(new Error("Service Google indisponible (connexion internet ?)"));
    }
    const client = google.accounts.oauth2.initTokenClient({
      client_id: state.mail.clientId,
      scope: "https://www.googleapis.com/auth/gmail.send",
      callback: (resp) => {
        if (resp.error) return reject(new Error(resp.error_description || resp.error));
        gmailToken = { token: resp.access_token, exp: Date.now() + (Number(resp.expires_in) || 3600) * 1000 };
        resolve(gmailToken.token);
      },
      error_callback: (err) => reject(new Error(err && err.type ? err.type : "Autorisation refusée")),
    });
    client.requestAccessToken();
  });
}

function bytesToBase64(bytes) {
  let bin = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    bin += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  }
  return btoa(bin);
}

function utf8Base64(str) {
  return bytesToBase64(new TextEncoder().encode(str));
}

function wrap76(b64) {
  return b64.replace(/(.{76})/g, "$1\r\n");
}

function buildMime({ to, subject, bodyText, filename, mime, bytes }) {
  const boundary = "capa_" + Math.random().toString(36).slice(2);
  return [
    `To: ${to}`,
    `Subject: =?UTF-8?B?${utf8Base64(subject)}?=`,
    "MIME-Version: 1.0",
    `Content-Type: multipart/mixed; boundary="${boundary}"`,
    "",
    `--${boundary}`,
    'Content-Type: text/plain; charset="UTF-8"',
    "Content-Transfer-Encoding: base64",
    "",
    wrap76(utf8Base64(bodyText)),
    `--${boundary}`,
    `Content-Type: ${mime}; name="${filename}"`,
    "Content-Transfer-Encoding: base64",
    `Content-Disposition: attachment; filename="${filename}"`,
    "",
    wrap76(bytesToBase64(bytes)),
    `--${boundary}--`,
    "",
  ].join("\r\n");
}

async function sendViaGmail(token, { kind, count, filename, bytes }) {
  const to = parseRecipients(state.mail.to).join(", ");
  const d = new Date();
  const when = d.toLocaleDateString("fr-FR") + " " + d.toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" });
  const label = kind === "prepa" ? "Prépa" : "Capa";
  const mimeStr = buildMime({
    to,
    subject: `${label} stock - ${when} (${count} ligne${count > 1 ? "s" : ""})`,
    bodyText: `Fichier ${label.toLowerCase()} en pièce jointe (${count} ligne${count > 1 ? "s" : ""}).\nGénéré le ${when} depuis l'appli Capa Stock.`,
    filename,
    mime: "application/vnd.ms-excel",
    bytes,
  });
  const raw = bytesToBase64(new TextEncoder().encode(mimeStr)).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  const res = await fetch("https://gmail.googleapis.com/gmail/v1/users/me/messages/send", {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ raw }),
  });
  if (!res.ok) {
    if (res.status === 401) gmailToken = null;
    let msg = `Erreur ${res.status}`;
    try {
      const j = await res.json();
      if (j.error && j.error.message) msg = j.error.message;
    } catch (e) {}
    throw new Error(msg);
  }
  return to;
}

async function exportAndShare(kind) {
  if (state.sending) return;
  const list = kind === "prepa" ? state.prepa : state.session;
  if (!list.length) {
    showToast(kind === "prepa" ? "La prépa est vide" : "Aucune capa à envoyer");
    return;
  }

  // Le popup Google doit s'ouvrir directement depuis le clic : on demande le jeton en premier.
  const tokenPromise = mailConfigured() ? getGmailToken() : null;
  if (tokenPromise) tokenPromise.catch(() => {});

  const ts = timestampForFile();
  const wb = buildWorkbook(kind);
  const out = XLSX.write(wb, { bookType: "biff5", type: "array" });
  const filename = `${kind}_${ts}.xls`;

  if (tokenPromise) {
    state.sending = true;
    try {
      const token = await tokenPromise;
      showToast("Envoi par Gmail en cours…", 8000);
      const to = await sendViaGmail(token, { kind, count: list.length, filename, bytes: new Uint8Array(out) });
      showToast(`✓ ${kind === "prepa" ? "Prépa" : "Capa"} envoyée à ${to}`, 4000);
    } catch (err) {
      console.error(err);
      showToast(`Échec de l'envoi Gmail : ${err.message}`, 6000);
    } finally {
      state.sending = false;
    }
    return;
  }

  // Gmail non configuré : repli sur le partage du téléphone / tablette (ou téléchargement).
  const blob = new Blob([out], { type: "application/vnd.ms-excel" });
  const file = new File([blob], filename, { type: "application/vnd.ms-excel" });

  if (navigator.canShare && navigator.canShare({ files: [file] })) {
    try {
      await navigator.share({
        files: [file],
        title: kind === "prepa" ? "Prépa" : "Capa stock",
        text: `Fichier ${kind} (${list.length} lignes) — ${filename}`,
      });
      return;
    } catch (err) {
      if (err && err.name === "AbortError") return;
      console.warn("Partage impossible, repli sur téléchargement", err);
    }
  }

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
  showToast("Fichier téléchargé — joins-le manuellement à ton mail", 3500);
}

/* ===================== Scanner caméra ===================== */
// Deux moteurs :
//  1) BarcodeDetector natif du navigateur (Chrome / Android / tablettes) : rapide et très fiable
//  2) ZXing en repli (Safari iOS, navigateurs sans détecteur natif)
// La caméra est ouverte en haute résolution, caméra arrière, avec mise au point continue.

const SCAN_FORMATS = ["ean_13", "ean_8", "upc_a", "upc_e", "code_128", "code_39", "itf"];

let zxingReader = null;
let scanStream = null;
let scanTimer = null;
let scanSession = 0; // permet d'ignorer un démarrage devenu obsolète (arrêt pendant l'ouverture)

function buildZxingHints() {
  const hints = new Map();
  hints.set(ZXing.DecodeHintType.POSSIBLE_FORMATS, [
    ZXing.BarcodeFormat.EAN_13,
    ZXing.BarcodeFormat.EAN_8,
    ZXing.BarcodeFormat.UPC_A,
    ZXing.BarcodeFormat.UPC_E,
    ZXing.BarcodeFormat.CODE_128,
    ZXing.BarcodeFormat.CODE_39,
    ZXing.BarcodeFormat.ITF,
  ]);
  hints.set(ZXing.DecodeHintType.TRY_HARDER, true);
  return hints;
}

function setScanStatus(text) {
  let el = document.getElementById("scan-status");
  if (!el) {
    el = document.createElement("p");
    el.id = "scan-status";
    el.style.cssText = "font-size:12px; color:var(--ink-soft); text-align:center; margin:6px 0 0; font-family:var(--mono);";
    document.getElementById("reader").appendChild(el);
  }
  el.textContent = text || "";
}

// Ouvre la caméra arrière en haute résolution ; réduit les exigences si l'appareil refuse.
async function openCameraStream() {
  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
    const e = new Error("getUserMedia indisponible");
    e.name = "NotSupportedError";
    throw e;
  }
  const attempts = [
    { facingMode: { ideal: "environment" }, width: { ideal: 1920 }, height: { ideal: 1080 } },
    { facingMode: { ideal: "environment" }, width: { ideal: 1280 }, height: { ideal: 720 } },
    { facingMode: { ideal: "environment" } },
    true,
  ];
  let lastErr = null;
  for (const video of attempts) {
    try {
      return await navigator.mediaDevices.getUserMedia({ audio: false, video });
    } catch (e) {
      lastErr = e;
      // Inutile de réessayer si l'accès est refusé ou s'il n'y a aucune caméra
      if (e.name === "NotAllowedError" || e.name === "SecurityError" || e.name === "NotFoundError") break;
    }
  }
  throw lastErr;
}

// Mise au point continue (et zoom léger) quand l'appareil le permet.
async function tuneCameraTrack(stream) {
  try {
    const track = stream.getVideoTracks()[0];
    if (!track || !track.getCapabilities) return;
    const caps = track.getCapabilities();
    const advanced = [];
    if (caps.focusMode && caps.focusMode.includes("continuous")) advanced.push({ focusMode: "continuous" });
    if (advanced.length) await track.applyConstraints({ advanced });
  } catch (e) {
    console.warn("Réglage caméra ignoré", e);
  }
}

async function makeNativeDetector() {
  if (!("BarcodeDetector" in window)) return null;
  try {
    let formats = SCAN_FORMATS;
    if (BarcodeDetector.getSupportedFormats) {
      const supported = await BarcodeDetector.getSupportedFormats();
      formats = SCAN_FORMATS.filter((f) => supported.includes(f));
    }
    if (!formats.length) return null;
    return new BarcodeDetector({ formats });
  } catch (e) {
    console.warn("BarcodeDetector indisponible", e);
    return null;
  }
}

function cameraErrorMessage(err) {
  switch (err && err.name) {
    case "NotAllowedError":
    case "SecurityError":
      return "Accès à la caméra refusé : autorise-la dans les réglages du navigateur";
    case "NotFoundError":
    case "OverconstrainedError":
      return "Aucune caméra trouvée : utilise la recherche par code";
    case "NotReadableError":
    case "AbortError":
      return "Caméra occupée par une autre appli : ferme-la et réessaie";
    case "NotSupportedError":
      return "Caméra indisponible ici : ouvre l'appli en https";
    default:
      return "Impossible d'accéder à la caméra";
  }
}

async function startScanner() {
  if (state.scannerActive) return;
  const session = ++scanSession;
  const readerEl = document.getElementById("reader");
  const video = document.getElementById("scanner-video");
  readerEl.style.display = "block";
  document.getElementById("scan-toggle-btn").textContent = "Arrêter le scan";
  state.scannerActive = true;
  setScanStatus("Ouverture de la caméra…");

  try {
    const stream = await openCameraStream();
    if (session !== scanSession) {
      // arrêté pendant l'ouverture
      stream.getTracks().forEach((t) => t.stop());
      return;
    }
    scanStream = stream;
    await tuneCameraTrack(stream);

    const detector = await makeNativeDetector();
    if (session !== scanSession) return;

    if (detector) {
      video.srcObject = stream;
      video.muted = true;
      await video.play();
      const s = stream.getVideoTracks()[0].getSettings ? stream.getVideoTracks()[0].getSettings() : {};
      setScanStatus(`Détecteur natif · ${s.width || "?"}×${s.height || "?"}`);

      const tick = async () => {
        if (session !== scanSession || !state.scannerActive) return;
        try {
          if (video.readyState >= 2) {
            const codes = await detector.detect(video);
            if (codes && codes.length) {
              onScanSuccess(codes[0].rawValue);
              return;
            }
          }
        } catch (e) {
          // erreurs ponctuelles de détection : on continue
        }
        scanTimer = setTimeout(tick, 120);
      };
      tick();
    } else {
      zxingReader = new ZXing.BrowserMultiFormatReader(buildZxingHints(), 150);
      const s = stream.getVideoTracks()[0].getSettings ? stream.getVideoTracks()[0].getSettings() : {};
      setScanStatus(`ZXing · ${s.width || "?"}×${s.height || "?"}`);
      await zxingReader.decodeFromStream(stream, video, (result, err) => {
        if (result && session === scanSession) onScanSuccess(result.getText());
        // err (NotFoundException) est déclenché en continu tant qu'aucun code
        // n'est détecté sur l'image courante : c'est normal, on l'ignore.
      });
    }
  } catch (err) {
    console.error(err);
    if (session === scanSession) {
      showToast(cameraErrorMessage(err), 4500);
      stopScanner();
    }
  }
}

function stopScanner() {
  scanSession++;
  state.scannerActive = false;
  document.getElementById("scan-toggle-btn").textContent = "📷 Scanner un article";
  document.getElementById("reader").style.display = "none";
  clearTimeout(scanTimer);
  scanTimer = null;
  if (zxingReader) {
    try {
      zxingReader.reset();
    } catch (e) {}
    zxingReader = null;
  }
  if (scanStream) {
    scanStream.getTracks().forEach((t) => {
      try {
        t.stop();
      } catch (e) {}
    });
    scanStream = null;
  }
  const video = document.getElementById("scanner-video");
  if (video) {
    try {
      video.pause();
    } catch (e) {}
    video.srcObject = null;
  }
}

let lastScanCode = null;
let lastScanTime = 0;
function onScanSuccess(text) {
  const norm = normalizeCode(text);
  const now = Date.now();
  if (norm === lastScanCode && now - lastScanTime < 2500) return;
  lastScanCode = norm;
  lastScanTime = now;

  if (navigator.vibrate) navigator.vibrate(60);
  selectManualCode(norm);
  stopScanner();
  document.getElementById("search-input").value = "";
}

/* ===================== Navigation ===================== */

function goToScreen(name) {
  document.querySelectorAll(".screen").forEach((s) => s.classList.remove("active"));
  document.getElementById(`screen-${name}`).classList.add("active");
  document.querySelectorAll(".tab").forEach((t) => t.classList.toggle("active", t.dataset.tab === name));
  if (name !== "scan" && state.scannerActive) stopScanner();
  if (name === "session") {
    state.editingSessionIdx = null;
    state.editingPrepaIdx = null;
    renderSessionTab();
    renderPrepaTab();
  }
  if (name === "settings") renderStockStatus();
}

/* ===================== Réglages : envoi Gmail ===================== */

function renderMailSettings() {
  document.getElementById("mail-to").value = state.mail.to || "";
  document.getElementById("mail-client-id").value = state.mail.clientId || "";
  const ok = mailConfigured();
  document.getElementById("mail-status").innerHTML = ok
    ? `<div class="chip-status ok">✓ Envoi Gmail activé</div>`
    : `<div class="chip-status off">Non configuré : envoi par le menu de partage</div>`;
  document.getElementById("export-capa-btn").textContent = ok ? "Envoyer la capa par Gmail" : "Envoyer la capa (XLS 5.0/95)";
  document.getElementById("export-prepa-btn").textContent = ok ? "Envoyer la prépa par Gmail" : "Envoyer la prépa (XLS 5.0/95)";
}

function saveMailSettings() {
  const to = document.getElementById("mail-to").value.trim();
  const clientId = document.getElementById("mail-client-id").value.trim();
  const bad = parseRecipients(to).filter((e) => !isValidEmail(e));
  if (bad.length) {
    showToast(`Adresse invalide : ${bad[0]}`, 3500);
    return;
  }
  if (clientId && !/\.apps\.googleusercontent\.com$/.test(clientId)) {
    showToast("L'identifiant client doit finir par .apps.googleusercontent.com", 4000);
    return;
  }
  state.mail = { clientId, to };
  gmailToken = null;
  saveJSON(LS_KEYS.mail, state.mail);
  renderMailSettings();
  showToast(mailConfigured() ? "Envoi Gmail enregistré" : "Réglages enregistrés (destinataire et identifiant requis)");
}

function disableMail() {
  state.mail = { clientId: "", to: "" };
  gmailToken = null;
  saveJSON(LS_KEYS.mail, state.mail);
  renderMailSettings();
  showToast("Envoi Gmail désactivé");
}

/* ===================== Réglages avancés : mapping des colonnes ===================== */

function renderColMapForm() {
  const cm = state.colMap;
  const fields = [
    ["code", "Colonne code interne"],
    ["libelle", "Colonne désignation"],
    ["codebarre", "Colonne code-barres"],
    ["pv", "Colonne prix de vente (PV)"],
    ["entrepotStock", "Colonne stock entrepôt"],
    ["sitePrefix", "Préfixe colonnes stock site (siteN)"],
    ["capaPrefix", "Préfixe colonnes capa site (capaN)"],
  ];
  const el = document.getElementById("colmap-form");
  el.innerHTML = fields
    .map(
      ([key, label]) => `
      <label>${label}
        <input type="text" data-key="${key}" value="${escapeHtml(cm[key])}">
      </label>`
    )
    .join("");
}

function saveColMapForm() {
  const inputs = document.querySelectorAll("#colmap-form input");
  const newMap = { ...state.colMap };
  inputs.forEach((inp) => {
    newMap[inp.dataset.key] = inp.value.trim();
  });
  state.colMap = newMap;
  saveJSON(LS_KEYS.colMap, newMap);
  showToast("Correspondance des colonnes enregistrée");
  const rows = loadJSON(LS_KEYS.stock, null);
  if (rows) buildStockIndex(rows);
}

/* ===================== Initialisation ===================== */

function init() {
  tryRestoreStockFromCache();
  renderStockStatus();
  renderSessionTab();
  renderPrepaTab();
  renderColMapForm();
  renderMailSettings();
  goToScreen("scan");

  document.querySelectorAll(".tab").forEach((t) => {
    t.onclick = () => goToScreen(t.dataset.tab);
  });

  document.getElementById("stock-file-input").addEventListener("change", (e) => {
    const f = e.target.files[0];
    if (f) handleStockFile(f);
    e.target.value = "";
  });

  document.getElementById("scan-toggle-btn").onclick = () => {
    if (state.scannerActive) stopScanner();
    else startScanner();
  };

  const searchInput = document.getElementById("search-input");
  searchInput.addEventListener("input", () => {
    const results = searchArticles(searchInput.value);
    const box = document.getElementById("suggestions");
    if (!results.length) {
      box.innerHTML = "";
      return;
    }
    box.innerHTML = results
      .map(
        (item, i) => `
        <div class="suggestion-item" data-idx="${i}">
          <div class="s-name">${item.libelle ? escapeHtml(item.libelle) : "(sans désignation)"}</div>
          <div class="s-meta">Code ${escapeHtml(item.code)}${item.barcode ? " · CB " + escapeHtml(item.barcode) : ""}</div>
        </div>`
      )
      .join("");
    box.querySelectorAll(".suggestion-item").forEach((elx) => {
      elx.onclick = () => {
        selectArticle(results[Number(elx.dataset.idx)]);
        box.innerHTML = "";
        searchInput.value = "";
      };
    });
  });

  document.getElementById("search-go-btn").onclick = () => {
    const val = searchInput.value.trim();
    if (!val) return;
    const found = findExact(val);
    if (found) {
      selectArticle(found);
      document.getElementById("suggestions").innerHTML = "";
    } else {
      selectManualCode(val);
    }
  };

  document.getElementById("clear-session-btn").onclick = clearSession;
  document.getElementById("clear-prepa-btn").onclick = clearPrepa;
  document.getElementById("save-mail-btn").onclick = saveMailSettings;
  document.getElementById("disable-mail-btn").onclick = disableMail;
  document.getElementById("export-capa-btn").onclick = () => exportAndShare("capa");
  document.getElementById("export-prepa-btn").onclick = () => exportAndShare("prepa");

  document.getElementById("save-colmap-btn").onclick = saveColMapForm;
  document.getElementById("reset-colmap-btn").onclick = () => {
    state.colMap = { ...DEFAULT_COLMAP };
    saveJSON(LS_KEYS.colMap, state.colMap);
    renderColMapForm();
    showToast("Mapping réinitialisé");
  };

  document.getElementById("forget-stock-btn").onclick = () => {
    if (!confirm("Oublier le fichier stock importé ?")) return;
    localStorage.removeItem(LS_KEYS.stock);
    localStorage.removeItem(LS_KEYS.stockMeta);
    state.stock = null;
    state.stockMeta = null;
    renderStockStatus();
    showToast("Fichier stock oublié");
  };

  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("./sw.js").catch(() => {});
  }
}

document.addEventListener("DOMContentLoaded", init);

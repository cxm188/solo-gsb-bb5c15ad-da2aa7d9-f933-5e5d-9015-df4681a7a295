/* 电子表格公式影响分析前端（原生 JS，无构建步骤） */
"use strict";

const state = {
  wbId: null,
  summary: null,
  direction: "dependents",
  selectedKey: null,
};

const $ = (sel) => document.querySelector(sel);

const els = {
  uploadView: $("#uploadView"),
  analysisView: $("#analysisView"),
  loading: $("#loading"),
  fileInput: $("#fileInput"),
  reuploadInput: $("#reuploadInput"),
  uploadError: $("#uploadError"),
  globalActions: $("#globalActions"),
  exportBtn: $("#exportBtn"),
  wbTitle: $("#wbTitle"),
  statsBar: $("#statsBar"),
  warningsBox: $("#warningsBox"),
  sheetFilter: $("#sheetFilter"),
  cellSearch: $("#cellSearch"),
  formulaOnly: $("#formulaOnly"),
  cellTbody: $("#cellTbody"),
  cellListHint: $("#cellListHint"),
  directionSeg: $("#directionSeg"),
  impactEmpty: $("#impactEmpty"),
  impactResult: $("#impactResult"),
  selectedCell: $("#selectedCell"),
  selectedFormula: $("#selectedFormula"),
  directPill: $("#directPill"),
  indirectPill: $("#indirectPill"),
  totalPill: $("#totalPill"),
  impactCycleNote: $("#impactCycleNote"),
  impactTbody: $("#impactTbody"),
  cyclesCard: $("#cyclesCard"),
  cyclesList: $("#cyclesList"),
};

/* ---------------- 上传 ---------------- */
$("#chooseBtn").addEventListener("click", (e) => {
  e.stopPropagation();
  els.fileInput.click();
});
$("#dropzone").addEventListener("click", () => els.fileInput.click());
els.fileInput.addEventListener("change", () => {
  if (els.fileInput.files[0]) upload(els.fileInput.files[0]);
});
els.reuploadInput.addEventListener("change", () => {
  if (els.reuploadInput.files[0]) upload(els.reuploadInput.files[0]);
});

const dz = $("#dropzone");
["dragover", "dragenter"].forEach((ev) =>
  dz.addEventListener(ev, (e) => {
    e.preventDefault();
    dz.classList.add("drag");
  })
);
["dragleave", "drop"].forEach((ev) =>
  dz.addEventListener(ev, (e) => {
    e.preventDefault();
    dz.classList.remove("drag");
  })
);
dz.addEventListener("drop", (e) => {
  const f = e.dataTransfer.files[0];
  if (f) upload(f);
});

async function upload(file) {
  els.uploadError.hidden = true;
  els.loading.hidden = false;
  const fd = new FormData();
  fd.append("file", file);
  try {
    const res = await fetch("/api/workbooks", { method: "POST", body: fd });
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail || "上传失败");
    state.wbId = data.id;
    state.summary = data.summary;
    state.selectedKey = null;
    renderSummary();
    renderSheetFilter();
    renderCellTable();
    renderCycles();
    resetImpact();
    els.uploadView.hidden = true;
    els.analysisView.hidden = false;
    els.globalActions.hidden = false;
  } catch (err) {
    els.uploadError.hidden = false;
    els.uploadError.textContent = err.message;
  } finally {
    els.loading.hidden = true;
  }
}

/* ---------------- 概要 ---------------- */
const STAT_LABELS = [
  ["sheet_count", "工作表"],
  ["node_count", "依赖图单元格"],
  ["formula_count", "公式单元格"],
  ["edge_count", "依赖边"],
  ["cross_sheet_edge_count", "跨表引用"],
  ["external_ref_count", "外部工作簿引用"],
  ["cycle_count", "循环引用"],
  ["defined_name_count", "定义名称"],
];

function renderSummary() {
  const s = state.summary;
  els.wbTitle.textContent = `${s.filename}`;
  els.statsBar.innerHTML = STAT_LABELS.map(([key, lbl]) => {
    const warn = key === "cycle_count" && s.stats[key] > 0 ? " warn" : "";
    return `<div class="stat${warn}"><div class="num">${s.stats[key] ?? 0}</div><div class="lbl">${lbl}</div></div>`;
  }).join("");

  els.warningsBox.innerHTML = "";
  if (s.warnings && s.warnings.length) {
    els.warningsBox.innerHTML = s.warnings
      .map((w) => `<div class="warn-line">⚠ ${escapeHtml(w)}</div>`)
      .join("");
  }
}

function renderSheetFilter() {
  const current = els.sheetFilter.value;
  els.sheetFilter.innerHTML =
    `<option value="">全部工作表</option>` +
    state.summary.sheets
      .map(
        (sh) =>
          `<option value="${escapeAttr(sh.name)}">${escapeHtml(sh.name)}（${sh.formula_cells} 公式）</option>`
      )
      .join("");
  if ([...els.sheetFilter.options].some((o) => o.value === current)) {
    els.sheetFilter.value = current;
  }
}

/* ---------------- 单元格表 ---------------- */
[els.sheetFilter, els.formulaOnly].forEach((el) =>
  el.addEventListener("change", renderCellTable)
);
let searchTimer = null;
els.cellSearch.addEventListener("input", () => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(renderCellTable, 200);
});

async function renderCellTable() {
  const params = new URLSearchParams();
  if (els.sheetFilter.value) params.set("sheet", els.sheetFilter.value);
  if (els.formulaOnly.checked) params.set("formula_only", "true");
  if (els.cellSearch.value.trim()) params.set("q", els.cellSearch.value.trim());

  const res = await fetch(`/api/workbooks/${state.wbId}/nodes?${params}`);
  const data = await res.json();
  const cycleKeys = new Set(
    state.summary.cycles.flatMap((c) => c)
  );

  els.cellTbody.innerHTML = data.nodes
    .map((n) => {
      const detail = n.formula
        ? `<span class="formula-text">${escapeHtml(n.formula)}</span>`
        : `<span class="value-text">${escapeHtml(n.value ?? "")}</span>`;
      return `<tr data-key="${escapeAttr(n.key)}" class="${
        state.selectedKey === n.key ? "selected" : ""
      }">
        <td><span class="sheet-tag">${escapeHtml(n.sheet)}</span><span class="cell-coord">${n.coord}</span>
          ${cycleKeys.has(n.key) ? '<span class="badge cycle">循环</span>' : ""}
        </td>
        <td>${n.formula ? '<span class="badge formula">公式</span>' : '<span class="badge value">值</span>'}</td>
        <td>${detail}</td>
        <td><button class="btn tiny" data-analyze="${escapeAttr(n.key)}">分析影响</button></td>
      </tr>`;
    })
    .join("");

  const parts = [`共 ${data.count} 个单元格`];
  if (data.truncated) parts.push("结果过多，仅显示前 20000 个，请使用工作表筛选或搜索");
  els.cellListHint.textContent = parts.join(" · ");
}

els.cellTbody.addEventListener("click", (e) => {
  const btn = e.target.closest("[data-analyze]");
  if (!btn) return;
  analyzeKey(btn.dataset.analyze);
});

/* ---------------- 影响分析 ---------------- */
els.directionSeg.addEventListener("click", (e) => {
  const btn = e.target.closest("button[data-dir]");
  if (!btn) return;
  [...els.directionSeg.children].forEach((b) =>
    b.classList.toggle("active", b === btn)
  );
  state.direction = btn.dataset.dir;
  if (state.selectedKey) analyzeKey(state.selectedKey);
});

function resetImpact() {
  els.impactEmpty.hidden = false;
  els.impactResult.hidden = true;
}

async function analyzeKey(key) {
  const idx = key.indexOf("!");
  const sheet = key.slice(0, idx);
  const coord = key.slice(idx + 1);
  state.selectedKey = key;
  renderCellTable();

  els.impactEmpty.hidden = true;
  els.impactResult.hidden = false;
  els.loading.hidden = false;
  try {
    const params = new URLSearchParams({
      sheet,
      coord,
      direction: state.direction,
    });
    const res = await fetch(
      `/api/workbooks/${state.wbId}/impact?${params}`
    );
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail || "分析失败");
    renderImpact(data);
  } catch (err) {
    els.impactEmpty.hidden = false;
    els.impactResult.hidden = true;
    els.impactEmpty.textContent = err.message;
  } finally {
    els.loading.hidden = true;
  }
}

function renderImpact(data) {
  const sel = data.selected;
  els.selectedCell.textContent = `${sel.sheet}!${sel.coord}`;
  els.selectedFormula.textContent = sel.formula
    ? `  ${sel.formula}`
    : sel.value != null
    ? `  值: ${sel.value}`
    : "";

  const dirLabels =
    data.direction === "dependents"
      ? { direct: "直接受影响", indirect: "间接受影响", total: "受影响合计" }
      : { direct: "直接依赖", indirect: "间接依赖", total: "依赖合计" };
  els.directPill.textContent = `${dirLabels.direct} ${data.direct_count}`;
  els.indirectPill.textContent = `${dirLabels.indirect} ${data.indirect_count}`;
  els.totalPill.textContent = `${dirLabels.total} ${data.affected_count}`;

  if (data.affected_count === 0) {
    els.impactCycleNote.hidden = false;
    els.impactCycleNote.textContent =
      data.direction === "dependents"
        ? "没有任何公式引用该单元格：修改它不会影响其他单元格。"
        : "该单元格不依赖其他单元格。";
  } else {
    els.impactCycleNote.hidden = !data.affected.some((a) => a.on_cycle);
    if (!els.impactCycleNote.hidden) {
      els.impactCycleNote.textContent =
        "⚠ 关联中包含处于循环引用上的单元格，传播路径可能重复回到自身。";
    }
  }

  els.impactTbody.innerHTML = data.affected
    .map((a, i) => {
      const n = a.node;
      const pathsHtml = a.paths
        .map((path, pi) => {
          const chain = path
            .map((p, j) => {
              const cls =
                j === path.length - 1
                  ? "path-node target"
                  : "path-node";
              return `<span class="${cls}" title="${escapeAttr(
                p.formula || p.sheet + "!" + p.coord
              )}">${escapeHtml(p.sheet)}!${p.coord}</span>`;
            })
            .join('<span class="path-arrow">→</span>');
          const forms = path
            .filter((p) => p.formula)
            .map(
              (p) =>
                `<span class="path-formula">${escapeHtml(p.sheet)}!${p.coord}: ${escapeHtml(p.formula)}</span>`
            )
            .join("");
          return `<div class="path-chain">${
            pi === 0 ? "" : '<span class="path-more">备选路径：</span>'
          }${chain}${forms}</div>`;
        })
        .join("");
      return `<tr>
        <td>${i + 1}</td>
        <td><span class="sheet-tag">${escapeHtml(n.sheet)}</span><span class="cell-coord">${n.coord}</span>
          ${a.on_cycle ? '<span class="badge cycle">循环</span>' : ""}
          ${n.formula ? `<div class="formula-text">${escapeHtml(n.formula)}</div>` : `<div class="value-text">${escapeHtml(n.value ?? "")}</div>`}
        </td>
        <td><span class="rel-tag ${a.relation}">${
        a.relation === "direct" ? "直接" : "间接"
      }</span><span class="dist">${a.distance} 跳</span></td>
        <td><div class="paths">${pathsHtml}</div></td>
      </tr>`;
    })
    .join("");

  if (data.truncated) {
    els.impactTbody.insertAdjacentHTML(
      "beforeend",
      `<tr><td colspan="4" class="list-hint">受影响单元格过多，仅展示前 ${data.affected_count} 个，可导出 JSON 获取完整结果。</td></tr>`
    );
  }
}

/* ---------------- 循环引用 ---------------- */
function renderCycles() {
  const cycles = state.summary.cycles;
  if (!cycles.length) {
    els.cyclesCard.hidden = true;
    return;
  }
  els.cyclesCard.hidden = false;
  els.cyclesList.innerHTML = cycles
    .map(
      (c, i) =>
        `<div class="cycle-item"><span class="idx">循环 ${i + 1}（${
          c.length - 1
        } 个单元格）</span>${c
          .map(escapeHtml)
          .join('<span class="path-arrow">→</span>')}</div>`
    )
    .join("");
}

/* ---------------- 导出 ---------------- */
els.exportBtn.addEventListener("click", () => {
  if (!state.wbId) return;
  const a = document.createElement("a");
  a.href = `/api/workbooks/${state.wbId}/export`;
  a.download = `analysis-${state.wbId}.json`;
  a.click();
});

/* ---------------- 工具 ---------------- */
function escapeHtml(v) {
  return String(v ?? "").replace(
    /[&<>"']/g,
    (ch) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch])
  );
}
function escapeAttr(v) {
  return escapeHtml(v);
}

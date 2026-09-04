
let dataA = null;
let dataB = null;
let currentResult = null;
let currentMissing = null;
let currentSummary = null;

const $ = (id) => document.getElementById(id);

$("fileA").addEventListener("change", async (e) => {
  dataA = await readInputFile(e.target.files[0]);
  $("fileAName").textContent = e.target.files[0]?.name || "尚未选择";
  refreshInputs();
});

$("fileB").addEventListener("change", async (e) => {
  dataB = await readInputFile(e.target.files[0]);
  $("fileBName").textContent = e.target.files[0]?.name || "尚未选择";
  refreshInputs();
});

$("runBtn").addEventListener("click", runProcessing);
$("downloadBtn").addEventListener("click", downloadWorkbook);

async function readInputFile(file) {
  if (!file) return null;
  try {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array", cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
    return normalizeRows(rows);
  } catch (err) {
    showError("读取文件失败：" + err.message);
    return null;
  }
}

function normalizeRows(rows) {
  return rows.map((row) => {
    const out = {};
    for (const [rawKey, rawValue] of Object.entries(row)) {
      const key = String(rawKey).trim();
      let value = rawValue;
      if (typeof value === "string") value = value.trim();
      out[key] = value;
    }
    return out;
  });
}

function refreshInputs() {
  hideError();
  if (!dataA || !dataB) return;

  $("previewWrap").classList.remove("hidden");
  renderTable("previewA", dataA.slice(0, 8));
  renderTable("previewB", dataB.slice(0, 8));

  const colsA = collectColumns(dataA);
  const colsB = new Set(collectColumns(dataB));
  const common = colsA.filter((c) => colsB.has(c));

  if (!common.length) {
    $("configCard").classList.add("hidden");
    showError("两份文件没有同名字段，暂时无法关联。请确保至少有一个共同字段，例如 order_id。");
    return;
  }

  $("joinKey").innerHTML = common.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join("");
  $("configCard").classList.remove("hidden");
}

function cleanRows(rows) {
  const normalized = normalizeRows(rows);
  const seen = new Set();
  const unique = [];

  for (const row of normalized) {
    const keys = Object.keys(row).sort();
    const signature = JSON.stringify(keys.map((k) => [k, normalizeKeyValue(row[k])]));
    if (!seen.has(signature)) {
      seen.add(signature);
      unique.push(row);
    }
  }

  return { rows: unique, removed: normalized.length - unique.length };
}

function normalizeKeyValue(v) {
  if (v instanceof Date) return v.toISOString();
  if (v === null || v === undefined) return null;
  return String(v);
}

function joinRows(aRows, bRows, key, mode) {
  const aCols = collectColumns(aRows);
  const bCols = collectColumns(bRows);
  const sharedNonKey = new Set(aCols.filter((c) => c !== key && bCols.includes(c)));

  const bMap = new Map();
  bRows.forEach((row, idx) => {
    const k = normalizeKeyValue(row[key]);
    if (!bMap.has(k)) bMap.set(k, []);
    bMap.get(k).push({ row, idx });
  });

  const usedB = new Set();
  const out = [];

  const merge = (a, b) => {
    const result = {};
    result[key] = a?.[key] ?? b?.[key] ?? null;

    for (const col of aCols) {
      if (col === key) continue;
      result[sharedNonKey.has(col) ? `${col}_A` : col] = a ? (a[col] ?? null) : null;
    }

    for (const col of bCols) {
      if (col === key) continue;
      result[sharedNonKey.has(col) ? `${col}_B` : col] = b ? (b[col] ?? null) : null;
    }

    return result;
  };

  for (const a of aRows) {
    const k = normalizeKeyValue(a[key]);
    const matches = bMap.get(k) || [];

    if (matches.length) {
      for (const match of matches) {
        usedB.add(match.idx);
        out.push(merge(a, match.row));
      }
    } else if (mode === "left" || mode === "outer") {
      out.push(merge(a, null));
    }
  }

  if (mode === "outer") {
    bRows.forEach((b, idx) => {
      if (!usedB.has(idx)) out.push(merge(null, b));
    });
  }

  return out;
}

function runProcessing() {
  hideError();

  try {
    const key = $("joinKey").value;
    const mode = $("joinMode").value;

    const cleanA = cleanRows(dataA);
    const cleanB = cleanRows(dataB);

    const merged = joinRows(cleanA.rows, cleanB.rows, key, mode);
    const columns = collectColumns(merged);

    const missing = columns.map((field) => {
      let count = 0;
      for (const row of merged) {
        if (isMissing(row[field])) count++;
      }
      const rate = merged.length ? (count / merged.length) * 100 : 0;
      return {
        "字段": field,
        "缺失值数量": count,
        "缺失率": rate.toFixed(2) + "%"
      };
    });

    const missingTotal = missing.reduce((sum, r) => sum + r["缺失值数量"], 0);
    const modeText = {
      left: "左连接（保留文件 A 全部数据）",
      inner: "内连接（只保留两边都匹配的数据）",
      outer: "外连接（保留两边全部数据）"
    }[mode];

    const summary = [
      { "指标": "文件 A 原始行数", "结果": dataA.length },
      { "指标": "文件 B 原始行数", "结果": dataB.length },
      { "指标": "文件 A 删除重复行", "结果": cleanA.removed },
      { "指标": "文件 B 删除重复行", "结果": cleanB.removed },
      { "指标": "最终结果行数", "结果": merged.length },
      { "指标": "最终结果列数", "结果": columns.length },
      { "指标": "缺失值总数", "结果": missingTotal },
      { "指标": "关联字段", "结果": key },
      { "指标": "合并方式", "结果": modeText },
      { "指标": "处理时间", "结果": new Date().toLocaleString("zh-CN") }
    ];

    currentResult = merged;
    currentMissing = missing;
    currentSummary = summary;

    $("mRows").textContent = merged.length;
    $("mCols").textContent = columns.length;
    $("mDup").textContent = cleanA.removed + cleanB.removed;
    $("mMissing").textContent = missingTotal;

    renderTable("resultTable", merged);
    renderTable("missingTable", missing);
    $("resultCard").classList.remove("hidden");
    $("resultCard").scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (err) {
    showError("处理失败：" + err.message);
  }
}

function isMissing(v) {
  return v === null || v === undefined || (typeof v === "string" && v.trim() === "");
}

function collectColumns(rows) {
  const out = [];
  const seen = new Set();
  for (const row of rows || []) {
    for (const key of Object.keys(row)) {
      if (!seen.has(key)) {
        seen.add(key);
        out.push(key);
      }
    }
  }
  return out;
}

function renderTable(targetId, rows) {
  const target = $(targetId);
  if (!rows || rows.length === 0) {
    target.innerHTML = '<div style="padding:16px;color:#aab3c2">暂无数据</div>';
    return;
  }

  const cols = collectColumns(rows);
  const head = cols.map((c) => `<th>${escapeHtml(c)}</th>`).join("");
  const body = rows.map((row) => {
    return "<tr>" + cols.map((c) => `<td>${escapeHtml(formatValue(row[c]))}</td>`).join("") + "</tr>";
  }).join("");

  target.innerHTML = `<table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
}

function formatValue(v) {
  if (v === null || v === undefined || v === "") return "—";
  if (v instanceof Date) return v.toLocaleDateString("zh-CN");
  return String(v);
}

function downloadWorkbook() {
  if (!currentResult) return;

  const wb = XLSX.utils.book_new();
  const wsResult = XLSX.utils.json_to_sheet(currentResult);
  const wsSummary = XLSX.utils.json_to_sheet(currentSummary);
  const wsMissing = XLSX.utils.json_to_sheet(currentMissing);

  setColumnWidths(wsResult, currentResult);
  setColumnWidths(wsSummary, currentSummary);
  setColumnWidths(wsMissing, currentMissing);

  XLSX.utils.book_append_sheet(wb, wsResult, "合并结果");
  XLSX.utils.book_append_sheet(wb, wsSummary, "处理摘要");
  XLSX.utils.book_append_sheet(wb, wsMissing, "缺失值报告");

  XLSX.writeFile(wb, "processed_result.xlsx");
}

function setColumnWidths(ws, rows) {
  const cols = collectColumns(rows);
  ws["!cols"] = cols.map((col) => {
    let max = String(col).length;
    for (const row of rows) {
      max = Math.max(max, formatValue(row[col]).length);
    }
    return { wch: Math.min(Math.max(max + 2, 12), 36) };
  });
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function showError(msg) {
  $("errorBox").textContent = msg;
  $("errorBox").classList.remove("hidden");
}

function hideError() {
  $("errorBox").classList.add("hidden");
}

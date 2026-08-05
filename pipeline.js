// 随访表格数据处理 pipeline（浏览器端移植自 pipeline.py）
// 用 SheetJS 解析 Excel，关键字映射 + 状态推导 + 脱敏。无 DOM 依赖（XLSX 为全局变量）。

(function () {
const M = (typeof window !== "undefined" ? window : globalThis).Mapping;
const XLSX = (typeof window !== "undefined" ? window : globalThis).XLSX;

function normHeader(h) { return M.normHeader(h); }

function asList(v) { return v == null ? [] : (Array.isArray(v) ? v : [v]); }

// 把单元格值清洗为字符串：Date 保留时分，null/空 -> null
function cellStr(v) {
  if (v == null) return null;
  if (v instanceof Date) return fmtDateTime(v);
  const s = String(v).trim();
  if (s === "" || s === "nan" || s === "NaN" || s === "None") return null;
  return s;
}
function cellStrSafe(v) { return cellStr(v) == null ? "" : cellStr(v); }

function pad2(n) { return String(n).padStart(2, "0"); }
function fmtDateTime(d) {
  const Y = d.getFullYear(), Mo = pad2(d.getMonth() + 1), D = pad2(d.getDate());
  const h = pad2(d.getHours()), mi = pad2(d.getMinutes()), s = pad2(d.getSeconds());
  const date = `${Y}-${Mo}-${D}`;
  return (h === "00" && mi === "00" && s === "00") ? date : `${date} ${h}:${mi}:${s}`;
}

// 单值字段关键字启发式映射；followup_time 需区分「执行时间」与「计划执行时间」
function mapColumns(rawCols) {
  const assigned = {};
  const used = new Set();
  for (const [field, kws] of M.KEYWORD_RULES) {
    let cand;
    if (field === "followup_time") {
      const clean = rawCols.filter(rc => !used.has(rc) &&
        normHeader(rc).includes("执行时间") && !normHeader(rc).includes("计划"));
      cand = (clean.length ? clean : rawCols)
        .filter(rc => !used.has(rc))
        .map(rc => [rc, normHeader(rc)]);
    } else {
      cand = rawCols.filter(rc => !used.has(rc)).map(rc => [rc, normHeader(rc)]);
    }
    let best = null, bestHits = 0;
    for (const [rc, nrc] of cand) {
      let hits = 0;
      for (const kw of kws) if (nrc.includes(kw)) hits++;
      if (hits > 0 && hits > bestHits) { bestHits = hits; best = rc; }
    }
    if (best !== null) { assigned[field] = best; used.add(best); }
  }
  for (const [field, kws] of M.KEYWORD_RULES_MULTI) {
    const cols = rawCols.filter(rc => !used.has(rc) && kws.some(kw => normHeader(rc).includes(kw)));
    if (cols.length) { assigned[field] = cols; cols.forEach(c => used.add(c)); }
  }
  return assigned;
}

function detectSource(columns, colmap) {
  if (colmap) {
    if (colmap._purchased_on_time) return "overdue_purchase";
    if (colmap._status_period) return "enrollment";
    if (colmap._nonstd_usage || colmap._is_dropout || colmap._follow_confirm) return "routine";
  }
  const normCols = columns.map(normHeader);
  for (const [stype, sig] of M.SIGNATURES) {
    const nsig = normHeader(sig);
    if (normCols.some(c => c.includes(nsig))) return stype;
  }
  return "unknown";
}

// 在原始数组（array-of-arrays，已转字符串）的前 maxScan 行中找最像表头的一行
function detectHeaderRow(aoaStr, maxScan = 15) {
  if (!aoaStr || !aoaStr.length) return 0;
  const kwPools = [];
  for (const [, kws] of M.KEYWORD_RULES) kwPools.push(...kws);
  for (const [, kws] of M.KEYWORD_RULES_MULTI) kwPools.push(...kws);
  for (const [, sig] of M.SIGNATURES) kwPools.push(sig);
  const kwLower = kwPools.map(k => k.toLowerCase());

  let bestRow = 0, bestScore = -1;
  const scan = Math.min(maxScan, aoaStr.length);
  for (let i = 0; i < scan; i++) {
    const row = aoaStr[i];
    const cells = row.map(c => String(c == null ? "" : c).trim())
      .filter(c => c !== "" && c !== "nan" && c !== "None");
    if (!cells.length) continue;
    let hit = 0;
    for (const c of cells) {
      const cl = c.toLowerCase();
      if (kwLower.some(kw => cl.includes(kw))) hit++;
    }
    const ratio = cells.length / Math.max(1, row.length);
    const score = hit * 3 + ratio * 2;
    if (score > bestScore) { bestScore = score; bestRow = i; }
  }
  return bestRow;
}

// 把一行数组转成 列名->清洗后字符串 的对象
function rowToObj(cols, r) {
  const o = {};
  cols.forEach((c, i) => { o[c] = cellStr(r[i]); });
  return o;
}

function normalizeRows(rows, cols, sourceFile, sheetName) {
  const colmap = mapColumns(cols);
  const sourceType = detectSource(cols, colmap);
  const records = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const rid = `${sourceFile}::${sourceType}::${sheetName || "sheet"}::${i}`;
    const rec = { source_file: sourceFile, source_type: sourceType, _row_id: rid };
    for (const field of M.CANONICAL_FIELDS) {
      if (["source_file", "source_type", "remarks", "medication_status",
           "irregularity_subtype", "stop_reduce_reason"].includes(field)) continue;
      const col = colmap[field];
      rec[field] = col ? cellStr(row[col]) : null;
    }
    // 备注：合并「自由文本」列 + 随访小结
    let rcols = asList(colmap.remarks);
    if (colmap.summary) rcols = rcols.concat([colmap.summary]);
    rcols = rcols.filter(c => cols.includes(c));
    const parts = [];
    for (const c of rcols) { const tv = cellStr(row[c]); if (tv) parts.push(tv); }
    rec.remarks = parts.length ? parts.join("\n") : null;
    // 停药/减量根本原因
    const reasonCols = asList(colmap.stop_reduce_reason).filter(c => cols.includes(c));
    const rparts = []; const seen = new Set();
    for (const c of reasonCols) {
      const tv = cellStr(row[c]);
      if (tv && !seen.has(tv)) { seen.add(tv); rparts.push(tv); }
    }
    rec.stop_reduce_reason = rparts.length ? rparts.join("\n") : null;
    // 用药状态 + 不规范下钻
    const status = M.deriveStatus(sourceType, row, colmap);
    rec.medication_status = status;
    rec.medication_status_raw = M._rawStatusText(sourceType, row, colmap) || status;
    const sub = status === "不规范用药" ? M.deriveIrregularitySubtype(sourceType, row, colmap) : null;
    rec.irregularity_subtype = sub;
    rec.irregularity_subtype_raw = sub ? M._rawSubtypeText(sourceType, row, colmap) : null;
    records.push(rec);
  }
  return records;
}

async function loadWorkbook(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: true });
  const out = [];
  for (const name of wb.SheetNames) {
    const sheet = wb.Sheets[name];
    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null, cellDates: true });
    const aoaStr = aoa.map(r => r.map(cellStrSafe));
    const hr = detectHeaderRow(aoaStr);
    let cols, rows;
    if (hr === 0) {
      cols = aoa[0].map((c, i) => String(c == null ? "" : c).trim() || `col_${i}`);
      rows = aoa.slice(1).map(r => rowToObj(cols, r));
    } else {
      cols = aoa[hr].map((c, i) => String(c == null ? "" : c).trim() || `col_${i}`);
      rows = aoa.slice(hr + 1).map(r => rowToObj(cols, r));
    }
    out.push({ source_file: file.name, sheet_name: name, cols, rows });
  }
  return out;
}

async function processFiles(files) {
  let records = [];
  for (const f of files) {
    try {
      const sheets = await loadWorkbook(f);
      for (const sh of sheets) {
        records = records.concat(normalizeRows(sh.rows, sh.cols, sh.source_file, sh.sheet_name));
      }
    } catch (e) {
      console.warn("处理文件失败", f.name, e);
    }
  }
  return records;
}

// 脱敏姓名与电话（展示层默认只给脱敏值）
function desensitize(rec, namePlain = false, phonePlain = false) {
  const out = Object.assign({}, rec);
  const name = out.patient_name;
  if (name && !namePlain) {
    out.patient_name = name.length >= 2 ? name[0] + "**" : "**";
  }
  const phone = out.phone;
  if (phone && !phonePlain) {
    const s = String(phone);
    const digits = s.replace(/\D/g, "");
    if (digits.length >= 7) out.phone = digits.slice(0, 3) + "****" + digits.slice(-4);
    else out.phone = s.replace(/\d/g, "*");
  }
  return out;
}

if (typeof window !== "undefined") {
  window.Pipeline = { mapColumns, detectSource, detectHeaderRow, normalizeRows,
    loadWorkbook, processFiles, desensitize, cellStr, fmtDateTime };
}
})();

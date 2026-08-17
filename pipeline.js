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
// 返回值统一为「列索引」（或索引数组），下游按索引读取，避免多模板合并表中重名列互相覆盖。
function mapColumns(rawCols) {
  const assigned = {};
  const used = new Set();
  for (const [field, kws] of M.KEYWORD_RULES) {
    const isCollective = field.charAt(0) === "_";
    let cand;
    if (field === "followup_time") {
      const clean = [];
      rawCols.forEach((rc, idx) => {
        if (!used.has(idx) && normHeader(rc).includes("执行时间") && !normHeader(rc).includes("计划"))
          clean.push([idx, normHeader(rc)]);
      });
      const base = clean.length ? clean : rawCols.map((rc, idx) => [idx, normHeader(rc)]);
      cand = base.filter(([idx]) => !used.has(idx));
    } else {
      cand = rawCols.map((rc, idx) => [idx, normHeader(rc)]).filter(([idx]) => !used.has(idx));
    }
    if (isCollective) {
      // 多模板合并表：同一逻辑字段可能有多个列副本（如「是否计划按时用药?」出现 3 次），
      // 收集所有命中列，deriveStatus 会跨这些列读取（每行仅其所属模板那一列有值）。
      const matched = cand.filter(([, nrc]) => kws.some(kw => nrc.includes(kw)));
      if (matched.length) {
        const idxs = matched.map(([idx]) => idx);
        assigned[field] = idxs.length === 1 ? idxs[0] : idxs;
        idxs.forEach(i => used.add(i));
      }
    } else {
      let best = null, bestHits = 0;
      for (const [idx, nrc] of cand) {
        let hits = 0;
        for (const kw of kws) if (nrc.includes(kw)) hits++;
        if (hits > 0 && hits > bestHits) { bestHits = hits; best = idx; }
      }
      if (best !== null) { assigned[field] = best; used.add(best); }
    }
  }
  for (const [field, kws] of M.KEYWORD_RULES_MULTI) {
    const cols = [];
    rawCols.forEach((rc, idx) => {
      if (!used.has(idx) && kws.some(kw => normHeader(rc).includes(kw))) cols.push(idx);
    });
    if (cols.length) { assigned[field] = cols; cols.forEach(i => used.add(i)); }
  }
  return assigned;
}

function detectSource(columns, colmap) {
  if (colmap) {
    if (colmap._purchased_on_time) return "overdue_purchase";
    if (colmap._status_period) return "enrollment";
    if (colmap._nonstd_usage || colmap._is_dropout || colmap._follow_confirm || colmap._near_usage) return "routine";
  }
  const normCols = columns.map(normHeader);
  for (const [stype, sig] of M.SIGNATURES) {
    const nsig = normHeader(sig);
    if (normCols.some(c => c.includes(nsig))) return stype;
  }
  return "unknown";
}

// 逐行判定来源：多模板合并表中，每行只属于其中一个模板（仅其模板的关键列有值）。
// 因此按「哪个来源的关键列在本行有实际取值」来归类，比文件级判定更准。
function detectSourceRow(row, colmap, fallback) {
  const filled = (field) => {
    const c = colmap[field];
    if (c == null) return false;
    const idxs = Array.isArray(c) ? c : [c];
    return idxs.some(i => cellStr(row[i]) != null);
  };
  // 日常随访：用药状态/易脱落/医嘱确认 任一有值
  if (filled("_usage_status") || filled("_is_dropout") || filled("_follow_confirm") || filled("_nonstd_usage") || filled("_near_usage")) return "routine";
  // 过期购药
  if (filled("_purchased_on_time")) return "overdue_purchase";
  // 入组
  if (colmap._status_period != null && cellStr(row[colmap._status_period]) != null) return "enrollment";
  return fallback;
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

// 把一行数组转成 列索引->清洗后字符串 的对象
// 用索引做 key 而非列名：多模板合并表中重名列（如「患者主诉」「是否计划按时用药?」多次出现）
// 若用列名做 key 会互相覆盖，导致只读到最后一列。
function rowToObj(cols, r) {
  const o = {};
  for (let i = 0; i < r.length; i++) o[i] = cellStr(r[i]);
  return o;
}

function normalizeRows(rows, cols, sourceFile, sheetName) {
  const colmap = mapColumns(cols);
  const fileSource = detectSource(cols, colmap);
  const records = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const sourceType = detectSourceRow(row, colmap, fileSource) || "unknown";
    const rid = `${sourceFile}::${sourceType}::${sheetName || "sheet"}::${i}`;
    const rec = { source_file: sourceFile, source_type: sourceType, _row_id: rid };
    for (const field of M.CANONICAL_FIELDS) {
      if (["source_file", "source_type", "remarks", "medication_status",
           "irregularity_subtype", "stop_reduce_reason"].includes(field)) continue;
      const col = colmap[field];
      if (col == null) { rec[field] = null; continue; }
      if (Array.isArray(col)) {
        const vs = col.map(c => cellStr(row[c])).filter(Boolean);
        rec[field] = vs.length ? vs.join("\n") : null;
      } else {
        rec[field] = cellStr(row[col]);
      }
    }
    // 备注：合并「自由文本」列 + 随访小结
    let rcols = asList(colmap.remarks);
    if (colmap.summary) rcols = rcols.concat([colmap.summary]);
    // colmap 存的是列索引（索引化行存储）；row 是对象无 length，colmap 索引本身即合法，仅排除 null
    rcols = rcols.filter(c => c != null);
    const parts = [];
    for (const c of rcols) { const tv = cellStr(row[c]); if (tv && !M.isPlaceholder(tv)) parts.push(tv); }
    rec.remarks = parts.length ? parts.join("\n") : null;
    // 停药/减量根本原因：colmap 存的是列索引（索引化行存储）；row 是对象无 length，仅排除 null
    const reasonCols = asList(colmap.stop_reduce_reason).filter(c => c != null);
    const rparts = []; const seen = new Set();
    for (const c of reasonCols) {
      const tv = cellStr(row[c]);
      if (tv && !seen.has(tv)) { seen.add(tv); rparts.push(tv); }
    }
    // 剔除「无/不详/未知」等占位值；若全被剔除则回退保留原值，避免丢失原文
    const rMeaningful = rparts.filter(p => !M.isPlaceholder(p));
    const rFinal = rMeaningful.length ? rMeaningful : rparts;
    rec.stop_reduce_reason = rFinal.length ? rFinal.join("\n") : null;
    // 用药状态 + 不规范下钻
    const status = M.deriveStatus(sourceType, row, colmap);
    rec.medication_status = status;
    rec.medication_status_raw = M._rawStatusText(sourceType, row, colmap) || status;
    const sub = status === "不规范用药" ? M.deriveIrregularitySubtype(sourceType, row, colmap) : null;
    rec.irregularity_subtype = sub;
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

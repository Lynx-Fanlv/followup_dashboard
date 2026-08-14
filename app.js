// 随访数据看板（纯本地版）前端编排层
// 全部计算在浏览器内完成：Excel 解析(pipeline.js) → 归一化 → 本地聚合 → 渲染。
// 无后端、无网络请求；数据从不离开本机。

(function () {
const M = window.Mapping;
const STATUS_TAXONOMY = M.STATUS_TAXONOMY;
const SUBTYPE_TAXONOMY = M.SUBTYPE_TAXONOMY;
const EXPORT_COLS = [
  ["patient_name", "患者"],
  ["phone", "电话"],
  ["gender", "性别"],
  ["age", "年龄"],
  ["followup_time", "随访时间"],
  ["drug_product", "药品"],
  ["indication", "适应症"],
  ["pharmacy", "药店"],
  ["medication_status_raw", "用药状态"],
  ["irregularity_subtype", "不规范类型"],
  ["stop_reduce_reason", "停药/减量根本原因"],
  ["remarks", "备注"],
];

// 内存存储（数据不出本机）。records=全部记录；files=已加载文件清单（文件管理）；seq=自增编号。
const STORE = { records: [], files: [], seq: 0 };
let pendingFiles = [];          // 待「开始分析」的 File 列表
let SNAP_MODE = false;          // 是否为快照打开模式（只读）

const TAXONOMY = STATUS_TAXONOMY;
const SRC_LABEL = { routine: "日常随访", enrollment: "入组", overdue_purchase: "超期未购药", unknown: "未知" };
const STATUS_COLOR = { "规范用药": "#0f9d6b", "不规范用药": "#e03131", "脱落停药": "#e8590c", "其他": "#868e96" };
const DETAIL = [["patient_name", "患者"], ["phone", "电话"], ["gender", "性别"], ["age", "年龄"],
  ["followup_time", "随访时间"], ["drug_product", "药品"], ["indication", "适应症"],
  ["pharmacy", "药店"], ["medication_status_raw", "用药状态"],
  ["irregularity_subtype", "不规范类型"],
  ["stop_reduce_reason", "停药/减量根本原因"], ["remarks", "备注"]];

const state = { status: new Set(), subtype: null, drugs: new Set(), pharmacies: new Set(), start: null, end: null,
  q: "", view: "detail", page: 1, pageSize: 50, hiddenCols: new Set(),
  plainName: false, plainPhone: false };
let CURRENT = { summary: null, global: null };
let DATA = { rows: [], patients: [] };

const $ = s => document.querySelector(s);
const drop = $("#drop"), fileInput = $("#fileInput");

/* ============ 工具函数 ============ */
function esc(s) { return String(s == null ? "" : s).replace(/[&<>]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c])); }
function fmtSize(n) {
  if (n < 1024) return n + " B";
  if (n < 1024 * 1024) return (n / 1024).toFixed(1) + " KB";
  return (n / 1024 / 1024).toFixed(2) + " MB";
}
function showLoading(txt) { $("#loadingTxt").textContent = txt || "正在处理…"; $("#loading").classList.remove("hidden"); }
function hideLoading() { $("#loading").classList.add("hidden"); }
function download(blob, name) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(a.href);
}
function cleanRec(r) {
  const c = Object.assign({}, r);
  delete c.source_file; delete c._file_id; delete c._file_label; delete c._row_id;
  return c;
}

/* ============ 本地数据层（移植自 app.py） ============ */
function _extract_date(s) {
  if (!s) return null;
  s = String(s).trim();
  if (s.length >= 10 && s[4] === "-" && s[7] === "-") return s.slice(0, 10);
  return null;
}
function _month(s) {
  const d = _extract_date(s);
  return d ? d.slice(0, 7) : null;
}
function countBy(arr, keyFn) {
  const m = {};
  for (const x of arr) { const k = keyFn(x); m[k] = (m[k] || 0) + 1; }
  return m;
}
function build_summary(records) {
  const months = records.map(r => _month(r.followup_time)).filter(Boolean).sort();
  const by_month = {};
  for (const m of months) by_month[m] = (by_month[m] || 0) + 1;
  return {
    total: records.length,
    by_status: countBy(records, r => r.medication_status),
    by_subtype: countBy(records.filter(r => r.irregularity_subtype), r => r.irregularity_subtype),
    by_source: countBy(records, r => r.source_type),
    by_task_status: countBy(records, r => r.task_status || "未知"),
    by_drug: countBy(records, r => r.drug_product || "未知"),
    by_pharmacy: countBy(records, r => r.pharmacy || "未知"),
    by_month,
    taxonomy: STATUS_TAXONOMY,
    subtypes: SUBTYPE_TAXONOMY,
  };
}
// kw: {status:[]|null, subtype, src, drug:[]|null, pharmacy:[]|null, start, end, q}
function filter_records(records, kw) {
  let recs = records;
  const { status, subtype, src, drug, pharmacy, start, end, q } = kw;
  if (status && status.length) {
    const sset = new Set(status);
    recs = recs.filter(r => sset.has(r.medication_status));
  }
  if (subtype) recs = recs.filter(r => r.irregularity_subtype === subtype);
  if (src) recs = recs.filter(r => r.source_type === src);
  if (drug && drug.length) {
    const dset = new Set(drug);
    recs = recs.filter(r => dset.has(r.drug_product || "未知"));
  }
  if (pharmacy && pharmacy.length) {
    const pset = new Set(pharmacy);
    recs = recs.filter(r => pset.has(r.pharmacy || "未知"));
  }
  if (q) {
    const ql = q.trim().toLowerCase();
    if (ql) {
      recs = recs.filter(r => {
        for (const f of ["patient_name", "phone", "pharmacy", "drug_product"]) {
          const v = r[f];
          if (v && String(v).toLowerCase().includes(ql)) return true;
        }
        return false;
      });
    }
  }
  if (start || end) {
    const kept = [];
    for (const r of recs) {
      const d = _extract_date(r.followup_time);
      if (d == null) continue;
      if (start && d < start) continue;
      if (end && d > end) continue;
      kept.push(r);
    }
    recs = kept;
  }
  return recs;
}
// 分面统计：exclude 维度不参与过滤（使各筛选项互不遮蔽）
function _facet(records, kw, exclude) {
  const k = Object.assign({}, kw);
  if (exclude === "status") k.status = null;
  else if (exclude === "subtype") k.subtype = null;
  else if (exclude === "source") k.src = null;
  else if (exclude === "drug") k.drug = null;
  else if (exclude === "pharmacy") k.pharmacy = null;
  else if (exclude === "time") { k.start = null; k.end = null; }
  return filter_records(records, k);
}
function currentKw() {
  return {
    status: state.status.size ? [...state.status] : null,
    subtype: state.subtype || null,
    src: null,
    drug: state.drugs.size ? [...state.drugs] : null,
    pharmacy: state.pharmacies.size ? [...state.pharmacies] : null,
    start: state.start || null,
    end: state.end || null,
    q: state.q || null,
  };
}
// 镜像 /api/summary：总记录数按全部筛选，各维度按「除自身外」筛选（faceted）
function summaryLocal(kw) {
  const base = build_summary(_facet(STORE.records, kw));
  base.by_status = build_summary(_facet(STORE.records, kw, "status")).by_status;
  base.by_subtype = build_summary(_facet(STORE.records, kw, "subtype")).by_subtype;
  base.by_source = build_summary(_facet(STORE.records, kw, "source")).by_source;
  base.by_drug = build_summary(_facet(STORE.records, kw, "drug")).by_drug;
  base.by_pharmacy = build_summary(_facet(STORE.records, kw, "pharmacy")).by_pharmacy;
  base.by_month = build_summary(_facet(STORE.records, kw, "time")).by_month;
  base.by_task_status = countBy(STORE.records, r => r.task_status || "未知");
  return base;
}
// 镜像 /api/patients：按 真实姓名+电话 分组；status/时间仅作「名单入选」门槛，不裁剪卡内记录
function patientsAgg(records, kw) {
  const status_sel = kw.status;
  const recs = filter_records(records, {
    status: null, subtype: kw.subtype, src: kw.src, drug: kw.drug, pharmacy: kw.pharmacy,
    start: null, end: null, q: kw.q,
  });
  const start = kw.start, end = kw.end;
  const pn = state.plainName, pp = state.plainPhone;
  const groups = {};
  for (const r of recs) {
    const key = (r.patient_name || "") + "\u0000" + (r.phone || "");
    (groups[key] = groups[key] || []).push(r);
  }
  const out = [];
  for (const key in groups) {
    const rs = groups[key];
    const dates = rs.map(x => _extract_date(x.followup_time)).filter(Boolean);
    const latest = dates.length ? dates.reduce((a, b) => a > b ? a : b) : "";
    if (start || end) {
      if (!latest) continue;
      if (start && latest < start) continue;
      if (end && latest > end) continue;
    }
    if (status_sel && status_sel.length) {
      const sset = new Set(status_sel);
      if (!rs.some(r => sset.has(r.medication_status))) continue;
    }
    const name = rs[0].patient_name || "";
    const phone = rs[0].phone || "";
    const maskedName = pn ? name : (name.length >= 2 ? name[0] + "**" : (name ? "**" : "(未知)"));
    const digits = String(phone).replace(/\D/g, "");
    const maskedPhone = pp ? phone : (digits.length >= 7 ? digits.slice(0, 3) + "****" + digits.slice(-4) : phone);
    out.push({
      patient_name: maskedName,
      phone: maskedPhone,
      count: rs.length,
      by_status: countBy(rs, r => r.medication_status),
      latest,
      drugs: [...new Set(rs.map(x => x.drug_product).filter(Boolean))].sort(),
      records: rs.map(r => window.Pipeline.desensitize(r, pn, pp)),
    });
  }
  out.sort((a, b) => (b.count - a.count) || (a.patient_name < b.patient_name ? -1 : 1));
  return out;
}

/* ============ 渲染层（移植自 dashboard.html，移除 fetch） ============ */
function renderSummary(d) {
  const total = d.total;
  let html = `<div class="total"><div class="n">${total}</div><div class="l">总记录数</div></div>`;
  for (const s of TAXONOMY) {
    const n = (d.by_status && d.by_status[s]) || 0;
    const active = state.status.has(s) ? "active" : "";
    html += `<div class="card c-${s} ${active}" data-status="${s}" title="点击筛选「${s}」，再次点击取消">
      <span class="dot"></span><div class="n">${n}</div><div class="l">${s}</div><span class="cue">点选</span></div>`;
  }
  $("#summary").innerHTML = html;
  document.querySelectorAll(".card").forEach(c => {
    c.onclick = () => {
      const s = c.dataset.status;
      if (state.status.has(s)) state.status.delete(s); else state.status.add(s);
      if (!state.status.has("不规范用药")) state.subtype = null;
      state.page = 1;
      refresh();
    };
  });
  buildMs("drugMs", d.by_drug, state.drugs, "药品");
  buildMs("pharmMs", d.by_pharmacy, state.pharmacies, "药店");
  renderSubtypeBar();
  updateFilterInfo();
}

function buildMs(msId, data, stateSet, label) {
  const entries = Object.entries(data || {});
  const list = $("#" + msId + "List");
  list.innerHTML = entries.map(([k, v]) =>
    `<label><input type="checkbox" value="${esc(k)}" ${stateSet.has(k) ? "checked" : ""}><span>${esc(k)}</span><span class="ms-cnt">${v}</span></label>`).join("");
  list.querySelectorAll("input").forEach(cb => {
    cb.onchange = () => { if (cb.checked) stateSet.add(cb.value); else stateSet.delete(cb.value); state.page = 1; refresh(); };
  });
  const panel = $("#" + msId + "Panel");
  let search = panel.querySelector(".ms-search");
  if (!search) {
    search = document.createElement("input");
    search.className = "ms-search";
    search.placeholder = "输入关键字检索…";
    list.parentNode.insertBefore(search, list);
  }
  const applyFilter = () => {
    const q = search.value.trim().toLowerCase();
    list.querySelectorAll("label").forEach(lb => {
      const k = lb.querySelector("input").value;
      const hit = !q || k.toLowerCase().includes(q) || stateSet.has(k);
      lb.style.display = hit ? "" : "none";
    });
  };
  search.oninput = applyFilter;
  applyFilter();
  panel.querySelectorAll(".ms-act").forEach(a => a.onclick = () => {
    if (a.dataset.act === "all") entries.forEach(([k]) => stateSet.add(k)); else stateSet.clear();
    state.page = 1; refresh();
  });
  const n = stateSet.size, tot = entries.length;
  const txt = (n === 0 || n === tot) ? "全部" : (n + "项");
  $("#" + msId + "Btn").innerHTML = `${label}：${txt} <span class="ms-caret">▾</span>`;
  $("#" + msId + "Btn").classList.toggle("has-sel", n > 0 && n < tot);
  $("#" + msId + "Count").textContent = n === 0 ? "未选（=全部）" : ("已选 " + n + "/" + tot);
}

function renderGlobal(d) {
  if (!d) return;
  const stCards = TAXONOMY.map(s =>
    `<div class="g-card ${s}"><span class="g-num">${(d.by_status && d.by_status[s]) || 0}</span><span class="g-cap">${s}</span></div>`).join("");
  const tc = d.by_task_status || {};
  const done = tc["已完成"] || 0;
  const ignored = Object.entries(tc).filter(([k]) => k !== "已完成").reduce((a, [, v]) => a + v, 0);
  const taskCards =
    `<div class="g-card task done"><span class="g-num">${done}</span><span class="g-cap">已完成</span></div>` +
    `<div class="g-card task ignored"><span class="g-num">${ignored}</span><span class="g-cap">忽略或改期</span></div>`;
  $("#globalRow").innerHTML =
    `<div class="g-head"><span class="g-title">全量统计</span><span class="g-sub">基于全部上传数据 · 不随筛选变化</span></div>` +
    `<div class="g-grid">${stCards}${taskCards}` +
    `<div class="g-card total"><span class="g-num">${d.total || 0}</span><span class="g-cap">合计</span></div>` +
    `</div>`;
}

function renderCharts(d) {
  $("#donutChart").innerHTML = donutChart(d.by_status || {});
  $("#drugChart").innerHTML = barChart(d.by_drug || {});
  $("#trendChart").innerHTML = lineChart(d.by_month || {});
}
function donutChart(data) {
  const entries = TAXONOMY.map(s => [s, data[s] || 0]).filter(([, v]) => v > 0);
  const total = entries.reduce((a, [, v]) => a + v, 0);
  if (!total) return '<div class="chart-empty">暂无数据</div>';
  const cx = 70, cy = 70, r = 54, ir = 31;
  let angle = -Math.PI / 2, paths = "";
  entries.forEach(([k, v]) => {
    const frac = v / total, a2 = angle + frac * 2 * Math.PI, large = frac > 0.5 ? 1 : 0;
    const x1 = cx + r * Math.cos(angle), y1 = cy + r * Math.sin(angle);
    const x2 = cx + r * Math.cos(a2), y2 = cy + r * Math.sin(a2);
    const x3 = cx + ir * Math.cos(a2), y3 = cy + ir * Math.sin(a2);
    const x4 = cx + ir * Math.cos(angle), y4 = cy + ir * Math.sin(angle);
    paths += `<path d="M${x1.toFixed(1)},${y1.toFixed(1)} A${r},${r} 0 ${large} 1 ${x2.toFixed(1)},${y2.toFixed(1)} L${x3.toFixed(1)},${y3.toFixed(1)} A${ir},${ir} 0 ${large} 0 ${x4.toFixed(1)},${y4.toFixed(1)} Z" fill="${STATUS_COLOR[k]}"><title>${k} ${v}</title></path>`;
    angle = a2;
  });
  const legend = entries.map(([k, v]) =>
    `<div class="lg-item"><span class="lg-dot" style="background:${STATUS_COLOR[k]}"></span>${k} ${v}</div>`).join("");
  return `<div class="donut-flex"><svg viewBox="0 0 140 140" width="118" height="118">${paths}<text x="70" y="76" text-anchor="middle" font-size="19" font-weight="700" fill="#1f2937">${total}</text></svg><div class="lg">${legend}</div></div>`;
}
function barChart(data) {
  const entries = Object.entries(data).sort((a, b) => b[1] - a[1]).slice(0, 5);
  if (!entries.length) return '<div class="chart-empty">暂无数据</div>';
  const max = Math.max(...entries.map(e => e[1]));
  const rows = entries.map(([k, v]) => {
    const w = Math.max(2, Math.round(v / max * 100));
    const lbl = k.length > 7 ? k.slice(0, 7) + "…" : k;
    return `<div class="bar-row"><span class="bar-label" title="${esc(k)}">${esc(lbl)}</span><span class="bar-track"><span class="bar-fill" style="width:${w}%"></span></span><span class="bar-val">${v}</span></div>`;
  }).join("");
  return `<div class="bars">${rows}</div>`;
}
function lineChart(data) {
  const entries = Object.entries(data);
  if (!entries.length) return '<div class="chart-empty">暂无数据</div>';
  const W = 320, H = 120, P = 26;
  const max = Math.max(...entries.map(e => e[1]), 1);
  const n = entries.length;
  const X = i => n === 1 ? W / 2 : P + i * (W - 2 * P) / (n - 1);
  const Y = v => H - P - (v / max) * (H - 2 * P);
  const pts = entries.map(([, v], i) => `${X(i).toFixed(1)},${Y(v).toFixed(1)}`).join(" ");
  const dots = entries.map(([k, v], i) => `<circle cx="${X(i).toFixed(1)}" cy="${Y(v).toFixed(1)}" r="3" fill="#3b5bdb"><title>${k}: ${v}条</title></circle>`).join("");
  const step = Math.ceil(n / 8);
  const labels = entries.map(([k], i) => (i % step === 0 || i === n - 1) ? `<text x="${X(i).toFixed(1)}" y="${H - 8}" font-size="9" text-anchor="middle" fill="#868e96">${k.slice(2)}</text>` : "").join("");
  return `<svg viewBox="0 0 ${W} ${H}" width="100%" height="120"><polyline points="${pts}" fill="none" stroke="#3b5bdb" stroke-width="2"/>${dots}${labels}</svg>`;
}

function renderSubtypeBar() {
  const bar = $("#subtypeBar");
  if (!state.status.has("不规范用药")) { bar.classList.add("hidden"); return; }
  const counts = (CURRENT.summary && CURRENT.summary.by_subtype) || {};
  let html = "";
  (CURRENT.summary.subtypes || []).forEach(k => {
    const n = counts[k] || 0;
    if (!n) return;
    const act = state.subtype === k ? "active" : "";
    html += `<span class="chip ${act}" data-sub="${k}">${k} (${n})</span>`;
  });
  $("#subtypeChips").innerHTML = html;
  bar.classList.remove("hidden");
  document.querySelectorAll("#subtypeChips .chip").forEach(c => c.onclick = () => {
    const k = c.dataset.sub; state.subtype = (state.subtype === k) ? null : k; state.page = 1; refresh();
  });
}

function updateFilterInfo() {
  const parts = [];
  if (state.q) parts.push("搜索=" + state.q);
  if (state.status.size) parts.push("用药状态=" + [...state.status].join("/"));
  if (state.subtype) parts.push("下钻=" + state.subtype);
  if (state.drugs.size) parts.push("药品=" + state.drugs.size + "项");
  if (state.pharmacies.size) parts.push("药店=" + state.pharmacies.size + "项");
  if (state.start || state.end) parts.push("时间=" + ((state.start || "…") + "~" + (state.end || "…")));
  $("#filterInfo").textContent = parts.length ? ("筛选：" + parts.join(" · ")) : "";
}

// 文件管理
function renderFiles(files) {
  $("#fileChips").innerHTML = (files || []).map(f =>
    `<span class="file-chip">${esc(f.label)} (${f.count})<span class="fx" data-fid="${f.id}" title="删除该文件">×</span></span>`).join("");
  document.querySelectorAll(".file-chip .fx").forEach(x => x.onclick = () => {
    if (!confirm("确定删除该文件的全部记录？")) return;
    const fid = x.dataset.fid;
    STORE.records = STORE.records.filter(r => r._file_id !== fid);
    STORE.files = STORE.files.filter(f => f.id !== fid);
    CURRENT.global = build_summary(STORE.records);
    renderGlobal(CURRENT.global);
    renderFiles(STORE.files);
    state.page = 1;
    refresh();
  });
}

const visibleCols = () => DETAIL.filter(([key]) => !state.hiddenCols.has(key));

function renderDetail() {
  const cols = visibleCols();
  $("#thead").innerHTML = cols.map(([, label]) => `<th>${label}</th>`).join("");
  const total = DATA.rows.length;
  const pages = Math.max(1, Math.ceil(total / state.pageSize));
  if (state.page > pages) state.page = pages;
  const start = (state.page - 1) * state.pageSize;
  const rows = DATA.rows.slice(start, start + state.pageSize);
  const tb = $("#tbody");
  if (!total) { tb.innerHTML = ""; $("#empty").classList.remove("hidden"); }
  else {
    $("#empty").classList.add("hidden");
    tb.innerHTML = rows.map((rec, idx) => rowHtml(rec, cols, start + idx)).join("");
  }
  tb.querySelectorAll("tr.data-row").forEach(tr => {
    tr.onclick = () => {
      const rec = DATA.rows[+tr.dataset.idx];
      const next = tr.nextElementSibling;
      if (next && next.classList.contains("expand-row")) { next.remove(); return; }
      tb.querySelectorAll("tr.expand-row").forEach(x => x.remove());
      tr.insertAdjacentHTML("afterend", expandHtml(rec, cols.length));
    };
  });
  renderPagination(total);
}
function rowHtml(rec, cols, idx) {
  return `<tr class="data-row" data-idx="${idx}">` + cols.map(([key]) => {
    let v = rec[key] || "";
    if (key === "medication_status_raw") {
      const norm = rec["medication_status"] || "";
      const txt = v || norm;
      return `<td><span class="tag ${norm}">${esc(txt)}</span></td>`;
    }
    if (key === "irregularity_subtype") {
      if (!v) return `<td></td>`;
      return `<td><div class="cell-clip">${esc(v)}</div></td>`;
    }
    if (key === "remarks" || key === "stop_reduce_reason") return `<td><div class="cell-clip">${esc(v)}</div></td>`;
    return `<td>${esc(v)}</td>`;
  }).join("") + "</tr>";
}
function expandHtml(rec, colspan) {
  const sec = (lbl, val) => val ? `<div><span class="eb-lbl">${lbl}：</span>${esc(val)}</div>` : "";
  return `<tr class="expand-row"><td colspan="${colspan}"><div class="expand-box">`
    + sec("不规范类型", rec.irregularity_subtype)
    + sec("停药/减量根本原因", rec.stop_reduce_reason)
    + sec("备注", rec.remarks)
    + sec("随访小结", rec.summary)
    + `</div></td></tr>`;
}

function renderPatients() {
  const total = DATA.patients.length;
  const pages = Math.max(1, Math.ceil(total / state.pageSize));
  if (state.page > pages) state.page = pages;
  const start = (state.page - 1) * state.pageSize;
  const pats = DATA.patients.slice(start, start + state.pageSize);
  const pv = $("#patientView");
  if (!total) { pv.innerHTML = '<div class="empty">无匹配患者</div>'; }
  else {
    pv.innerHTML = pats.map((p, i) => patCardHtml(p, start + i)).join("");
    pv.querySelectorAll(".pat-head").forEach(h => h.onclick = () => {
      const body = h.nextElementSibling;
      body.classList.toggle("hidden");
      h.querySelector(".pat-arrow").textContent = body.classList.contains("hidden") ? "▸" : "▾";
    });
  }
  renderPagination(total);
}
function patCardHtml(p, idx) {
  const tags = Object.entries(p.by_status || {}).map(([s, n]) => `<span class="tag ${s}">${s} ${n}</span>`).join("");
  const recs = (p.records || []).slice().sort((a, b) => {
    const ta = a.followup_time || "", tb = b.followup_time || "";
    return tb.localeCompare(ta);
  });
  const recRows = recs.map(r => {
    const norm = r["medication_status"] || "";
    const sub = r["irregularity_subtype"] || "";
    return `<tr><td>${esc(r.followup_time || "")}</td><td>${esc(r.drug_product || "")}</td><td>${esc(r.pharmacy || "")}</td>`
      + `<td><span class="tag ${norm}">${esc(r.medication_status_raw || norm)}</span></td>`
      + `<td style="white-space:pre-wrap">${esc(sub)}</td><td style="white-space:pre-wrap;max-width:280px">${esc(r.remarks || "")}</td></tr>`;
  }).join("");
  return `<div class="pat-card">
    <div class="pat-head">
      <span class="pat-name">${esc(p.patient_name)}</span>
      <span class="pat-phone">${esc(p.phone || "")}</span>
      <span class="pat-tags">${tags}</span>
      <span class="pat-count">${p.count} 次随访 · 最近 ${esc(p.latest || "—")}</span>
      <span class="pat-arrow">▸</span>
    </div>
    <div class="pat-body hidden">
      <table class="mini-table"><thead><tr><th>随访时间</th><th>药品</th><th>药店</th><th>用药状态</th><th>不规范类型</th><th>备注</th></tr></thead>
      <tbody>${recRows}</tbody></table>
    </div>
  </div>`;
}

function renderPagination(total) {
  const pages = Math.max(1, Math.ceil(total / state.pageSize));
  $("#pagination").innerHTML = `
    <span>共 ${total} 条 · 每页 ${state.pageSize} 条 · 第 ${state.page}/${pages} 页</span>
    <span class="pg-btns">
      <button class="btn ghost sm" id="prevPg" ${state.page <= 1 ? "disabled" : ""}>‹ 上一页</button>
      <button class="btn ghost sm" id="nextPg" ${state.page >= pages ? "disabled" : ""}>下一页 ›</button>
    </span>`;
  $("#prevPg").onclick = () => { if (state.page > 1) { state.page--; renderCurrentView(); window.scrollTo({ top: 0, behavior: "smooth" }); } };
  $("#nextPg").onclick = () => { if (state.page < pages) { state.page++; renderCurrentView(); window.scrollTo({ top: 0, behavior: "smooth" }); } };
}

function renderCurrentView() {
  if (state.view === "detail") { $("#detailView").classList.remove("hidden"); $("#patientView").classList.add("hidden"); renderDetail(); }
  else { $("#detailView").classList.add("hidden"); $("#patientView").classList.remove("hidden"); renderPatients(); }
}

/* ============ 刷新流程（本地计算，无任何网络请求） ============ */
async function refresh() {
  await loadSummary();
  if (state.view === "detail") await fetchDetail(); else await fetchPatients();
  renderCurrentView();
}
async function loadSummary() {
  CURRENT.summary = summaryLocal(currentKw());
  renderSummary(CURRENT.summary);
  renderCharts(CURRENT.summary);
}
async function fetchDetail() {
  const recs = filter_records(STORE.records, currentKw());
  DATA.rows = recs.map(r => window.Pipeline.desensitize(r, state.plainName, state.plainPhone));
}
async function fetchPatients() {
  DATA.patients = patientsAgg(STORE.records, currentKw());
}

/* ============ 上传 / 待分析 / 开始分析 ============ */
function addToPending(files) {
  let added = 0;
  for (const f of files) {
    // 去重：同名同大小视为同一文件
    if (!pendingFiles.some(p => p.name === f.name && p.size === f.size)) {
      pendingFiles.push(f); added++;
    }
  }
  renderPending();
  if (added) $("#pendingArea").classList.remove("hidden");
}
function renderPending() {
  const list = $("#pendingList");
  list.innerHTML = pendingFiles.map((f, i) =>
    `<span class="file-chip">${esc(f.name)} (${fmtSize(f.size)})<span class="fx" data-idx="${i}" title="移除">×</span></span>`).join("");
  list.querySelectorAll(".fx").forEach(x => x.onclick = () => {
    pendingFiles.splice(+x.dataset.idx, 1);
    renderPending();
  });
  $("#pendingCount").textContent = pendingFiles.length ? `共 ${pendingFiles.length} 个文件待分析` : "";
  $("#startBtn").disabled = pendingFiles.length === 0;
  if (pendingFiles.length === 0) $("#pendingArea").classList.add("hidden");
}

$("#pickBtn").onclick = () => fileInput.click();
$("#addMoreBtn").onclick = () => fileInput.click();
fileInput.onchange = e => { if (e.target.files.length) { addToPending(e.target.files); e.target.value = ""; } };
["dragover", "dragenter"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add("drag"); }));
["dragleave", "drop"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove("drag"); }));
drop.addEventListener("drop", e => { const f = e.dataTransfer.files; if (f.length) addToPending(f); });

$("#clearPendingBtn").onclick = () => {
  pendingFiles = [];
  renderPending();
  $("#pendingArea").classList.add("hidden");
};

$("#startBtn").onclick = async () => {
  if (!pendingFiles.length) return;
  showLoading("正在解析 Excel 并归一化计算…");
  $("#startBtn").disabled = true;
  try {
    for (const f of pendingFiles) {
      STORE.seq++;
      const fid = "f" + STORE.seq;
      const recs = await window.Pipeline.processFiles([f]);
      if (recs.length) {
        recs.forEach(r => { r._file_id = fid; r._file_label = f.name; });
        STORE.records.push(...recs);
        STORE.files.push({ id: fid, label: f.name, count: recs.length });
      }
    }
    pendingFiles = [];
    renderPending();
    CURRENT.global = build_summary(STORE.records);
    renderGlobal(CURRENT.global);
    renderFiles(STORE.files);
    $("#board").classList.remove("hidden");
    state.page = 1;
    await refresh();
  } catch (err) {
    alert("分析失败：" + (err && err.message ? err.message : err));
  } finally {
    hideLoading();
    $("#startBtn").disabled = pendingFiles.length === 0;
  }
};

$("#clearAllBtn").onclick = () => {
  if (!confirm("确定清空全部已加载数据？")) return;
  STORE.records = []; STORE.files = []; STORE.seq = 0;
  CURRENT = { summary: null, global: null }; DATA = { rows: [], patients: [] };
  $("#board").classList.add("hidden");
  $("#fileChips").innerHTML = "";
};

/* ============ 筛选交互 ============ */
$("#clearBtn").onclick = () => {
  state.status.clear(); state.subtype = null;
  state.drugs.clear(); state.pharmacies.clear();
  state.start = state.end = null; state.q = ""; state.page = 1;
  $("#searchInput").value = ""; $("#startDate").value = ""; $("#endDate").value = "";
  refresh();
};
$("#startDate").onchange = e => { state.start = e.target.value || null; state.page = 1; refresh(); };
$("#endDate").onchange = e => { state.end = e.target.value || null; state.page = 1; refresh(); };
let searchTimer = null;
$("#searchInput").addEventListener("input", e => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => { state.q = e.target.value.trim(); state.page = 1; refresh(); }, 350);
});
$("#viewDetailBtn").onclick = () => { if (state.view !== "detail") { state.view = "detail"; state.page = 1; syncViewBtns(); refresh(); } };
$("#viewPatientBtn").onclick = () => { if (state.view !== "patients") { state.view = "patients"; state.page = 1; syncViewBtns(); refresh(); } };
function syncViewBtns() {
  $("#viewDetailBtn").classList.toggle("active", state.view === "detail");
  $("#viewPatientBtn").classList.toggle("active", state.view === "patients");
}
document.querySelectorAll(".dt-btn").forEach(b => b.onclick = () => {
  const field = b.dataset.field, mode = b.dataset.mode;
  const plain = (mode === "plain");
  if (field === "name") state.plainName = plain; else state.plainPhone = plain;
  b.parentElement.querySelectorAll(".dt-btn").forEach(x => x.classList.toggle("active", x === b));
  state.page = 1;
  refresh();
});

// 弹出层统一管理：多选面板 + 列显隐，点击外部关闭
function closeAllPopovers() {
  document.querySelectorAll(".ms-panel:not(.hidden)").forEach(p => p.classList.add("hidden"));
  $("#colPanel").classList.add("hidden");
}
function togglePanel(panel) {
  const open = !panel.classList.contains("hidden");
  closeAllPopovers();
  if (!open) panel.classList.remove("hidden");
}
$("#drugMsBtn").onclick = e => { e.stopPropagation(); togglePanel($("#drugMsPanel")); };
$("#pharmMsBtn").onclick = e => { e.stopPropagation(); togglePanel($("#pharmMsPanel")); };
document.addEventListener("click", e => {
  if (e.target.closest(".ms") || e.target.closest("#colPanel") || e.target.closest("#colBtn")) return;
  closeAllPopovers();
});
$("#colBtn").onclick = e => {
  e.stopPropagation();
  const p = $("#colPanel");
  const open = !p.classList.contains("hidden");
  closeAllPopovers();
  if (open) return;
  p.innerHTML = DETAIL.map(([key, label]) =>
    `<label><input type="checkbox" data-col="${key}" ${state.hiddenCols.has(key) ? "" : "checked"}> ${label}</label>`).join("");
  p.querySelectorAll("input").forEach(cb => cb.onchange = () => {
    if (cb.checked) state.hiddenCols.delete(cb.dataset.col); else state.hiddenCols.add(cb.dataset.col);
    renderDetail();
  });
  p.classList.remove("hidden");
};

/* ============ 导出（客户端 SheetJS） ============ */
function doExport(desen) {
  const recs = filter_records(STORE.records, currentKw());
  if (!recs.length) { alert("当前筛选无数据可导出"); return; }
  const rows = desen ? recs.map(r => window.Pipeline.desensitize(r)) : recs;
  const aoa = [EXPORT_COLS.map(([, l]) => l)];
  for (const r of rows) {
    aoa.push(EXPORT_COLS.map(([k]) => {
      let v = r[k];
      if (k === "medication_status_raw") v = r.medication_status_raw || r.medication_status;
      return v == null ? "" : v;
    }));
  }
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "随访明细");
  const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  download(new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    desen ? "随访明细_脱敏.xlsx" : "随访明细_未脱敏.xlsx");
}
$("#exportDesenBtn").onclick = () => doExport(true);
$("#exportPlainBtn").onclick = () => {
  $("#keyInput").value = ""; $("#keyErr").textContent = "";
  $("#keyMask").classList.remove("hidden"); $("#keyInput").focus();
};
$("#keyCancel").onclick = () => $("#keyMask").classList.add("hidden");
$("#keyConfirm").onclick = () => {
  const key = $("#keyInput").value.trim();
  if (!key) { $("#keyErr").textContent = "请输入任意内容以确认"; return; }
  $("#keyMask").classList.add("hidden");
  doExport(false);
};
$("#keyInput").addEventListener("keydown", e => { if (e.key === "Enter") $("#keyConfirm").click(); });
$("#keyMask").addEventListener("click", e => { if (e.target === $("#keyMask")) $("#keyMask").classList.add("hidden"); });

/* ============ 快照生成（自包含 HTML，可离线打开/分享） ============ */
async function doSnapshot(desen) {
  if (!STORE.records.length) { alert("暂无数据可生成快照"); return; }
  showLoading(desen ? "正在生成脱敏快照…" : "正在生成不脱敏快照…");
  try {
    // 脱敏快照：写入文件的是「已脱敏」记录，原始 PII 不存在于文件中 → 无法还原 / 无法按全名检索。
    const records = desen ? STORE.records.map(r => window.Pipeline.desensitize(r)).map(cleanRec)
                           : STORE.records.map(cleanRec);
    const snap = { desen, records, files: STORE.files, buildAt: new Date().toISOString() };
    // 转义 JSON 中的尖括号，防止内联 HTML 时提前闭合 script 标签
    const dataJson = JSON.stringify(snap).replace(/</g, "\\u003c");
    const dataScript = `<\script>window.__SNAP__=${dataJson};<\/script>`;
    const bootstrap = `<\script>(function(){
      if(window.__SNAP__&&window.AppCore){window.AppCore.loadSnapshot(window.__SNAP__);}
      else{var ld=document.getElementById('loading');if(ld){ld.querySelector('.txt').textContent='快照加载失败：脚本未内联。请用网页版（index.html）生成快照，不要用 index.template.html。';ld.classList.remove('hidden');}}
    })();<\/script>`;
    // 直接序列化当前页面：逻辑脚本已内联在页面中，无需 fetch，
    // 因此 https 与 file://（双击打开）都能生成可离线打开的自包含快照。
    let html = "<!DOCTYPE html>\n" + document.documentElement.outerHTML;
    // 快照为只读视图，不需要 xlsx 库；剥离外链，避免 file:// 下加载失败
    html = html.replace('<' + 'script src="vendor/xlsx.full.min.js"></sc' + 'ript>', "");
    // 注意：本文件源码中已含有字符串 "</body>"，若用 html.replace("</body>", ...) 会命中源码里的那个，
    // 把数据脚本塞进 app.js 源码字符串、而非文档真正的 </body> 前。必须用 lastIndexOf 定位文档末尾的真实 </body>。
    const bodyIdx = html.lastIndexOf("</body>");
    html = html.slice(0, bodyIdx) + dataScript + bootstrap + "\n" + html.slice(bodyIdx);
    download(new Blob([html], { type: "text/html" }),
      desen ? "随访看板_脱敏快照.html" : "随访看板_不脱敏快照.html");
  } catch (err) {
    alert("快照生成失败：" + (err && err.message ? err.message : err));
  } finally {
    hideLoading();
  }
}
$("#snapshotDesenBtn").onclick = () => doSnapshot(true);
$("#snapshotPlainBtn").onclick = () => {
  // 未脱敏快照含明文个人信息，生成前必须明确提示隐私风险，禁止外部转发
  const ok = confirm(
    "⚠️ 隐私风险提示\n\n" +
    "「未脱敏快照」包含明文姓名、电话等个人敏感信息。\n" +
    "仅可分享给可信的团队成员，切勿通过外部渠道（邮件 / 微信 / 公网 / 不可信接收方）转发。\n\n" +
    "如仅需对外分享，请改用「脱敏快照」。\n\n" +
    "确定要生成未脱敏快照吗？"
  );
  if (ok) doSnapshot(false);
};

/* ============ 快照打开模式（仅供生成的快照内部调用） ============ */
function loadSnapshot(snap) {
  hideLoading(); // 快照可能序列化时拍入了可见的「正在生成」遮罩，打开后必须先关掉
  SNAP_MODE = true;
  STORE.records = (snap.records || []).map(r => Object.assign({}, r));
  STORE.files = snap.files || [];
  state.plainName = !snap.desen;
  state.plainPhone = !snap.desen;
  $("#drop").classList.add("hidden");
  $("#pendingArea").classList.add("hidden");
  $("#snapBanner").classList.remove("hidden");
  $("#snapBanner").textContent = snap.desen
    ? "📄 这是一份脱敏快照：姓名、电话已脱敏，文件中不含任何明文个人信息，可安全分享。"
    : "⚠️ 这是一份「不脱敏」快照：包含明文姓名、电话等个人信息，仅可分享给可信接收方。";
  // 快照为只读分享件：隐藏上传、文件管理、导出、再次快照等按钮
  const fb = document.querySelector(".filebar"); if (fb) fb.classList.add("hidden");
  ["#exportDesenBtn", "#exportPlainBtn", "#snapshotDesenBtn", "#snapshotPlainBtn"].forEach(s => $(s).classList.add("hidden"));
  if (snap.desen) {
    document.querySelectorAll('.dt-btn[data-mode="plain"]').forEach(b => b.disabled = true);
  }
  CURRENT.global = build_summary(STORE.records);
  renderGlobal(CURRENT.global);
  $("#board").classList.remove("hidden");
  state.page = 1;
  refresh();
}

// 调试/测试用：暴露核心计算与快照接口
window.AppCore = { loadSnapshot, build_summary, filter_records, patientsAgg, summaryLocal, STORE };
})();

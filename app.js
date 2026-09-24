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
  ["executor", "执行人"],
  ["medication_status_raw", "用药状态"],
  ["irregularity_subtype", "不规范类型"],
  ["stop_reduce_reason", "停药/减量根本原因"], ["reason_bucket", "根本原因分类"],
  ["dosage_raw", "用法用量"],
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
  ["pharmacy", "药店"], ["executor", "执行人"], ["medication_status_raw", "用药状态"],
  ["irregularity_subtype", "不规范类型"],
  ["stop_reduce_reason", "停药/减量根本原因"], ["reason_bucket", "根本原因分类"],
  ["dosage_raw", "用法用量"], ["remarks", "备注"]];

const state = { status: new Set(), subtype: null, drugs: new Set(), pharmacies: new Set(), executors: new Set(),
  reasons: new Set(), start: null, end: null,
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
  delete c._tutorial;   // 教程注入的示例数据标记，不应写进快照
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

/* ============ 「停药/减量根本原因」统计口径 ============ */
// 口径（用户确认）：只统计「任务状态＝已完成」的记录；并把误入原因列的正常用药描述
//（reason_bucket = 非停减…）从分子分母中剔除，避免污染占比。
const REASON_SKIP = "非停减（仍在用药／已购药）";
function isDoneTask(r) { return /已完成/.test(String(r.task_status || "")); }
function isReasonCounted(r) {
  return isDoneTask(r) && r.reason_bucket && r.reason_bucket !== REASON_SKIP;
}
// 返回 { by_reason, meta }：meta 用于在卡片脚注里交代口径，便于核对
function reasonStats(records) {
  const done = records.filter(isDoneTask);
  const withReason = done.filter(r => r.reason_bucket);
  const counted = done.filter(r => r.reason_bucket && r.reason_bucket !== REASON_SKIP);
  return {
    by_reason: countBy(counted, r => r.reason_bucket),
    meta: {
      all: records.length,
      done: done.length,
      withReason: withReason.length,
      skipped: withReason.filter(r => r.reason_bucket === REASON_SKIP).length,
      counted: counted.length,
    },
  };
}
function build_summary(records) {
  const months = records.map(r => _month(r.followup_time)).filter(Boolean).sort();
  const by_month = {};
  for (const m of months) by_month[m] = (by_month[m] || 0) + 1;
  const days = records.map(r => _extract_date(r.followup_time)).filter(Boolean).sort();
  const by_day = {};
  for (const d of days) by_day[d] = (by_day[d] || 0) + 1;
  const rs = reasonStats(records);
  return {
    total: records.length,
    by_status: countBy(records, r => r.medication_status),
    by_subtype: countBy(records.filter(r => r.irregularity_subtype), r => r.irregularity_subtype),
    by_source: countBy(records, r => r.source_type),
    by_task_status: countBy(records, r => r.task_status || "未知"),
    by_drug: countBy(records, r => r.drug_product || "未知"),
    by_pharmacy: countBy(records, r => r.pharmacy || "未知"),
    by_executor: countBy(records, r => r.executor || "未知"),
    by_reason: rs.by_reason,
    reason_meta: rs.meta,
    by_month,
    by_day,
    taxonomy: STATUS_TAXONOMY,
    subtypes: SUBTYPE_TAXONOMY,
  };
}
// kw: {status:[]|null, subtype, src, drug:[]|null, pharmacy:[]|null, reason:[]|null, start, end, q}
function filter_records(records, kw) {
  let recs = records;
  const { status, subtype, src, drug, pharmacy, executor, reason, start, end, q } = kw;
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
  if (executor && executor.length) {
    const eset = new Set(executor);
    recs = recs.filter(r => eset.has(r.executor || "未知"));
  }
  // 根本原因：口径同为「已完成任务」，与图表/脚注保持一致
  if (reason && reason.length) {
    const rset = new Set(reason);
    recs = recs.filter(r => isDoneTask(r) && rset.has(r.reason_bucket || ""));
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
  else if (exclude === "executor") k.executor = null;
  else if (exclude === "reason") k.reason = null;
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
    executor: state.executors.size ? [...state.executors] : null,
    reason: state.reasons.size ? [...state.reasons] : null,
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
  base.by_executor = build_summary(_facet(STORE.records, kw, "executor")).by_executor;
  // 根本原因：排除自身筛选，使各分桶计数互相可见（与其它维度一致）
  const rf = build_summary(_facet(STORE.records, kw, "reason"));
  base.by_reason = rf.by_reason;
  base.reason_meta = rf.reason_meta;
  base.by_month = build_summary(_facet(STORE.records, kw, "time")).by_month;
  base.by_day = build_summary(_facet(STORE.records, kw, "time")).by_day;
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
  buildMs("execMs", d.by_executor, state.executors, "执行人");
  buildMs("reasonMs", d.by_reason, state.reasons, "根本原因");
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
  renderDrugChart(d);
  // 随访时间趋势：时间跨度 ≤ 31 天（约一个月）时按「天」画折线，否则按「月」
  const bd = d.by_day || {};
  const bm = d.by_month || {};
  const dates = Object.keys(bd).sort();
  let useDay = dates.length > 0;
  if (dates.length >= 2) {
    const span = +new Date(dates[dates.length - 1]) - +new Date(dates[0]);
    useDay = span <= 31 * 24 * 3600 * 1000;
  }
  $("#trendChart").innerHTML = lineChart(useDay ? bd : bm, useDay ? "day" : "month");
  bindTrendHover();
  renderReasonChart(d);
}

// 趋势图悬停：鼠标移入整张图时，自动吸附到最近的月份点并显示数量
function bindTrendHover() {
  const svg = document.querySelector("#trendChart svg");
  if (!svg) return;
  const raw = svg.getAttribute("data-series");
  if (!raw) return;
  let series;
  try { series = JSON.parse(raw); } catch (e) { return; }
  if (!series.length) return;
  const W = 360, H = 140, P = 28;
  const overlay = svg.querySelector(".trend-overlay");
  const hover = svg.querySelector(".trend-hover");
  if (!overlay || !hover) return;
  const line = hover.querySelector(".th-line");
  const dot = hover.querySelector(".th-dot");
  const bg = hover.querySelector(".th-bg");
  const txt = hover.querySelector(".th-txt");

  overlay.addEventListener("mousemove", e => {
    const rect = svg.getBoundingClientRect();
    if (!rect.width) return;
    const scaleX = W / rect.width;          // 客户端 px -> viewBox 坐标
    const mx = (e.clientX - rect.left) * scaleX;
    let best = 0, bd = Infinity;
    series.forEach((p, i) => { const dd = Math.abs(p.x - mx); if (dd < bd) { bd = dd; best = i; } });
    const p = series[best];
    hover.style.display = "";
    line.setAttribute("x1", p.x);
    line.setAttribute("x2", p.x);
    dot.setAttribute("cx", p.x);
    dot.setAttribute("cy", p.y);
    const ty = Math.max(8, Math.min(p.y - 11, H - P));
    const bx = Math.max(32, Math.min(p.x, W - 32));
    bg.setAttribute("x", bx - 32);
    bg.setAttribute("y", ty - 14);
    txt.setAttribute("x", bx);
    txt.setAttribute("y", ty - 2);
    txt.textContent = `${p.k}：${p.v} 条`;
  });
  overlay.addEventListener("mouseleave", () => { hover.style.display = "none"; });
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
// 药品记录数：呈现**全部**药品（不止 Top5），纵向滚动；点击条形即按该药品筛选（再点取消）。
// 卡片尺寸保持不变——滚动区内高由 CSS #drugChart{max-height} 固定为原 Top5 的高度。
function barChart(data) {
  const entries = Object.entries(data || {}).sort((a, b) => b[1] - a[1]);
  if (!entries.length) return '<div class="chart-empty">暂无数据</div>';
  const max = Math.max(...entries.map(e => e[1]));
  const rows = entries.map(([k, v]) => {
    const w = Math.max(2, Math.round(v / max * 100));
    const lbl = k.length > 8 ? k.slice(0, 8) + "…" : k;
    const active = state.drugs.has(k) ? " active" : "";
    return `<div class="bar-row clickable${active}" data-drug="${esc(k)}" title="${esc(k)}：${v} 条 · 点击筛选，再次点击取消">
      <span class="bar-label">${esc(lbl)}</span>
      <span class="bar-track"><span class="bar-fill" style="width:${w}%"></span></span>
      <span class="bar-val">${v}</span></div>`;
  }).join("");
  return `<div class="bars bars-scroll">${rows}</div>`;
}
function renderDrugChart(d) {
  const el = $("#drugChart");
  if (!el) return;
  // 重渲染后保持滚动位置：否则点选列表深处的药品会把列表弹回顶部
  const keepTop = el.scrollTop;
  el.innerHTML = barChart(d.by_drug || {});
  el.scrollTop = keepTop;
  el.querySelectorAll(".bar-row.clickable").forEach(row => {
    row.onclick = () => {
      const k = row.dataset.drug;
      if (state.drugs.has(k)) state.drugs.delete(k); else state.drugs.add(k);
      state.page = 1;
      refresh();
    };
  });
}

// 「停药/减量根本原因」分布：横向条形图，点击条形即加入/取消筛选（下钻）
// 排版：两列（按名次自上而下、先填左列）；标签用短名（全名放 title 悬停提示）；
// 桶很多时由 #reasonChart 的 max-height 出滚动条。
function reasonBarChart(byReason) {
  const entries = Object.entries(byReason || {}).sort((a, b) => b[1] - a[1]);
  if (!entries.length) return '<div class="chart-empty">当前筛选下暂无「已完成任务」的根本原因数据</div>';
  const SHORT = M.REASON_BUCKET_SHORT || {};
  const max = Math.max(...entries.map(e => e[1]));
  const total = entries.reduce((a, [, v]) => a + v, 0);
  const rows = Math.ceil(entries.length / 2);
  const body = entries.map(([k, v]) => {
    const w = Math.max(2, Math.round(v / max * 100));
    const pct = (v / total * 100).toFixed(1);
    const active = state.reasons.has(k) ? " active" : "";
    const label = SHORT[k] || k;
    return `<div class="bar-row clickable${active}" data-reason="${esc(k)}" title="${esc(k)}：${v} 条（${pct}%）· 点击筛选，再次点击取消">
      <span class="bar-label">${esc(label)}</span>
      <span class="bar-track"><span class="bar-fill" style="width:${w}%"></span></span>
      <span class="bar-val">${v}</span>
      <span class="bar-pct">${pct}%</span></div>`;
  }).join("");
  return `<div class="bars bars-lg" style="grid-template-rows:repeat(${rows},auto)">${body}</div>`;
}
function renderReasonChart(d) {
  const el = $("#reasonChart");
  if (!el) return;
  el.innerHTML = reasonBarChart(d.by_reason);
  const m = d.reason_meta || {};
  const sub = $("#reasonSub");
  if (sub) {
    sub.textContent = m.counted
      ? `已完成 ${m.done} 条 · 有原因 ${m.counted} 条 · 点击条形下钻`
      : "基于当前筛选 · 暂无「已完成任务」的停减量原因";
  }
  const foot = $("#reasonFoot");
  if (foot) {
    const bits = [`记录数口径`, `筛选 ${m.all || 0}`, `已完成 ${m.done || 0}`, `有原因 ${m.withReason || 0}`];
    if (m.skipped) bits.push(`剔除非停减 ${m.skipped}`);
    const pending = (m.withReason || 0) - (m.counted || 0) - (m.skipped || 0);
    if (pending > 0) bits.push(`未计入未完成 ${pending}`);
    foot.textContent = bits.join(" · ");
    foot.title = "仅统计任务状态为「已完成」的记录；「非停减（仍在用药／已购药）」类描述已剔除，不计入分子分母。";
  }
  el.querySelectorAll(".bar-row.clickable").forEach(row => {
    row.onclick = () => {
      const k = row.dataset.reason;
      if (state.reasons.has(k)) state.reasons.delete(k); else state.reasons.add(k);
      state.page = 1;
      refresh();
    };
  });
}
function lineChart(data, unit) {
  const entries = Object.entries(data);
  if (!entries.length) return '<div class="chart-empty">暂无数据</div>';
  const W = 360, H = 140, P = 28, PT = 12;
  const max = Math.max(...entries.map(e => e[1]), 1);
  const n = entries.length;
  const X = i => n === 1 ? W / 2 : P + i * (W - 2 * P) / (n - 1);
  const Y = v => H - P - (v / max) * (H - 2 * P - PT);
  const pts = entries.map(([, v], i) => `${X(i).toFixed(1)},${Y(v).toFixed(1)}`).join(" ");
  const dots = entries.map(([k, v], i) =>
    `<circle cx="${X(i).toFixed(1)}" cy="${Y(v).toFixed(1)}" r="3" fill="#3b5bdb"></circle>`).join("");
  // 横坐标标签：按月显示「YY-MM」，按天显示「MM-DD」
  const fmt = k => unit === "day" ? k.slice(5) : k.slice(2);
  const step = Math.ceil(n / 8);
  const labels = entries.map(([k], i) =>
    (i % step === 0 || i === n - 1) ? `<text x="${X(i).toFixed(1)}" y="${H - 8}" font-size="9" text-anchor="middle" fill="#868e96">${esc(fmt(k))}</text>` : "").join("");
  // 基线
  const baseline = `<line x1="${P}" y1="${H - P}" x2="${W - P}" y2="${H - P}" stroke="#e9ecef" stroke-width="1"/>`;
  // 预计算每个点的 viewBox 坐标，供悬停吸附使用
  const series = entries.map(([k, v], i) => ({ x: +X(i).toFixed(1), y: +Y(v).toFixed(1), k, v }));
  const hover = `<g class="trend-hover" style="display:none">
      <line class="th-line" x1="0" y1="${PT}" x2="0" y2="${H - P}" stroke="#adb5bd" stroke-dasharray="3 2" stroke-width="1"/>
      <circle class="th-dot" r="4.5" fill="#fff" stroke="#3b5bdb" stroke-width="2"/>
      <rect class="th-bg" x="0" y="0" width="64" height="16" rx="3" fill="#212529" opacity="0.92"/>
      <text class="th-txt" text-anchor="middle" font-size="10" fill="#fff"></text>
    </g>`;
  const overlay = `<rect class="trend-overlay" x="0" y="0" width="${W}" height="${H}" fill="transparent" style="cursor:crosshair"/>`;
  return `<svg viewBox="0 0 ${W} ${H}" width="100%" height="140" data-series='${JSON.stringify(series)}' style="overflow:visible">`
    + baseline + `<polyline points="${pts}" fill="none" stroke="#3b5bdb" stroke-width="2"/>` + dots + labels
    + hover + overlay + `</svg>`;
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
  if (state.executors.size) parts.push("执行人=" + state.executors.size + "项");
  if (state.reasons.size) parts.push("根本原因=" + state.reasons.size + "项");
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
  const sec = (lbl, val) => val ? `<div class="eb-row"><span class="eb-lbl">${lbl}：</span><span class="eb-val">${esc(val)}</span></div>` : "";
  // 推导字段（与主表互补，不在 DETAIL 重复列）
  const derived = sec("不规范类型", rec.irregularity_subtype) + sec("停药/减量根本原因", rec.stop_reduce_reason);
  // 专项原文：按随访项目展示各自的结构化原列（真实列名 + 原值）
  const pf = (rec.project_fields || []).filter(f => f && f.value);
  const rawHtml = pf.map(f => `<div class="eb-row"><span class="eb-lbl">${esc(f.label)}：</span><span class="eb-val">${esc(f.value)}</span></div>`).join("");
  const srcLabel = SRC_LABEL[rec.source_type] || rec.source_type || "其它";
  return `<tr class="expand-row"><td colspan="${colspan}"><div class="expand-box">`
    + (derived ? `<div class="eb-sec">${derived}</div>` : "")
    + (rawHtml ? `<div class="eb-sec eb-sec-proj"><div class="eb-sec-title">${esc(srcLabel)} · 原始字段</div>${rawHtml}</div>` : "")
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
  // 适应症：取该患者最末一次随访记录；若无则逐次往前（记录已按时间降序）
  let latestInd = "";
  for (const r of recs) { if (r.indication) { latestInd = r.indication; break; } }
  const recRows = recs.map(r => {
    const norm = r["medication_status"] || "";
    const sub = r["irregularity_subtype"] || "";
    return `<tr><td>${esc(r.followup_time || "")}</td><td>${esc(r.drug_product || "")}</td><td>${esc(r.pharmacy || "")}</td>`
      + `<td><span class="tag ${norm}">${esc(r.medication_status_raw || norm)}</span></td>`
      + `<td style="white-space:pre-wrap">${esc(sub)}</td>`
      + `<td style="white-space:pre-wrap;max-width:240px">${esc(r.stop_reduce_reason || "")}</td>`
      + `<td style="white-space:pre-wrap;max-width:240px">${esc(r.remarks || "")}</td></tr>`;
  }).join("");
  return `<div class="pat-card">
    <div class="pat-head">
      <span class="pat-name">${esc(p.patient_name)}</span>
      ${latestInd ? `<span class="pat-ind">${esc(latestInd)}</span>` : ""}
      <span class="pat-phone">${esc(p.phone || "")}</span>
      <span class="pat-tags">${tags}</span>
      <span class="pat-count">${p.count} 次随访 · 最近 ${esc(p.latest || "—")}</span>
      <span class="pat-arrow">▸</span>
    </div>
    <div class="pat-body hidden">
      <table class="mini-table"><thead><tr><th>随访时间</th><th>药品</th><th>药店</th><th>用药状态</th><th>不规范类型</th><th>停药/减量根本原因</th><th>备注</th></tr></thead>
      <tbody>${recRows}</tbody></table>
    </div>
  </div>`;
}

function renderPagination(total) {
  const pages = Math.max(1, Math.ceil(total / state.pageSize));
  const opts = [50, 100, 200].map(n => `<option value="${n}" ${state.pageSize === n ? "selected" : ""}>${n} 条/页</option>`).join("");
  $("#pagination").innerHTML = `
    <span>共 ${total} 条 · 第 ${state.page}/${pages} 页 · 每页
      <select id="pageSizeSel" class="pg-size" title="每页显示条数">${opts}</select>
    </span>
    <span class="pg-btns">
      <button class="btn ghost sm" id="prevPg" ${state.page <= 1 ? "disabled" : ""}>‹ 上一页</button>
      <button class="btn ghost sm" id="nextPg" ${state.page >= pages ? "disabled" : ""}>下一页 ›</button>
    </span>`;
  // 翻页后只把表格容器（明细表）滚回顶部，页面滚动条保持原位
  const scrollTableTop = () => { const w = $("#detailView"); if (w) w.scrollTop = 0; };
  $("#prevPg").onclick = () => { if (state.page > 1) { state.page--; renderCurrentView(); scrollTableTop(); } };
  $("#nextPg").onclick = () => { if (state.page < pages) { state.page++; renderCurrentView(); scrollTableTop(); } };
  $("#pageSizeSel").onchange = e => {
    const n = parseInt(e.target.value, 10);
    if (n && n !== state.pageSize) { state.pageSize = n; state.page = 1; renderCurrentView(); }
  };
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

// 解析并并入一个或多个文件（上传按钮与教程共用同一条路径）
// opts.tutorial = true 时给记录打标，教程结束时据此清理示例数据
async function ingestFiles(files, opts) {
  const tutorial = !!(opts && opts.tutorial);
  let added = 0;
  for (const f of files) {
    STORE.seq++;
    const fid = "f" + STORE.seq;
    const recs = await window.Pipeline.processFiles([f]);
    if (recs.length) {
      recs.forEach(r => {
        r._file_id = fid; r._file_label = f.name;
        if (tutorial) r._tutorial = true;
      });
      STORE.records.push(...recs);
      STORE.files.push({ id: fid, label: f.name, count: recs.length, _tutorial: tutorial || undefined });
      added += recs.length;
    }
  }
  CURRENT.global = build_summary(STORE.records);
  renderGlobal(CURRENT.global);
  renderFiles(STORE.files);
  $("#board").classList.remove("hidden");
  state.page = 1;
  await refresh();
  return added;
}

$("#startBtn").onclick = async () => {
  if (!pendingFiles.length) return;
  showLoading("正在解析 Excel 并归一化计算…");
  $("#startBtn").disabled = true;
  try {
    await ingestFiles(pendingFiles);
    pendingFiles = [];
    renderPending();
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
  state.drugs.clear(); state.pharmacies.clear(); state.executors.clear(); state.reasons.clear();
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
$("#execMsBtn").onclick = e => { e.stopPropagation(); togglePanel($("#execMsPanel")); };
$("#reasonMsBtn").onclick = e => { e.stopPropagation(); togglePanel($("#reasonMsPanel")); };
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

/* ============ 导出（客户端，ExcelJS 带样式；SheetJS 兜底） ============ */
// 用药状态着色（Excel 惯例浅底深字；argb 需 8 位）
const EXPORT_STATUS_STYLE = {
  "规范用药": ["FFC6EFCE", "FF006100"],   // 绿底 深绿字
  "不规范用药": ["FFFFC7CE", "FF9C0006"], // 红底 深红字
  "脱落停药": ["FFFFE2C7", "FFC55A11"],  // 橙底 深橙字
  "其他": ["FFF2F2F2", "FF595959"],      // 灰底 深灰字
};
function exportRowVals(r) {
  return EXPORT_COLS.map(([k]) => {
    let v = r[k];
    if (k === "medication_status_raw") v = r.medication_status_raw || r.medication_status;
    return v == null ? "" : String(v);
  });
}
async function doExport(desen) {
  const recs = filter_records(STORE.records, currentKw());
  if (!recs.length) { alert("当前筛选无数据可导出"); return; }
  const rows = desen ? recs.map(r => window.Pipeline.desensitize(r)) : recs;
  const fname = desen ? "随访明细_脱敏.xlsx" : "随访明细_未脱敏.xlsx";
  const mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  // 首选 ExcelJS：表头加粗+底色+冻结、用药状态按四态着色
  if (window.ExcelJS) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("随访明细");
    ws.addRow(EXPORT_COLS.map(([, l]) => l));
    const hrow = ws.getRow(1);
    hrow.height = 20;
    hrow.eachCell(c => {
      c.font = { bold: true, color: { argb: "FF1F2937" } };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9E1F2" } };
      c.alignment = { horizontal: "center", vertical: "middle" };
      c.border = { bottom: { style: "thin", color: { argb: "FFB0BEC5" } } };
    });
    ws.views = [{ state: "frozen", ySplit: 1 }];
    const stateColIdx = EXPORT_COLS.findIndex(([k]) => k === "medication_status_raw");
    for (const r of rows) {
      ws.addRow(exportRowVals(r));
      const ridx = ws.rowCount;
      if (stateColIdx >= 0) {
        const st = r.medication_status || "其他";
        const [fg, fc] = EXPORT_STATUS_STYLE[st] || EXPORT_STATUS_STYLE["其他"];
        const cell = ws.getCell(ridx, stateColIdx + 1);
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fg } };
        cell.font = { bold: true, color: { argb: fc } };
        cell.alignment = { horizontal: "center" };
      }
    }
    // 列宽：文本列放宽，其余按表头
    const wideCols = new Set(["remarks", "stop_reduce_reason", "summary", "dosage_raw", "medication_status_raw"]);
    EXPORT_COLS.forEach(([k, l], i) => {
      const w = wideCols.has(k) ? 46 : Math.max(8, Math.min(20, l.length + 6));
      ws.getColumn(i + 1).width = w;
    });
    const buf = await wb.xlsx.writeBuffer();
    download(new Blob([buf], { type: mime }), fname);
    return;
  }
  // 兜底 SheetJS（无样式，仅当 exceljs 未加载时）
  const aoa = [EXPORT_COLS.map(([, l]) => l)];
  for (const r of rows) aoa.push(exportRowVals(r));
  const ws2 = XLSX.utils.aoa_to_sheet(aoa);
  ws2["!cols"] = EXPORT_COLS.map(([, l], i) => ({ wch: Math.max(8, l.length + 6) }));
  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws2, "随访明细");
  const out = XLSX.write(wb2, { bookType: "xlsx", type: "array" });
  download(new Blob([out], { type: mime }), fname);
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

/* ============ 快照数据范围（2026-09-24：快照 = 当前筛选结果，不再是全量） ============ */
// 文件名标签：只取「药品」「时间」这两个最能一眼认出的维度，避免文件名过长；
// 若用户只用其它维度筛选（状态/药店/执行人等），退回通用标签「按筛选」。
function scopeFileTag(kw) {
  const t = [];
  if (kw.drug && kw.drug.length) {
    t.push(kw.drug.length <= 2 ? kw.drug.join("+") : (kw.drug[0] + "等" + kw.drug.length + "个品种"));
  }
  if (kw.start || kw.end) {
    t.push((kw.start || "最早").replace(/-/g, "") + "-" + (kw.end || "最新").replace(/-/g, ""));
  }
  return t.length ? t : null;
}
// 范围明细：写进快照内部的「数据范围」说明行，顺序与筛选栏一致
function scopeLines(kw) {
  const out = [];
  if (kw.status && kw.status.length) out.push("用药状态 = " + kw.status.join("、"));
  if (kw.subtype) out.push("不规范下钻 = " + kw.subtype);
  if (kw.drug && kw.drug.length) out.push("药品 = " + kw.drug.join("、"));
  if (kw.pharmacy && kw.pharmacy.length) out.push("药店 = " + kw.pharmacy.join("、"));
  if (kw.executor && kw.executor.length) out.push("执行人 = " + kw.executor.join("、"));
  if (kw.reason && kw.reason.length) out.push("根本原因 = " + kw.reason.join("、"));
  if (kw.start || kw.end) out.push("随访时间 = " + (kw.start || "不限") + " ~ " + (kw.end || "不限"));
  if (kw.q) out.push("关键词 = " + kw.q);
  return out;
}
const RE_FNAME_BAD = /[\\/:*?"<>|]/g;
const sanitizeFname = s => String(s).replace(RE_FNAME_BAD, "_");

/* ============ 快照生成（自包含 HTML，可离线打开/分享） ============ */
async function doSnapshot(desen) {
  if (!STORE.records.length) { alert("暂无数据可生成快照"); return; }
  // 快照靠序列化当前页面实现：教程的遮罩/说明卡也会被一起序列化进去，必须先退出教程
  if (TUT.on) await tutEnd("close");
  // 快照的数据范围 = 当前筛选结果（与「导出明细」同一套 filter_records，口径完全一致）。
  // 想「只发 A 品种给他人」，就在筛选栏选好 A 品种再点下载。
  const kw = currentKw();
  const recs = filter_records(STORE.records, kw);
  if (!recs.length) {
    alert("当前筛选条件下没有记录，无法生成快照。\n请调整筛选条件，或点「清除筛选」后重试。");
    return;
  }
  const total = STORE.records.length;
  const lines = scopeLines(kw);
  const scope = { filtered: lines.length > 0, lines, count: recs.length, total };
  const tag = scope.filtered ? (scopeFileTag(kw) || ["按筛选"]) : null;
  showLoading(desen ? "正在生成脱敏快照…" : "正在生成不脱敏快照…");
  try {
    // 脱敏快照：写入文件的是「已脱敏」记录，原始 PII 不存在于文件中 → 无法还原 / 无法按全名检索。
    const records = (desen ? recs.map(r => window.Pipeline.desensitize(r)) : recs).map(cleanRec);
    // 文件清单同步收窄：只保留真正贡献了这批记录的文件；
    // 只取 id/label/count 三个字段，避免把内部标记（如 _tutorial）写进快照
    const keepIds = new Set(recs.map(r => r._file_id));
    const snap = {
      desen, records, scope, buildAt: new Date().toISOString(),
      files: (STORE.files || []).filter(f => keepIds.has(f.id))
        .map(f => ({ id: f.id, label: f.label, count: f.count })),
    };
    // 转义 JSON 中的尖括号，防止内联 HTML 时提前闭合 script 标签
    const dataJson = JSON.stringify(snap).replace(/</g, "\\u003c");
    const dataScript = `<\script>window.__SNAP__=${dataJson};<\/script>`;
    const bootstrap = `<\script>(function(){
      if(window.__SNAP__&&window.AppCore){window.AppCore.loadSnapshot(window.__SNAP__);}
      else{var ld=document.getElementById('loading');if(ld){ld.querySelector('.txt').textContent='快照加载失败：脚本未内联。请用网页版（index.html）生成快照，不要用 index.template.html。';ld.classList.remove('hidden');}}
    })();<\/script>`;
    // 快照里不再提供「姓名 / 电话 是否脱敏」的切换按钮：
    // 脱敏与否在生成那一刻已经固化进数据，快照期内再切换既无意义（脱敏快照的原始姓名根本没写进文件，
    // 切「不脱敏」也只能看到掩码），又容易让接收方误解数据口径。
    // 这里在序列化前把该控件从 DOM 临时摘除，使它既不进入快照 HTML、也不留下可被取消隐藏的死按钮；
    // 序列化完成后立刻原位还原，不影响当前页面。
    const desenTog = document.querySelector(".desen-tog");
    let desenAnchor = null;
    if (desenTog && desenTog.parentNode) {
      desenAnchor = document.createComment("desen-tog-removed-in-snapshot");
      desenTog.parentNode.replaceChild(desenAnchor, desenTog);
    }
    // 教程用过的浮层（遮罩/说明卡/提示条）虽然已隐藏，也不应进入快照：临时摘出，序列化后放回
    const tutNodes = [TUT.root, TUT.card, document.querySelector(".tut-toast")].filter(Boolean);
    tutNodes.forEach(n => { if (n.parentNode) n.parentNode.removeChild(n); });
    // 直接序列化当前页面：逻辑脚本已内联在页面中，无需 fetch，
    // 因此 https 与 file://（双击打开）都能生成可离线打开的自包含快照。
    let html = "<!DOCTYPE html>\n" + document.documentElement.outerHTML;
    tutNodes.forEach(n => document.body.appendChild(n));
    if (desenTog && desenAnchor && desenAnchor.parentNode) {
      desenAnchor.parentNode.replaceChild(desenTog, desenAnchor);
    }
    // 快照为只读视图，不需要 xlsx / exceljs 库；剥离外链，避免 file:// 下加载失败
    html = html.replace('<' + 'script src="vendor/xlsx.full.min.js"></sc' + 'ript>', "");
    html = html.replace('<' + 'script src="vendor/exceljs.min.js"></sc' + 'ript>', "");
    // 注意：本文件源码中已含有字符串 "</body>"，若用 html.replace("</body>", ...) 会命中源码里的那个，
    // 把数据脚本塞进 app.js 源码字符串、而非文档真正的 </body> 前。必须用 lastIndexOf 定位文档末尾的真实 </body>。
    const bodyIdx = html.lastIndexOf("</body>");
    html = html.slice(0, bodyIdx) + dataScript + bootstrap + "\n" + html.slice(bodyIdx);
    download(new Blob([html], { type: "text/html" }),
      "随访看板_" + (desen ? "脱敏" : "不脱敏") + "快照"
      + (tag ? "_" + tag.map(sanitizeFname).join("_") : "") + ".html");
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
  // 快照可能只是「当前筛选结果」的一个子集，必须把范围写清楚，避免接收方误以为拿到全量
  const scopeHtml = (snap.scope && snap.scope.filtered)
    ? `<div class="snap-scope">📊 数据范围（非全量）：${esc(snap.scope.lines.join("；"))}；共 ${snap.scope.count} 条（全量 ${snap.scope.total} 条）</div>`
    : "";
  $("#snapBanner").innerHTML = (snap.desen
    ? "📄 这是一份脱敏快照：姓名、电话已脱敏，文件中不含任何明文个人信息，可安全分享。"
    : "⚠️ 这是一份「不脱敏」快照：包含明文姓名、电话等个人信息，仅可分享给可信接收方。") + scopeHtml;
  // 快照为只读分享件：隐藏上传、文件管理、导出、再次快照等按钮
  const fb = document.querySelector(".filebar"); if (fb) fb.classList.add("hidden");
  ["#exportDesenBtn", "#exportPlainBtn", "#snapshotDesenBtn", "#snapshotPlainBtn", "#tutBtn"]
    .forEach(s => { const el = $(s); if (el) el.classList.add("hidden"); });
  // 快照不提供「是否脱敏 / 脱敏方式」切换：脱敏口径已在生成时固化，
  // 新快照生成阶段就摘除了该控件，这里再兜底移除一次，保证旧快照打开后同样看不到这些按钮。
  document.querySelectorAll(".desen-tog").forEach(el => el.remove());
  CURRENT.global = build_summary(STORE.records);
  renderGlobal(CURRENT.global);
  $("#board").classList.remove("hidden");
  state.page = 1;
  refresh();
}

/* ============ 使用方法教程（聚光灯引导式 onboarding） ============ */
// 素材是一份内置的「虚构」随访表，走与真实上传**完全相同**的解析路径
//（SheetJS 生成 xlsx → File → Pipeline.processFiles），所以教程里看到的分布、
//  状态判定、根本原因分桶都是真实算出来的，映射规则演进时教程不会失真。
const DEMO_HEADERS = ["患者姓名", "联系电话", "药品名称", "适应症", "药店名称", "执行人",
  "任务状态", "执行时间", "用药周期状态", "未按计划持续用药原因", "用药依从性", "备注"];
const DEMO_SHEET = "入组随访（示例）";
const DEMO_FILE = "示例数据_入组随访（虚构，仅供教程）.xlsx";
const DEMO_DRUGS = ["百泽安", "百悦泽", "索托克拉"];
const DEMO_PHARMS = ["阳光大药房（示例）", "康泰药房（示例）", "惠民药房（示例）"];
const DEMO_EXECS = ["张护士", "李药师", "王随访"];
const DEMO_INDICATIONS = ["非小细胞肺癌", "肝细胞癌", "尿路上皮癌", "套细胞淋巴瘤", "华氏巨球蛋白血症"];
const DEMO_REMARKS = [
  "患者询问下次复查时间，已告知",
  "自述近期食欲一般，已建议复诊时反馈医生",
  "",
  "已提醒按时用药并做好记录",
  "",
  "家属代为接听，表示会按医嘱继续执行",
];
// 计划表：period 写入「用药周期状态」列（enrollment 来源的关键列），
// reason 写入「未按计划持续用药原因」→ 会被自动归入 13 个根本原因分桶之一。
function demoPlan() {
  return [
    { period: "按计划持续用药", n: 26, reason: "" },
    { period: "减量用药", n: 4, reason: "患者自行减少用药剂量，自觉症状缓解" },
    { period: "减量用药", n: 4, reason: "遵医嘱减量，因血象指标偏低" },
    { period: "推迟购药", n: 4, reason: "未到用药时间，用药周期尚未开始" },
    { period: "当期未复购", n: 4, reason: "患者自述购药不方便，家离门店较远" },
    { period: "持续用药中--流失", n: 5, reason: "家属反馈多次联系不上患者" },
    { period: "完全停用", n: 3, reason: "出现皮疹，患者不耐受而停药" },
    { period: "当期随访确认脱落", n: 3, reason: "经济负担较重，暂时停止用药" },
    { period: "停药----脱落", n: 2, reason: "病情稳定，已按疗程结束用药" },
    { period: "停药----脱落", n: 2, reason: "复查提示疗效不佳，患者想更换方案" },
    // 无「用药周期状态」取值 + 任务未执行 → 落到「其他」，用来演示「其他」的细分标注
    { period: "", n: 3, reason: "", task: "待执行" },
  ];
}
function demoRows() {
  const rows = [];
  let n = 0;
  for (const p of demoPlan()) {
    for (let k = 0; k < p.n; k++) {
      n++;
      const empty = !p.period;
      const month = 7 + (n % 3);                 // 跨 7/8/9 三个月，趋势图有 3 个点
      const day = 1 + ((n * 7) % 27);
      const hh = 9 + (n % 8);
      const date = `2026-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")} `
        + `${String(hh).padStart(2, "0")}:${n % 2 ? "30" : "00"}`;
      rows.push([
        "示例患者" + String(n).padStart(2, "0"),
        "138" + String(10000000 + n * 137).slice(-8),
        DEMO_DRUGS[n % 3],
        DEMO_INDICATIONS[n % 5],
        DEMO_PHARMS[n % 3],
        DEMO_EXECS[n % 3],
        p.task || "已完成",
        date,
        p.period,
        p.reason || "",
        empty ? "" : (n % 3 === 0 ? "良好" : "一般"),
        empty ? "" : DEMO_REMARKS[n % DEMO_REMARKS.length],
      ]);
    }
  }
  return rows;
}
// 生成一个真实的 .xlsx File 对象（教程里唯一需要外部库的地方）
function buildDemoFile() {
  const ws = XLSX.utils.aoa_to_sheet([DEMO_HEADERS].concat(demoRows()));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, DEMO_SHEET);
  const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  return new File([buf], DEMO_FILE,
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}
// 不经 xlsx 往返、直接跑归一化（供自动化测试核对示例数据的分布，不参与页面流程）
function demoRecordsDirect() {
  const cols = DEMO_HEADERS.slice();
  const rows = demoRows().map(r => {
    const o = {};
    for (let i = 0; i < r.length; i++) o[i] = window.Pipeline.cellStr(r[i]);
    return o;
  });
  return window.Pipeline.normalizeRows(rows, cols, DEMO_FILE, DEMO_SHEET);
}

const TUT = { on: false, i: 0, steps: [], target: null, snap: null, injected: false,
              raf: 0, root: null, card: null, masks: null, ring: null, timer: 0 };

function tutSetBox(el, x, y, w, h) {
  el.style.left = x + "px"; el.style.top = y + "px";
  el.style.width = Math.max(0, w) + "px"; el.style.height = Math.max(0, h) + "px";
}
function tutEnsureDom() {
  if (TUT.root) return;
  const root = document.createElement("div");
  root.className = "tut-root hidden";
  ["t", "b", "l", "r"].forEach(k => {
    const d = document.createElement("div");
    d.className = "tut-mask"; d.dataset.k = k; root.appendChild(d);
  });
  const ring = document.createElement("div");
  ring.className = "tut-ring"; root.appendChild(ring);
  document.body.appendChild(root);

  const card = document.createElement("div");
  card.className = "tut-card hidden";
  card.innerHTML =
    '<div class="tut-head"><span class="tut-step"></span><b class="tut-title"></b>' +
      '<button class="tut-x" type="button" title="退出教程">✕</button></div>' +
    '<div class="tut-body"></div><div class="tut-acts hidden"></div>' +
    '<div class="tut-foot"><span class="tut-dots"></span><span class="tut-ctl">' +
      '<button class="tut-prev" type="button">上一步</button>' +
      '<button class="tut-next" type="button">下一步</button>' +
      '<a class="tut-skip">跳过教程</a></span></div>';
  document.body.appendChild(card);

  TUT.root = root; TUT.card = card; TUT.ring = ring;
  TUT.masks = {};
  ["t", "b", "l", "r"].forEach(k => { TUT.masks[k] = root.querySelector('[data-k="' + k + '"]'); });
  card.querySelector(".tut-x").onclick = () => tutEnd("close");
  card.querySelector(".tut-prev").onclick = () => tutGo(TUT.i - 1);
  card.querySelector(".tut-next").onclick = () => tutGo(TUT.i + 1);
  card.querySelector(".tut-skip").onclick = () => tutEnd("skip");
}
// 四块遮罩围出「聚光区」：区内不放置任何元素，所以高亮的目标仍可被直接点选
function tutLayout() {
  if (!TUT.on || !TUT.card) return;
  const vw = window.innerWidth, vh = window.innerHeight, pad = 8;
  const r = TUT.target ? TUT.target.getBoundingClientRect() : null;
  let x = 0, y = 0, w = 0, h = 0;
  if (r && r.width > 2 && r.height > 2) {
    x = Math.max(0, r.left - pad);
    y = Math.max(0, r.top - pad);
    w = Math.min(vw, r.right + pad) - x;
    h = Math.min(vh, r.bottom + pad) - y;
    if (w < 8 || h < 8) { x = y = w = h = 0; }
  }
  const m = TUT.masks;
  if (w && h) {
    tutSetBox(m.t, 0, 0, vw, y);
    tutSetBox(m.b, 0, y + h, vw, vh - (y + h));
    tutSetBox(m.l, 0, y, x, h);
    tutSetBox(m.r, x + w, y, vw - (x + w), h);
    tutSetBox(TUT.ring, x, y, w, h);
    TUT.ring.style.display = "";
  } else {                                   // 无目标 / 目标不可见：整屏压暗
    tutSetBox(m.t, 0, 0, vw, vh);
    tutSetBox(m.b, 0, vh, 0, 0);
    tutSetBox(m.l, 0, vh, 0, 0);
    tutSetBox(m.r, 0, vh, 0, 0);
    TUT.ring.style.display = "none";
  }
  const card = TUT.card;
  const cw = Math.min(400, vw - 32);
  card.style.width = cw + "px";
  const ch = card.offsetHeight || 240;
  let cx, cy;
  if (w && h) {
    cx = Math.min(Math.max(12, x), Math.max(12, vw - cw - 12));
    if (y + h + 16 + ch <= vh - 10) cy = y + h + 16;              // 优先放下方
    else if (y - 16 - ch >= 10) cy = y - 16 - ch;                 // 其次放上方
    else cy = Math.max(10, Math.min(vh - ch - 10, (vh - ch) / 2)); // 都不够就居中
  } else {
    cx = Math.max(12, (vw - cw) / 2);
    cy = Math.max(20, Math.min(vh - ch - 20, (vh - ch) / 2 - 40));
  }
  card.style.left = cx + "px";
  card.style.top = cy + "px";
}
function tutSchedule() {
  if (TUT.raf) return;
  TUT.raf = requestAnimationFrame(() => { TUT.raf = 0; tutLayout(); });
}
// 平滑滚动可能被用户操作/布局变化打断，停在「目标刚好贴住视口边缘」的位置。
// 这里用即时 scrollBy 做一次矫正（不再走动画，避免和正在进行的平滑滚动互相追赶）。
function tutEnsureVisible() {
  if (!TUT.on || !TUT.target) return;
  const vh = window.innerHeight;
  const r = TUT.target.getBoundingClientRect();
  if (r.width < 2 || r.height < 2) return;
  const margin = 96;
  let dy = 0;
  if (r.top < margin && r.bottom < vh - margin) dy = r.top - margin;
  else if (r.bottom > vh - margin) dy = r.bottom - (vh - margin);
  if (dy) {
    const y0 = window.scrollY;
    window.scrollBy(0, dy);
    // 贴到文档顶/底后 scrollY 不再变化，说明已经到头，无需再纠
    if (Math.abs(window.scrollY - y0) < 1) return;
  }
  tutLayout();
}
window.addEventListener("scroll", () => { if (TUT.on) tutSchedule(); }, true);
window.addEventListener("resize", () => { if (TUT.on) tutLayout(); });

function tutToast(msg) {
  let el = document.querySelector(".tut-toast");
  if (!el) { el = document.createElement("div"); el.className = "tut-toast"; document.body.appendChild(el); }
  el.textContent = msg;
  el.classList.add("on");
  clearTimeout(TUT.timer);
  TUT.timer = setTimeout(() => el.classList.remove("on"), 2800);
}

function tutSteps() {
  const hasData = () => STORE.records.length > 0;
  return [
    { sel: null, t: "欢迎使用随访数据看板",
      body: `这是一个<b>纯本地</b>工具：Excel 在你的浏览器里解析和计算，数据不会上传到任何服务器。<br>
        接下来约 1 分钟，我用一份<b>虚构的示例数据</b>带你把每个功能点一遍 —— 包括哪些地方可以点、
        以及<b>点完之后哪些数字会跟着变</b>。<br>
        <span class="tut-muted">随时可按 Esc 退出。教程结束后示例数据会自动清空，不会混进你自己的数据。</span>` },

    { sel: "#drop", t: "上传随访表：把表格拖进来",
      body: `支持 .xls / .xlsx，可以一次选多个文件；<b>列名不需要事先统一</b>，表头会自动识别、映射成统一字段。<br>
        选中文件后先进入「待分析」列表，再点「开始分析」才真正解析。
        ${hasData() ? '<br><span class="tut-muted">你已经有数据了，直接点「下一步」继续即可。</span>' : ""}`,
      acts: hasData() ? null : [{
        label: "用示例数据演示", primary: true,
        fn: async () => { await tutLoadDemo(); await tutGo(TUT.i + 1); },
      }] },

    { sel: ".filebar", t: "已加载的文件",
      body: `每个文件一个标签，标签后面是它贡献的记录数。可以多次加载多个文件，数据会自动合并去重。<br>
        右侧「<b>清空全部</b>」一次性清空所有数据、回到上传界面。` },

    { sel: "#globalRow", t: "全局概览：这批数据整体长什么样",
      body: `总记录数、四种用药状态的构成、数据来源、任务完成情况都在这里。<br>
        注意「<b>其他</b>」这一态：它表示<b>没有足够的结构化数据可判定</b>的记录
        （例如任务未执行 / 已取消 / 联系失败），而不是"另外一种用药状态"。` },

    { sel: "#summary", t: "用药状态卡片：卡片本身就是按钮",
      body: `点任意一张卡片 → 立刻按该状态筛选（可以多选叠加），再点一次取消。<br>
        <b>点完之后页面上几乎所有数字都会跟着变</b>：上方图表、右侧各筛选项的计数、
        下方明细表，以及导出和快照的内容。` },

    { sel: "#subtypeBar", t: "试一试：点选一张卡片会怎样",
      before: async () => {
        state.status = new Set(["不规范用药"]); state.subtype = null; state.page = 1;
        await refresh();
      },
      body: `我已经替你点了「<b>不规范用药</b>」，注意三处变化：<br>
        <ul><li>卡片变成选中态；上方图表、下方明细同时只剩这一部分；</li>
        <li>这里多出一行「<b>不规范下钻</b>」，可以再按 <code>自行减量</code> / <code>医嘱减量</code> /
            <code>延迟未按时用药</code> 等具体类型细分；</li>
        <li>筛选栏下方会列出当前生效的全部筛选条件，方便核对。</li></ul>
        <span class="tut-muted">下一步我会把它取消掉。</span>` },

    { sel: ".charts", t: "图表区：四张卡里的条形都能点",
      before: async () => { state.status = new Set(); state.subtype = null; state.page = 1; await refresh(); },
      body: `用药状态占比（环形）、药品记录数、随访时间趋势、停药／减量根本原因。<br>
        每张卡里的<b>条形</b>都是按钮：点一次筛选，再点一次取消 —— 和卡片一样，全页联动。` },

    { sel: "#drugChart", t: "按品种筛选 —— 也是「只发某个品种」的办法",
      body: `这里列出的是<b>全部品种</b>（放不下时可滚动），不是只显示前几名。<br>
        点某一条 → 只看该药品。<br>
        更实用的是：<b>快照按当前筛选范围导出</b> —— 先筛好某个品种再下载快照，
        对方拿到的就只是这个品种的数据，文件名还会自动带上品种名。` },

    { sel: "#reasonCard", t: "停药／减量根本原因：自由文本自动归类",
      body: `原始原因列是自由填写的，几百种措辞各不相同，这里按 <b>13 个桶</b>自动归一化
        （医嘱调整 / 自主调整 / 经济费用 / 联系失败 / 不良反应 …），点条形即可下钻。<br>
        <b>统计口径</b>：只统计「任务状态 = 已完成」且确实填了根本原因的记录 ——
        卡片脚注里写明了具体条数，方便和明细核对。` },

    { sel: ".filterbar", t: "统一筛选栏：所有维度都能多选",
      body: `搜索（患者 / 电话 / 药店 / 药品）、药品、药店、执行人、根本原因、时间范围。<br>
        多个维度之间是「<b>且</b>」的关系；时间按<b>随访时间</b>筛选，起止两天都包含在内。<br>
        最右侧是「清除筛选」，一键复位所有条件。` },

    { sel: "#colPanel", t: "两种视角 + 自选列",
      before: async () => { $("#colPanel").classList.remove("hidden"); },
      body: `「<b>明细</b>」逐条记录、「<b>患者聚合</b>」把同一患者的多条记录合并起来看用药轨迹，随时切换。<br>
        下面这块是「<b>列显隐</b>」：表格太宽时，只留下你关心的列。` },

    { sel: "#detailView", t: "明细表：点一行看结构化原文",
      before: async () => { $("#colPanel").classList.add("hidden"); },
      body: `点任意一行会<b>展开该条记录的专项原文</b>（各项结构化字段的真实取值），
        而不是让你去读随访小结的自由文本。<br>
        表格可横向滚动，底部是分页和每页条数。` },

    { sel: ".fb-export", t: "导出与分享",
      body: `<ul>
        <li><b>导出脱敏 / 未脱敏明细</b>：生成带表头样式、按用药状态着色的 Excel（未脱敏需二次确认）。</li>
        <li><b>下载脱敏 / 不脱敏快照</b>：生成一个<b>可离线打开的单文件网页</b>，
            对方双击就能看，不用装任何环境。</li>
        <li>快照按<b>当前筛选范围</b>导出，并在页面里写明数据范围；脱敏快照不含任何明文姓名和电话。</li></ul>` },

    { sel: null, t: "就到这里，记住一个口诀",
      body: `<b>卡片、条形、下拉都能点；点完之后，所有数字都会跟着变。</b><br>
        教程结束，示例数据会自动清空，你可以把自己的表拖进来试试。<br>
        <span class="tut-muted">以后想再看一遍，点标题右边的「使用方法教程」即可。</span>` },
  ];
}

async function tutGo(i) {
  if (!TUT.on) return;
  if (i < 0) i = 0;
  if (i >= TUT.steps.length) return tutEnd("done");
  TUT.i = i;
  const st = TUT.steps[i];
  if (typeof st.before === "function") {
    try { await st.before(); } catch (e) { console.warn("教程步骤预处理失败", e); }
  }
  TUT.target = st.sel ? document.querySelector(st.sel) : null;
  if (TUT.target && typeof TUT.target.scrollIntoView === "function") {
    try { TUT.target.scrollIntoView({ block: "center", inline: "nearest", behavior: "smooth" }); } catch (_) {}
  }
  const c = TUT.card;
  c.querySelector(".tut-step").textContent = "第 " + (i + 1) + " / " + TUT.steps.length + " 步";
  c.querySelector(".tut-title").textContent = st.t;
  c.querySelector(".tut-body").innerHTML = st.body || "";
  const acts = c.querySelector(".tut-acts");
  acts.innerHTML = "";
  if (st.acts && st.acts.length) {
    acts.classList.remove("hidden");
    st.acts.forEach(a => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "tut-actbtn" + (a.primary ? " primary" : "");
      b.textContent = a.label;
      b.onclick = async () => {
        const old = b.textContent;
        b.disabled = true; b.textContent = "处理中…";
        try { await a.fn(); }
        catch (e) { alert("操作失败：" + (e && e.message ? e.message : e)); }
        finally { if (b.isConnected) { b.disabled = false; b.textContent = old; } }
      };
      acts.appendChild(b);
    });
  } else {
    acts.classList.add("hidden");
  }
  c.querySelector(".tut-dots").innerHTML =
    TUT.steps.map((_, k) => '<i class="tut-dot' + (k === i ? " on" : "") + '"></i>').join("");
  c.querySelector(".tut-prev").disabled = (i === 0);
  c.querySelector(".tut-next").textContent = (i === TUT.steps.length - 1) ? "完成" : "下一步";
  TUT.root.classList.remove("hidden");
  c.classList.remove("hidden");
  tutLayout();
  // scrollIntoView 是平滑滚动，途中位置在变：多补几次布局，让聚光区跟住目标，
  // 并在动画大致结束后做一次「别贴住视口边缘」的矫正。
  [60, 180, 340].forEach(d => setTimeout(() => { if (TUT.on && TUT.i === i) tutLayout(); }, d));
  [520, 760].forEach(d => setTimeout(() => { if (TUT.on && TUT.i === i) tutEnsureVisible(); }, d));
}

function tutSnapshotState() {
  return {
    status: [...state.status], subtype: state.subtype,
    drugs: [...state.drugs], pharmacies: [...state.pharmacies],
    executors: [...state.executors], reasons: [...state.reasons],
    start: state.start, end: state.end, q: state.q, page: state.page, view: state.view,
    search: $("#searchInput").value, sd: $("#startDate").value, ed: $("#endDate").value,
  };
}
async function tutRestoreState() {
  const s = TUT.snap;
  if (!s) return;
  state.status = new Set(s.status); state.subtype = s.subtype || null;
  state.drugs = new Set(s.drugs); state.pharmacies = new Set(s.pharmacies);
  state.executors = new Set(s.executors); state.reasons = new Set(s.reasons);
  state.start = s.start; state.end = s.end; state.q = s.q;
  state.page = s.page; state.view = s.view || "detail";
  $("#searchInput").value = s.search; $("#startDate").value = s.sd; $("#endDate").value = s.ed;
  $("#colPanel").classList.add("hidden");
  syncViewBtns();
  await refresh();
}
async function tutLoadDemo() {
  showLoading("正在生成示例数据并解析…");
  try {
    const n = await ingestFiles([buildDemoFile()], { tutorial: true });
    if (!n) { alert("示例数据生成失败，请改用你自己的 Excel 文件。"); return; }
    TUT.injected = true;
  } finally {
    hideLoading();
  }
  await new Promise(r => setTimeout(r, 120));
}
async function tutClearDemo() {
  if (!TUT.injected) return false;
  TUT.injected = false;
  if (!STORE.records.some(r => r._tutorial)) return false;
  STORE.records = STORE.records.filter(r => !r._tutorial);
  STORE.files = STORE.files.filter(f => !f._tutorial);
  if (!STORE.records.length) {
    STORE.seq = 0;
    CURRENT = { summary: null, global: null };
    DATA = { rows: [], patients: [] };
    $("#board").classList.add("hidden");
    $("#fileChips").innerHTML = "";
  } else {
    CURRENT.global = build_summary(STORE.records);
    renderGlobal(CURRENT.global);
    renderFiles(STORE.files);
    await refresh();
  }
  return true;
}
async function tutStart() {
  tutEnsureDom();
  closeAllPopovers();
  TUT.snap = tutSnapshotState();
  TUT.injected = false;
  TUT.steps = tutSteps();
  TUT.on = true;
  await tutGo(0);
}
async function tutEnd(reason) {
  if (!TUT.on) return;
  TUT.on = false;
  TUT.target = null;
  if (TUT.root) TUT.root.classList.add("hidden");
  if (TUT.card) TUT.card.classList.add("hidden");
  closeAllPopovers();
  try { await tutRestoreState(); } catch (e) { console.warn(e); }
  let cleared = false;
  try { cleared = await tutClearDemo(); } catch (e) { console.warn(e); }
  if (reason === "done") {
    tutToast(cleared ? "教程结束 · 示例数据已清空" : "教程结束，随时可以再点标题右边的按钮重看");
    if (cleared) { try { window.scrollTo({ top: 0, behavior: "smooth" }); } catch (_) {} }
  }
}
$("#tutBtn").onclick = () => { if (TUT.on) { tutEnd("close"); } else { tutStart(); } };
document.addEventListener("keydown", e => {
  if (!TUT.on) return;
  if (e.key === "Escape") { e.preventDefault(); tutEnd("close"); }
  else if (e.key === "ArrowRight" || e.key === "Enter") { e.preventDefault(); tutGo(TUT.i + 1); }
  else if (e.key === "ArrowLeft") { e.preventDefault(); tutGo(TUT.i - 1); }
});

// 调试/测试用：暴露核心计算与快照接口
window.AppCore = { loadSnapshot, build_summary, filter_records, patientsAgg, summaryLocal, STORE,
  tutStart, tutGo, tutEnd, tutLoadDemo, tutClearDemo, TUT,
  buildDemoFile, demoRows, demoRecordsDirect, DEMO_HEADERS, DEMO_FILE };
})();

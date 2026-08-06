// 随访表格映射与状态归一化配置（浏览器端移植自 mapping_config.py）
// 全部为纯函数 / 常量，无 DOM 依赖，可在浏览器与 Node 共用。

const CANONICAL_FIELDS = [
  "source_file", "source_type", "task_status",
  "patient_name", "patient_id", "phone", "gender", "age",
  "followup_time", "task_type", "drug_product", "indication", "pharmacy",
  "is_key_patient", "medication_status", "irregularity_subtype",
  "stop_reduce_reason",
  "adherence", "summary", "remarks",
];

const STATUS_TAXONOMY = ["规范用药", "不规范用药", "脱落停药", "其他"];
const SUBTYPE_TAXONOMY = ["自行减量", "医嘱减量", "医嘱停药", "延迟/未按时用药", "其他不规范"];

// 来源类型识别兜底：按顺序匹配「签名列」是否出现在表头中（关键字子串，抗尾标）
const SIGNATURES = [
  ["overdue_purchase", "据上次随访后，是否已按时购药？"],
  ["enrollment", "用药周期状态"],
  ["routine", "是否判定患者属于易脱落"],
];

// 关键字启发式映射（单值字段）
const KEYWORD_RULES = [
  ["patient_name",  ["姓名", "名字", "患者", "患者姓名", "会员姓名", "客户姓名", "name", "xm"]],
  ["patient_id",    ["患者编号", "患者编码", "会员号", "编号", "编码"]],
  ["phone",         ["电话", "手机", "联系", "phone", "tel"]],
  ["gender",        ["性别", "sex", "gender"]],
  ["age",           ["年龄", "岁", "age"]],
  ["followup_time", ["随访时间", "随访日期", "购药日期", "执行时间", "计划执行",
                     "记录时间", "回访日期", "日期", "时间"]],
  ["drug_product",  ["药品名称", "药物", "药品", "产品", "商品"]],
  ["indication",    ["适应症", "病种", "诊断", "病症"]],
  ["pharmacy",      ["药店名称", "门店", "药房名称", "购药药店", "药店", "药房"]],
  ["is_key_patient",["是否重点患者", "重点患者", "重点"]],
  ["adherence",     ["用药依从性", "依从性"]],
  ["summary",       ["随访小结", "小结", "随访记录"]],
  ["task_type",     ["任务类型", "任务摘要", "服务摘要", "随访类型"]],
  ["task_status",    ["任务状态", "随访任务状态", "任务执行状态"]],
  // —— 状态推导中间列（按来源关键字识别，不依赖具体列名）——
  ["_status_period",    ["用药周期状态", "周期状态"]],
  ["_nonstd_usage",     ["非标准用法用量的类型", "非标准用法", "非标准"]],
  ["_usage_status",     ["是否计划按时用药", "计划按时用药", "当前用药情况"]],
  ["_purchased_on_time",["已按时购药", "按时购药"]],
  ["_follow_confirm",   ["医生医嘱的确认", "按照医嘱服用", "医嘱的确认"]],
  ["_is_dropout",       ["是否判定患者属于易脱落", "易脱落"]],
  ["_dropout_reason",   ["判断为易脱落的原因", "易脱落的原因"]],
  ["_reduce_reason",    ["减量的具体原因", "停药的具体原因"]],
];

// 多值字段：收集所有命中列（用于合并备注、合并根因）
const KEYWORD_RULES_MULTI = [
  ["remarks", [
    "备注", "患者反馈", "用户反馈", "药师备注", "反馈",
  ]],
  ["stop_reduce_reason", [
    "未按计划持续用药原因", "脱落/流失原因", "未购药的原因",
    "未在门店购药的原因", "未遵医嘱", "未按计划",
    "根本原因", "根因", "遵医嘱减量",
    "遵医嘱推迟用药/停药的原因", "自行推迟用药/停药的原因",
    "停药的原因", "推迟用药的原因", "详细医嘱", "更改方案/更换药房",
  ]],
];

function normHeader(h) {
  return String(h == null ? "" : h).replace(/\s+/g, "").toLowerCase();
}

function _cell(row, col) {
  if (col == null) return null;
  const v = row[col];
  if (v == null) return null;
  if (typeof v === "number" && Number.isNaN(v)) return null;
  const s = String(v).trim();
  return (s === "" || s === "nan" || s === "NaN" || s === "None") ? null : s;
}

function _gtext(row, colmap, field) {
  let cols = colmap[field];
  if (!cols) return "";
  if (typeof cols === "string") cols = [cols];
  const vals = cols.map(c => _cell(row, c)).filter(Boolean);
  return vals.join("\n") || "";
}

function deriveStatus(sourceType, row, colmap) {
  const g = f => _gtext(row, colmap, f);
  const has = f => Boolean(g(f));

  if (sourceType === "enrollment") {
    const v = g("_status_period");
    if (v === "按计划持续用药") return "规范用药";
    if (v === "延迟用药" || v === "推迟购药" || v === "未按计划持续用药-不依从") return "不规范用药";
    if (v === "停药----脱落" || v === "随访失败") return "脱落停药";
    // 兜底：_status_period 缺失时按根因文本判别
    const reason = g("stop_reduce_reason") || "";
    if (/停药|流失|拒接|未拨通|拒绝随访|自主停药|脱落|换方案|转渠道/.test(reason)) return "脱落停药";
    if (/减量|延迟|推迟|不依从|不规律|减药/.test(reason)) return "不规范用药";
    const adh = g("adherence") || "";
    if (adh.includes("良好") || adh.includes("遵医嘱")) return "规范用药";
    return "其他";
  }
  if (sourceType === "routine") {
    // 优先用「是否计划按时用药?」等状态列（真实主力表列名）
    const usage = g("_usage_status");
    if (usage) {
      if (/停/.test(usage) && /(药|换)/.test(usage)) return "脱落停药";
      if (/推迟|延迟|延后|提前/.test(usage)) return "不规范用药";
      if (/减量|减药/.test(usage)) return "不规范用药";
      if (/按时|按医嘱正常|正常用药/.test(usage)) return "规范用药";
      return "其他";
    }
    // 回退：分级示例的「非标准用法用量的类型」
    const nonstd = g("_nonstd_usage");
    if (nonstd) {
      if (/停/.test(nonstd) && /(药|换)/.test(nonstd)) return "脱落停药";
      if (/减量|减药|推迟|延迟/.test(nonstd)) return "不规范用药";
      return "不规范用药";
    }
    // 根因含停药/转药也判脱落（「易脱落」标记仅作风险提示，不参与状态判定）
    const selfStop = g("stop_reduce_reason");
    if (selfStop && /停/.test(selfStop) && /(药|换)/.test(selfStop)) return "脱落停药";
    const follow = g("_follow_confirm");
    if (follow && follow !== "医生确认，按医嘱执行") return "不规范用药";
    const summ = g("summary") || "";
    if (summ.includes("足量") || summ.includes("按医嘱执行") ||
       (summ.includes("用药规范") && !summ.includes("减量") && !summ.includes("停"))) return "规范用药";
    // 兜底：已排除脱落/不规范且随访已完成（有小结）→ 规范；无小结才落「其他」
    if (summ) return "规范用药";
    return "其他";
  }
  if (sourceType === "overdue_purchase") {
    const onTime = g("_purchased_on_time");
    if (onTime === "是") return "规范用药";
    // 未按时购药：根因含停药/转药/换方案 → 脱落；其余（延迟未购等）→ 不规范
    const reason = g("stop_reduce_reason") || g("_reduce_reason") || "";
    if (/停\s*药|停药|转药|转渠道|换方案|换用|改用|改方案|使用其他|其他药|流失|脱落/.test(reason)) return "脱落停药";
    if (onTime === "否" || reason) return "不规范用药";
    return "其他";
  }
  return "其他";
}

function deriveIrregularitySubtype(sourceType, row, colmap) {
  const g = f => _gtext(row, colmap, f);
  const has = f => Boolean(g(f));

  if (sourceType === "routine") {
    const usage = g("_usage_status");
    if (usage && /推迟|延迟|延后|提前/.test(usage)) return "延迟/未按时用药";
    const t = g("_nonstd_usage");
    if (t && t.includes("自行减量")) return "自行减量";
    if (t && t.includes("医嘱减量")) return "医嘱减量";
    if (t && t.includes("医嘱停药")) return "医嘱停药";
    if (t && t.includes("自行停药")) return "其他不规范";
    const dr = g("_dropout_reason");
    if (dr) {
      if (dr.includes("延迟") || dr.includes("未按时") || dr.includes("推迟")) return "延迟/未按时用药";
      if (dr.includes("减量")) return "医嘱减量";
      if (dr.includes("停药") || dr.includes("换药")) return "医嘱停药";
      return "其他不规范";
    }
    return "其他不规范";
  }
  if (sourceType === "overdue_purchase") {
    const reduceReason = g("_reduce_reason");
    if (reduceReason && reduceReason.includes("自行减量")) return "自行减量";
    if (reduceReason && reduceReason.includes("医嘱减量")) return "医嘱减量";
    if (reduceReason && reduceReason.includes("停药")) return "医嘱停药";
    if (has("stop_reduce_reason")) {
      const rt = g("stop_reduce_reason");
      if (rt && (rt.includes("未购药") || rt.includes("未在门店"))) return "延迟/未按时用药";
      return "自行减量";
    }
    return "其他不规范";
  }
  if (sourceType === "enrollment") {
    const reason = g("stop_reduce_reason");
    const status = g("_status_period");
    if (reason && (reason.includes("自行减少用药剂量") || (reason.includes("减量") && reason.includes("自行")))) return "自行减量";
    if (reason && reason.includes("遵医嘱") && reason.includes("减量")) return "医嘱减量";
    if (reason && reason.includes("停药")) return "医嘱停药";
    if (reason && (reason.includes("依从性差") || reason.includes("未拨通"))) return "其他不规范";
    if (reason && (reason.includes("购药不方便") || reason.includes("配送不方便"))) return "延迟/未按时用药";
    if (status === "延迟用药" || status === "推迟购药") return "延迟/未按时用药";
    return "其他不规范";
  }
  return "其他不规范";
}

function _rawStatusText(sourceType, row, colmap) {
  const g = f => _gtext(row, colmap, f);
  if (sourceType === "enrollment") return g("_status_period") || null;
  if (sourceType === "routine") {
    const t = g("_nonstd_usage");
    if (t) return t;
    if (g("_is_dropout") === "是") return "判定为易脱落";
    const f = g("_follow_confirm");
    if (f && f !== "医生确认，按医嘱执行") return f;
    return null;
  }
  if (sourceType === "overdue_purchase") {
    const v = g("_purchased_on_time");
    if (v === "否") return "未按时购药";
    if (v === "是") return "已按时购药";
    return v || null;
  }
  return null;
}

function _rawSubtypeText(sourceType, row, colmap) {
  const g = f => _gtext(row, colmap, f);
  if (sourceType === "routine") {
    const parts = [g("_nonstd_usage"), g("_dropout_reason")].filter(Boolean);
    return parts.length ? Array.from(new Set(parts)).join("；") : null;
  }
  if (sourceType === "overdue_purchase") return g("_reduce_reason") || g("stop_reduce_reason") || null;
  if (sourceType === "enrollment") return g("stop_reduce_reason") || null;
  return null;
}

// 浏览器端：导出到全局，供 pipeline.js / app.js 使用（非模块化脚本共享词法作用域）
if (typeof window !== "undefined") {
  window.Mapping = { CANONICAL_FIELDS, STATUS_TAXONOMY, SUBTYPE_TAXONOMY, SIGNATURES,
    KEYWORD_RULES, KEYWORD_RULES_MULTI, normHeader, _cell, _gtext,
    deriveStatus, deriveIrregularitySubtype, _rawStatusText, _rawSubtypeText };
}

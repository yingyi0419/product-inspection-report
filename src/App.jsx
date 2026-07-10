
import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { AlertTriangle, CheckCircle2, Download, FileSpreadsheet, RefreshCw, UploadCloud } from "lucide-react";
import { Bar, BarChart, CartesianGrid, Cell, Legend, Pie, PieChart, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";

const HEADER_ROW_INDEX = 6;
const DETAIL_START_INDEX = 7;
const DETAIL_END_INDEX = 22;
const EXPECTED_SAMPLE_COUNT = 16;

const STANDARD_COLUMNS = [
  "來源工作表", "月份", "週別", "抽驗餐廳", "抽驗日期", "廠商", "統計廠商", "子公司", "產品效期", "整包重量", "抽驗者", "區塊序號", "樣本序號", "每片重量", "長", "寬", "其他問題", "重量判定", "長度判定", "寬度判定", "整體判定", "資料品質狀態", "資料品質備註"
];

const DIAGNOSIS_ROWS = [
  ["P01","表首","報表上方包含標題、餐廳、日期、月份等資訊，不是每筆明細資料。","欄位錯位或產生無效資料列。","轉成每筆明細的共同欄位。","人工檢查第 1 到第 6 列。","Power Query、Excel 樞紐分析"],
  ["P02","表尾","第 24 列之後可能包含規格說明、分切要求或表尾文字。","文字說明會混入資料表。","明細只保留第 8 到第 23 列。","人工檢查第 24 列之後。","Power Query 篩選列"],
  ["P03","頁首","列印用報表可能在每頁重複欄位標題。","重複頁首會被誤認為資料列。","偵測欄位名稱列，只保留樣本列。","檢查重複欄位名稱。","Power Query 移除重複列"],
  ["P04","頁尾","頁尾、簽核、頁碼或備註可能出現在資料區下方。","造成資料型態混亂。","排除非明細區。","確認頁尾是否排除。","Excel 篩選"],
  ["P05","小計","報表中可能含有小計列或區塊統計。","小計與明細混合會重複計算。","只保留一列一樣本的原始明細。","搜尋小計、平均、合計文字。","樞紐分析表"],
  ["P06","總計","彙整報告或總計資料不屬於明細資料。","總計混入會放大統計值。","工作表名稱含彙整或報告自動排除。","檢查工作表名稱。","Power Pivot"],
  ["P07","空白列","樣本區可能有空白列或未填量測值。","影響筆數檢核與缺漏率。","保留樣本列但標記缺漏，不補 0。","檢查重量、長、寬是否空白。","條件式格式"],
  ["P08","分類值缺漏","廠商、抽驗者、效期可能放在上方而非每列。","無法依分類做樞紐分析。","將區塊上方資訊帶入每筆明細。","人工比對表首與廠商區塊。","Power Query 填滿"],
  ["P09","資料型態錯誤","日期與數值欄位可能混入文字、單位或錯誤格式。","無法加總、平均、篩選或排序。","轉換數值與日期，失敗列異常。","檢查日期與數值轉型。","資料驗證"],
  ["P10","欄位名稱重複","橫向廠商區塊會重複每片重量、長、寬、其他問題。","重複欄名不適合標準資料表。","每個區塊轉直向明細。","檢查第 7 列每片重量出現次數。","取消樞紐"],
  ["P11","合併儲存格","主管閱讀報表常有合併儲存格。","會讓部分列或欄變空白。","以二維陣列解析並標記缺漏。","人工檢查合併儲存格。","取消合併"],
  ["P12","橫向廠商區塊","不同廠商資料橫向排列，每 4 欄為一個量測區塊。","不利依廠商、週別或日期統計。","將每個廠商區塊轉直向資料。","以每片重量欄作為區塊起點。","資料正規化"],
  ["P13","表單式資料","原始表格是給人填寫與列印，不是資料庫格式。","表單欄位與明細欄位混在一起。","轉成一列一筆資料。","確認可直接做樞紐分析。","Excel Table"],
  ["P14","彙整資料與明細資料混合","同一活頁簿可能同時有週別明細與彙整報告。","混合計算造成重複統計。","排除名稱含彙整或報告的工作表。","檢查排除清單。","Power Query"],
  ["P15","日期格式不一致","日期可能是序號、YYYYMMDD、YYYY/MM/DD 或多個日期在同一格。","日期無法正確排序或分組。","嘗試轉 yyyy-mm-dd，失敗標異常。","抽樣比對原始日期。","DATEVALUE"],
  ["P16","數值欄位混入文字","重量、長、寬可能混入文字、單位或錯誤輸入。","無法計算平均、異常率與合格率。","可轉則轉數值，失敗保留原值並標異常。","檢查數值異常清單。","資料驗證"],
].map(([問題編號, 問題類型, 問題描述, 對分析的影響, 建議處理方式, 檢核方式, 後續可用工具]) => ({問題編號, 問題類型, 問題描述, 對分析的影響, 建議處理方式, 檢核方式, 後續可用工具}));

const RISK_REMINDERS = [
  "AI 與程式只能協助初步診斷，不能取代人工檢核。",
  "清理後資料必須用筆數、數值範圍、分類完整性與抽樣比對檢查。",
  "若資料包含真實餐廳、廠商、人員或產品資訊，使用前應注意資料隱私。",
  "日期、廠商名稱、效期與數值異常需人工確認。",
  "彙整報告不可與明細資料混合計算，避免重複統計。",
  "程式產出的判定結果應視為初步檢查，不可直接作為正式稽核結論。",
];

const PROCESS_RULES = [
  ["表首", "第 1 到第 6 列視為表首或區塊上方資訊來源。"],
  ["表尾", "第 24 列之後視為規格說明或表尾，不納入明細資料。"],
  ["頁首", "以第 7 列作為主要欄位名稱列，避免重複頁首混入。"],
  ["頁尾", "頁尾、簽核、備註文字不列入清理後明細。"],
  ["空白列", "第 8 到第 23 列保留為樣本列，空白值標記缺漏，不補 0。"],
  ["彙整報告", "工作表名稱包含彙整或報告者自動排除。"],
  ["橫向廠商區塊", "偵測第 7 列每片重量欄位，每 4 欄轉成直向資料。"],
  ["日期格式", "支援 Excel 日期序號、YYYYMMDD、YYYY.MM.DD、YYYY/MM/DD、YYYY-MM-DD。"],
  ["數值格式", "重量、長、寬、整包重量嘗試轉成數值，失敗保留原值並標異常。"],
  ["異常值", "使用可調整預設範圍進行初步判定，結果需人工確認。"],
  ["缺漏值", "廠商、抽驗者、整包重量、數值欄位空白時標記缺漏待確認。"],
  ["廠商括號規則", "廠商名稱若為「母公司(子公司)」或「母公司（子公司）」，明細保留完整名稱；週別統計、月統計與老闆彙總表以括號前母公司作為統計廠商，括號內子公司一併納入計算，不排除。"],
].map(([處理項目, 處理方式]) => ({處理項目, 處理方式}));

const CHART_COLORS = ["#2563eb", "#16a34a", "#f97316", "#dc2626", "#7c3aed", "#0891b2", "#475569"];

function cleanText(value) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return formatDate(value);
  return String(value).replace(/\r?\n/g, " ").replace(/\s+/g, " ").trim();
}
function toHalfWidth(value) {
  return cleanText(value).replace(/[０-９Ａ-Ｚａ-ｚ．，－]/g, (char) => {
    const code = char.charCodeAt(0);
    if ((code >= 0xff10 && code <= 0xff19) || (code >= 0xff21 && code <= 0xff3a) || (code >= 0xff41 && code <= 0xff5a)) return String.fromCharCode(code - 0xfee0);
    if (char === "．") return ".";
    if (char === "，") return ",";
    if (char === "－") return "-";
    return char;
  });
}
function normalizeHeader(value) { return cleanText(value).replace(/\s|　/g, "").replace(/[()（）:：]/g, "").toLowerCase(); }
function isBlank(value) { return cleanText(value) === ""; }
function safeCell(row, colIndex) { return row && colIndex >= 0 ? row[colIndex] ?? "" : ""; }
function formatDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  return `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,"0")}-${String(date.getDate()).padStart(2,"0")}`;
}
function isValidYMD(y, m, d) {
  const date = new Date(y, m - 1, d);
  return date.getFullYear() === y && date.getMonth() + 1 === m && date.getDate() === d;
}
function parseDateValue(value, required = false) {
  if (isBlank(value)) return { value: "", ok: !required, status: required ? "missing" : "blank", note: required ? "日期空白" : "" };
  if (value instanceof Date) return { value: formatDate(value), ok: true, status: "ok", note: "" };
  if (typeof value === "number" && value > 20000 && value < 80000) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed && isValidYMD(parsed.y, parsed.m, parsed.d)) return { value: `${parsed.y}-${String(parsed.m).padStart(2,"0")}-${String(parsed.d).padStart(2,"0")}`, ok: true, status: "ok", note: "" };
  }
  const text = toHalfWidth(value);
  const dates = [];
  let match;
  const separated = /((?:19|20)\d{2})[./\-年](\d{1,2})[./\-月](\d{1,2})日?/g;
  while ((match = separated.exec(text))) {
    const y = Number(match[1]), m = Number(match[2]), d = Number(match[3]);
    if (isValidYMD(y,m,d)) dates.push(`${y}-${String(m).padStart(2,"0")}-${String(d).padStart(2,"0")}`);
  }
  const compact = /((?:19|20)\d{2})(\d{2})(\d{2})/g;
  while ((match = compact.exec(text))) {
    const y = Number(match[1]), m = Number(match[2]), d = Number(match[3]);
    if (isValidYMD(y,m,d)) dates.push(`${y}-${String(m).padStart(2,"0")}-${String(d).padStart(2,"0")}`);
  }
  const unique = [...new Set(dates)];
  if (unique.length === 1) return { value: unique[0], ok: true, status: "ok", note: "" };
  if (unique.length > 1) return { value: unique.join("、"), ok: true, status: "multi", note: `同一格包含多個日期：${text}` };
  return { value: text, ok: false, status: "invalid", note: `日期無法解析：${text}` };
}
function parseNumberValue(value, fieldName, optional = false) {
  if (isBlank(value)) return { value: "", number: null, ok: optional, status: optional ? "blank" : "missing", note: optional ? "" : `${fieldName}空白` };
  if (typeof value === "number" && Number.isFinite(value)) return { value, number: value, ok: true, status: "ok", note: "" };

  const original = toHalfWidth(value);
  let text = original.replace(/,/g, "").trim();

  // 自動修正常見人工輸入錯誤：連續小數點。
  // 例如 67..2、67...2 會轉成 67.2；但 6.7.2 這類不明確格式仍列為異常。
  let autoCorrectNote = "";
  let autoCorrectedText = text;
  if (/\.{2,}/.test(autoCorrectedText)) {
    autoCorrectedText = autoCorrectedText.replace(/\.{2,}/g, ".");
    const originalWithoutUnit = text.replace(/(公克|克|g|G|公分|厘米|cm|CM)/g, "").trim();
    const correctedWithoutUnit = autoCorrectedText.replace(/(公克|克|g|G|公分|厘米|cm|CM)/g, "").trim();
    if (/^[+-]?\d+(\.\d+)?$/.test(correctedWithoutUnit)) {
      text = autoCorrectedText;
      autoCorrectNote = `${fieldName}疑似多輸入小數點：${originalWithoutUnit}，已自動修正為 ${correctedWithoutUnit}，請人工確認`;
    }
  }

  if ((text.match(/\./g) || []).length > 1 || text.includes("..")) return { value: original, number: null, ok: false, status: "invalid", note: `${fieldName}數值格式錯誤：${original}` };

  const noUnit = text.replace(/(公克|克|g|G|公分|厘米|cm|CM)/g, "").trim();
  if (!/^[+-]?\d+(\.\d+)?$/.test(noUnit)) return { value: original, number: null, ok: false, status: "invalid", note: `${fieldName}不是數值：${original}` };
  const num = Number(noUnit);

  if (autoCorrectNote) {
    return { value: num, number: num, ok: true, status: "autoCorrected", note: autoCorrectNote };
  }

  return { value: num, number: num, ok: true, status: noUnit !== text ? "convertedWithUnit" : "ok", note: noUnit !== text ? `${fieldName}含單位文字，已嘗試轉換，請人工確認` : "" };
}
function judgeRange(parsed, min, max, label) {
  if (parsed.number === null) return { 判定: "無法判定", note: `${label}無法判定` };
  if (parsed.number < min) return { 判定: "低於標準", note: `${label}低於預設範圍 ${min}～${max}` };
  if (parsed.number > max) return { 判定: "高於標準", note: `${label}高於預設範圍 ${min}～${max}` };
  return { 判定: "正常", note: "" };
}
function isSheetEmpty(aoa) { return !aoa.some(row => row && row.some(cell => !isBlank(cell))); }
function parseSheetName(sheetName) {
  const text = cleanText(sheetName);
  return { 月份: text.match(/((?:19|20)\d{4})/)?.[1] || "", 週別: (text.match(/(第?[一二三四五六七八九十0-9]+[周週])/)?.[1] || "").replace("週", "周") };
}

function worksheetToAOA(worksheet) {
  const aoa = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", raw: true });

  // 若原始報表使用合併儲存格，xlsx 只會保留左上角值。
  // 這裡把合併範圍內的空白格補成左上角值，讓「廠商、效期、抽驗者」等欄位比較容易被正確讀取。
  (worksheet["!merges"] || []).forEach(range => {
    const startValue = safeCell(aoa[range.s.r], range.s.c);
    if (isBlank(startValue)) return;
    for (let r = range.s.r; r <= range.e.r; r++) {
      if (!aoa[r]) aoa[r] = [];
      for (let c = range.s.c; c <= range.e.c; c++) {
        if (isBlank(aoa[r][c])) aoa[r][c] = startValue;
      }
    }
  });

  return aoa;
}
function normalizeWeekCode(value) {
  const text = cleanText(value).replace(/週/g, "周").toUpperCase();
  const digit = text.match(/W\s*([1-9]\d*)/)?.[1] || text.match(/第?([1-9]\d*)周/)?.[1];
  if (digit) return `W${Number(digit)}`;
  const zh = text.match(/第?([一二三四五六七八九十]+)周/)?.[1];
  const map = { 一:1, 二:2, 三:3, 四:4, 五:5, 六:6, 七:7, 八:8, 九:9, 十:10 };
  if (zh) {
    if (zh === "十") return "W10";
    if (zh.startsWith("十")) return `W${10 + (map[zh.slice(1)] || 0)}`;
    if (zh.endsWith("十")) return `W${(map[zh[0]] || 1) * 10}`;
    if (zh.includes("十")) return `W${(map[zh[0]] || 1) * 10 + (map[zh.slice(-1)] || 0)}`;
    return `W${map[zh] || zh}`;
  }
  return text || "";
}
function periodKey(month, week) {
  return `${cleanText(month)}__${normalizeWeekCode(week)}`;
}
function parsePeriodCell(value) {
  const text = toHalfWidth(value).replace(/\s/g, "").replace(/週/g, "周");
  if (!text) return null;
  const slash = text.match(/((?:19|20)\d{2})[\/.-]?(\d{1,2})[\/.-]?W?([1-9]\d*)/i);
  if (slash) {
    const month = `${slash[1]}${String(Number(slash[2])).padStart(2, "0")}`;
    return { month, week: `W${Number(slash[3])}` };
  }
  const named = parseSheetName(text);
  if (named.月份 && named.週別) return { month: named.月份, week: normalizeWeekCode(named.週別) };
  return null;
}
function findPeriodNearCell(aoa, rowIndex, colIndex) {
  // 先找同一欄附近，再往左找最近的期間欄位；可支援彙整報告上方合併週期標題。
  for (let r = Math.max(0, rowIndex - 2); r <= Math.min(5, rowIndex + 1); r++) {
    for (let offset = 0; offset <= 8; offset++) {
      for (const direction of [1, -1]) {
        const c = colIndex + offset * direction;
        if (c < 0) continue;
        const period = parsePeriodCell(safeCell(aoa[r], c));
        if (period) return period;
      }
    }
  }
  return null;
}
function buildReportVendorHints(workbook) {
  const hints = {};

  workbook.SheetNames.forEach(sheetName => {
    if (!/彙整|報告|summary/i.test(sheetName)) return;

    try {
      const worksheet = workbook.Sheets[sheetName];
      if (!worksheet) return;
      const aoa = worksheetToAOA(worksheet);
      const maxRows = Math.min(8, aoa.length);

      for (let r = 0; r < maxRows; r++) {
        const row = aoa[r] || [];
        for (let c = 0; c < row.length; c++) {
          const label = normalizeHeader(row[c]);
          if (!["廠商", "廠商名稱", "供應商", "製造商"].includes(label)) continue;

          const vendor = findNextNonLabelValue(row, c, 8);
          if (!vendor || isInvalidVendorValue(vendor)) continue;

          const period = findPeriodNearCell(aoa, r, c);
          if (!period?.month || !period?.week) continue;

          const key = periodKey(period.month, period.week);
          hints[key] = hints[key] || [];
          if (!hints[key].includes(vendor)) hints[key].push(vendor);
        }
      }
    } catch (error) {
      // 彙整報告只是輔助補值來源；解析失敗不可影響主流程。
    }
  });

  return hints;
}
function escapeRegex(text) { return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }
function extractAfterLabel(text, label) { return cleanText(text).match(new RegExp(`${escapeRegex(label)}\\s*[:：]?\\s*(.*)$`))?.[1]?.trim() || ""; }
function isGenericLabelRemainder(value) {
  const text = normalizeHeader(value);
  return !text || [
    "名稱", "日期", "效期", "有效日期", "批號", "重量", "整包重量", "抽驗者", "檢驗者", "人員", "結果", "result"
  ].includes(text);
}
function isInvalidVendorValue(value) {
  const text = normalizeHeader(value);
  return !text || ["廠商", "廠商名稱", "供應商", "製造商", "名稱", "缺漏待確認", "未填"].includes(text);
}
function parseVendorCategory(value) {
  const raw = cleanText(value).replace(/（/g, "(").replace(/）/g, ")");
  if (!raw || raw === "缺漏待確認") {
    return { 原始廠商: raw || "缺漏待確認", 統計廠商: "缺漏待確認", 子公司: "" };
  }

  // 規則：括號前視為統計用母公司，括號內視為子公司或廠別。
  // 例如「泰安(泰立富)」與「泰安」在週別、月統計、老闆彙總表中都歸到「泰安」。
  const match = raw.match(/^(.+?)\s*\((.+?)\)\s*$/);
  if (match) {
    const parent = cleanText(match[1]);
    const child = cleanText(match[2]);
    return {
      原始廠商: raw,
      統計廠商: parent || raw,
      子公司: child || "",
    };
  }

  return { 原始廠商: raw, 統計廠商: raw, 子公司: "" };
}
function findNextNonLabelValue(row, col, maxOffset = 6) {
  for (let offset = 1; offset <= maxOffset; offset++) {
    const candidate = cleanText(safeCell(row, col + offset));
    if (!candidate) continue;
    const norm = normalizeHeader(candidate);
    if (["廠商", "廠商名稱", "供應商", "製造商", "產品效期", "效期", "有效日期", "整包重量", "包重量", "抽驗者", "抽測者", "檢驗者", "名稱"].includes(norm)) continue;
    return candidate;
  }
  return "";
}
function findLabelValue(aoa, labels, options = {}) {
  const maxRows = options.maxRows ?? 6, minCol = options.minCol ?? 0, maxCol = options.maxCol ?? 999;
  const normalizedLabels = labels.map(normalizeHeader).sort((a, b) => b.length - a.length);
  const labelMap = new Map(labels.map(label => [normalizeHeader(label), label]));
  for (let r = 0; r < Math.min(maxRows, aoa.length); r++) {
    const row = aoa[r] || [];
    for (let c = minCol; c <= Math.min(maxCol, row.length - 1); c++) {
      const text = cleanText(row[c]);
      if (!text) continue;
      const normText = normalizeHeader(text);
      const matchedNorm = normalizedLabels.find(label => normText.includes(label));
      if (!matchedNorm) continue;
      const matchedLabel = labelMap.get(matchedNorm) || matchedNorm;
      const direct = extractAfterLabel(text, matchedLabel);
      if (direct && !isGenericLabelRemainder(direct)) return direct;
      const adjacent = findNextNonLabelValue(row, c, 6);
      if (adjacent) return adjacent;
    }
  }
  return "";
}
function findFirstDateInTopRows(aoa) {
  for (let r = 0; r < Math.min(6, aoa.length); r++) for (let c = 0; c < (aoa[r] || []).length; c++) if (parseDateValue(aoa[r][c], false).ok && parseDateValue(aoa[r][c], false).value) return aoa[r][c];
  return "";
}
function findTopTextNearBlock(aoa, startCol) {
  const excluded = ["產品", "簡易", "抽驗", "抽測", "紀錄", "報表", "月份", "日期", "週", "周", "效期", "重量", "抽驗者", "每片", "長", "寬", "其他問題", "廠商", "名稱", "有效日期", "批號", "result", "結果"];
  for (let r = 0; r < Math.min(6, aoa.length); r++) for (let c = startCol; c <= startCol + 3; c++) {
    const text = cleanText(safeCell(aoa[r], c));
    if (!text) continue;
    const norm = normalizeHeader(text);
    if (!excluded.some(k => norm.includes(normalizeHeader(k))) && !isInvalidVendorValue(text) && Number.isNaN(Number(text))) return text;
  }
  return "";
}
function extractSheetInfo(sheetName, aoa) {
  const nameInfo = parseSheetName(sheetName);
  const dateRaw = findLabelValue(aoa, ["抽驗日期", "日期", "抽測日期", "檢驗日期"]) || findFirstDateInTopRows(aoa);
  const dateParsed = parseDateValue(dateRaw, true);
  return {
    來源工作表: sheetName,
    月份: nameInfo.月份 || findLabelValue(aoa, ["月份", "年月"]) || "缺漏待確認",
    週別: nameInfo.週別 || findLabelValue(aoa, ["週別", "周別", "週次", "周次"]) || sheetName,
    抽驗餐廳: findLabelValue(aoa, ["抽驗餐廳", "餐廳", "門市", "分店"]) || "缺漏待確認",
    抽驗日期: dateParsed.value || "缺漏待確認",
    日期解析: dateParsed,
    工作表名稱解析備註: nameInfo.月份 && nameInfo.週別 ? "" : "工作表名稱無法完整解析月份或週別",
  };
}
function extractBlockMeta(aoa, startCol) {
  const minCol = Math.max(0, startCol - 2), maxCol = startCol + 6;
  const vendorByLabelRaw = findLabelValue(aoa, ["廠商名稱", "廠商", "供應商", "製造商"], { minCol, maxCol });
  const vendorByLabel = isInvalidVendorValue(vendorByLabelRaw) ? "" : vendorByLabelRaw;
  const vendorByTopTextRaw = vendorByLabel ? "" : findTopTextNearBlock(aoa, startCol);
  const vendorByTopText = isInvalidVendorValue(vendorByTopTextRaw) ? "" : vendorByTopTextRaw;
  return {
    廠商: vendorByLabel || vendorByTopText || "缺漏待確認",
    廠商擷取方式: vendorByLabel ? "標籤擷取" : vendorByTopText ? "區塊上方文字擷取，請人工確認" : "缺漏",
    產品效期: findLabelValue(aoa, ["產品效期", "效期", "有效日期", "保存期限"], { minCol, maxCol }) || "缺漏待確認",
    整包重量: findLabelValue(aoa, ["整包重量", "包重量", "整包重", "總重量"], { minCol, maxCol }) || "缺漏待確認",
    抽驗者: findLabelValue(aoa, ["抽驗者", "抽測者", "檢驗者", "人員"], { minCol, maxCol }) || "缺漏待確認",
  };
}
function detectVendorBlocks(headerRow) {
  const blocks = [], insufficientStarts = [];
  for (let c = 0; c < headerRow.length; c++) if (normalizeHeader(headerRow[c]).includes("每片重量")) c + 3 < headerRow.length ? blocks.push({ startCol: c }) : insufficientStarts.push(c + 1);
  return { blocks, insufficientStarts };
}
function detectSampleColumn(headerRow) {
  const i = headerRow.findIndex(cell => { const text = normalizeHeader(cell); return text.includes("樣本") || text.includes("序號") || text === "no"; });
  return i >= 0 ? i : 0;
}
function hasEffectiveDetailValues(detailRows, blocks) {
  return detailRows.some(row => blocks.some(block => [0,1,2].some(offset => parseNumberValue(safeCell(row, block.startCol + offset), "檢核欄位", true).number !== null)));
}
function parseSampleNo(rawValue, expectedNo) {
  if (isBlank(rawValue)) return { value: "缺漏待確認", ok: false, note: `樣本序號空白，預期第 ${expectedNo} 筆` };
  const parsed = parseNumberValue(rawValue, "樣本序號");
  if (parsed.number === null || !Number.isInteger(parsed.number)) return { value: cleanText(rawValue), ok: false, note: `樣本序號不是整數：${cleanText(rawValue)}` };
  if (parsed.number < 1 || parsed.number > 16) return { value: parsed.number, ok: false, note: `樣本序號不是 1 到 16：${parsed.number}` };
  return { value: parsed.number, ok: true, note: "" };
}
function resolveQualityStatus(types) {
  return ["日期格式異常", "數值格式異常", "樣本序號異常", "疑似異常值", "缺漏值", "待人工確認", "非明細資料"].find(type => types.includes(type)) || "正常";
}
function buildCleanRow({ sheetInfo, blockMeta, block, blockSerialNo, sourceRow, sampleCol, expectedNo, specs }) {
  const notes = [], types = [];
  const add = (type, note) => { if (note) { types.push(type); notes.push(note); } };
  if (sheetInfo.抽驗餐廳 === "缺漏待確認") add("缺漏值", "抽驗餐廳空白");
  if (sheetInfo.月份 === "缺漏待確認") add("缺漏值", "月份空白");
  if (sheetInfo.工作表名稱解析備註) add("待人工確認", sheetInfo.工作表名稱解析備註);
  if (sheetInfo.日期解析.status === "invalid") add("日期格式異常", sheetInfo.日期解析.note);
  if (sheetInfo.日期解析.status === "missing") add("缺漏值", "抽驗日期空白");
  if (sheetInfo.日期解析.status === "multi") add("待人工確認", sheetInfo.日期解析.note);
  if (blockMeta.廠商 === "缺漏待確認") add("缺漏值", "廠商空白");
  if (blockMeta.產品效期 === "缺漏待確認") add("缺漏值", "產品效期空白");
  if (blockMeta.抽驗者 === "缺漏待確認") add("缺漏值", "抽驗者空白");
  if (blockMeta.廠商擷取方式.includes("人工確認")) add("待人工確認", `${blockMeta.廠商擷取方式}`);
  const packWeight = parseNumberValue(blockMeta.整包重量, "整包重量");
  if (packWeight.status === "missing") add("缺漏值", "整包重量空白");
  if (packWeight.status === "invalid") add("數值格式異常", packWeight.note);
  if (packWeight.status === "convertedWithUnit") add("待人工確認", packWeight.note);
  if (packWeight.status === "autoCorrected") add("待人工確認", packWeight.note);
  const sample = parseSampleNo(safeCell(sourceRow, sampleCol), expectedNo);
  if (!sample.ok) add("樣本序號異常", sample.note);
  const weight = parseNumberValue(safeCell(sourceRow, block.startCol), "每片重量");
  const length = parseNumberValue(safeCell(sourceRow, block.startCol + 1), "長");
  const width = parseNumberValue(safeCell(sourceRow, block.startCol + 2), "寬");
  const issue = cleanText(safeCell(sourceRow, block.startCol + 3));
  [weight, length, width].forEach(p => { if (p.status === "missing") add("缺漏值", p.note); if (p.status === "invalid") add("數值格式異常", p.note); if (p.status === "convertedWithUnit") add("待人工確認", p.note); if (p.status === "autoCorrected") add("待人工確認", p.note); });
  if (weight.value === "" && length.value === "" && width.value === "") add("缺漏值", "此樣本列的每片重量、長、寬皆空白");
  const wj = judgeRange(weight, specs.weightMin, specs.weightMax, "每片重量");
  const lj = judgeRange(length, specs.lengthMin, specs.lengthMax, "長度");
  const xj = judgeRange(width, specs.widthMin, specs.widthMax, "寬度");
  [wj, lj, xj].forEach(j => { if (["低於標準", "高於標準"].includes(j.判定)) add("疑似異常值", j.note); });
  if (issue) add("待人工確認", `其他問題欄有文字：${issue}`);
  const overall = [wj.判定, lj.判定, xj.判定].includes("無法判定") || issue ? "待人工確認" : [wj.判定, lj.判定, xj.判定].some(v => ["低於標準", "高於標準"].includes(v)) ? "異常" : "正常";
  const vendorCategory = parseVendorCategory(blockMeta.廠商);
  return {
    來源工作表: sheetInfo.來源工作表, 月份: sheetInfo.月份, 週別: sheetInfo.週別, 抽驗餐廳: sheetInfo.抽驗餐廳, 抽驗日期: sheetInfo.抽驗日期,
    廠商: vendorCategory.原始廠商, 統計廠商: vendorCategory.統計廠商, 子公司: vendorCategory.子公司, 產品效期: blockMeta.產品效期, 整包重量: packWeight.number !== null ? packWeight.value : blockMeta.整包重量, 抽驗者: blockMeta.抽驗者,
    區塊序號: blockSerialNo, 樣本序號: sample.value, 每片重量: weight.value, 長: length.value, 寬: width.value, 其他問題: issue,
    重量判定: wj.判定, 長度判定: lj.判定, 寬度判定: xj.判定, 整體判定: overall,
    資料品質狀態: resolveQualityStatus(types), 資料品質備註: notes.length ? [...new Set(notes)].join("；") : "無",
  };
}
function countBy(rows, key) { return rows.reduce((acc, row) => { const value = row[key] || "未填"; acc[value] = (acc[value] || 0) + 1; return acc; }, {}); }
function chartData(obj) { return Object.entries(obj).map(([名稱, 筆數]) => ({ 名稱, 筆數 })); }
function percent(n, d) { return d ? `${((n / d) * 100).toFixed(1)}%` : "0.0%"; }
function buildKpis({ workbookSheetCount, validSheets, excludedSheets, rows, totalBlocks }) {
  const theoretical = totalBlocks * EXPECTED_SAMPLE_COUNT;
  const note = row => cleanText(row.資料品質備註);
  const missing = rows.filter(row => /空白|缺漏/.test(note(row))).length;
  const numeric = rows.filter(row => /不是數值|數值格式錯誤|無法轉成數值/.test(note(row))).length;
  const date = rows.filter(row => /日期無法解析/.test(note(row))).length;
  const suspicious = rows.filter(row => /低於預設範圍|高於預設範圍/.test(note(row))).length;
  const normal = rows.filter(row => row.資料品質狀態 === "正常" && row.整體判定 === "正常").length;
  const abnormal = rows.length - normal;
  const weightAbnormal = rows.filter(row => ["低於標準", "高於標準"].includes(row.重量判定)).length;
  const lengthAbnormal = rows.filter(row => ["低於標準", "高於標準"].includes(row.長度判定)).length;
  const widthAbnormal = rows.filter(row => ["低於標準", "高於標準"].includes(row.寬度判定)).length;
  return [
    ["讀取工作表數", workbookSheetCount, "活頁簿內所有工作表數", "neutral"],
    ["有效工作表數", validSheets.length, "成功解析且有抽驗明細的工作表", validSheets.length ? "success" : "danger"],
    ["排除工作表數", excludedSheets.length, "不納入清理的工作表", excludedSheets.length ? "warning" : "success"],
    ["有效廠商區塊數", totalBlocks, "依每片重量欄位偵測到的區塊數", totalBlocks ? "success" : "danger"],
    ["清理後明細筆數", rows.length, "橫向轉直向後的實際筆數", rows.length ? "success" : "danger"],
    ["理論明細筆數", theoretical, "有效廠商區塊數 × 16", "neutral"],
    ["筆數檢核結果", rows.length === theoretical ? "通過" : "未通過", rows.length === theoretical ? "清理後明細筆數符合預期" : "筆數檢核未通過，請人工確認是否有空白區塊、缺漏樣本或格式變動", rows.length === theoretical ? "success" : "danger"],
    ["缺漏值筆數", missing, "含空白或缺漏待確認的資料列", missing ? "warning" : "success"],
    ["數值格式異常筆數", numeric, "重量、長、寬或整包重量無法轉數值", numeric ? "danger" : "success"],
    ["日期格式異常筆數", date, "抽驗日期無法解析的資料列", date ? "danger" : "success"],
    ["疑似異常值筆數", suspicious, "低於或高於預設範圍", suspicious ? "warning" : "success"],
    ["正常筆數", normal, "資料品質與整體判定皆正常", "success"],
    ["異常筆數", abnormal, "需要人工確認或異常的資料列", abnormal ? "warning" : "success"],
    ["整體合格率", percent(rows.filter(row => row.整體判定 === "正常").length, rows.length), "整體判定為正常的比例", "neutral"],
    ["重量異常率", percent(weightAbnormal, rows.length), "重量低於或高於標準的比例", weightAbnormal ? "warning" : "success"],
    ["長度異常率", percent(lengthAbnormal, rows.length), "長度低於或高於標準的比例", lengthAbnormal ? "warning" : "success"],
    ["寬度異常率", percent(widthAbnormal, rows.length), "寬度低於或高於標準的比例", widthAbnormal ? "warning" : "success"],
  ].map(([KPI名稱, KPI數值, 狀態說明, level]) => ({KPI名稱, KPI數值, 狀態說明, level}));
}
function parseWorkbook(workbook, specs) {
  const rows = [], excludedSheets = [], validSheets = [];
  const reportVendorHints = buildReportVendorHints(workbook);
  let totalBlocks = 0;
  workbook.SheetNames.forEach(sheetName => {
    try {
      if (/彙整|報告/.test(sheetName)) return excludedSheets.push({工作表名稱: sheetName, 排除原因: "工作表名稱包含「彙整」或「報告」", 檢核方式: "依工作表名稱判斷為彙整或主管閱讀報告，不納入明細清理"});
      const worksheet = workbook.Sheets[sheetName];
      const aoa = worksheetToAOA(worksheet);
      if (!aoa?.length || isSheetEmpty(aoa)) return excludedSheets.push({工作表名稱: sheetName, 排除原因: "完全空白的工作表", 檢核方式: "二維陣列沒有任何非空白儲存格"});
      const headerRow = aoa[HEADER_ROW_INDEX];
      if (!headerRow) return excludedSheets.push({工作表名稱: sheetName, 排除原因: "第 7 列欄位名稱列不存在", 檢核方式: "找不到預期第 7 列欄位名稱"});
      const { blocks, insufficientStarts } = detectVendorBlocks(headerRow);
      if (!blocks.length && insufficientStarts.length) return excludedSheets.push({工作表名稱: sheetName, 排除原因: "找到每片重量欄位，但後方欄位數不足", 檢核方式: `每片重量起始欄位：${insufficientStarts.join("、")}`});
      if (!blocks.length) return excludedSheets.push({工作表名稱: sheetName, 排除原因: "找不到「每片重量」欄位", 檢核方式: "檢查第 7 列是否有每片重量、長、寬、其他問題欄位"});
      const detailRows = aoa.slice(DETAIL_START_INDEX, DETAIL_END_INDEX + 1);
      if (!hasEffectiveDetailValues(detailRows, blocks)) return excludedSheets.push({工作表名稱: sheetName, 排除原因: "第 8 到第 23 列沒有任何有效抽驗數值", 檢核方式: "檢查廠商區塊的每片重量、長、寬是否有可轉換數值"});
      const sheetInfo = extractSheetInfo(sheetName, aoa);
      const sampleCol = detectSampleColumn(headerRow);
      const sheetRows = [];
      const hintVendors = reportVendorHints[periodKey(sheetInfo.月份, sheetInfo.週別)] || [];
      blocks.forEach((block, idx) => {
        const blockMeta = extractBlockMeta(aoa, block.startCol);

        // 若週別明細表的廠商列因合併儲存格、範本複製或欄位偏移導致抓不到，
        // 嘗試用同一活頁簿「彙整/報告」工作表的相同月份週別、相同區塊順序補入。
        // 這只補廠商名稱，不會創造量測資料；清理後備註會保留待確認。
        if (isInvalidVendorValue(blockMeta.廠商) && hintVendors[idx]) {
          blockMeta.廠商 = hintVendors[idx];
          blockMeta.廠商擷取方式 = "由彙整報告補入，請人工確認";
        }

        for (let i = 0; i < EXPECTED_SAMPLE_COUNT; i++) sheetRows.push(buildCleanRow({ sheetInfo, blockMeta, block, blockSerialNo: idx + 1, sourceRow: aoa[DETAIL_START_INDEX + i] || [], sampleCol, expectedNo: i + 1, specs }));
      });
      rows.push(...sheetRows);
      totalBlocks += blocks.length;
      validSheets.push({ 工作表名稱: sheetName, 廠商區塊數: blocks.length, 明細筆數: sheetRows.length });
    } catch (err) {
      excludedSheets.push({ 工作表名稱: sheetName, 排除原因: "單一工作表解析失敗", 檢核方式: err?.message || "未知錯誤，請人工檢查格式" });
    }
  });
  const vendorAbnormal = rows.reduce((acc, row) => { if (row.資料品質狀態 !== "正常" || row.整體判定 !== "正常") acc[row.統計廠商 || row.廠商 || "未填"] = (acc[row.統計廠商 || row.廠商 || "未填"] || 0) + 1; return acc; }, {});
  const kpis = buildKpis({ workbookSheetCount: workbook.SheetNames.length, validSheets, excludedSheets, rows, totalBlocks });
  return { workbookSheetCount: workbook.SheetNames.length, rows, excludedSheets, validSheets, totalBlocks, kpis, sheetChartData: validSheets.map(s => ({名稱:s.工作表名稱, 筆數:s.明細筆數})), qualityChartData: chartData(countBy(rows, "資料品質狀態")), overallChartData: chartData(countBy(rows, "整體判定")), vendorAbnormalChartData: chartData(vendorAbnormal).sort((a,b)=>b.筆數-a.筆數).slice(0,12) };
}
function uniqueValues(rows, key) { return [...new Set(rows.map(row => row[key]).filter(Boolean))].sort(); }
function vendorValues(rows) { return uniqueValues(rows, "統計廠商").filter(value => !isInvalidVendorValue(value)); }
function rowsByMonthWeek(rows, month, week) {
  return rows.filter(row => (!month || row.月份 === month) && (!week || row.週別 === week));
}
function getBadgeClass(value) { const text = cleanText(value); if (/正常|通過/.test(text)) return "badge success"; if (/異常|低於|高於|未通過|無法/.test(text)) return "badge danger"; if (/待|缺漏|空白/.test(text)) return "badge warning"; return "badge neutral"; }
function DataTable({ rows, columns, emptyText="尚無資料" }) {
  if (!rows?.length) return <div className="empty-state">{emptyText}</div>;
  return <div className="table-wrapper"><table className="data-table"><thead><tr>{columns.map(col=><th key={col}>{col}</th>)}</tr></thead><tbody>{rows.map((row, i)=><tr key={i}>{columns.map(col=>{ const value = row[col] ?? ""; const badge = ["資料品質狀態","整體判定","重量判定","長度判定","寬度判定","KPI數值"].includes(col); return <td key={col} title={cleanText(value)}>{badge ? <span className={getBadgeClass(value)}>{cleanText(value)}</span> : cleanText(value)}</td>; })}</tr>)}</tbody></table></div>;
}
function ChartCard({ title, data, children }) { return <div className="chart-card"><h3>{title}</h3>{!data?.length ? <div className="empty-state">上傳並成功解析 Excel 後顯示圖表</div> : <div className="chart-box">{children}</div>}</div>; }


function parsePlainNumber(value, fieldName = "數值") {
  const parsed = parseNumberValue(value, fieldName, true);
  return parsed.number;
}
function percentValue(numerator, denominator) {
  if (!denominator) return 0;
  return Number(((numerator / denominator) * 100).toFixed(1));
}
function percentTextFromNumber(value) {
  return `${Number(value || 0).toFixed(1)}%`;
}
function uniqueJoined(rows, key) {
  return uniqueValues(rows, key).join("、") || "缺漏待確認";
}
function filterRowsBySummary(rows, filters) {
  return rows.filter(row =>
    (!filters.month || row.月份 === filters.month) &&
    (!filters.week || row.週別 === filters.week) &&
    (!filters.vendor || row.統計廠商 === filters.vendor)
  );
}
function compareStandard(ruleType, resultPct) {
  if (ruleType === "min100") {
    if (resultPct >= 100) return { text: "符合標準", level: "ok" };
    return { text: "低於標準規格", level: "low" };
  }
  if (ruleType === "max0") {
    if (resultPct <= 0) return { text: "符合標準", level: "ok" };
    return { text: "高於標準規格", level: "high" };
  }
  if (ruleType === "max10") {
    if (resultPct < 10) return { text: "符合標準", level: "ok" };
    return { text: "高於標準規格", level: "high" };
  }
  return { text: "待人工確認", level: "warn" };
}
function buildBossSummaryRows(rows) {
  const weights = rows.map(row => parsePlainNumber(row.每片重量, "每片重量")).filter(v => v !== null);
  const lengths = rows.map(row => parsePlainNumber(row.長, "長")).filter(v => v !== null);
  const widths = rows.map(row => parsePlainNumber(row.寬, "寬")).filter(v => v !== null);
  const issueTotal = rows.length;
  const issueCount = rows.filter(row => cleanText(row.其他問題) !== "").length;

  const defs = [
    { 項目: "重量", 規格: "60~70g", 標準: "100.0%", type: "min100", total: weights.length, count: weights.filter(v => v >= 60 && v <= 70).length },
    { 項目: "重量", 規格: "<60g", 標準: "0.0%", type: "max0", total: weights.length, count: weights.filter(v => v < 60).length },
    { 項目: "重量", 規格: "60~62g", 標準: "<10%", type: "max10", total: weights.length, count: weights.filter(v => v >= 60 && v <= 62).length },
    { 項目: "重量", 規格: "68~70g", 標準: "<10%", type: "max10", total: weights.length, count: weights.filter(v => v >= 68 && v <= 70).length },
    { 項目: "重量", 規格: ">70g", 標準: "0.0%", type: "max0", total: weights.length, count: weights.filter(v => v > 70).length },
    { 項目: "長度", 規格: "12~16cm", 標準: "100.0%", type: "min100", total: lengths.length, count: lengths.filter(v => v >= 12 && v <= 16).length },
    { 項目: "長度", 規格: "12~13cm", 標準: "<10%", type: "max10", total: lengths.length, count: lengths.filter(v => v >= 12 && v <= 13).length },
    { 項目: "長度", 規格: ">16cm", 標準: "0.0%", type: "max0", total: lengths.length, count: lengths.filter(v => v > 16).length },
    { 項目: "長度", 規格: "<12cm", 標準: "0.0%", type: "max0", total: lengths.length, count: lengths.filter(v => v < 12).length },
    { 項目: "寬度", 規格: "3.5~4.5cm", 標準: "100.0%", type: "min100", total: widths.length, count: widths.filter(v => v >= 3.5 && v <= 4.5).length },
    { 項目: "寬度", 規格: "3.5~3.6cm", 標準: "<10%", type: "max10", total: widths.length, count: widths.filter(v => v >= 3.5 && v <= 3.6).length },
    { 項目: "寬度", 規格: "<3.5cm", 標準: "0.0%", type: "max0", total: widths.length, count: widths.filter(v => v < 3.5).length },
    { 項目: "寬度", 規格: ">4.5cm", 標準: "0.0%", type: "max0", total: widths.length, count: widths.filter(v => v > 4.5).length },
    { 項目: "其他問題", 規格: "其他問題", 標準: "0.0%", type: "max0", total: issueTotal, count: issueCount },
  ];

  return defs.map(def => {
    const resultPct = percentValue(def.count, def.total);
    const compared = compareStandard(def.type, resultPct);
    return {
      項目: def.項目,
      規格: def.規格,
      標準: def.標準,
      Result: percentTextFromNumber(resultPct),
      分子筆數: def.count,
      分母筆數: def.total,
      比較結果: compared.text,
      level: compared.level,
    };
  });
}
function buildInspectionStats(rows) {
  const total = rows.length;
  const normalOverall = rows.filter(row => row.整體判定 === "正常").length;
  const abnormalOverall = rows.filter(row => row.整體判定 !== "正常" || row.資料品質狀態 !== "正常").length;
  const weightNums = rows.map(row => parsePlainNumber(row.每片重量, "每片重量")).filter(v => v !== null);
  const lengthNums = rows.map(row => parsePlainNumber(row.長, "長")).filter(v => v !== null);
  const widthNums = rows.map(row => parsePlainNumber(row.寬, "寬")).filter(v => v !== null);
  const weightOk = weightNums.filter(v => v >= 60 && v <= 70).length;
  const lengthOk = lengthNums.filter(v => v >= 12 && v <= 16).length;
  const widthOk = widthNums.filter(v => v >= 3.5 && v <= 4.5).length;
  const issueCount = rows.filter(row => cleanText(row.其他問題) !== "").length;
  return [
    { 指標: "明細筆數", 數值: total, 說明: "目前篩選條件下的樣本筆數", level: total ? "neutral" : "danger" },
    { 指標: "廠商數", 數值: uniqueValues(rows, "統計廠商").length, 說明: "目前篩選條件下的不同廠商數", level: "neutral" },
    { 指標: "有效日期數", 數值: uniqueValues(rows, "產品效期").length, 說明: "產品效期不重複數量", level: "neutral" },
    { 指標: "整體正常筆數", 數值: normalOverall, 說明: "整體判定為正常", level: "success" },
    { 指標: "異常/待確認筆數", 數值: abnormalOverall, 說明: "含異常或資料品質非正常", level: abnormalOverall ? "warning" : "success" },
    { 指標: "整體合格率", 數值: percent(total ? normalOverall : 0, total), 說明: "整體正常筆數 ÷ 明細筆數", level: "neutral" },
    { 指標: "重量 60~70g 比率", 數值: percent(weightOk, weightNums.length), 說明: "每片重量落在 60~70g 的比例", level: weightNums.length && weightOk < weightNums.length ? "warning" : "success" },
    { 指標: "長度 12~16cm 比率", 數值: percent(lengthOk, lengthNums.length), 說明: "長度落在 12~16cm 的比例", level: lengthNums.length && lengthOk < lengthNums.length ? "warning" : "success" },
    { 指標: "寬度 3.5~4.5cm 比率", 數值: percent(widthOk, widthNums.length), 說明: "寬度落在 3.5~4.5cm 的比例", level: widthNums.length && widthOk < widthNums.length ? "warning" : "success" },
    { 指標: "其他問題率", 數值: percent(issueCount, total), 說明: "其他問題欄有文字的比例", level: issueCount ? "warning" : "success" },
  ];
}
function buildMonthlyStats(rows) {
  const monthMap = rows.reduce((acc, row) => {
    const month = row.月份 || "缺漏待確認";
    if (!acc[month]) acc[month] = [];
    acc[month].push(row);
    return acc;
  }, {});
  return Object.entries(monthMap).sort(([a], [b]) => String(a).localeCompare(String(b))).map(([月份, monthRows]) => {
    const stats = buildInspectionStats(monthRows);
    const get = name => stats.find(item => item.指標 === name)?.數值 ?? "";
    return {
      月份,
      明細筆數: get("明細筆數"),
      週別數: uniqueValues(monthRows, "週別").length,
      廠商數: get("廠商數"),
      有效日期數: get("有效日期數"),
      整體正常筆數: get("整體正常筆數"),
      異常待確認筆數: get("異常/待確認筆數"),
      整體合格率: get("整體合格率"),
      重量6070比率: get("重量 60~70g 比率"),
      長度1216比率: get("長度 12~16cm 比率"),
      寬度3545比率: get("寬度 3.5~4.5cm 比率"),
      其他問題率: get("其他問題率"),
    };
  });
}
function buildBossSummaryAOA({ filters, rows, bossRows }) {
  const dateList = uniqueJoined(rows, "產品效期");
  const inspectionDates = uniqueJoined(rows, "抽驗日期");
  const originalVendors = uniqueJoined(rows, "廠商");
  const subsidiaries = uniqueValues(rows, "子公司").filter(Boolean).join("、") || "無";
  const titleWeek = filters.week || "全部週別";
  const titleVendor = filters.vendor || "全部統計廠商";
  const titleMonth = filters.month || "全部月份";
  const aoa = [
    ["月份", titleMonth, "", "", ""],
    ["週別", titleWeek, "", "", ""],
    ["統計廠商", titleVendor, "", "", ""],
    ["原始廠商/子公司", originalVendors, "", "", ""],
    ["子公司清單", subsidiaries, "", "", ""],
    ["有效日期", dateList, "", "", ""],
    ["抽驗日期", inspectionDates, "", "", ""],
    ["", "", "規格", "標準", "Result", "比較結果", "分子筆數", "分母筆數"],
    ...bossRows.map(row => [row.項目, "", row.規格, row.標準, row.Result, row.比較結果, row.分子筆數, row.分母筆數]),
    [],
    ["低於標準規格", "代表結果未達應為 100% 的要求，例如 60~70g 未達 100%。"],
    ["高於標準規格", "代表結果超過應為 0% 或 <10% 的要求，例如 >4.5cm 過高。"],
  ];
  return aoa;
}
function buildAllBossSummaryRows(rows) {
  const groups = {};
  rows.forEach(row => {
    const key = `${row.月份 || "缺漏待確認"}__${row.週別 || "缺漏待確認"}__${row.統計廠商 || row.廠商 || "缺漏待確認"}`;
    if (!groups[key]) groups[key] = [];
    groups[key].push(row);
  });
  return Object.entries(groups).flatMap(([key, groupRows]) => {
    const [月份, 週別, 廠商] = key.split("__");
    const dates = uniqueJoined(groupRows, "產品效期");
    const originalVendors = uniqueJoined(groupRows, "廠商");
    const subsidiaries = uniqueValues(groupRows, "子公司").filter(Boolean).join("、") || "無";
    return buildBossSummaryRows(groupRows).map(summary => ({
      月份,
      週別,
      統計廠商: 廠商,
      原始廠商清單: originalVendors,
      子公司清單: subsidiaries,
      有效日期: dates,
      項目: summary.項目,
      規格: summary.規格,
      標準: summary.標準,
      Result: summary.Result,
      比較結果: summary.比較結果,
      分子筆數: summary.分子筆數,
      分母筆數: summary.分母筆數,
    }));
  });
}
function BossSummaryTable({ rows }) {
  if (!rows?.length) return <div className="empty-state">目前沒有可彙總的資料</div>;
  return <div className="table-wrapper boss-wrapper"><table className="boss-table"><thead><tr><th>項目</th><th>規格</th><th>標準</th><th>Result</th><th>比較結果</th><th>計算筆數</th></tr></thead><tbody>{rows.map((row, index) => <tr key={`${row.項目}-${row.規格}-${index}`}><td>{row.項目}</td><td>{row.規格}</td><td>{row.標準}</td><td className={`result-cell ${row.level}`}>{row.Result}</td><td><span className={row.level === "ok" ? "badge success" : row.level === "low" ? "badge warning" : row.level === "high" ? "badge danger" : "badge neutral"}>{row.比較結果}</span></td><td>{row.分子筆數} / {row.分母筆數}</td></tr>)}</tbody></table></div>;
}
function saveWorkbookWithFallback(workbook, fileName) {
  try {
    XLSX.writeFile(workbook, fileName);
  } catch (writeFileError) {
    const arrayBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([arrayBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }
}


function isExcelFile(file) {
  if (!file?.name) return false;
  const name = file.name.toLowerCase();
  if (name.startsWith("~$")) return false;
  return [".xlsx", ".xls", ".xlsm"].some(ext => name.endsWith(ext));
}
function displayFilePath(file) {
  return file?.webkitRelativePath || file?.relativeFolderPath || file?.name || "未命名檔案";
}
function readWorkbookFromFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = event => {
      try {
        const workbook = XLSX.read(event.target.result, { type: "array", cellDates: true, cellNF: true, cellText: false });
        resolve(workbook);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(new Error("檔案讀取失敗"));
    reader.readAsArrayBuffer(file);
  });
}
function decorateParsedResult(parsed, file) {
  const filePath = displayFilePath(file);
  return {
    ...parsed,
    rows: parsed.rows.map(row => ({ ...row, 來源工作表: `${filePath} / ${row.來源工作表}` })),
    validSheets: parsed.validSheets.map(sheet => ({ ...sheet, 工作表名稱: `${filePath} / ${sheet.工作表名稱}` })),
    excludedSheets: parsed.excludedSheets.map(sheet => ({ ...sheet, 工作表名稱: `${filePath} / ${sheet.工作表名稱}` })),
  };
}
function buildCombinedResult(parsedResults, additionalExcludedSheets = []) {
  const rows = parsedResults.flatMap(item => item.rows || []);
  const validSheets = parsedResults.flatMap(item => item.validSheets || []);
  const excludedSheets = [
    ...parsedResults.flatMap(item => item.excludedSheets || []),
    ...additionalExcludedSheets,
  ];
  const workbookSheetCount = parsedResults.reduce((sum, item) => sum + (item.workbookSheetCount || 0), 0);
  const totalBlocks = parsedResults.reduce((sum, item) => sum + (item.totalBlocks || 0), 0);
  const vendorAbnormal = rows.reduce((acc, row) => {
    if (row.資料品質狀態 !== "正常" || row.整體判定 !== "正常") {
      const vendor = row.統計廠商 || row.廠商 || "未填";
      acc[vendor] = (acc[vendor] || 0) + 1;
    }
    return acc;
  }, {});
  const kpis = buildKpis({ workbookSheetCount, validSheets, excludedSheets, rows, totalBlocks });
  kpis.unshift({ KPI名稱: "讀取 Excel 檔案數", KPI數值: parsedResults.length, 狀態說明: "本次成功讀取的 Excel 檔案數；可由資料匣或多檔上傳", level: parsedResults.length ? "success" : "danger" });
  return {
    workbookSheetCount,
    parsedExcelFileCount: parsedResults.length,
    rows,
    excludedSheets,
    validSheets,
    totalBlocks,
    kpis,
    sheetChartData: validSheets.map(s => ({ 名稱: s.工作表名稱, 筆數: s.明細筆數 })),
    qualityChartData: chartData(countBy(rows, "資料品質狀態")),
    overallChartData: chartData(countBy(rows, "整體判定")),
    vendorAbnormalChartData: chartData(vendorAbnormal).sort((a,b)=>b.筆數-a.筆數).slice(0,12),
  };
}
async function readDirectoryEntry(entry, path = "") {
  if (!entry) return [];
  if (entry.isFile) {
    return new Promise(resolve => {
      entry.file(file => {
        file.relativeFolderPath = `${path}${file.name}`;
        resolve([file]);
      }, () => resolve([]));
    });
  }
  if (entry.isDirectory) {
    const reader = entry.createReader();
    const allEntries = [];
    while (true) {
      const batch = await new Promise(resolve => reader.readEntries(resolve, () => resolve([])));
      if (!batch.length) break;
      allEntries.push(...batch);
    }
    const nested = await Promise.all(allEntries.map(child => readDirectoryEntry(child, `${path}${entry.name}/`)));
    return nested.flat();
  }
  return [];
}
async function filesFromDataTransfer(dataTransfer) {
  const items = Array.from(dataTransfer?.items || []);
  const entries = items.map(item => item.webkitGetAsEntry?.()).filter(Boolean);
  if (entries.length) {
    const files = await Promise.all(entries.map(entry => readDirectoryEntry(entry, "")));
    return files.flat();
  }
  return Array.from(dataTransfer?.files || []);
}

export default function App() {
  const [fileName, setFileName] = useState("");
  const [status, setStatus] = useState("idle");
  const [errorMessage, setErrorMessage] = useState("");
  const [result, setResult] = useState(null);
  const [lastFiles, setLastFiles] = useState([]);
  const [isDragging, setIsDragging] = useState(false);
  const [specs, setSpecs] = useState({ weightMin: 60, weightMax: 70, lengthMin: 12, lengthMax: 16, widthMin: 3.5, widthMax: 4.5 });
  const [filters, setFilters] = useState({ sheet:"", vendor:"", quality:"", overall:"" });
  const [summaryFilters, setSummaryFilters] = useState({ month:"", week:"", vendor:"" });

  const rows = result?.rows || [];
  const filteredRows = useMemo(() => rows.filter(row => (!filters.sheet || row.來源工作表 === filters.sheet) && (!filters.vendor || row.統計廠商 === filters.vendor) && (!filters.quality || row.資料品質狀態 === filters.quality) && (!filters.overall || row.整體判定 === filters.overall)), [rows, filters]);
  const abnormalRows = useMemo(() => rows.filter(row => row.資料品質狀態 !== "正常" || row.整體判定 !== "正常"), [rows]);
  const summaryRows = useMemo(() => filterRowsBySummary(rows, summaryFilters), [rows, summaryFilters]);
  const weeklyStats = useMemo(() => buildInspectionStats(summaryRows), [summaryRows]);
  const monthlyStats = useMemo(() => buildMonthlyStats(rows), [rows]);
  const bossRows = useMemo(() => buildBossSummaryRows(summaryRows), [summaryRows]);

  const sheetOptions = useMemo(() => uniqueValues(rows, "來源工作表"), [rows]);
  const vendorOptions = useMemo(() => vendorValues(rows), [rows]);
  const qualityOptions = useMemo(() => uniqueValues(rows, "資料品質狀態"), [rows]);
  const overallOptions = useMemo(() => uniqueValues(rows, "整體判定"), [rows]);
  const monthOptions = useMemo(() => uniqueValues(rows, "月份"), [rows]);
  const weekOptions = useMemo(() => uniqueValues(rowsByMonthWeek(rows, summaryFilters.month, ""), "週別"), [rows, summaryFilters.month]);
  const summaryVendorOptions = useMemo(() => vendorValues(rowsByMonthWeek(rows, summaryFilters.month, summaryFilters.week)), [rows, summaryFilters.month, summaryFilters.week]);

  async function processFiles(fileList, activeSpecs = specs) {
    setErrorMessage("");
    setResult(null);

    const selectedFiles = Array.from(fileList || []).filter(Boolean);
    if (!selectedFiles.length) {
      setFileName("");
      setLastFiles([]);
      setErrorMessage("請先選擇 Excel 檔案或資料匣。");
      setStatus("error");
      return;
    }

    const excelFiles = selectedFiles.filter(isExcelFile);
    const skippedFiles = selectedFiles
      .filter(file => !isExcelFile(file))
      .map(file => ({
        工作表名稱: displayFilePath(file),
        排除原因: "非 Excel 檔案，資料匣上傳時自動略過",
        檢核方式: "只處理 .xlsx、.xls、.xlsm；暫存檔與其他格式不納入解析",
      }));

    if (!excelFiles.length) {
      setFileName(selectedFiles.length === 1 ? displayFilePath(selectedFiles[0]) : `${selectedFiles.length} 個檔案`);
      setLastFiles([]);
      setResult(buildCombinedResult([], skippedFiles));
      setErrorMessage("此資料匣或選取檔案中沒有可讀取的 Excel 檔案，請選擇 .xlsx、.xls 或 .xlsm。 ");
      setStatus("error");
      return;
    }

    setFileName(excelFiles.length === 1 ? displayFilePath(excelFiles[0]) : `${excelFiles.length} 個 Excel 檔案 / 資料匣批次上傳`);
    setLastFiles(excelFiles);
    setStatus("reading");

    const parsedResults = [];
    const failedFiles = [];

    for (const file of excelFiles) {
      try {
        const workbook = await readWorkbookFromFile(file);
        if (!workbook.SheetNames?.length) {
          failedFiles.push({ 工作表名稱: displayFilePath(file), 排除原因: "Excel 沒有可讀取的工作表", 檢核方式: "請確認檔案是否損毀或是否為有效 Excel 活頁簿" });
          continue;
        }
        parsedResults.push(decorateParsedResult(parseWorkbook(workbook, activeSpecs), file));
      } catch (err) {
        failedFiles.push({ 工作表名稱: displayFilePath(file), 排除原因: "單一 Excel 檔案讀取或解析失敗", 檢核方式: err?.message || "未知錯誤，請人工檢查檔案格式" });
      }
    }

    const combined = buildCombinedResult(parsedResults, [...skippedFiles, ...failedFiles]);
    setResult(combined);

    if (!combined.rows.length) {
      setErrorMessage("找不到有效明細資料。請確認 Excel 第 7 列是否有「每片重量」，且第 8 到第 23 列是否有實際抽驗數值。");
      setStatus("error");
      return;
    }

    const firstMonth = uniqueValues(combined.rows, "月份")[0] || "";
    const firstWeek = uniqueValues(rowsByMonthWeek(combined.rows, firstMonth, ""), "週別")[0] || "";
    const firstVendor = vendorValues(rowsByMonthWeek(combined.rows, firstMonth, firstWeek))[0] || "";
    setSummaryFilters({ month: firstMonth, week: firstWeek, vendor: firstVendor });
    setFilters({ sheet:"", vendor:"", quality:"", overall:"" });
    setStatus("success");
  }

  function handleSelectedFiles(files) {
    processFiles(files);
  }

  function handleFileChange(event) {
    handleSelectedFiles(event.target.files);
    event.target.value = "";
  }

  function handleFolderChange(event) {
    handleSelectedFiles(event.target.files);
    event.target.value = "";
  }

  function handleDragEnter(event) {
    event.preventDefault();
    event.stopPropagation();
    setIsDragging(true);
  }

  function handleDragOver(event) {
    event.preventDefault();
    event.stopPropagation();
    setIsDragging(true);
  }

  function handleDragLeave(event) {
    event.preventDefault();
    event.stopPropagation();
    if (event.currentTarget.contains(event.relatedTarget)) return;
    setIsDragging(false);
  }

  async function handleDrop(event) {
    event.preventDefault();
    event.stopPropagation();
    setIsDragging(false);
    const droppedFiles = await filesFromDataTransfer(event.dataTransfer);
    handleSelectedFiles(droppedFiles);
  }
  function handleSpecChange(key, value) { setSpecs(prev => ({...prev, [key]: Number(value)})); }
  function reapplySpecs() {
    if (!lastFiles?.length) { setErrorMessage("請先上傳 Excel 檔案或資料匣，再套用規格判定。"); setStatus("error"); return; }
    processFiles(lastFiles, specs);
  }

  function downloadExcel() {
    try {
      if (!result?.rows?.length) { setErrorMessage("目前沒有可下載的清理後資料。"); return; }
      const wb = XLSX.utils.book_new();
      const bossAoa = buildBossSummaryAOA({ filters: summaryFilters, rows: summaryRows, bossRows });
      const bossSheet = XLSX.utils.aoa_to_sheet(bossAoa);
      bossSheet["!cols"] = [{ wch: 14 }, { wch: 30 }, { wch: 18 }, { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 12 }, { wch: 12 }];
      bossSheet["!autofilter"] = { ref: "A8:H22" };
      XLSX.utils.book_append_sheet(wb, bossSheet, "老闆彙總表_目前篩選");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(weeklyStats.map(({level, ...item}) => item)), "週別統計_目前篩選");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(monthlyStats), "月統計摘要");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildAllBossSummaryRows(result.rows)), "彙總檔格式_全部");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.rows, { header: STANDARD_COLUMNS }), "清理後明細資料");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(DIAGNOSIS_ROWS), "報表問題診斷表");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.kpis.map(({level, ...k}) => k)), "KPI摘要");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(abnormalRows, { header: STANDARD_COLUMNS }), "異常資料清單");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.excludedSheets), "排除工作表清單");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(PROCESS_RULES), "處理規則說明");
      saveWorkbookWithFallback(wb, "產品簡易抽驗紀錄表_清理結果_含週月統計與老闆彙總.xlsx");
    } catch (err) { setErrorMessage(`下載檔案失敗：${err?.message || "未知錯誤，請稍後再試。"}`); }
  }

  return <div className="app-shell">
    <header className="hero"><div><p className="eyebrow">Excel 報表二次分析工具</p><h1>產品簡易抽驗紀錄表清理與診斷工具</h1><p className="subtitle">先診斷，再自動化</p></div><div className="hero-badge"><FileSpreadsheet size={28}/><span>真實 xlsx 解析</span></div></header>
    <main className="main-grid">
      <section className="panel upload-panel"><div className="panel-title-row"><h2>1. 上傳 Excel 原始檔</h2><UploadCloud size={22}/></div><div className={`upload-box ${isDragging ? "drag-active" : ""}`} onDragEnter={handleDragEnter} onDragOver={handleDragOver} onDragLeave={handleDragLeave} onDrop={handleDrop}><span>{isDragging ? "放開即可上傳 Excel 檔案或資料匣" : "選擇檔案、選擇資料匣，或直接拖拉到這裡"}</span><small>支援 .xlsx、.xls、.xlsm；可多檔上傳，也可選擇整個資料匣批次解析</small><div className="upload-actions"><label className="upload-action-button"><input type="file" accept=".xlsx,.xls,.xlsm" multiple onChange={handleFileChange}/>選擇 Excel 檔案</label><label className="upload-action-button folder"><input type="file" webkitdirectory="true" directory="true" multiple onChange={handleFolderChange}/>選擇資料匣</label></div></div><div className="file-info"><p><strong>上傳來源：</strong>{fileName || "尚未上傳"}</p><p><strong>讀取狀態：</strong>{status === "idle" && "等待上傳"}{status === "reading" && "讀取與解析中"}{status === "success" && "處理成功"}{status === "error" && "需要檢查"}</p>{result?.parsedExcelFileCount > 1 && <p><strong>批次檔案數：</strong>{result.parsedExcelFileCount} 個 Excel 檔案已納入清理</p>}</div>{status === "success" && <div className="message success-message"><CheckCircle2 size={18}/>Excel 已成功解析，已產生清理後標準資料表、週別統計、月統計與老闆彙總表。</div>}{status === "success" && result?.rows?.length > 0 && <button className="primary-button download-now" onClick={downloadExcel}><Download size={18}/>立即下載含彙總報表 Excel</button>}{errorMessage && <div className="message error-message"><AlertTriangle size={18}/>{errorMessage}</div>}</section>
      <section className="panel rules-panel"><div className="panel-title-row"><h2>2. 預設檢核規則</h2><RefreshCw size={22}/></div><p className="hint">預設檢核規則，可人工調整；67..2 這類連續小數點手誤會自動修正為 67.2，並在資料品質備註標示待人工確認；廠商篩選保留母公司/子公司統計，若明細表廠商欄空白會嘗試參考彙整報告補入。老闆彙總表固定使用圖示規格：重量 60~70g、長度 12~16cm、寬度 3.5~4.5cm。</p><div className="spec-grid">{[["weightMin","每片重量下限"],["weightMax","每片重量上限"],["lengthMin","長度下限"],["lengthMax","長度上限"],["widthMin","寬度下限"],["widthMax","寬度上限"]].map(([key,label])=><label key={key}>{label}<input type="number" step={key.startsWith("width") ? "0.1" : "1"} value={specs[key]} onChange={e=>handleSpecChange(key,e.target.value)}/></label>)}</div><button className="secondary-button" onClick={reapplySpecs}>套用規格並重新判定</button></section>

      <section className="panel full-span priority-panel"><div className="panel-title-row"><div><h2>3. 週別統計、月統計與老闆彙總表</h2><p className="hint">先選月份、週別、統計廠商；括號內子公司會併入母公司統計，下方會帶出該週期與該統計廠商的產品有效日期，並依規格計算百分比。</p></div><button className="primary-button" onClick={downloadExcel}><Download size={18}/>下載含彙總 Excel</button></div>
        <div className="summary-toolbar">
          <label>月份<select value={summaryFilters.month} onChange={e=>setSummaryFilters(prev=>({...prev, month:e.target.value, week:"", vendor:""}))}><option value="">全部月份</option>{monthOptions.map(o=><option key={o}>{o}</option>)}</select></label>
          <label>週別<select value={summaryFilters.week} onChange={e=>setSummaryFilters(prev=>({...prev, week:e.target.value, vendor:""}))}><option value="">全部週別</option>{weekOptions.map(o=><option key={o}>{o}</option>)}</select></label>
          <label>統計廠商<select value={summaryFilters.vendor} onChange={e=>setSummaryFilters(prev=>({...prev, vendor:e.target.value}))}><option value="">全部統計廠商</option>{summaryVendorOptions.map(o=><option key={o}>{o}</option>)}</select></label>
        </div>
        {!result ? <div className="empty-state">上傳 Excel 後顯示週別統計、月統計與老闆彙總表</div> : <>
          <div className="report-header-grid">
            <div><span>目前月份</span><strong>{summaryFilters.month || "全部月份"}</strong></div>
            <div><span>目前週別</span><strong>{summaryFilters.week || "全部週別"}</strong></div>
            <div><span>目前統計廠商</span><strong>{summaryFilters.vendor || "全部統計廠商"}</strong></div>
            <div><span>原始廠商/子公司</span><strong>{uniqueJoined(summaryRows, "廠商")}</strong></div><div><span>有效日期</span><strong>{uniqueJoined(summaryRows, "產品效期")}</strong></div>
          </div>
          <h3 className="subsection-title">週別統計 KPI（可篩選）</h3>
          <div className="kpi-grid compact-kpis">{weeklyStats.map(kpi=><div className={`kpi-card ${kpi.level}`} key={kpi.指標}><p>{kpi.指標}</p><strong>{kpi.數值}</strong><small>{kpi.說明}</small></div>)}</div>
          <h3 className="subsection-title">老闆彙總檔格式</h3>
          <div className="manager-report"><div className="manager-title"><strong>{summaryFilters.week || "全部週別"}</strong><span>統計廠商：{summaryFilters.vendor || "全部統計廠商"}</span><span>原始廠商/子公司：{uniqueJoined(summaryRows, "廠商")}</span><span>有效日期：{uniqueJoined(summaryRows, "產品效期")}</span></div><BossSummaryTable rows={bossRows}/><div className="legend-row"><span className="legend low"></span>低於標準規格 <span className="legend high"></span>高於標準規格</div></div>
          <h3 className="subsection-title">月統計摘要</h3>
          <DataTable rows={monthlyStats} columns={["月份","明細筆數","週別數","廠商數","有效日期數","整體正常筆數","異常待確認筆數","整體合格率","重量6070比率","長度1216比率","寬度3545比率","其他問題率"]} emptyText="尚無月統計資料" />
        </>}
      </section>

      <section className="panel full-span"><h2>4. 資料品質風險提醒</h2><div className="risk-grid">{RISK_REMINDERS.map((item,i)=><div className="risk-item" key={item}><span>{i+1}</span><p>{item}</p></div>)}</div></section>
      <section className="panel full-span"><h2>5. 整體 KPI 儀表板</h2>{!result ? <div className="empty-state">上傳 Excel 後顯示 KPI</div> : <><div className="kpi-grid">{result.kpis.map(kpi=><div className={`kpi-card ${kpi.level}`} key={kpi.KPI名稱}><p>{kpi.KPI名稱}</p><strong>{kpi.KPI數值}</strong><small>{kpi.狀態說明}</small></div>)}</div></>}</section>
      <section className="panel full-span"><h2>6. 圖表區</h2><div className="charts-grid"><ChartCard title="各工作表明細筆數" data={result?.sheetChartData}><ResponsiveContainer width="100%" height="100%"><BarChart data={result?.sheetChartData || []}><CartesianGrid strokeDasharray="3 3"/><XAxis dataKey="名稱"/><YAxis allowDecimals={false}/><Tooltip/><Bar dataKey="筆數"/></BarChart></ResponsiveContainer></ChartCard><ChartCard title="各資料品質狀態筆數" data={result?.qualityChartData}><ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={result?.qualityChartData || []} dataKey="筆數" nameKey="名稱" outerRadius={90} label>{(result?.qualityChartData || []).map((e,i)=><Cell key={e.名稱} fill={CHART_COLORS[i % CHART_COLORS.length]}/>)}</Pie><Tooltip/><Legend/></PieChart></ResponsiveContainer></ChartCard><ChartCard title="各廠商異常筆數" data={result?.vendorAbnormalChartData}><ResponsiveContainer width="100%" height="100%"><BarChart data={result?.vendorAbnormalChartData || []}><CartesianGrid strokeDasharray="3 3"/><XAxis dataKey="名稱"/><YAxis allowDecimals={false}/><Tooltip/><Bar dataKey="筆數"/></BarChart></ResponsiveContainer></ChartCard><ChartCard title="整體判定分布" data={result?.overallChartData}><ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={result?.overallChartData || []} dataKey="筆數" nameKey="名稱" outerRadius={90} label>{(result?.overallChartData || []).map((e,i)=><Cell key={e.名稱} fill={CHART_COLORS[i % CHART_COLORS.length]}/>)}</Pie><Tooltip/><Legend/></PieChart></ResponsiveContainer></ChartCard></div></section>
      <section className="panel full-span"><h2>7. 報表問題診斷表</h2><DataTable rows={DIAGNOSIS_ROWS} columns={["問題編號","問題類型","問題描述","對分析的影響","建議處理方式","檢核方式","後續可用工具"]}/></section>
      <section className="panel full-span"><div className="panel-title-row"><h2>8. 清理後資料表預覽</h2><button className="primary-button" onClick={downloadExcel}><Download size={18}/>下載清理後 Excel</button></div><div className="filters">{[["sheet","來源工作表",sheetOptions],["vendor","統計廠商",vendorOptions],["quality","資料品質狀態",qualityOptions],["overall","整體判定",overallOptions]].map(([key,label,options])=><label key={key}>{label}<select value={filters[key]} onChange={e=>setFilters(prev=>({...prev,[key]:e.target.value}))}><option value="">全部</option>{options.map(o=><option key={o}>{o}</option>)}</select></label>)}</div><p className="hint">目前篩選後共 {filteredRows.length} 筆，預覽前 100 筆。</p><DataTable rows={filteredRows.slice(0,100)} columns={STANDARD_COLUMNS} emptyText="尚無清理後資料"/></section>
      <section className="panel full-span"><h2>9. 異常資料清單</h2><p className="hint">包含缺漏值、日期格式錯誤、數值格式錯誤、疑似異常值、樣本序號異常與待人工確認資料。</p><DataTable rows={abnormalRows.slice(0,200)} columns={STANDARD_COLUMNS} emptyText="目前沒有異常資料"/></section>
      <section className="panel full-span"><h2>10. 排除工作表清單</h2><DataTable rows={result?.excludedSheets || []} columns={["工作表名稱","排除原因","檢核方式"]} emptyText="尚無排除工作表"/></section>
      <section className="panel full-span"><h2>11. 人工檢核規則</h2><ol className="check-list"><li>清理後明細筆數 = 有效廠商區塊數 × 16。</li><li>每筆明細都應有來源工作表。</li><li>每筆明細都應有樣本序號。</li><li>每筆明細都應有廠商欄位。</li><li>每筆明細都應有週別欄位。</li><li>重量、長、寬若不是數值，需列入異常清單。</li><li>日期無法轉換需列入異常清單。</li><li>廠商空白需標示缺漏待確認。</li><li>整包重量空白需標示缺漏待確認。</li><li>抽驗者空白需標示缺漏待確認。</li><li>匯出 Excel 必須包含明細資料、週別統計、月統計與老闆彙總表。</li><li>清理後資料表應能直接用於 Excel 樞紐分析表。</li></ol></section>
    </main>
  </div>;
}

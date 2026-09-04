#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

const [payloadPath, outputPath, previewPath = ""] = process.argv.slice(2);
if (!payloadPath || !outputPath) {
  throw new Error("用法：invoice_excel.mjs <payload.json> <output.xlsx> [preview.png]");
}

const artifactEntry = process.env.INVOICE_ARTIFACT_TOOL_ENTRY;
if (!artifactEntry) {
  throw new Error("缺少 INVOICE_ARTIFACT_TOOL_ENTRY，无法加载电子表格运行时");
}

const { FileBlob, SpreadsheetFile, Workbook } = await import(pathToFileURL(artifactEntry).href);
const payload = JSON.parse(await fs.readFile(payloadPath, "utf8"));
const records = Array.isArray(payload.records) ? payload.records : [];
const sheetName = "发票登记";
const headers = [
  "序号",
  "发票日期",
  "支付日期",
  "收款姓名",
  "发票金额",
  "报销金额",
  "费用类型",
  "开票方",
  "发票编号",
  "发票下载时间",
  "备注",
  "是否存在报销凭证",
  "报销凭证",
  "凭证匹配状态",
  "凭证校验说明",
  "发票与实付差额",
  "校验状态",
  "",
];

function dateOnly(value) {
  if (!value) return null;
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(value));
  if (!match) return value;
  return new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
}

function dateTime(value) {
  if (!value) return null;
  const match = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?/.exec(String(value));
  if (match) {
    return new Date(
      Date.UTC(
        Number(match[1]),
        Number(match[2]) - 1,
        Number(match[3]),
        Number(match[4]),
        Number(match[5]),
        Number(match[6] ?? 0),
      ),
    );
  }
  const parsed = new Date(value);
  return Number.isNaN(parsed.valueOf()) ? value : parsed;
}

function amountValue(value) {
  if (value === null || value === undefined || value === "") return null;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : value;
}

async function readManualValues() {
  try {
    await fs.access(outputPath);
  } catch {
    return new Map();
  }
  let existing;
  try {
    existing = await SpreadsheetFile.importXlsx(await FileBlob.load(outputPath));
  } catch (error) {
    throw new Error(`现有发票登记表无法读取，为保护人工数据未覆盖：${error.message}`);
  }
  let sheet;
  try {
    sheet = existing.worksheets.getItem(sheetName);
  } catch {
    return new Map();
  }
  const used = sheet.getUsedRange(true);
  const values = used ? used.values : [];
  const headerIndex = values.findIndex((row) => row?.[0] === "序号" && row?.[8] === "发票编号");
  if (headerIndex < 0) return new Map();
  const header = values[headerIndex] ?? [];
  const indexOf = (name, fallback) => {
    const index = header.indexOf(name);
    return index >= 0 ? index : fallback;
  };
  const manualColumns = {
    paymentDate: indexOf("支付日期", 2),
    recipient: indexOf("收款姓名", 3),
    reimbursementAmount: indexOf("报销金额", 5),
    expenseType: indexOf("费用类型", 6),
    note: indexOf("备注", 10),
  };
  const validationColumn = header.indexOf("校验状态");
  const keyColumn = validationColumn >= 0 ? validationColumn + 1 : 12;
  const saved = new Map();
  for (const row of values.slice(headerIndex + 1)) {
    const key = String(row?.[keyColumn] ?? "").trim();
    if (!key) continue;
    saved.set(key, {
      paymentDate: row[manualColumns.paymentDate] ?? null,
      recipient: row[manualColumns.recipient] ?? "",
      reimbursementAmount: row[manualColumns.reimbursementAmount] ?? null,
      expenseType: row[manualColumns.expenseType] ?? "",
      note: row[manualColumns.note] ?? "",
    });
  }
  return saved;
}

const savedManual = await readManualValues();
records.sort((left, right) =>
  String(left.downloaded_at ?? "").localeCompare(String(right.downloaded_at ?? "")) ||
  String(left.invoice_date ?? "").localeCompare(String(right.invoice_date ?? "")) ||
  String(left.record_key ?? "").localeCompare(String(right.record_key ?? "")),
);

const workbook = Workbook.create();
const sheet = workbook.worksheets.add(sheetName);
sheet.showGridLines = false;
sheet.freezePanes.freezeRows(3);

sheet.getRange("A1:Q1").merge();
sheet.getRange("A1").values = [["发票登记表"]];
sheet.getRange("A2:Q2").merge();
sheet.getRange("A2").values = [["浅蓝色为自动填写字段；浅黄色为人工填写字段；凭证由 TextIn XParse 公有服务识别；敏感凭证请勿放入投放目录。"]];
sheet.getRange("A3:R3").values = [headers];

const rows = records.map((record, index) => {
  const manual = savedManual.get(String(record.record_key)) ?? null;
  return [
    index + 1,
    dateOnly(record.invoice_date),
    manual?.paymentDate || dateOnly(record.receipt_payment_date),
    manual ? manual.recipient : "",
    amountValue(record.invoice_amount),
    manual && manual.reimbursementAmount !== null && manual.reimbursementAmount !== ""
      ? manual.reimbursementAmount
      : amountValue(record.invoice_amount),
    manual ? manual.expenseType : "",
    record.seller ?? "",
    record.invoice_number ?? "",
    dateTime(record.downloaded_at),
    manual ? manual.note : "",
    record.has_receipt ?? "否",
    record.receipt_path ?? "",
    record.receipt_status ?? "未匹配",
    record.receipt_detail ?? "尚未找到对应报销凭证",
    amountValue(record.receipt_amount_difference),
    record.validation_status ?? "待确认",
    record.record_key ?? "",
  ];
});
if (rows.length) {
  sheet.getRange(`I4:I${rows.length + 3}`).setNumberFormat("@");
  sheet.getRange(`A4:R${rows.length + 3}`).values = rows;
}

const title = sheet.getRange("A1:Q1");
title.format = {
  fill: "#234E52",
  font: { bold: true, color: "#FFFFFF", size: 18 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
title.format.rowHeight = 34;

const note = sheet.getRange("A2:Q2");
note.format = {
  fill: "#E6FFFA",
  font: { color: "#285E61", italic: true, size: 10 },
  horizontalAlignment: "left",
  verticalAlignment: "center",
};
note.format.rowHeight = 24;

const header = sheet.getRange("A3:R3");
header.format = {
  fill: "#319795",
  font: { bold: true, color: "#FFFFFF", size: 11 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
  borders: { preset: "all", style: "thin", color: "#A0AEC0" },
};
header.format.rowHeight = 30;

const lastDataRow = Math.max(4, rows.length + 3);
if (rows.length) {
  const data = sheet.getRange(`A4:R${lastDataRow}`);
  data.format = {
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: "#D7E0E5" },
  };
  data.format.rowHeight = 32;
  sheet.getRange(`A4:A${lastDataRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`B4:C${lastDataRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`E4:F${lastDataRow}`).format.horizontalAlignment = "right";
  sheet.getRange(`G4:G${lastDataRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`I4:J${lastDataRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`H4:H${lastDataRow}`).format.wrapText = true;
  sheet.getRange(`K4:P${lastDataRow}`).format.wrapText = true;
  sheet.getRange(`B4:B${lastDataRow}`).setNumberFormat("yyyy/m/d");
  sheet.getRange(`C4:C${lastDataRow}`).setNumberFormat("yyyy/m/d");
  sheet.getRange(`E4:F${lastDataRow}`).setNumberFormat("¥#,##0.00");
  sheet.getRange(`P4:P${lastDataRow}`).setNumberFormat("¥#,##0.00");
  sheet.getRange(`I4:I${lastDataRow}`).setNumberFormat("@");
  sheet.getRange(`J4:J${lastDataRow}`).setNumberFormat("yyyy/m/d hh:mm");

  for (const column of ["A", "B", "E", "H", "I", "J", "L", "M", "N", "O", "P", "Q"]) {
    sheet.getRange(`${column}4:${column}${lastDataRow}`).format.fill = "#EAF4F8";
  }
  for (const column of ["C", "D", "F", "G", "K"]) {
    sheet.getRange(`${column}4:${column}${lastDataRow}`).format.fill = "#FFF7D6";
  }

  const table = sheet.tables.add(`A3:Q${lastDataRow}`, true, "InvoiceRegisterTable");
  table.style = "TableStyleLight9";
  table.showFilterButton = true;
}

const validationLastRow = Math.max(200, lastDataRow + 50);
sheet.getRange(`G4:G${validationLastRow}`).dataValidation = {
  rule: { type: "list", values: ["部门营销费用", "企业文化费用", "出差报销费用"] },
};

if (rows.length) {
  const expenseRange = sheet.getRange(`G4:G${lastDataRow}`);
  expenseRange.conditionalFormats.addCustom('=$G4="部门营销费用"', { fill: "#B2F5EA", font: { color: "#234E52" } });
  expenseRange.conditionalFormats.addCustom('=$G4="企业文化费用"', { fill: "#C6F6D5", font: { color: "#22543D" } });
  expenseRange.conditionalFormats.addCustom('=$G4="出差报销费用"', { fill: "#FEEBC8", font: { color: "#7B341E" } });
  const receiptStatusRange = sheet.getRange(`N4:N${lastDataRow}`);
  receiptStatusRange.conditionalFormats.addCustom('=OR($N4="未匹配",LEFT($N4,4)="匹配冲突",LEFT($N4,4)="识别失败",LEFT($N4,3)="待确认")', { fill: "#FED7D7", font: { color: "#9B2C2C", bold: true } });
  receiptStatusRange.conditionalFormats.addCustom('=OR($N4="已匹配",$N4="人工匹配")', { fill: "#C6F6D5", font: { color: "#22543D", bold: true } });
  const statusRange = sheet.getRange(`Q4:Q${lastDataRow}`);
  statusRange.conditionalFormats.addCustom('=LEFT($Q4,3)="待确认"', { fill: "#FED7D7", font: { color: "#9B2C2C", bold: true } });
  statusRange.conditionalFormats.addCustom('=$Q4="通过"', { fill: "#C6F6D5", font: { color: "#22543D", bold: true } });
}

const widths = {
  A: 8,
  B: 13,
  C: 13,
  D: 15,
  E: 14,
  F: 14,
  G: 18,
  H: 40,
  I: 24,
  J: 20,
  K: 42,
  L: 16,
  M: 48,
  N: 18,
  O: 52,
  P: 18,
  Q: 30,
  R: 0.1,
};
for (const [column, width] of Object.entries(widths)) {
  sheet.getRange(`${column}:${column}`).format.columnWidth = width;
}
sheet.getRange(`R1:R${lastDataRow}`).format.font = { color: "#FFFFFF", size: 1 };

const outputDir = path.dirname(outputPath);
await fs.mkdir(outputDir, { recursive: true });
const temporaryOutput = `${outputPath}.tmp-${process.pid}-${Date.now()}.xlsx`;
try {
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(temporaryOutput);
  await fs.rename(temporaryOutput, outputPath);
} finally {
  await fs.rm(temporaryOutput, { force: true });
}

if (previewPath) {
  await fs.mkdir(path.dirname(previewPath), { recursive: true });
  const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));
}

let inspection;
let formulaErrors;
try {
  inspection = await workbook.inspect({
    kind: "region",
    sheetId: sheetName,
    range: `A1:Q${lastDataRow}`,
    maxChars: 5000,
    tableMaxRows: 12,
    tableMaxCols: 17,
    tableMaxCellChars: 100,
  });
  formulaErrors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 300 },
    summary: "final formula error scan",
  });
} finally {
  await fs.rm(`${temporaryOutput}.inspect.ndjson`, { force: true });
}
console.log(JSON.stringify({
  output: outputPath,
  rows: records.length,
  inspection: inspection.ndjson ?? inspection,
  formulaErrors: formulaErrors.ndjson ?? formulaErrors,
}, null, 2));

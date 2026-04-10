import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import ExcelJS from "exceljs";

const worksheetName = "Progress";
const expectedHeaders = [
  "RowID",
  "KPI番号",
  "担当者名",
  "登録日",
  "更新日",
  "ステータス",
  "社外関係者",
  "社内関連部署",
  "顧客名",
  "内容",
  "NextAction",
  "更新者",
  "Version"
];

function toIsoString(value) {
  if (!value) {
    return "";
  }

  if (value instanceof Date) {
    return value.toISOString();
  }

  return new Date(value).toISOString();
}

function normalizeRecord(record) {
  return {
    id: record.RowID,
    kpiNumber: record.KPI番号 ?? "",
    assignee: record.担当者名 ?? "",
    createdAt: record.登録日 ?? "",
    updatedAt: record.更新日 ?? "",
    status: record.ステータス ?? "",
    externalStakeholders: record.社外関係者 ?? "",
    internalDepartments: record.社内関連部署 ?? "",
    customer: record.顧客名 ?? "",
    content: record.内容 ?? "",
    nextAction: record.NextAction ?? "",
    updatedBy: record.更新者 ?? "",
    version: Number(record.Version ?? 1)
  };
}

function toWorksheetRecord(payload, currentRecord) {
  const now = new Date().toISOString();

  return {
    RowID: currentRecord?.RowID ?? crypto.randomUUID(),
    KPI番号: payload.kpiNumber,
    担当者名: payload.assignee,
    登録日: currentRecord?.登録日 ?? now,
    更新日: now,
    ステータス: payload.status,
    社外関係者: payload.externalStakeholders ?? "",
    社内関連部署: payload.internalDepartments ?? "",
    顧客名: payload.customer ?? "",
    内容: payload.content ?? "",
    NextAction: payload.nextAction ?? "",
    更新者: payload.updatedBy ?? payload.assignee,
    Version: currentRecord ? Number(currentRecord.Version ?? 1) + 1 : 1
  };
}

async function loadWorkbook(excelFilePath) {
  const workbook = new ExcelJS.Workbook();
  let fileExists = true;

  try {
    await fs.access(excelFilePath);
  } catch {
    fileExists = false;
  }

  try {
    if (fileExists) {
      await workbook.xlsx.readFile(excelFilePath);
    } else {
      await fs.mkdir(path.dirname(excelFilePath), { recursive: true });
      const worksheet = workbook.addWorksheet(worksheetName);
      worksheet.addRow(expectedHeaders);
      worksheet.getRow(1).font = { bold: true };
      worksheet.views = [{ state: "frozen", ySplit: 1 }];
      await workbook.xlsx.writeFile(excelFilePath);
    }
  } catch (error) {
    if (error.code !== "ENOENT" && !String(error.message).includes("File not found")) {
      throw error;
    }

    await fs.mkdir(path.dirname(excelFilePath), { recursive: true });
    const worksheet = workbook.addWorksheet(worksheetName);
    worksheet.addRow(expectedHeaders);
    worksheet.getRow(1).font = { bold: true };
    worksheet.views = [{ state: "frozen", ySplit: 1 }];
    await workbook.xlsx.writeFile(excelFilePath);
  }

  let worksheet = workbook.getWorksheet(worksheetName);

  if (!worksheet) {
    worksheet = workbook.worksheets[0];
  }

  if (!worksheet) {
    worksheet = workbook.addWorksheet(worksheetName);
    worksheet.addRow(expectedHeaders);
  }

  const headerValues = worksheet.getRow(1).values.slice(1);
  const headersMatch = expectedHeaders.every((header, index) => headerValues[index] === header);

  if (!headersMatch) {
    throw new Error("Excel ファイルのヘッダーが想定と一致しません。テンプレートを利用してください。");
  }

  return { workbook, worksheet };
}

function getRecordsFromWorksheet(worksheet) {
  const records = [];

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) {
      return;
    }

    const values = row.values.slice(1);
    const isEmpty = values.every((value) => value === null || value === undefined || value === "");

    if (isEmpty) {
      return;
    }

    const record = expectedHeaders.reduce((result, header, index) => {
      const cellValue = row.getCell(index + 1).value;
      result[header] = typeof cellValue === "object" && cellValue?.text ? cellValue.text : cellValue;
      return result;
    }, {});

    record.登録日 = toIsoString(record.登録日);
    record.更新日 = toIsoString(record.更新日);
    records.push(record);
  });

  return records;
}

async function saveWorkbook(workbook, excelFilePath) {
  await fs.mkdir(path.dirname(excelFilePath), { recursive: true });
  await workbook.xlsx.writeFile(excelFilePath);
}

function assertPayload(payload) {
  const requiredFields = [
    ["assignee", "担当者名"],
    ["status", "ステータス"],
    ["content", "内容"]
  ];

  const missingField = requiredFields.find(([key]) => !String(payload[key] ?? "").trim());

  if (missingField) {
    throw new Error(`${missingField[1]} は必須です。`);
  }
}

export async function ensureWorkbook(excelFilePath) {
  const { workbook } = await loadWorkbook(excelFilePath);
  await saveWorkbook(workbook, excelFilePath);
}

export async function listProgress(excelFilePath) {
  const { worksheet } = await loadWorkbook(excelFilePath);
  return getRecordsFromWorksheet(worksheet).map(normalizeRecord).sort((left, right) => {
    return right.updatedAt.localeCompare(left.updatedAt);
  });
}

export async function getSummary(excelFilePath) {
  const items = await listProgress(excelFilePath);

  return items.reduce((summary, item) => {
    summary.total += 1;
    summary.byStatus[item.status] = (summary.byStatus[item.status] ?? 0) + 1;
    return summary;
  }, { total: 0, byStatus: {} });
}

export async function createProgress(excelFilePath, payload) {
  assertPayload(payload);
  const { workbook, worksheet } = await loadWorkbook(excelFilePath);
  const record = toWorksheetRecord(payload);

  worksheet.addRow(expectedHeaders.map((header) => record[header]));
  await saveWorkbook(workbook, excelFilePath);

  return normalizeRecord(record);
}

export async function updateProgress(excelFilePath, id, payload) {
  assertPayload(payload);
  const { workbook, worksheet } = await loadWorkbook(excelFilePath);
  const records = getRecordsFromWorksheet(worksheet);
  const targetIndex = records.findIndex((record) => record.RowID === id);

  if (targetIndex === -1) {
    throw new Error("対象データが見つかりません。");
  }

  const currentRecord = records[targetIndex];
  const expectedVersion = Number(payload.version ?? 0);
  const currentVersion = Number(currentRecord.Version ?? 1);

  if (expectedVersion !== currentVersion) {
    const conflictError = new Error("他の更新が先に保存されています。最新データを再読み込みしてください。");
    conflictError.code = "VERSION_CONFLICT";
    throw conflictError;
  }

  const nextRecord = toWorksheetRecord(payload, currentRecord);
  const rowNumber = targetIndex + 2;

  expectedHeaders.forEach((header, index) => {
    worksheet.getRow(rowNumber).getCell(index + 1).value = nextRecord[header];
  });

  await saveWorkbook(workbook, excelFilePath);

  return normalizeRecord(nextRecord);
}

export { expectedHeaders };

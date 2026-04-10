import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import ExcelJS from '../backend/node_modules/exceljs/excel.js';

const SHEET_NAME = 'Progress';
const HEADERS = [
  'RowID',
  'KPI番号',
  'カテゴリー',
  '担当者名',
  '登録日',
  '更新日',
  'ステータス',
  'ランク',
  'ディールサイズ',
  '社外関係者',
  '社内関連部署',
  '顧客名',
  '内容',
  'NextAction',
  '報告メモ',
  '更新者',
  'Version'
];

const settingsPath = path.join(os.homedir(), 'AppData', 'Roaming', 'ProgressTrackerPoc', 'settings.json');
const defaultExcelPath = path.join(os.homedir(), 'Documents', 'ProgressTrackerPoc', 'progress.xlsx');

function isoDaysAgo(days) {
  return new Date(Date.now() - days * 24 * 60 * 60 * 1000).toISOString();
}

function buildSamples() {
  const categories = ['営業','マーケティング','開発','カスタマーサクセス','経理'];
  const names = ['佐藤','鈴木','高橋','伊藤','中村','小林','加藤','吉田','山田','木村','斎藤','松本','井上','橋本','清水'];
  const statuses = ['進捗中','計画中','保留','クローズ','失注'];
  const dealSizes = ['','50万円','1200万円','300万円','2500万円','800万円','150万円'];
  const external = ['田中部長、ABC商事購買','広告代理店X','DEF工業 情報システム部','制作会社Y','GHI物流 役員会','顧客J','パートナーK'];
  const internal = ['営業企画、法務','広報、デザイン','導入支援','インサイドセールス','カスタマーサクセス、経理','開発','経理'];
  const customers = ['ABC商事','DEF工業','GHI物流','JKLホールディングス','MNOシステム','顧客X','顧客Y'];
  const contents = ['提案内容の詰め','LP制作','予算確認待ち','ABテスト実施','契約締結完了','仕様調整中','オンボーディング準備中'];
  const nexts = ['来週に提示','スケジュール調整','再提案を準備','素材制作依頼','稟議確認','キックオフ設定','詳細見積作成'];
  const notes = ['最終調整中','要承認','先方調整中','初動良好','受注済','懸念あり','フォロー中'];

  const samples = [];
  for (let i = 1; i <= 25; i++) {
    const idx = i - 1;
    const id = `sample-${String(i).padStart(3, '0')}`;
    const kpi = String(i);
    const cat = categories[idx % categories.length];
    const name = names[idx % names.length];
    const status = statuses[idx % statuses.length];
    const rank = (i % 7 === 0) ? 'S' : (i % 3 === 0) ? 'B' : (i % 4 === 0) ? 'A' : '';
    const deal = rank ? dealSizes[i % dealSizes.length] : '';
    const ext = external[idx % external.length];
    const intd = internal[idx % internal.length];
    const cust = customers[idx % customers.length];
    const content = contents[idx % contents.length];
    const next = nexts[idx % nexts.length];
    const note = notes[idx % notes.length];
    const registerDaysAgo = 30 - (i % 28);
    const updateDaysAgo = Math.max(0, registerDaysAgo - (i % 10));

    samples.push({
      RowID: id,
      'KPI番号': kpi,
      'カテゴリー': cat,
      '担当者名': name,
      '登録日': isoDaysAgo(registerDaysAgo),
      '更新日': isoDaysAgo(updateDaysAgo),
      'ステータス': status,
      'ランク': rank,
      'ディールサイズ': deal,
      '社外関係者': ext,
      '社内関連部署': intd,
      '顧客名': cust,
      '内容': content,
      NextAction: next,
      報告メモ: note,
      更新者: 'サンプル投入',
      Version: 1
    });
  }

  return samples;
}

async function resolveExcelPath() {
  try {
    const raw = await fs.readFile(settingsPath, 'utf8');
    const settings = JSON.parse(raw);
    return settings.excelFilePath || defaultExcelPath;
  } catch {
    return defaultExcelPath;
  }
}

async function loadWorkbook(excelPath) {
  const workbook = new ExcelJS.Workbook();
  try {
    await fs.access(excelPath);
    await workbook.xlsx.readFile(excelPath);
  } catch {
    await fs.mkdir(path.dirname(excelPath), { recursive: true });
    const worksheet = workbook.addWorksheet(SHEET_NAME);
    worksheet.addRow(HEADERS);
    worksheet.getRow(1).font = { bold: true };
    worksheet.views = [{ state: 'frozen', ySplit: 1 }];
    await workbook.xlsx.writeFile(excelPath);
  }

  let worksheet = workbook.getWorksheet(SHEET_NAME) || workbook.worksheets[0];
  if (!worksheet) {
    worksheet = workbook.addWorksheet(SHEET_NAME);
    worksheet.addRow(HEADERS);
  }

  const row1 = worksheet.getRow(1);
  const actualHeaders = HEADERS.map((_, index) => row1.getCell(index + 1).value);
  const hasHeaders = HEADERS.every((header, index) => actualHeaders[index] === header);
  if (!hasHeaders) {
    worksheet.spliceRows(1, worksheet.rowCount || 1);
    worksheet.addRow(HEADERS);
  }

  return { workbook, worksheet };
}

function rowToRecord(row) {
  const record = {};
  HEADERS.forEach((header, index) => {
    const value = row.getCell(index + 1).value;
    record[header] = typeof value === 'object' && value && 'text' in value ? value.text : value ?? '';
  });
  return record;
}

async function main() {
  const excelPath = await resolveExcelPath();
  const { workbook, worksheet } = await loadWorkbook(excelPath);
  const existing = new Map();

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) {
      return;
    }
    const record = rowToRecord(row);
    if (String(record.RowID || '').trim()) {
      existing.set(String(record.RowID), { rowNumber, record });
    }
  });

  let inserted = 0;
  let updated = 0;
  for (const sample of buildSamples()) {
    const current = existing.get(sample.RowID);
    const values = HEADERS.map((header) => sample[header]);
    if (current) {
      const row = worksheet.getRow(current.rowNumber);
      values.forEach((value, index) => {
        row.getCell(index + 1).value = value;
      });
      row.commit();
      updated += 1;
    } else {
      worksheet.addRow(values);
      inserted += 1;
    }
  }

  await fs.mkdir(path.dirname(excelPath), { recursive: true });
  await workbook.xlsx.writeFile(excelPath);

  console.log(`seeded_path=${excelPath}`);
  console.log(`inserted=${inserted}`);
  console.log(`updated=${updated}`);
  console.log(`total=${worksheet.rowCount - 1}`);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

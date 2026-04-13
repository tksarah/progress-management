import ExcelJS from 'exceljs';
import fs from 'node:fs/promises';
import path from 'node:path';

const OUT_PATH = path.join(process.cwd(), 'progress.xlsx');
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
  'リード元',
  '社外関係者',
  '社内関連部署',
  '顧客名',
  '内容',
  'NextAction',
  '報告メモ',
  '更新者',
  'Version'
];

function isoDaysAgo(days) {
  return new Date(Date.now() - days * 24 * 60 * 60 * 1000).toISOString();
}

function buildSamples() {
  const categories = ['営業','マーケティング','開発','カスタマーサクセス','経理'];
  const names = ['佐藤','鈴木','高橋','伊藤','中村','小林','加藤','吉田','山田','木村','斎藤','松本','井上','橋本','清水'];
  const statuses = ['進捗中','計画中','保留','クローズ','失注'];
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
    const deal = ((i * 31) % 491) + 10; // deterministic
    const ext = external[idx % external.length];
    const intd = internal[idx % internal.length];
    const cust = customers[idx % customers.length];
    const content = contents[idx % contents.length];
    const next = nexts[idx % nexts.length];
    const note = notes[idx % notes.length];
    const registerDaysAgo = 30 - (i % 28);
    const updateDaysAgo = Math.max(0, registerDaysAgo - (i % 10));

    samples.push([
      id,
      kpi,
      cat,
      name,
      isoDaysAgo(registerDaysAgo),
      isoDaysAgo(updateDaysAgo),
      status,
      rank,
      String(deal),
      '',
      ext,
      intd,
      cust,
      content,
      next,
      note,
      'サンプル投入',
      '1'
    ]);
  }

  return samples;
}

async function main() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(SHEET_NAME);

  worksheet.addRow(HEADERS);
  worksheet.getRow(1).font = { bold: true };

  const samples = buildSamples();
  for (const row of samples) {
    worksheet.addRow(row);
  }

  await workbook.xlsx.writeFile(OUT_PATH);
  console.log(`Wrote ${OUT_PATH}`);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});

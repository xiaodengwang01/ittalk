import fs from 'node:fs/promises';
import { constants } from 'node:fs';
import path from 'node:path';
import assert from 'node:assert/strict';
import os from 'node:os';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const bundledModules = process.env.TAILKI_ARTIFACT_MODULES || path.join(os.homedir(), '.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules');
const moduleResolver = createRequire(path.join(bundledModules, 'tailki-loader.cjs'));
const { Workbook, SpreadsheetFile } = await import(pathToFileURL(moduleResolver.resolve('@oai/artifact-tool')).href);

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const workDir = await fs.mkdtemp(path.join(os.tmpdir(), 'tailki-preview-'));
const destination = path.resolve(process.argv[2] || path.join(projectRoot, 'Tailki_模拟.xlsx'));
try {
  await fs.access(destination);
  throw new Error('目标文件已存在，停止以避免覆盖。');
} catch (error) {
  if (error.code !== 'ENOENT') throw error;
}
const workbook = Workbook.create();
const font = 'Arial';
const headerColor = '#263D57';

const specs = [
  {
    name: '达人名单',
    tableName: 'TailkiContacts',
    endColumn: 'K',
    headers: ['达人ID', '达人', '联系人', '邮箱', '联系状态', '最近跟进日期', '人工报价', '策略', '备注', '邮件称呼', '允许本批次发送'],
    widths: [14, 20, 21, 40, 14, 22, 16, 24, 39, 17, 24],
    rows: [
      ['SIM-001', '模拟达人 Alpha', 'Alpha', 'alpha@tailki-demo.invalid', '未联系', null, null, null, '模拟数据：直接联系达人。', 'Alpha', '否'],
      ['SIM-002', '模拟达人 Beta', 'Morgan（经纪人）', 'morgan@tailki-demo.invalid', '未联系', null, null, null, '模拟数据：联系人为经纪人。', 'Morgan', '否'],
      ['SIM-003', '模拟达人 Gamma', 'Gamma', 'gamma@tailki-demo.invalid', '未联系', null, null, null, '模拟数据：直接联系达人。', 'Gamma', '否'],
    ],
  },
  {
    name: '发送记录',
    tableName: 'TailkiSentLog',
    endColumn: 'H',
    headers: ['批次ID', '达人ID', '达人', '邮箱', '使用模板', '实际发送时间', '发送结果', '邮件唯一标识'],
    widths: [20, 14, 20, 40, 27, 26, 22, 40],
    rows: [],
  },
  {
    name: '邮件往来',
    tableName: 'TailkiMailLog',
    endColumn: 'G',
    headers: ['邮件唯一标识', '对应达人', '发件人', '主题', '收件时间', '回复正文', '归属状态'],
    widths: [38, 20, 40, 40, 26, 64, 20],
    rows: [],
  },
];

for (const spec of specs) {
  const sheet = workbook.worksheets.add(spec.name);
  sheet.showGridLines = false;
  const lastRow = spec.rows.length + 1;
  sheet.getRange(`A1:${spec.endColumn}${lastRow}`).values = [spec.headers, ...spec.rows];
  const table = sheet.tables.add(`A1:${spec.endColumn}${lastRow}`, true, spec.tableName);
  table.showFilterButton = true;
  table.showTotals = false;

  const view = sheet.getRange(`A1:${spec.endColumn}8`);
  view.format.font = { name: font, size: 11, color: '#243042' };
  view.format.verticalAlignment = 'center';
  view.format.rowHeight = 30;
  view.format.wrapText = false;

  const header = sheet.getRange(`A1:${spec.endColumn}1`);
  header.format.fill = headerColor;
  header.format.font = { name: font, size: 11, bold: true, color: '#FFFFFF' };
  header.format.horizontalAlignment = 'center';
  header.format.rowHeight = 36;
  header.format.borders = { insideVertical: { style: 'thin', color: '#FFFFFF' } };

  for (let index = 0; index < spec.widths.length; index++) {
    sheet.getRangeByIndexes(0, index, 8, 1).format.columnWidth = spec.widths[index];
  }
  sheet.freezePanes.freezeRows(1);

  if (spec.name === '达人名单') {
    sheet.tabColor = headerColor;
    sheet.getRange('A2:K4').format.rowHeight = 40;
    sheet.getRange('J2:K4').format.fill = '#FFF5DE';
    sheet.getRange('K2:K4').format.horizontalAlignment = 'center';
    sheet.getRange('K2:K1000').dataValidation = {
      rule: { type: 'list', values: ['是', '否'] },
    };
    sheet.getRange('F2:F1000').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('G2:G1000').setNumberFormat('#,##0.00');
  }
  if (spec.name === '发送记录') {
    sheet.getRange('F2:F1000').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
  if (spec.name === '邮件往来') {
    sheet.getRange('E2:E1000').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange('F2:F1000').format.wrapText = true;
    sheet.getRange('F2:F1000').format.verticalAlignment = 'top';
  }
}

workbook.recalculate();

for (const spec of specs) {
  const sheet = workbook.worksheets.getItem(spec.name);
  assert.deepEqual(sheet.getRange(`A1:${spec.endColumn}1`).values[0], spec.headers);
  if (spec.rows.length) {
    const rows = sheet.getRange(`A2:${spec.endColumn}4`).values;
    assert.equal(rows.length, 3);
    assert.ok(rows.every(row => row[10] === '否' && row[3].endsWith('.invalid')));
    assert.ok(rows.every(row => row.slice(5, 8).every(value => value == null || value === '')));
  } else {
    assert.ok(sheet.getRange(`A2:${spec.endColumn}8`).values.flat().every(value => value == null || value === ''));
  }
  const inspection = await workbook.inspect({
    kind: 'table',
    range: `${spec.name}!A1:${spec.endColumn}${spec.rows.length + 1}`,
    include: 'values,formulas',
    tableMaxRows: 4,
    tableMaxCols: spec.headers.length,
    maxChars: 1800,
  });
  console.log(JSON.stringify({ sheet: spec.name, inspection: inspection.ndjson }));
  const preview = await workbook.render({
    sheetName: spec.name,
    range: `A1:${spec.endColumn}${spec.rows.length ? 5 : 4}`,
    scale: 1.5,
    format: 'png',
  });
  await fs.writeFile(path.join(workDir, `${spec.name}.png`), new Uint8Array(await preview.arrayBuffer()));
}

const errors = await workbook.inspect({
  kind: 'match',
  searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',
  options: { useRegex: true, maxResults: 20 },
  maxChars: 1500,
  summary: 'formula error scan',
});
console.log(JSON.stringify({ formulaErrorScan: errors.ndjson }));

const exported = await SpreadsheetFile.exportXlsx(workbook);
const staged = path.join(workDir, 'Tailki_模拟.xlsx');
await exported.save(staged);
await fs.copyFile(staged, destination, constants.COPYFILE_EXCL);
console.log(JSON.stringify({ saved: destination, bytes: (await fs.stat(destination)).size }));

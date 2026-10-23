// Author only the requested cells. daily_workbook.py preserves native XLSX parts.
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import {createRequire} from 'node:module';
import {pathToFileURL} from 'node:url';

const modules = process.env.TAILKI_ARTIFACT_MODULES ||
  path.join(os.homedir(), '.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules');
const resolver = createRequire(path.join(modules, 'tailki-loader.cjs'));
const {Workbook, SpreadsheetFile} = await import(pathToFileURL(resolver.resolve('@oai/artifact-tool')).href);
const [input, output, fixture] = process.argv.slice(2);
const cells = JSON.parse(await fs.readFile(input, 'utf8'));
const book = Workbook.create();
const sheet = book.worksheets.add('达人开发主表');
if (fixture === '--fixture') {
  book.worksheets.add('长期合作意愿').getRange('A1').values = [['达人ID']];
  sheet.getRange('A1:AW5').format.font = {name: 'Arial', size: 10};
  sheet.getRange('A1:AW1').format.fill = '#705CBC';
  sheet.getRange('A1:AW1').format.font.color = '#FFFFFF';
  sheet.getRange('A1:AW5').format.columnWidth = 20;
  sheet.getRange('A1:AW5').format.rowHeight = 28;
  sheet.getRange('G2:G5').setNumberFormat('yyyy-mm-dd');
}
for (const [address, spec] of Object.entries(cells)) {
  const range = sheet.getRange(address);
  if (spec.formula) range.formulas = [[spec.formula]];
  else {
    let value = spec.value;
    if (spec.date) {
      value = new Date(`${value}T00:00:00Z`);
      range.setNumberFormat('yyyy-mm-dd');
    } else if (typeof value === 'string' && value.startsWith('=')) value = `'${value}`;
    range.values = [[value]];
  }
}
book.recalculate();
for (const [address, spec] of Object.entries(cells)) {
  const range = sheet.getRange(address);
  // Unscored research rows intentionally display an empty aggregate score.
  if (spec.formula && typeof range.values[0][0] !== 'number' && range.values[0][0] !== '')
    throw Error(`新增公式未得到数值或空值：${address}`);
}
await (await SpreadsheetFile.exportXlsx(book)).save(output);

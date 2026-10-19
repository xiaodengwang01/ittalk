import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';

const bundledModules = process.env.TAILKI_ARTIFACT_MODULES || path.join(os.homedir(), '.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules');
const resolver = createRequire(path.join(bundledModules, 'tailki-loader.cjs'));
const { Workbook, SpreadsheetFile, FileBlob } = await import(pathToFileURL(resolver.resolve('@oai/artifact-tool')).href);
const [input, output, previewDir] = process.argv.slice(2);
if (!input || !output) throw Error('需要输入数据与输出文件。');
if (input.endsWith('.xlsx')) {
  const book = await SpreadsheetFile.importXlsx(await FileBlob.load(input));
  const [sheetName,range]=previewDir?previewDir.split('!'):['达人开发主表','D1:I12'];
  const blob = await book.render({sheetName,range,scale:1,format:'png'});
  await fs.writeFile(output,new Uint8Array(await blob.arrayBuffer()));
  process.exit(0);
}
const spec = JSON.parse(await fs.readFile(input,'utf8'));
for (const tab of spec.sheets) {
  if (tab.headers.includes('历史截图')) throw Error('纯文本工作簿不能包含历史截图列。');
  if (tab.rows.some(row => (row._images || []).length)) throw Error('纯文本工作簿不能写入图片，请先转录正文。');
}
const wb = Workbook.create();
const purple = '#705CBC';
const col = n => { let s=''; for(n++;n;n=Math.floor((n-1)/26)) s=String.fromCharCode(65+(n-1)%26)+s; return s; };
const dates = new Set(['首次联系日期','历史跟进日期','计划跟进日期','邮件时间','最近收信时间','最近发信时间','最近抓取时间','确认时间','草稿生成时间','发布日期']);
const numbers = new Set(['粉丝数','近10条中位播放','近10条最高播放','内容能力(1-5)','宠物人群匹配度(1-5)','播放能力(1-5)','联系方式完整度(1-5)','综合评分','首次报价(USD)','我们的报价(USD)2']);
const choices = {
  '首联内容分类':['宠物剧情类','宠物用品类','非宠物垂类','待核实'],
  '分类核实状态':['初判','已确认','待核实'],
  '平台':['TikTok','Instagram','YouTube','Other'],
  '联系状态':['待联系','已联系','跟进中','已回复','拒绝','合作中','已完成','勿联系','待核实'],
  '下一步动作':['补资料','核实分类','首联草稿','回复草稿','跟进草稿','等待对方','暂停联系'],
  '方向':['收信','发信','待核实'],
  '处理状态':['待处理','待回复','已归档','待人工归属','历史文本待核实','已发送'],
  '正文完整性':['完整','部分','待核实','缺失'],
  '摘要状态':['有效','待更新'],
  '沟通摘要状态':['有效','待更新'],
  '实际发送结果':['成功'],
  '邮件用途':['收信','回复'],
  '宠物类型':['Dog','Cat','Dog+Cat','Cat+Dog','Other'],
  '是否低粉高播放':['是','否'], '是否带货类似产品':['是','否'],
  '是否寄样':['否','准备寄出','已寄出','已签收'],
  '到货情况':['未寄出','运输中','已签收','异常'],
  '是否发布':['否','待发布','是'],'是否授权素材':['否','待确认','是'],
};
const widths = {
  '沟通摘要':58, '沟通摘要状态':16, '邮件摘要':62, '邮件正文':90,
  '邮件记录ID':18, '达人ID':17, '达人昵称':24, '方向':10,
  '邮件时间':21, '主题':52, '处理状态':22, '正文完整性':16, '摘要状态':15,
  '发件人':32, '收件人':32, '原邮件定位':40, '邮件原始时间':24,
  '核实说明':62, '记录来源':24,
};
for (const tab of spec.sheets) {
  const sheet = wb.worksheets.add(tab.name);
  sheet.showGridLines = false;
  sheet.tabColor = tab.name === '达人开发主表' ? purple : '#9687C8';
  const headers = tab.headers;
  const last = col(headers.length-1);
  const lastRow = Math.max(2,tab.rows.length+1);
  const matrix = tab.rows.map(row => headers.map(h => {
    let value = row[h] ?? null;
    if (dates.has(h) && value && /^\d{4}-\d{2}-\d{2}/.test(String(value))) {
      const normalized = /^\d{4}-\d{2}-\d{2}$/.test(value) ? `${value}T00:00:00Z` : String(value).replace(' ','T')+'Z';
      return new Date(normalized);
    }
    // The value setter consumes this escape, keeping formula-like source text literal.
    if (typeof value === 'string' && value.startsWith('=')) return `'${value}`;
    return value;
  }));
  sheet.getRange(`A1:${last}${lastRow}`).values = [headers,...matrix,...(matrix.length?[]:[headers.map(()=>null)])];
  const table = sheet.tables.add(`A1:${last}${lastRow}`,true,tab.name === '达人开发主表'?'TailkiCreators':'TailkiMessages');
  table.style = 'TableStyleMedium2';
  table.showFilterButton = true;
  const area = sheet.getRange(`A1:${last}${lastRow}`);
  area.format.font = {name:'Arial',size:10,color:'#243042'};
  area.format.verticalAlignment = 'center';
  area.format.rowHeight = tab.name==='达人开发主表'?64:50;
  area.format.wrapText = false;
  const header = sheet.getRange(`A1:${last}1`);
  header.format.fill = purple;
  header.format.font = {name:'Arial',size:10,bold:true,color:'#FFFFFF'};
  header.format.horizontalAlignment = 'center';
  header.format.wrapText = true;
  header.format.rowHeight = 36;
  for (let i=0;i<headers.length;i++) {
    const h=headers[i];
    const range=sheet.getRange(`${col(i)}1:${col(i)}${lastRow}`);
    range.format.columnWidth = tab.widths?.[h] || widths[h] || 20;
    const body=sheet.getRange(`${col(i)}2:${col(i)}${Math.max(lastRow,500)}`);
    if (dates.has(h)) body.setNumberFormat(h==='计划跟进日期'||h==='首次联系日期'||h==='历史跟进日期'||h==='发布日期'?'yyyy-mm-dd':'yyyy-mm-dd hh:mm');
    if (dates.has(h)) for(let r=0;r<tab.rows.length;r++) {
      if (/^\d{4}-\d{2}-\d{2}$/.test(String(tab.rows[r][h]||''))) sheet.getRange(`${col(i)}${r+2}`).setNumberFormat('yyyy-mm-dd');
    }
    if (numbers.has(h)) body.setNumberFormat(h==='综合评分'?'0.00':'#,##0');
    if (choices[h]) body.dataValidation={rule:{type:'list',values:choices[h]}};
    if (['资料核实说明','核实说明','沟通摘要','邮件摘要','邮件正文','主题'].includes(h)) {
      range.format.wrapText = true;
      range.format.verticalAlignment = 'top';
    }
    if (h==='下一步动作'||h==='计划跟进日期'||h==='分类核实状态') body.format.fill='#FFF6DE';
    if (h==='分类核实状态'||h==='下一步动作') {
      body.conditionalFormats.add('containsText',{text:'待核实',format:{fill:'#FFF1D6',font:{color:'#8A5700'}}});
      if(h==='下一步动作') body.conditionalFormats.add('containsText',{text:'暂停',format:{fill:'#FCE8E6',font:{color:'#9D2823'}}});
    }
    if (h==='摘要状态'||h==='沟通摘要状态') {
      body.conditionalFormats.add('containsText',{text:'待更新',format:{fill:'#FFF1D6',font:{color:'#8A5700'}}});
    }
  }
  if (tab.name==='达人开发主表') {
    const a=headers.indexOf('内容能力(1-5)'),b=headers.indexOf('联系方式完整度(1-5)'),target=headers.indexOf('综合评分');
    if(a>=0&&b>=0&&target>=0) {
      for(let r=2;r<=lastRow;r++) sheet.getRange(`${col(target)}${r}`).formulas=[[`=IF(COUNTA(${col(a)}${r}:${col(b)}${r})=0,"",AVERAGE(${col(a)}${r}:${col(b)}${r}))`]];
      sheet.getRange(`${col(target)}2:${col(target)}${lastRow}`).conditionalFormats.add('colorScale',{colors:['#FEE2E2','#FEF3C7','#DCFCE7'],thresholds:['min',{type:'percentile',value:50},'max']});
    }
  }
  for (let r=0;r<tab.rows.length;r++) {
    for(const [h,text] of Object.entries(tab.rows[r]._notes||{})) {
      const c=headers.indexOf(h);
      if(c<0||!text) continue;
      const address=`${col(c)}${r+2}`;
      wb.notes.add({id:`${tab.name}:${address}`,target:{cell:{sheetName:sheet.name,sheetId:sheet.sheetId,address}},authorId:'',createdAt:'',body:{plainText:String(text)}});
    }
  }
  sheet.freezePanes.freezeRows(1);
  sheet.freezePanes.freezeColumns(tab.name==='邮件记录'?3:2);
}
wb.recalculate();
const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#NULL!',options:{useRegex:true,maxResults:10},maxChars:1000});
if (!errors.ndjson.includes('matched 0')) console.error(errors.ndjson);
if(previewDir) {
  await fs.mkdir(previewDir,{recursive:true});
  for(const [name,range,file] of [['达人开发主表','A1:G7','main.png'],['达人开发主表','H1:P8','main-data.png'],['邮件记录','A1:J6','mail.png']]) {
    const blob=await wb.render({sheetName:name,range,scale:1,format:'png'});
    await fs.writeFile(path.join(previewDir,file),new Uint8Array(await blob.arrayBuffer()));
  }
}
const file=await SpreadsheetFile.exportXlsx(wb);
await file.save(output);
console.log(JSON.stringify({saved:true,sheets:spec.sheets.map(s=>({name:s.name,rows:s.rows.length}))}));

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
  '处理状态':['待处理','待回复','已归档','待人工归属','历史截图待核实','待确认草稿','已确认待发送','已发送','发送失败','发送结果待核实','已作废'],
  '实际发送结果':['未发送','成功','失败','结果待核实'],
  '草稿用途':['首联','回复','跟进','历史'],
  '宠物类型':['Dog','Cat','Dog+Cat','Cat+Dog','Other'],
  '是否低粉高播放':['是','否'], '是否带货类似产品':['是','否'],
  '是否寄样':['否','准备寄出','已寄出','已签收'],
  '到货情况':['未寄出','运输中','已签收','异常'],
  '是否发布':['否','待发布','是'],'是否授权素材':['否','待确认','是'],
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
  area.format.rowHeight = tab.name==='达人开发主表'?36:42;
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
    range.format.columnWidth = tab.widths?.[h] || (h.includes('ID')?17: h.includes('时间')||h.includes('日期')?20:20);
    const body=sheet.getRange(`${col(i)}2:${col(i)}${Math.max(lastRow,500)}`);
    if (dates.has(h)) body.setNumberFormat(h==='计划跟进日期'||h==='首次联系日期'||h==='历史跟进日期'||h==='发布日期'?'yyyy-mm-dd':'yyyy-mm-dd hh:mm');
    if (dates.has(h)) for(let r=0;r<tab.rows.length;r++) {
      if (/^\d{4}-\d{2}-\d{2}$/.test(String(tab.rows[r][h]||''))) sheet.getRange(`${col(i)}${r+2}`).setNumberFormat('yyyy-mm-dd');
    }
    if (numbers.has(h)) body.setNumberFormat(h==='综合评分'?'0.00':'#,##0');
    if (choices[h]) body.dataValidation={rule:{type:'list',values:choices[h]}};
    if (h==='资料核实说明'||h==='核实说明') {
      range.format.wrapText = true;
      range.format.verticalAlignment = 'top';
    }
    if (h==='正文或草稿') {
      range.format.wrapText = true;
      range.format.verticalAlignment = 'top';
      for(let r=0;r<tab.rows.length;r++) {
        const text=String(tab.rows[r][h]||'');
        const lines=text.split('\n').reduce((n,line)=>n+Math.max(1,Math.ceil(line.length/85)),0);
        sheet.getRange(`A${r+2}:${last}${r+2}`).format.rowHeight=Math.min(405,Math.max(72,lines*12+12));
      }
    }
    if (h==='下一步动作'||h==='计划跟进日期'||h==='分类核实状态') body.format.fill='#FFF6DE';
    if (h==='分类核实状态'||h==='下一步动作') {
      body.conditionalFormats.add('containsText',{text:'待核实',format:{fill:'#FFF1D6',font:{color:'#8A5700'}}});
      if(h==='下一步动作') body.conditionalFormats.add('containsText',{text:'暂停',format:{fill:'#FCE8E6',font:{color:'#9D2823'}}});
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
    const imageIndex=headers.indexOf('历史截图');
    for(const picture of tab.rows[r]._images||[]) {
      if(imageIndex<0) throw Error('截图缺少目标列');
      const bytes=await fs.readFile(picture.path);
      const width=Math.min(480,picture.width||480,500*(picture.width||480)/(picture.height||240));
      const height=width*(picture.height||240)/(picture.width||480);
      sheet.images.add({dataUrl:`data:image/png;base64,${bytes.toString('base64')}`,anchor:{from:{row:r+1,col:imageIndex},extent:{widthPx:width,heightPx:height}}});
      sheet.getRange(`${col(imageIndex)}${r+2}`).format.rowHeight=Math.max(sheet.getRange(`${col(imageIndex)}${r+2}`).format.rowHeight||0,Math.min(405,height*.75+8));
    }
  }
  sheet.freezePanes.freezeRows(1);
  sheet.freezePanes.freezeColumns(2);
}
wb.recalculate();
const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#NULL!',options:{useRegex:true,maxResults:10},maxChars:1000});
if (!errors.ndjson.includes('matched 0')) console.error(errors.ndjson);
if(previewDir) {
  await fs.mkdir(previewDir,{recursive:true});
  for(const [name,range,file] of [['达人开发主表','A1:H10','main.png'],['达人开发主表','H1:P8','main-data.png'],['邮件记录','B1:I4','mail.png']]) {
    const blob=await wb.render({sheetName:name,range,scale:1,format:'png'});
    await fs.writeFile(path.join(previewDir,file),new Uint8Array(await blob.arrayBuffer()));
  }
}
const file=await SpreadsheetFile.exportXlsx(wb);
await file.save(output);
console.log(JSON.stringify({saved:true,sheets:spec.sheets.map(s=>({name:s.name,rows:s.rows.length}))}));

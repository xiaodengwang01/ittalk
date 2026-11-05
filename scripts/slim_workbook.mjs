// Author only the requested main-sheet view; the Python wrapper preserves other parts.
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import {createRequire} from 'node:module';
import {pathToFileURL} from 'node:url';
const modules=process.env.TAILKI_ARTIFACT_MODULES || path.join(os.homedir(),'.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules');
const resolver=createRequire(path.join(modules,'tailki-loader.cjs'));
const {Workbook,SpreadsheetFile,FileBlob}=await import(pathToFileURL(resolver.resolve('@oai/artifact-tool')).href);
const [input,output,preview,mode]=process.argv.slice(2);
if(mode==='--before'){
  const old=await SpreadsheetFile.importXlsx(await FileBlob.load(input));
  const png=await old.render({sheetName:'达人开发主表',range:'A1:F4',scale:1});
  await fs.writeFile(preview,new Uint8Array(await png.arrayBuffer()));
  console.log('Existing view rendered.');
  process.exit(0);
}
const spec=JSON.parse(await fs.readFile(input,'utf8'));
const wb=Workbook.create(),sheet=wb.worksheets.add('达人开发主表');
const col=n=>{let s='';for(;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
const end=spec.rows.length+1,last=col(spec.headers.length);
const widths={'达人ID':14,'达人昵称':25,'新增批次':23,'账号链接':38,'首联内容分类':22,'内容类型':38,'粉丝数':16,
 '当前建联状况':21,'沟通摘要':52,'Email':35,'WhatsApp联系方式':31,'可沟通时间':36,'收件信息':35,'待发送回复梗概':50,'往来邮件梗概':120};
const values=[spec.headers,...spec.rows.map(row=>spec.headers.map(h=>{
 let v=row[h]??null;
 if(h==='新增批次'&&/^\d{4}-\d{2}-\d{2}$/.test(v))return (Date.parse(v+'T00:00:00Z')-Date.UTC(1899,11,30))/86400000;
 return typeof v==='string'&&v.startsWith('=')?"'"+v:v;
}))];
sheet.getRange(`A1:${last}${end}`).values=values;
sheet.showGridLines=false;sheet.tabColor='#315774';
const all=sheet.getRange(`A1:${last}${end}`);
all.format.font={name:'Arial',size:11,color:'#233444'};
all.format.verticalAlignment='top';all.format.wrapText=true;
sheet.getRange(`A1:${last}1`).format={fill:'#315774',font:{name:'Arial',size:11,bold:true,color:'#FFFFFF'},rowHeight:40,verticalAlignment:'center',horizontalAlignment:'center',wrapText:true};
for(const [i,h] of spec.headers.entries())sheet.getRange(`${col(i+1)}1:${col(i+1)}${end}`).format.columnWidth=widths[h];
sheet.getRange(`C2:C${end}`).setNumberFormat('yyyy-mm-dd');
sheet.getRange(`C2:C${end}`).format.horizontalAlignment='center';
sheet.getRange(`G2:G${end}`).setNumberFormat('#,##0');
sheet.getRange(`G2:G${end}`).format.horizontalAlignment='right';
for(const [i,row] of spec.rows.entries()){
 const n=i+2;
 if(i%2===1)sheet.getRange(`A${n}:${last}${n}`).format.fill='#F3F6F8';
 let needed=54;
 for(const h of spec.headers){
  const text=String(row[h]??'');
  const lines=text.split('\n').reduce((sum,line)=>sum+Math.max(1,Math.ceil([...line].reduce((a,c)=>a+(c.charCodeAt(0)>255?2:1),0)/(widths[h]-3))),0);
  needed=Math.max(needed,lines*16+22);
 }
 sheet.getRange(`A${n}:${last}${n}`).format.rowHeight=Math.min(409,Math.max(row._height||0,needed));
}
const status=col(spec.headers.indexOf('当前建联状况')+1);
for(const [text,fill,color] of [['已结束','#ECEFF1','#687581'],['确认','#FFF0CE','#7A5610'],['已回复','#E3F0E9','#285E42']])
 sheet.getRange(`${status}2:${status}${end}`).conditionalFormats.add('containsText',{text,format:{fill,font:{color}}});
sheet.freezePanes.freezeRows(1);sheet.freezePanes.freezeColumns(3);
wb.recalculate();
const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#NULL!',options:{useRegex:true,maxResults:20},maxChars:1200});
console.log(errors.ndjson);
await(await SpreadsheetFile.exportXlsx(wb)).save(output);
if(preview){
 for(const [name,range] of [['identity','A1:G4'],['followup',`${status}1:${last}3`],['latest',`A${Math.max(2,end-4)}:H${end}`]]){
  const png=await wb.render({sheetName:sheet.name,range,scale:1});
  await fs.writeFile(path.join(preview,name+'.png'),new Uint8Array(await png.arrayBuffer()));
 }
}

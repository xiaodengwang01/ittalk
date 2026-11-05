// Build the compact two-tab follow-up layout from already reviewed summaries.
// Input contains no complete mail bodies; ordinary updates use daily_workbook.py.
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import {createRequire} from 'node:module';
import {pathToFileURL} from 'node:url';
const modules=process.env.TAILKI_ARTIFACT_MODULES || path.join(os.homedir(),'.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules');
const resolver=createRequire(path.join(modules,'tailki-loader.cjs'));
const {Workbook,SpreadsheetFile}=await import(pathToFileURL(resolver.resolve('@oai/artifact-tool')).href);
const [input,output,previewDir]=process.argv.slice(2);
const spec=JSON.parse(await fs.readFile(input,'utf8'));
const wb=Workbook.create();
const col=n=>{let s='';for(;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
const widths=[14,25,21,52,35,31,36,35,50,120,32,22,16,38,15,16,24,16,20,20,19,24,19,23,16,21,24,90];
for(const tab of spec.sheets){
 const s=wb.worksheets.add(tab.name), end=tab.rows.length+1, last=col(spec.headers.length);
 s.showGridLines=false; s.tabColor=tab.name==='达人开发主表'?'#315774':'#577F7E';
 s.getRange(`A1:${last}${end}`).values=[spec.headers,...tab.rows.map(row=>spec.headers.map(h=>{
  const v=row[h]===''?null:(row[h]??null);return typeof v==='string'&&v.startsWith('=')?"'"+v:v;
 }))];
 const all=s.getRange(`A1:${last}${end}`);
 all.format.font={name:'Arial',size:11,color:'#233444'};
 all.format.verticalAlignment='top';all.format.wrapText=true;
 s.getRange(`A1:${last}1`).format={fill:'#315774',font:{name:'Arial',size:11,bold:true,color:'#FFFFFF'},rowHeight:40,verticalAlignment:'center',horizontalAlignment:'center',wrapText:true};
 widths.forEach((w,i)=>s.getRange(`${col(i+1)}1:${col(i+1)}${end}`).format.columnWidth=w);
 tab.rows.forEach((row,i)=>{
  const n=i+2;
  if(i%2===1)s.getRange(`A${n}:${last}${n}`).format.fill='#F3F6F8';
  const height=Math.max(54,...spec.headers.slice(0,10).map((h,j)=>{
   const text=String(row[h]??'');
   const lines=text.split('\n').reduce((total,line)=>total+Math.max(1,Math.ceil([...line].reduce((a,c)=>a+(c.charCodeAt(0)>255?2:1),0)/(widths[j]-3))),0);
   return lines*16+22;
  }));
  s.getRange(`A${n}:${last}${n}`).format.rowHeight=Math.min(409,height);
  if(row['初次建联日期']){
   const r=s.getRange(`${col(spec.headers.indexOf('初次建联日期')+1)}${n}`);
   r.values=[[new Date(String(row['初次建联日期']).slice(0,10)+'T00:00:00Z')]];r.setNumberFormat('yyyy-mm-dd');
  }
  const score=spec.headers.indexOf('综合评分');
  if(score>=0&&row._scoreFormula){
   const refs=['内容能力(1-5)','宠物人群匹配度(1-5)','播放能力(1-5)','联系方式完整度(1-5)'].map(h=>`${col(spec.headers.indexOf(h)+1)}${n}`).join(',');
   s.getRange(`${col(score+1)}${n}`).formulas=[[`=IF(COUNTA(${refs})=0,"",AVERAGE(${refs}))`]];
  }
 });
 s.getRange(`R2:T${end}`).setNumberFormat('#,##0');
 s.getRange(`U2:Y${end}`).setNumberFormat('0.0');
 s.getRange(`K2:AB${end}`).format.wrapText=false;
 s.getRange(`C2:C${end}`).conditionalFormats.add('containsText',{text:'已结束',format:{fill:'#ECEFF1',font:{color:'#687581'}}});
 s.getRange(`C2:C${end}`).conditionalFormats.add('containsText',{text:'确认',format:{fill:'#FFF0CE',font:{color:'#7A5610'}}});
 s.getRange(`C2:C${end}`).conditionalFormats.add('containsText',{text:'已回复',format:{fill:'#E3F0E9',font:{color:'#285E42'}}});
 s.freezePanes.freezeRows(1);s.freezePanes.freezeColumns(2);
}
wb.recalculate();
const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!',options:{useRegex:true,maxResults:20},maxChars:1000});
console.log(errors.ndjson);
await(await SpreadsheetFile.exportXlsx(wb)).save(output);
if(previewDir){
 for(const [i,tab] of spec.sheets.entries()){
  for(const [label,range] of [['daily','A1:F'+Math.min(tab.rows.length+1,5)],['contacts','F1:J'+Math.min(tab.rows.length+1,4)]]){
   const png=await wb.render({sheetName:tab.name,range,scale:1});
   await fs.writeFile(path.join(previewDir,`after-${i}-${label}.png`),new Uint8Array(await png.arrayBuffer()));
  }
 }
}

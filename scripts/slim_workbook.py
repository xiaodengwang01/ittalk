"""Explicit local migration of the main tracker; never touches the secondary sheet."""
from __future__ import annotations

import copy
from datetime import datetime
from io import BytesIO
import json
import os
from pathlib import Path
import re
import subprocess
import tempfile
import xml.etree.ElementTree as ET
from zipfile import ZipFile

import daily_workbook as w
import workbook_metadata as metadata

HEADERS = ['达人ID','达人昵称','新增批次','账号链接','首联内容分类','内容类型','粉丝数','当前建联状况',
           '沟通摘要','Email','往来邮件梗概']
S=w.S; NS=w.NS


def batch_label(value):
    text=str(value or '').strip()
    if '之前' in text:
        match=re.search(r'\d{4}-\d{2}-\d{2}',text)
        if not match: raise ValueError('无法确定早期批次范围：'+text)
        return match[0]+'之前'
    match=re.search(r'\d{4}-\d{2}-\d{2}',text)
    if match:
        return datetime.strptime(match[0],'%Y-%m-%d').date().isoformat()
    if text=='10-30新增': return '2026-10-30'
    raise ValueError('无法明确转换新增批次：'+text)


def merge_main(original, authored, *, end, columns):
    """Append authored styles without renumbering any existing styles."""
    with ZipFile(BytesIO(original)) as src, ZipFile(authored) as new:
        part=w.main_part(src)
        native=ET.fromstring(src.read('xl/styles.xml'))
        styles=ET.fromstring(new.read('xl/styles.xml'))
        order=['numFmts','fonts','fills','borders','cellStyleXfs','cellXfs','cellStyles','dxfs','tableStyles','colors','extLst']
        def container(name):
            node=native.find('s:'+name,NS)
            if node is None:
                node=ET.Element('{'+S+'}'+name)
                at=next((i for i,x in enumerate(native) if x.tag.split('}')[-1] in order and order.index(x.tag.split('}')[-1])>order.index(name)),len(native))
                native.insert(at,node)
            return node
        nums=container('numFmts'); num_map={}
        next_num=max([163]+[int(n.get('numFmtId')) for n in nums])+1
        for n in styles.findall('s:numFmts/s:numFmt',NS):
            existing=next((x for x in nums if x.get('formatCode')==n.get('formatCode')),None)
            if existing is None:
                existing=copy.deepcopy(n);existing.set('numFmtId',str(next_num));next_num+=1;nums.append(existing)
            num_map[int(n.get('numFmtId'))]=int(existing.get('numFmtId'))
        nums.set('count',str(len(nums)))
        maps={}
        for name in ['fonts','fills','borders','cellStyleXfs','cellXfs','dxfs']:
            target=container(name); mapping={}; nodes=styles.find('s:'+name,NS)
            for i,n in enumerate(nodes if nodes is not None else []):
                child=copy.deepcopy(n)
                if name in ('cellStyleXfs','cellXfs'):
                    for attr,source in [('fontId','fonts'),('fillId','fills'),('borderId','borders'),('xfId','cellStyleXfs')]:
                        if attr in child.attrib and source in maps: child.set(attr,str(maps[source][int(child.get(attr))]))
                    fmt=int(child.get('numFmtId','0'));child.set('numFmtId',str(num_map.get(fmt,fmt)))
                mapping[i]=len(target);target.append(child)
            target.set('count',str(len(target)));maps[name]=mapping
        root=ET.fromstring(new.read('xl/worksheets/sheet1.xml'))
        strings=list(ET.fromstring(new.read('xl/sharedStrings.xml'))) if 'xl/sharedStrings.xml' in new.namelist() else []
        for c in root.findall('.//s:c',NS):
            c.set('s',str(maps['cellXfs'][int(c.get('s','0'))]))
            if c.get('t')=='s':
                value=c.find('s:v',NS);inline=copy.deepcopy(strings[int(value.text)])
                inline.tag='{'+S+'}is';c.remove(value);c.append(inline);c.set('t','inlineStr')
        for rule in root.findall('.//s:cfRule',NS):
            if 'dxfId' in rule.attrib: rule.set('dxfId',str(maps['dxfs'][int(rule.get('dxfId'))]))
        for cols in root.findall('s:cols/s:col',NS):
            for attr in ('hidden','outlineLevel','collapsed'):cols.attrib.pop(attr,None)
        data=root.find('s:sheetData',NS)
        ref=f'A1:{w.workflow.col_name(columns)}{end}'
        dimension=root.find('s:dimension',NS)
        if dimension is None:
            dimension=ET.Element('{'+S+'}dimension')
            root.insert(1 if root.find('s:sheetPr',NS) is not None else 0,dimension)
        dimension.set('ref',ref)
        root.insert(list(root).index(data)+1,ET.Element('{'+S+'}autoFilter',{'ref':ref}))
        for view in root.findall('s:sheetViews/s:sheetView',NS):
            for node in list(view):
                if node.tag in ('{'+S+'}pane','{'+S+'}selection'):view.remove(node)
            view.set('topLeftCell','A1')
            view.insert(0,ET.Element('{'+S+'}pane',{'xSplit':'3','ySplit':'1','topLeftCell':'D2','activePane':'bottomRight','state':'frozen'}))
            ET.SubElement(view,'{'+S+'}selection',{'pane':'bottomRight','activeCell':'D2','sqref':'D2'})
        book=ET.fromstring(src.read('xl/workbook.xml'))
        names=book.find('s:definedNames',NS)
        if names is None: names=ET.SubElement(book,'{'+S+'}definedNames')
        for name in list(names):
            if name.get('name')=='_xlnm._FilterDatabase' and name.get('localSheetId')=='0':names.remove(name)
        ET.SubElement(names,'{'+S+'}definedName',{'name':'_xlnm._FilterDatabase','localSheetId':'0','hidden':'1'}).text=f"'达人开发主表'!$A$1:${w.workflow.col_name(columns)}${end}"
        edits={part:ET.tostring(root,encoding='utf-8',xml_declaration=True),
               'xl/styles.xml':ET.tostring(native,encoding='utf-8',xml_declaration=True),
               'xl/workbook.xml':ET.tostring(book,encoding='utf-8',xml_declaration=True)}
        result=BytesIO()
        with ZipFile(result,'w') as out:
            for info in src.infolist():out.writestr(info,edits.get(info.filename,src.read(info.filename)))
        return result.getvalue()


def migrate(state, book, backup, *, expected_hash, work, headers=None):
    book,backup,work=Path(book),Path(backup),Path(work)
    original=book.read_bytes()
    if w.mail.digest(original)!=expected_hash or backup.read_bytes()!=original:
        raise w.mail.DraftError('原表或备份已变化，停止覆盖。')
    before=w.read(book);main=before['sheets'][0];headers=headers or HEADERS
    if '新增批次' in main['headers']:raise w.mail.DraftError('主表已经精简，停止重复迁移。')
    with ZipFile(book) as z:
        root=ET.fromstring(z.read(w.main_part(z)))
        if root.find('s:tableParts',NS) is not None:raise w.mail.DraftError('主表新增了表对象，需单独处理。')
        heights={int(r.get('r')):float(r.get('ht','54')) for r in root.findall('s:sheetData/s:row',NS)}
        if w.workflow.rels(z,w.main_part(z)):raise w.mail.DraftError('主表含批注或图形关系，需先保留原生对象。')
    rows=[]
    for row in main['creators']:
        value={h:row.get(h) for h in headers};value['新增批次']=batch_label(row.get('历史批次'))
        value['_height']=heights.get(row['_row'],54);rows.append(value)
    spec=work/'spec.json';spec.write_text(json.dumps({'headers':headers,'rows':rows},ensure_ascii=False))
    authored=work/'authored.xlsx'
    node=Path.home()/'.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node'
    subprocess.run([str(node),str(Path(__file__).with_suffix('.mjs')),str(spec),str(authored),str(work)],check=True,timeout=120)
    final=merge_main(original,authored,end=len(rows)+1,columns=len(headers))
    staging=work/'verified.xlsx';staging.write_bytes(final)
    after=w.read(staging)
    assert after['sheets'][0]['headers']==headers
    assert len(after['main_creators'])==len(rows)
    for old,new in zip(main['creators'],after['main_creators']):
        for h in headers:
            expected=batch_label(old.get('历史批次')) if h=='新增批次' else old.get(h)
            assert (new.get(h) or None)==(expected or None),(old['达人ID'],h)
    with ZipFile(BytesIO(original)) as a,ZipFile(BytesIO(final)) as b:
        changed=[n for n in a.namelist() if a.read(n)!=b.read(n)]
        assert set(changed)=={w.main_part(a),'xl/styles.xml','xl/workbook.xml'},changed
        old_styles=ET.fromstring(a.read('xl/styles.xml'));new_styles=ET.fromstring(b.read('xl/styles.xml'))
        for kind in ('fonts','fills','borders','cellStyleXfs','cellXfs','dxfs'):
            old_nodes=old_styles.find('s:'+kind,NS);new_nodes=new_styles.find('s:'+kind,NS)
            for i,item in enumerate(old_nodes if old_nodes is not None else []):
                assert ET.tostring(item)==ET.tostring(new_nodes[i]),kind
    if book.read_bytes()!=original:raise w.mail.DraftError('检查期间表格发生人工修改，停止覆盖。')
    for row in main['creators']:
        metadata.update(state,book,row['达人ID'],{h:row.get(h) for h in main['headers'] if h not in headers and h in metadata.FIELDS},archive_hash=expected_hash)
    fd,replacement=tempfile.mkstemp(prefix='.tailki-slim-',suffix='.xlsx',dir=book.parent)
    try:
        with os.fdopen(fd,'wb') as out:out.write(final);out.flush();os.fsync(out.fileno())
        os.chmod(replacement,book.stat().st_mode&0o777)
        if book.read_bytes()!=original:raise w.mail.DraftError('替换前表格变化，停止覆盖。')
        os.replace(replacement,book)
    finally:
        if os.path.exists(replacement):os.unlink(replacement)
    return {'rows':len(rows),'columns':len(headers),'hash':w.mail.digest(final),'changed_parts':changed}

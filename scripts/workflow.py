#!/usr/bin/env python3
"""Local XLSX adapter and commands. It never logs into or sends from a mailbox."""
from __future__ import annotations

import argparse
import copy
from datetime import datetime, timedelta, timezone
import hashlib
import json
import os
from pathlib import Path
import posixpath
import re
import shutil
import subprocess
import tempfile
import xml.etree.ElementTree as E
from zipfile import ZipFile, ZIP_DEFLATED

import workflow_rules as rules

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_BOOK = ROOT/'KOL搜集与邮件回复/达人跟进表.xlsx'
RUNTIME = Path.home()/'.cache/codex-runtimes/codex-primary-runtime/dependencies'
NODE = Path(os.environ.get('TAILKI_NODE',RUNTIME/'node/bin/node'))
NS = {'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'p':'http://schemas.openxmlformats.org/package/2006/relationships',
      'd':'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
      'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
M = NS['s']
MASTER = ['达人ID','达人昵称','首联内容分类','分类核实状态','联系状态','下一步动作','计划跟进日期',
 'Email','账号链接','平台','宠物类型','联系页链接','最近收信时间','最近发信时间','来源链接','最近抓取时间',
 '内容类型','粉丝数','近10条中位播放','近10条最高播放','来源','首次联系日期','历史跟进日期','资料核实说明',
 '联系方式备注','Instagram','TikTok DM','内容能力(1-5)','宠物人群匹配度(1-5)','播放能力(1-5)',
 '联系方式完整度(1-5)','综合评分','是否低粉高播放','是否带货类似产品','首次报价(USD)',
 '对方回复风格判断','应对策略选择','我们的报价或合作方式','我们的报价(USD)2','备注',
 '是否寄样','收件人','地址信息','电话','物流单号','到货情况','推荐脚本方向','是否发布','发布日期',
 '视频链接','是否授权素材','Bio外链','UTM链接','优惠码']
MAIL = ['邮件记录ID','达人ID','达人昵称','方向','邮件时间','发件人','收件人','主题','正文或草稿',
 '处理状态','实际发送结果','草稿用途','模板及版本','草稿版本','草稿生成时间','原邮件定位',
 '邮件原始时间','核实说明','历史截图','记录来源','邮箱消息ID','会话ID','草稿指纹','发送确认批次',
 '确认时间','确认指纹','首联邮件记录ID','回复邮件记录ID']
DATE_FIELDS = {'首次联系日期','历史跟进日期','计划跟进日期','邮件时间','最近收信时间','最近发信时间',
 '最近抓取时间','确认时间','草稿生成时间','发布日期'}
ALIASES = {'编号':'达人ID','TikTok账号链接':'账号链接','来源路径':'来源','最近跟进日期':'历史跟进日期'}
MOVED = {'第一次回复情况截图','第一次回复内容文本','第二次回复情况截图2'}

def text(value):
    return '' if value is None else str(value).strip()

def col_number(address):
    n=0
    for c in re.match(r'[A-Z]+',address).group(): n=n*26+ord(c)-64
    return n

def col_name(n):
    out=''
    while n: n,i=divmod(n-1,26); out=chr(65+i)+out
    return out

def resolve_part(owner,target):
    return target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(owner),target))

def rels(z,part):
    name=posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')
    if name not in z.namelist(): return {}
    return {r.get('Id'):r for r in E.fromstring(z.read(name))}

def read_book(path,workdir):
    """Read values, dates, notes and all original PNG screenshots without third-party writers."""
    Path(workdir).mkdir(parents=True,exist_ok=True)
    with ZipFile(path) as z:
        strings=[''.join(x.itertext()) for x in E.fromstring(z.read('xl/sharedStrings.xml'))] if 'xl/sharedStrings.xml' in z.namelist() else []
        style=E.fromstring(z.read('xl/styles.xml'))
        custom={int(n.get('numFmtId')):n.get('formatCode') for n in style.findall('s:numFmts/s:numFmt',NS)}
        date_styles=set(); date_only_styles=set()
        for i,x in enumerate(style.findall('s:cellXfs/s:xf',NS)):
            code=re.sub(r'"[^"]*"','',custom.get(int(x.get('numFmtId','0')),''))
            fmt=int(x.get('numFmtId','0'))
            if fmt in range(14,23) or re.search('[ydh]',code,re.I): date_styles.add(i)
            if fmt in range(14,18) or (re.search('[yd]',code,re.I) and not re.search('[hs]',code,re.I)): date_only_styles.add(i)
        wb=E.fromstring(z.read('xl/workbook.xml')); relationships=rels(z,'xl/workbook.xml')
        result={'sheets':[]}
        for tab in wb.findall('s:sheets/s:sheet',NS):
            part=resolve_part('xl/workbook.xml',relationships[tab.get('{'+NS['r']+'}id')].get('Target'))
            root=E.fromstring(z.read(part)); rows=[]; by_row={}; header=None
            for r in root.findall('s:sheetData/s:row',NS):
                cells={}
                for c in r.findall('s:c',NS):
                    v=c.find('s:v',NS); value=v.text if v is not None else None
                    typ=c.get('t')
                    if typ=='s' and value is not None: value=strings[int(value)]
                    elif typ=='inlineStr': value=''.join(c.find('s:is',NS).itertext())
                    elif typ=='d' and value:
                        # Excel dates have no time zone. The writer uses UTC construction
                        # to encode the business wall-clock value without a host-zone shift.
                        dt=datetime.fromisoformat(value.replace('Z','+00:00')).replace(tzinfo=None)
                        value=dt.date().isoformat() if int(c.get('s','0')) in date_only_styles else dt.isoformat(timespec='seconds')
                    elif typ=='b' and value is not None: value=value=='1'
                    elif value not in (None,'') and typ not in ('str','e'):
                        value=float(value)
                        if int(c.get('s','0')) in date_styles:
                            dt=datetime(1899,12,30)+timedelta(seconds=round(value*86400))
                            value=dt.date().isoformat() if int(c.get('s','0')) in date_only_styles else dt.isoformat(timespec='seconds')
                        elif value.is_integer(): value=int(value)
                    cells[col_number(c.get('r'))]=value
                if header is None:
                    header=[cells.get(i) for i in range(1,max(cells,default=0)+1)]; continue
                # Empty score formulas are not creator records.
                record={str(h):cells.get(i+1) for i,h in enumerate(header) if h}
                if not any(v not in (None,'') for k,v in record.items() if k!='综合评分'): continue
                record['_row']=int(r.get('r')); record['_notes']={}; record['_images']=[]
                rows.append(record); by_row[record['_row']]=record
            for rel in rels(z,part).values():
                typ=rel.get('Type',''); target=resolve_part(part,rel.get('Target'))
                if typ.endswith('/comments'):
                    for note in E.fromstring(z.read(target)).findall('s:commentList/s:comment',NS):
                        address=note.get('ref'); rn=int(re.search(r'\d+',address).group()); cn=col_number(address)
                        if rn in by_row and cn<=len(header): by_row[rn]['_notes'][header[cn-1]]=''.join(note.find('s:text',NS).itertext())
                if typ.endswith('/drawing'):
                    drawing=E.fromstring(z.read(target)); drawing_rels=rels(z,target)
                    for i,anchor in enumerate(drawing):
                        origin=anchor.find('d:from',NS); picture=anchor.find('d:pic',NS)
                        if origin is None or picture is None: continue
                        rn=int(origin.find('d:row',NS).text)+1
                        if rn not in by_row: raise ValueError('截图没有对应记录，停止以避免丢失')
                        blip=picture.find('.//a:blip',NS); rid=blip.get('{'+NS['r']+'}embed')
                        media=resolve_part(target,drawing_rels[rid].get('Target'))
                        raw=z.read(media); saved=Path(workdir)/f'picture-{len(result["sheets"])}-{i}.png'; saved.write_bytes(raw)
                        if raw[:8]!=b'\x89PNG\r\n\x1a\n': raise ValueError('当前适配器仅支持PNG截图，已停止以保留原文件')
                        width=int.from_bytes(raw[16:20],'big'); height=int.from_bytes(raw[20:24],'big')
                        by_row[rn]['_images'].append({'path':str(saved),'width':width,'height':height,'source_media_path':media,'sha256':hashlib.sha256(raw).hexdigest()})
            result['sheets'].append({'name':tab.get('name'),'headers':header or [],'rows':rows})
        return result

def append_note(record,key,value):
    if not value: return
    notes=record.setdefault('_notes',{}); old=notes.get(key,'')
    if value not in old: notes[key]=(old+'\n'+value).strip()

def issue(record,value):
    append_note(record,'资料核实说明',value)
    short=text(record.get('资料核实说明'))
    if value not in short and len(short)<90: record['资料核实说明']=(short+'；'+value).strip('；')

def new_mail_id(messages):
    nums=[int(text(m.get('邮件记录ID')).split('-')[-1]) for m in messages if re.fullmatch(r'MAIL-\d+',text(m.get('邮件记录ID')))]
    return f'MAIL-{max(nums,default=0)+1:06d}'

def migrate(state,evidence,template_path):
    if len(state['sheets'])!=1 or state['sheets'][0]['name']!='达人开发主表': raise ValueError('迁移只适用于旧单页主表；新格式请使用refresh')
    source=state['sheets'][0]; creators=[]; messages=[]; aliases={}; by_url={}
    evidence=evidence or {}
    incoming=evidence.get('incoming_screenshot_records',[])
    for old in source['rows']:
        if '示例' in text(old.get('达人昵称')) and 'example' in text(old.get('TikTok账号链接')): continue
        creator={ALIASES.get(k,k):v for k,v in old.items() if k not in MOVED and not k.startswith('_')}
        creator['_notes']={ALIASES.get(k,k):v for k,v in old['_notes'].items() if k not in MOVED}
        raw_name=text(creator.get('达人昵称')); found=re.search(r'https?://[^\s]+',raw_name)
        if found and not creator.get('账号链接'):
            creator['账号链接']=found.group(); creator['达人昵称']=raw_name.split('\t')[0]
            append_note(creator,'达人昵称','原昵称单元格：'+raw_name)
            issue(creator,'已拆分昵称中混入的账号链接')
        if not creator.get('账号链接'):
            candidates={m.get('account_url') for m in incoming if m.get('original_anchor_row')==old['_row'] and m.get('account_url') and m.get('sender_email')==creator.get('Email')}
            if len(candidates)==1:
                creator['账号链接']=candidates.pop()
                matching=next(m for m in incoming if m.get('account_url')==creator['账号链接'])
                creator['达人昵称']=matching.get('sender_display_name') or creator['账号链接'].split('@')[-1]
                issue(creator,'账号与昵称依据历史邮件主题补齐')
        url=text(creator.get('账号链接'))
        if 'tiktok.com/' in url: creator['平台']='TikTok'
        creator['账号链接']=rules.canonical_account_url(creator.get('平台'),url)
        creator['宠物类型']=text(creator.get('宠物类型')).replace('other','Other')
        if not rules.is_valid_email(creator.get('Email')):
            contact=text(creator.get('Email'))
            if contact.startswith(('http://','https://')): creator['联系页链接']=contact
            elif contact: append_note(creator,'Email','原联系方式：'+contact)
            creator['Email']=None
            issue(creator,'缺少有效邮箱，联系页已单独保存')
        for field in ('粉丝数','近10条中位播放','近10条最高播放'):
            raw=creator.get(field)
            # A trailing apostrophe was a literal typing artifact, not a scale qualifier.
            parsed=rules.parse_metric(text(raw).rstrip("'’")) if isinstance(raw,str) else rules.parse_metric(raw)
            if raw not in (None,'') and parsed is None: issue(creator,f'{field}格式待核实'); append_note(creator,field,'原值：'+text(raw))
            creator[field]=parsed
        if creator.get('近10条中位播放') is not None and creator.get('近10条最高播放') is not None and creator['近10条中位播放']>creator['近10条最高播放']:
            issue(creator,'中位播放高于最高播放，保留原值待核实')
        creator['分类核实状态']='待核实' if creator.get('首联内容分类')=='待核实' else '初判'
        note=creator['_notes'].get('首联内容分类','')
        urls=re.findall(r'https?://[^\s]+',note)
        creator['来源链接']=urls[0] if urls else None
        creator['最近抓取时间']=None
        creator['最近收信时间']=None; creator['最近发信时间']=None; creator['计划跟进日期']=None
        key=(creator.get('平台'),creator.get('账号链接'))
        if key[1] and key in by_url:
            kept=by_url[key]; aliases[old['_row']]=kept
            for h,v in creator.items():
                if h.startswith('_'): continue
                if kept.get(h) in (None,'') and v not in (None,''): kept[h]=v
                elif kept.get(h)!=v and v not in (None,''):
                    append_note(kept,h,'合并重复记录原值：'+text(v))
            for h,v in creator['_notes'].items(): append_note(kept,h,v)
            issue(kept,'已按相同账号合并重复记录并保留历史')
        else:
            if creator.get('达人ID'):
                creator['达人ID']=str(creator['达人ID'])
            else:
                reserved={str(r.get('编号')) for r in source['rows'] if r.get('编号')}
                number=len(creators)+1
                while f'KOL-{number:06d}' in reserved or any(c.get('达人ID')==f'KOL-{number:06d}' for c in creators): number+=1
                creator['达人ID']=f'KOL-{number:06d}'
            creators.append(creator); aliases[old['_row']]=creator
            if key[1]: by_url[key]=creator
        if old.get('第一次回复内容文本'):
            messages.append({'邮件记录ID':new_mail_id(messages),'达人ID':aliases[old['_row']]['达人ID'],
              '方向':'收信','正文或草稿':old['第一次回复内容文本'],'处理状态':'历史截图待核实',
              '记录来源':f'原表第{old["_row"]}行回复文本','核实说明':'历史文本缺少原邮件定位与时间。'})
    all_images=[(row,picture) for row in source['rows'] for picture in row['_images']]
    attached=set()
    for item in incoming:
        url=rules.canonical_account_url('TikTok',item.get('account_url'))
        creator=by_url.get(('TikTok',url))
        # A contradictory subject alias is not resolved using an email/anchor alone.
        if any('别名建议核实' in x for x in item.get('issues',[])): creator=None
        record={'邮件记录ID':new_mail_id(messages),'达人ID':creator['达人ID'] if creator else None,
          '达人昵称':creator.get('达人昵称') if creator else None,'方向':'收信','邮件时间':None,
          '发件人':item.get('sender_email'),'收件人':item.get('recipient_email'),'主题':item.get('subject'),
          '正文或草稿':item.get('body'),'处理状态':'历史截图待核实' if creator else '待人工归属',
          '邮件原始时间':item.get('time_raw'),'核实说明':'\n'.join(item.get('issues',[])),
          '历史截图':'保留原图，可放大查看','记录来源':f'原表第{item.get("original_anchor_row")}行截图（{item.get("evidence_id")}）',
          '原邮件定位':None,'邮箱消息ID':None,'会话ID':None,'_images':[],'_notes':{}}
        media=item.get('source_media_path')
        matches=[i for i,(row,picture) in enumerate(all_images) if i not in attached and picture['source_media_path']==media]
        preferred=[i for i in matches if all_images[i][0]['_row']==item.get('original_anchor_row')]
        if preferred or matches:
            image_index=(preferred or matches)[0]; record['_images']=[all_images[image_index][1]]; attached.add(image_index)
        messages.append(record)
        if creator:
            creator['联系状态']='已回复'; creator['下一步动作']='回复草稿'
            issue(creator,'历史收信时间未完整显示，需到邮箱核实')
    for item in evidence.get('quoted_outgoing_evidence',[]):
        creator=by_url.get(('TikTok',rules.canonical_account_url('TikTok',item.get('account_url'))))
        messages.append({'邮件记录ID':new_mail_id(messages),'达人ID':creator.get('达人ID') if creator else None,
          '达人昵称':creator.get('达人昵称') if creator else None,'方向':'发信','邮件时间':None,
          '发件人':item.get('sender_email'),'收件人':item.get('recipient_email'),'主题':item.get('subject'),
          '正文或草稿':item.get('body'),'处理状态':'历史截图待核实','实际发送结果':None,
          '草稿用途':'历史','邮件原始时间':item.get('time_raw'),'核实说明':'\n'.join(item.get('issues',[])),
          '记录来源':'历史来信中的引用头；不视为已核实的发送回执','_images':[],'_notes':{}})
    for image_index,(old,picture) in enumerate(all_images):
        if image_index in attached: continue
        creator=aliases.get(old['_row'])
        messages.append({'邮件记录ID':new_mail_id(messages),'达人ID':creator.get('达人ID') if creator else None,
          '方向':'待核实','处理状态':'历史截图待核实','历史截图':'保留原图，尚未转录',
          '记录来源':f'原表第{old["_row"]}行截图','核实说明':'缺少可核实的具体邮件内容与时间。','_images':[picture],'_notes':{}})
    for old in source['rows']:
        moved_notes={k:v for k,v in old.get('_notes',{}).items() if k in MOVED and v}
        if not moved_notes: continue
        creator=aliases.get(old['_row'])
        annotation='\n'.join(f'{key}：{value}' for key,value in moved_notes.items())
        messages.append({'邮件记录ID':new_mail_id(messages),'达人ID':creator.get('达人ID') if creator else None,
          '方向':'待核实','处理状态':'历史截图待核实' if creator else '待人工归属',
          '核实说明':'迁移原回复字段的备注，未视为一封已核实的邮件。',
          '记录来源':f'原表第{old["_row"]}行回复字段备注','_notes':{'核实说明':annotation},'_images':[]})
    if sum(len(m.get('_images',[])) for m in messages)!=len(all_images): raise ValueError('迁移截图实例数量不一致')
    for creator in creators:
        if not creator.get('下一步动作'):
            creator['下一步动作']='核实分类' if creator.get('分类核实状态')=='待核实' else '补资料'
        if creator.get('历史跟进日期') or creator.get('首次联系日期'):
            append_note(creator,'历史跟进日期','原表手工记录，未与完整原邮件核对；不据此自动催复。')
        if not any(m.get('达人ID')==creator['达人ID'] and m.get('方向')=='收信' for m in messages):
            issue(creator,'历史联系状态尚未与原邮件核对')
    for item in evidence.get('document_drafts',[]):
        creator=by_url.get(('TikTok',rules.canonical_account_url('TikTok',item.get('account_url'))))
        if not creator: continue
        draft=rules.create_draft(creators,messages,creator['达人ID'],item.get('subject',''),item.get('body',''),
          f'{Path(item.get("source_document",template_path)).name} · 现有草稿',purpose='回复')
        draft['记录来源']=f'{Path(item.get("source_document",template_path)).name}：{item.get("section","")}'
        draft['核实说明']='原有文档拟稿，尚未发送；发送前核实最新原邮件及草稿内容。'
        draft['达人昵称']=creator.get('达人昵称')
    state={'sheets':[{'name':'达人开发主表','headers':MASTER,'rows':creators},
                     {'name':'邮件记录','headers':MAIL,'rows':messages}]}
    return state

def prepare(state):
    if [s['name'] for s in state['sheets']]!=['达人开发主表','邮件记录']: raise ValueError('需要两张工作表：达人开发主表、邮件记录')
    creators,messages=[s['rows'] for s in state['sheets']]
    rules.ensure_creator_ids(creators)
    ids={c['达人ID']:c for c in creators}
    for message in messages:
        if message.get('达人ID') in ids: message['达人昵称']=ids[message['达人ID']].get('达人昵称')
        if message.get('方向')=='发信' and message.get('处理状态')=='已确认待发送' and not rules.confirmation_valid(message):
            message['处理状态']='待确认草稿'; message['发送确认批次']=None; message['确认时间']=None; message['确认指纹']=None
    for row in creators+messages:
        for field in DATE_FIELDS:
            value=row.get(field)
            if not value: continue
            if hasattr(value,'isoformat'): value=value.isoformat()
            if not isinstance(value,str): raise ValueError('日期必须为可识别的日期或时间')
            try: dt=datetime.fromisoformat(value.replace('Z','+00:00'))
            except ValueError: raise ValueError(f'{field}不是有效日期')
            if dt.tzinfo: dt=dt.astimezone(timezone(timedelta(hours=8))).replace(tzinfo=None)
            row[field]=dt.date().isoformat() if len(value)==10 else dt.isoformat(timespec='seconds')
    for tab,expected in zip(state['sheets'],[MASTER,MAIL]):
        # Preserve extra columns added manually rather than silently dropping them.
        extra=[h for h in tab['headers'] if h and h not in expected]
        extra += [h for row in tab['rows'] for h in row if not h.startswith('_') and h not in expected+extra]
        tab['headers']=expected+list(dict.fromkeys(extra))
    state['sheets'][0]['widths']={'达人ID':17,'达人昵称':26,'首联内容分类':17,'分类核实状态':15,
      '联系状态':13,'下一步动作':16,'计划跟进日期':17,'Email':34,'账号链接':42,'平台':13,'宠物类型':13,
      '联系页链接':36,'来源链接':42,'内容类型':31,'资料核实说明':46,'备注':38}
    state['sheets'][1]['widths']={'达人昵称':24,'方向':10,'邮件时间':21,'发件人':32,'收件人':32,
      '主题':52,'正文或草稿':90,'处理状态':22,'实际发送结果':19,'模板及版本':38,'原邮件定位':40,
      '邮件原始时间':45,'核实说明':65,'历史截图':74,'记录来源':42}
    return creators,messages

def validate(state):
    creators,messages=prepare(state)
    ids=[c['达人ID'] for c in creators]
    if len(set(ids))!=len(ids): raise ValueError('达人ID重复')
    keys=[(c.get('平台'),rules.canonical_account_url(c.get('平台'),c.get('账号链接'))) for c in creators if c.get('账号链接')]
    if len(set(keys))!=len(keys): raise ValueError('账号仍有重复')
    for c in creators:
        if c.get('Email') and not rules.is_valid_email(c['Email']): raise ValueError('Email列存在无效邮箱')
        if c.get('下一步动作') not in rules.NEXT_ACTIONS: raise ValueError('下一步动作不在约定范围')
        for field in ('粉丝数','近10条中位播放','近10条最高播放'):
            if c.get(field) is not None and not isinstance(c[field],(int,float)): raise ValueError('粉丝或播放不是数字')
    mail_ids=[m.get('邮件记录ID') for m in messages]
    if None in mail_ids or len(set(mail_ids))!=len(mail_ids): raise ValueError('邮件记录ID缺失或重复')
    for m in messages:
        if m.get('达人ID') and m['达人ID'] not in ids: raise ValueError('邮件关联到不存在的达人ID')
        if not m.get('达人ID') and m.get('处理状态')!='待人工归属': raise ValueError('未关联邮件必须待人工归属')
    return {'creators':len(creators),'messages':len(messages),
      'screenshots':sum(len(r.get('_images',[])) for r in messages),
      'drafts':sum(m.get('处理状态') in ('待确认草稿','已确认待发送') for m in messages),
      'unassigned':sum(not m.get('达人ID') for m in messages)}

def patch_groups(path,state):
    """Preserve Excel column outline groups, currently missing from the authoring API."""
    edits={}
    with ZipFile(path) as z:
        for index,tab in enumerate(state['sheets'],1):
            part=f'xl/worksheets/sheet{index}.xml'; root=E.fromstring(z.read(part))
            E.register_namespace('',M); E.register_namespace('r',NS['r'])
            start=tab['headers'].index('是否寄样')+1 if index==1 else tab['headers'].index('邮箱消息ID')+1
            cols=root.find('s:cols',NS)
            replacements=[]
            for c in list(cols):
                for n in range(int(c.get('min')),int(c.get('max'))+1):
                    new=copy.deepcopy(c); new.set('min',str(n));new.set('max',str(n))
                    if n>=start: new.set('hidden','1');new.set('outlineLevel','1')
                    elif n==start-1: new.set('collapsed','1')
                    replacements.append(new)
            cols[:]=replacements
            props=root.find('s:sheetPr',NS)
            if props is None: props=E.Element('{'+M+'}sheetPr');root.insert(0,props)
            outline=props.find('s:outlinePr',NS)
            if outline is None: outline=E.SubElement(props,'{'+M+'}outlinePr')
            outline.set('summaryRight','0')
            sf=root.find('s:sheetFormatPr',NS);sf.set('outlineLevelCol','1')
            edits[part]=E.tostring(root,encoding='utf-8',xml_declaration=True)
        temp=Path(str(path)+'.grouped')
        with ZipFile(temp,'w',ZIP_DEFLATED) as out:
            for item in z.infolist(): out.writestr(item,edits.get(item.filename,z.read(item.filename)))
    os.replace(temp,path)

def write_book(path,state,workdir,original_hash,preview=False):
    Path(workdir).mkdir(parents=True,exist_ok=True)
    report=validate(state)
    manifest=Path(workdir)/'workbook.json'; staged=Path(workdir)/'staged.xlsx'
    manifest.write_text(json.dumps(state,ensure_ascii=False,default=str),encoding='utf-8')
    command=[str(NODE),str(ROOT/'scripts/write_workbook.mjs'),str(manifest),str(staged)]
    if preview: command.append(str(Path(workdir)/'previews'))
    subprocess.run(command,check=True,cwd=ROOT,capture_output=True,text=True)
    patch_groups(staged,state)
    saved=read_book(staged,Path(workdir)/'verify')
    if validate(saved)!=report: raise ValueError('保存后的记录数量校验失败')
    # Screenshots must retain their original bytes; they are evidence, not new illustrations.
    before=sorted(p['sha256'] for s in state['sheets'] for r in s['rows'] for p in r.get('_images',[]))
    after=sorted(p['sha256'] for s in saved['sheets'] for r in s['rows'] for p in r.get('_images',[]))
    if before!=after: raise ValueError('截图文件校验失败，原表未覆盖')
    for a,b in zip(state['sheets'],saved['sheets']):
        for old,new in zip(a['rows'],b['rows']):
            for h in a['headers']:
                if h=='综合评分': continue  # Recalculated from the preserved ratings.
                x,y=old.get(h),new.get(h)
                if x in (None,'') and y in (None,''): continue
                if h in DATE_FIELDS and x and y:
                    if str(x).replace('T',' ')[:16]==str(y).replace('T',' ')[:16]: continue
                if x!=y: raise ValueError(f'字段保存不一致：{a["name"]}/{h}')
            for h,note in old.get('_notes',{}).items():
                if h in a['headers'] and note!=new.get('_notes',{}).get(h): raise ValueError('单元格备注保存失败')
    if hashlib.sha256(Path(path).read_bytes()).hexdigest()!=original_hash: raise ValueError('原表在处理中被修改，已停止覆盖')
    fd,name=tempfile.mkstemp(prefix='.tailki-',suffix='.xlsx',dir=Path(path).parent);os.close(fd)
    shutil.copyfile(staged,name);shutil.copystat(path,name);os.replace(name,path)
    return report

def templates(path):
    with ZipFile(path) as z:
        paragraphs=[''.join(p.itertext()) for p in E.fromstring(z.read('word/document.xml')).findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')]
    def section(start,end):
        a=next(i for i,p in enumerate(paragraphs) if start in p);b=next(i for i in range(a+1,len(paragraphs)) if end in paragraphs[i])
        chosen=paragraphs[a:b]; body='\n\n'.join(p for p in chosen if not p.startswith(('选择策略：','Subject:')) and p!=chosen[0])
        subject=next((p.split(':',1)[1].strip() for p in chosen if p.startswith('Subject:')),'')
        body=body.replace('Best,SaraTailki','Best,\nSara\nTailki')
        return subject,body
    return {'first':section('初次发送 合作询价','A 类'),'followup':section('未回复时 跟进询价','常见产品问题'),
            'version':Path(path).name+' · '+hashlib.sha256(Path(path).read_bytes()).hexdigest()[:12]}

def fill_template(value,creator):
    email=text(creator.get('Email')); handle=text(creator.get('账号链接')).split('@')[-1].rstrip('/')
    description={'宠物剧情类':'pet personality and everyday interaction videos','宠物用品类':'pet product reviews and demonstrations',
                 '非宠物垂类':'product-focused short-form content'}.get(creator.get('首联内容分类'),'pet-focused short-form content')
    for key,replacement in {'name':email.split('@')[0],'tiktok_handle':handle,'content_description':description}.items(): value=value.replace('{{'+key+'}}',replacement)
    if re.search(r'\{\{.*?\}\}',value): raise ValueError('邮件模板仍有未填写字段')
    return value

def refresh(creators,messages,template_path,today):
    t=templates(template_path)
    for c in creators:
        if c.get('下一步动作')=='暂停联系' or c.get('联系状态') in ('拒绝','勿联系','已完成'):
            c['下一步动作']='暂停联系';c['计划跟进日期']=None;continue
        records=[m for m in messages if m.get('达人ID')==c['达人ID']]
        if any(m.get('方向')=='收信' for m in records):
            c['联系状态']='已回复';c['下一步动作']='回复草稿';c['计划跟进日期']=None
            for m in records:
                if m.get('草稿用途')=='跟进' and m.get('处理状态') in ('待确认草稿','已确认草稿'):
                    m['处理状态']='已作废';m['发送确认批次']=None;m['确认时间']=None;m['确认指纹']=None
            continue
        if not rules.is_valid_email(c.get('Email')): c['下一步动作']='补资料';continue
        if c.get('分类核实状态')=='待核实' or c.get('首联内容分类') in (None,'','待核实'): c['下一步动作']='核实分类';continue
        if c.get('联系状态') in ('待联系','未联系') and not any(m.get('方向')=='发信' for m in records):
            rules.create_draft(creators,messages,c['达人ID'],fill_template(t['first'][0],c),fill_template(t['first'][1],c),t['version'],purpose='首联')
            c['下一步动作']='首联草稿'
    added=[]
    for creator in creators:
        drafts=rules.generate_followup_drafts([creator],messages,today=today,template=fill_template(t['followup'][1],creator))
        for draft in drafts: draft['模板及版本']=t['version']+' · 未回复跟进'
        added.extend(drafts)
    return added

def main():
    parser=argparse.ArgumentParser(description='本地达人工作流：归档和草稿，不包含实际发送。')
    parser.add_argument('--book',type=Path,default=DEFAULT_BOOK)
    parser.add_argument('--preview',action='store_true')
    sub=parser.add_subparsers(dest='command',required=True)
    m=sub.add_parser('migrate');m.add_argument('--evidence',type=Path)
    sub.add_parser('check')
    p=sub.add_parser('preview-batch');p.add_argument('record_ids',nargs='+')
    a=sub.add_parser('assign-mail');a.add_argument('record_id');a.add_argument('creator_id')
    r=sub.add_parser('refresh');r.add_argument('--today',default=datetime.now().date().isoformat())
    for command in ('import-mail','import-scrape'):
        p=sub.add_parser(command);p.add_argument('input',type=Path)
    d=sub.add_parser('draft-reply');d.add_argument('creator_id');d.add_argument('subject');d.add_argument('body_file',type=Path);d.add_argument('--template-version',required=True);d.add_argument('--reply-to',default='')
    c=sub.add_parser('confirm-batch');c.add_argument('batch_id');c.add_argument('record_ids',nargs='+')
    s=sub.add_parser('record-result');s.add_argument('record_id');s.add_argument('result',choices=['成功','失败','结果待核实']);s.add_argument('--at',required=True)
    args=parser.parse_args();book=args.book.resolve()
    workdir=Path(tempfile.mkdtemp(prefix='tailki-workflow-'))
    (workdir/'verify').mkdir()
    original_hash=hashlib.sha256(book.read_bytes()).hexdigest()
    state=read_book(book,workdir)
    template_path=book.parent/'邮件模板.docx'
    if args.command=='migrate':
        evidence=json.loads(args.evidence.read_text()) if args.evidence else None
        state=migrate(state,evidence,template_path)
    else:
        creators,messages=prepare(state)
        if args.command=='check': print(json.dumps(validate(state),ensure_ascii=False));return
        if args.command=='preview-batch':
            for identifier in args.record_ids:
                matches=[m for m in messages if m.get('邮件记录ID')==identifier]
                if len(matches)!=1: raise ValueError('邮件记录不存在或不唯一')
                m=matches[0]
                print(json.dumps({k:m.get(k) for k in ['邮件记录ID','达人ID','收件人','主题','正文或草稿','草稿版本','处理状态','实际发送结果']},ensure_ascii=False,indent=2))
            return
        if args.command=='refresh': refresh(creators,messages,template_path,args.today)
        elif args.command=='import-scrape':
            data=json.loads(args.input.read_text()); candidates=data if isinstance(data,list) else data.get('creators',[])
            event_time=None if isinstance(data,list) else data.get('抓取时间') or data.get('event_time')
            result=rules.merge_creator_candidates(creators,candidates,event_time=event_time)
            for conflict in result.get('conflicts',[]):
                cid=conflict.get('达人ID') or conflict.get('creator_id')
                for creator in creators:
                    if creator.get('达人ID')==cid: issue(creator,'抓取内容与已确认资料冲突，需人工核实')
            refresh(creators,messages,template_path,datetime.now().date().isoformat())
        elif args.command=='import-mail':
            data=json.loads(args.input.read_text()); entries=data if isinstance(data,list) else data.get('messages',[])
            for incoming in entries: rules.import_message(creators,messages,incoming)
            refresh(creators,messages,template_path,datetime.now().date().isoformat())
        elif args.command=='draft-reply': rules.create_draft(creators,messages,args.creator_id,args.subject,args.body_file.read_text(),args.template_version,purpose='回复',in_reply_to=args.reply_to)
        elif args.command=='confirm-batch': rules.confirm_batch(messages,args.record_ids,args.batch_id)
        elif args.command=='record-result': rules.record_send_result(creators,messages,args.record_id,args.result,args.at)
        elif args.command=='assign-mail':
            creator=next((c for c in creators if c['达人ID']==args.creator_id),None)
            record=next((m for m in messages if m['邮件记录ID']==args.record_id),None)
            if not creator or not record: raise ValueError('达人ID或邮件记录ID不存在')
            if record.get('处理状态')!='待人工归属': raise ValueError('仅可人工关联待归属记录')
            record['达人ID']=creator['达人ID'];record['处理状态']='历史截图待核实' if not record.get('邮件时间') else '待回复'
            record['核实说明']=(text(record.get('核实说明'))+'\n已由人工明确达人归属。').strip()
            refresh(creators,messages,template_path,datetime.now().date().isoformat())
    report=write_book(book,state,workdir,original_hash,args.preview)
    print(json.dumps({'saved':str(book),**report,'preview_dir':str(workdir/'previews') if args.preview else None},ensure_ascii=False))

if __name__=='__main__':
    try: main()
    except (ValueError,OSError,subprocess.CalledProcessError) as error:
        if isinstance(error,subprocess.CalledProcessError): print('表格生成失败：'+(error.stderr or error.stdout or str(error))[-1600:])
        else: print('未保存：'+str(error))
        raise SystemExit(1)

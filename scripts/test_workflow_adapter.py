"""Integration checks using simulated data, including native XLSX round-trips."""
import copy
from datetime import date
import hashlib
import json
from pathlib import Path
import struct
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch
import zlib

sys.path.insert(0,str(Path(__file__).resolve().parent))
import workflow as app
import workflow_rules as rules


def png():
    def chunk(name,data):
        return struct.pack('!I',len(data))+name+data+struct.pack('!I',zlib.crc32(name+data)&0xffffffff)
    return b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('!2I5B',1,1,8,2,0,0,0))+chunk(b'IDAT',zlib.compress(b'\x00\x20\x40\x60'))+chunk(b'IEND',b'')


def creator():
    return {'达人ID':'SIM-000001','达人昵称':'模拟达人 Alpha','平台':'TikTok','账号链接':'https://tiktok.com/@sim_alpha',
      'Email':'alpha@example.invalid','首联内容分类':'宠物剧情类','分类核实状态':'初判','联系状态':'待联系','下一步动作':'首联草稿',
      '粉丝数':8500000,'近10条中位播放':21000,'近10条最高播放':50000,'_notes':{'首联内容分类':'模拟来源，待人工核实。'}}


def state(creators=None,messages=None):
    return {'sheets':[{'name':'达人开发主表','headers':list(app.MASTER),'rows':creators or [creator()]},
                      {'name':'邮件记录','headers':list(app.MAIL),'rows':messages or []}]}


class MigrationTests(unittest.TestCase):
    def old(self,rows):
        return {'sheets':[{'name':'达人开发主表','headers':['编号','达人昵称','TikTok账号链接','Email'], 'rows':rows}]}

    def row(self,index=3,**fields):
        r={'_row':index,'_notes':{},'_images':[],'编号':'SIM-KEEP-42','达人昵称':'模拟达人 Alpha',
          'TikTok账号链接':'https://www.tiktok.com/@sim_alpha?tracking=one','Email':'alpha@example.invalid','首联内容分类':'宠物剧情类'}
        r.update(fields);return r

    def test_migration_preserves_existing_id_and_normalizes_counts(self):
        r=self.row(粉丝数='8.5m',近10条中位播放='2k',近10条最高播放='1k')
        result=app.migrate(self.old([r]),None,Path('templates.docx'))
        out=result['sheets'][0]['rows'][0]
        self.assertEqual(out['达人ID'],'SIM-KEEP-42')
        self.assertEqual(out['粉丝数'],8500000)
        self.assertEqual(out['近10条中位播放'],2000)
        self.assertIn('中位播放高于',out['资料核实说明'])
        self.assertFalse(out['最近发信时间'])

    def test_reused_media_preserves_two_anchor_instances(self):
        picture={'source_media_path':'xl/media/shared.png','path':'unused','sha256':'simulated'}
        a=self.row();a['_images']=[dict(picture)]
        b=self.row(4,编号=None,达人昵称='模拟达人 Beta',TikTok账号链接='https://tiktok.com/@sim_beta',Email='beta@example.invalid');b['_images']=[dict(picture)]
        result=app.migrate(self.old([a,b]),None,Path('templates.docx'))
        records=result['sheets'][1]['rows']
        self.assertEqual(len(records),2)
        self.assertEqual(sum(len(m['_images']) for m in records),2)

    def test_old_reply_notes_migrate_with_history(self):
        r=self.row();r['_notes']={'第一次回复内容文本':'模拟历史备注，不可丢失'}
        result=app.migrate(self.old([r]),None,Path('templates.docx'))
        self.assertIn('模拟历史备注',result['sheets'][1]['rows'][0]['_notes']['核实说明'])

    def test_contact_page_is_not_email(self):
        result=app.migrate(self.old([self.row(Email='https://example.invalid/contact')]),None,Path('templates.docx'))
        out=result['sheets'][0]['rows'][0]
        self.assertFalse(out['Email'])
        self.assertEqual(out['联系页链接'],'https://example.invalid/contact')

    def test_prepare_invalidates_excel_edits_and_preserves_extra_fields(self):
        c=creator();messages=[]
        draft=rules.create_draft([c],messages,c['达人ID'],'Sim subject','Sim body','sim/v1')
        rules.confirm_batch(messages,[draft['邮件记录ID']],'SIM-BATCH','2026-10-19T10:00:00+08:00')
        draft['正文或草稿']='Edited in Excel'
        draft['自定义字段']='Must survive'
        s=state([c],messages);app.prepare(s)
        self.assertEqual(draft['处理状态'],'待确认草稿')
        self.assertFalse(draft['发送确认批次'])
        self.assertIn('自定义字段',s['sheets'][1]['headers'])

    def test_refresh_initial_draft_is_idempotent(self):
        c=creator();messages=[]
        t={'first':('Hello {{tiktok_handle}}','Hi {{name}}\n{{content_description}}'),
           'followup':('','Hi {{name}}\nFollowing up'), 'version':'sim/v1'}
        with patch.object(app,'templates',return_value=t):
            app.refresh([c],messages,Path('unused'),'2026-10-19')
            app.refresh([c],messages,Path('unused'),'2026-10-19')
        self.assertEqual(len(messages),1)
        self.assertNotIn('{{',messages[0]['正文或草稿'])

    def test_refresh_followup_uses_current_docx_template(self):
        c=creator();messages=[]
        draft=rules.create_draft([c],messages,c['达人ID'],'Original subject','Original body','sim/v1')
        rules.confirm_batch(messages,[draft['邮件记录ID']],'SIM-BATCH','2026-10-01T10:00:00+08:00')
        rules.record_send_result([c],messages,draft['邮件记录ID'],'成功','2026-10-01T10:05:00+08:00')
        t={'first':('Hello','Body'),'followup':('','Hi {{name}}\nFollowing up on @{{tiktok_handle}}'), 'version':'sim/v2'}
        with patch.object(app,'templates',return_value=t):
            app.refresh([c],messages,Path('unused'),'2026-10-19')
        self.assertEqual(len(messages),2)
        self.assertIn('alpha',messages[1]['正文或草稿'])
        self.assertEqual(messages[1]['模板及版本'],'sim/v2 · 未回复跟进')

    def test_refresh_keeps_pause_even_after_a_new_reply(self):
        c=creator();c['下一步动作']='暂停联系';messages=[]
        rules.import_message([c],messages,{'方向':'收信','达人ID':c['达人ID'],'邮箱消息ID':'SIM-IN-PAUSED',
          '发件人':'alpha@example.invalid','主题':'Sim reply','正文或草稿':'Hello'})
        t={'first':('',''),'followup':('','Hi {{name}}'), 'version':'sim/v1'}
        with patch.object(app,'templates',return_value=t): app.refresh([c],messages,Path('unused'),'2026-10-19')
        self.assertEqual(c['下一步动作'],'暂停联系')


class RoundTripTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp=tempfile.TemporaryDirectory(prefix='tailki-adapter-test-');cls.work=Path(cls.temp.name)
        image=cls.work/'sim.png';image.write_bytes(png())
        cls.expected=state()
        c=cls.expected['sheets'][0]['rows'][0]
        c['首次联系日期']='2026-10-19';c['最近收信时间']='2026-10-19T00:00:00';c['最近发信时间']='2026-10-19T10:00:00'
        c['自定义字段']='Extra field'
        cls.expected['sheets'][0]['headers'].append('自定义字段')
        cls.expected['sheets'][1]['rows']=[{'邮件记录ID':'MAIL-000001','达人ID':c['达人ID'],'方向':'收信',
          '处理状态':'历史截图待核实','邮件原始时间':'10月19日（年份待核实）','正文或草稿':'Simulation text',
          '_images':[{'path':str(image),'width':1,'height':1,'sha256':hashlib.sha256(image.read_bytes()).hexdigest()}]}]
        app.prepare(cls.expected)
        manifest=cls.work/'sim.json';manifest.write_text(json.dumps(cls.expected,ensure_ascii=False),encoding='utf-8')
        cls.book=cls.work/'sim.xlsx'
        subprocess.run([str(app.NODE),str(app.ROOT/'scripts/write_workbook.mjs'),str(manifest),str(cls.book)],check=True,capture_output=True,text=True)
        app.patch_groups(cls.book,cls.expected)
        cls.actual=app.read_book(cls.book,cls.work/'read')

    @classmethod
    def tearDownClass(cls): cls.temp.cleanup()

    def test_dates_preserve_wall_clock_and_precision(self):
        row=self.actual['sheets'][0]['rows'][0]
        self.assertEqual(row['首次联系日期'],'2026-10-19')
        self.assertEqual(row['最近收信时间'],'2026-10-19T00:00:00')
        self.assertEqual(row['最近发信时间'],'2026-10-19T10:00:00')

    def test_notes_images_extra_columns_and_schema_survive(self):
        self.assertEqual(app.validate(self.actual),{'creators':1,'messages':1,'screenshots':1,'drafts':0,'unassigned':0})
        self.assertEqual(self.actual['sheets'][0]['rows'][0]['自定义字段'],'Extra field')
        self.assertEqual(self.actual['sheets'][0]['rows'][0]['_notes']['首联内容分类'],'模拟来源，待人工核实。')
        self.assertEqual(self.actual['sheets'][1]['rows'][0]['_images'][0]['sha256'],hashlib.sha256(png()).hexdigest())

    def test_atomic_save_preserves_verified_data(self):
        destination=self.work/'atomic.xlsx';destination.write_bytes(self.book.read_bytes())
        before=hashlib.sha256(destination.read_bytes()).hexdigest()
        report=app.write_book(destination,copy.deepcopy(self.actual),self.work/'save',before)
        self.assertEqual(report['screenshots'],1)


if __name__=='__main__': unittest.main()

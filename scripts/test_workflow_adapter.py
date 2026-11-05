"""Simulated integration checks, including native XLSX round-trips."""
import copy
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch
import xml.etree.ElementTree as E
from zipfile import ZipFile

sys.path.insert(0, str(Path(__file__).resolve().parent))
import workflow as app


def creator(number=1):
    handle = 'sim_alpha' if number == 1 else 'sim_beta'
    return {'达人ID': 'SIM-%06d' % number, '达人昵称': '模拟达人 ' + handle,
            '平台': 'TikTok', '账号链接': 'https://tiktok.com/@' + handle,
            'Email': handle + '@example.invalid', '首联内容分类': '宠物剧情类',
            '分类核实状态': '初判', '联系状态': '待联系', '下一步动作': '首联草稿',
            '粉丝数': 8500000, '近10条中位播放': 21000, '近10条最高播放': 50000,
            '_notes': {'首联内容分类': '模拟来源，待人工核实。'}}


def incoming(number=1, creator_id='SIM-000001', **fields):
    row = {'邮件记录ID': 'MAIL-%06d' % number, '达人ID': creator_id, '方向': '收信',
           '主题': '模拟来信', '邮件正文': '模拟原始邮件正文。', '邮件摘要': '模拟摘要',
           '正文完整性': '完整', '邮件用途': '收信', '处理状态': '待回复', '记录来源': '原邮件',
           '发件人': 'sim_alpha@example.invalid', '收件人': 'brand@example.invalid',
           '_images': [], '_notes': {}}
    row.update(fields)
    return row


def state(creators=None, messages=None):
    return {'sheets': [
        {'name': '达人开发主表', 'headers': list(app.MASTER),
         'rows': [creator()] if creators is None else creators},
        {'name': '邮件记录', 'headers': list(app.MAIL),
         'rows': [] if messages is None else messages}]}


class MigrationTests(unittest.TestCase):
    def old(self, rows):
        return {'sheets': [{'name': '达人开发主表',
                           'headers': ['编号', '达人昵称', 'TikTok账号链接', 'Email'], 'rows': rows}]}

    def row(self, index=3, **fields):
        row = {'_row': index, '_notes': {}, '_images': [], '编号': 'SIM-KEEP-42',
               '达人昵称': '模拟达人 Alpha',
               'TikTok账号链接': 'https://www.tiktok.com/@sim_alpha?tracking=one',
               'Email': 'alpha@example.invalid', '首联内容分类': '宠物剧情类'}
        row.update(fields)
        return row

    def test_legacy_id_counts_and_unverified_dates(self):
        result = app.migrate(self.old([self.row(粉丝数='8.5m', 近10条中位播放='2k',
                                               近10条最高播放='1k')]), None, Path('templates.docx'))
        out = result['sheets'][0]['rows'][0]
        self.assertEqual(out['达人ID'], 'SIM-KEEP-42')
        self.assertEqual(out['粉丝数'], 8500000)
        self.assertEqual(out['近10条中位播放'], 2000)
        self.assertIn('中位播放高于', out['资料核实说明'])
        self.assertFalse(out['最近发信时间'])

    def test_legacy_keeps_each_image_until_text_conversion(self):
        picture = {'source_media_path': 'xl/media/shared.png', 'path': 'unused', 'sha256': 'simulated'}
        first = self.row(); first['_images'] = [dict(picture)]
        second = self.row(4, 编号=None, 达人昵称='模拟达人 Beta',
                          TikTok账号链接='https://tiktok.com/@sim_beta', Email='beta@example.invalid')
        second['_images'] = [dict(picture)]
        records = app.migrate(self.old([first, second]), None, Path('templates.docx'))['sheets'][1]['rows']
        self.assertEqual(len(records), 2)
        self.assertEqual(sum(len(m['_images']) for m in records), 2)

    def test_legacy_reply_notes_survive(self):
        row = self.row(); row['_notes'] = {'第一次回复内容文本': '模拟历史备注，不可丢失'}
        result = app.migrate(self.old([row]), None, Path('templates.docx'))
        self.assertIn('模拟历史备注', result['sheets'][1]['rows'][0]['_notes']['核实说明'])

    def test_contact_page_is_not_email(self):
        out = app.migrate(self.old([self.row(Email='https://example.invalid/contact')]),
                          None, Path('templates.docx'))['sheets'][0]['rows'][0]
        self.assertFalse(out['Email'])
        self.assertEqual(out['联系页链接'], 'https://example.invalid/contact')


class PureTextTests(unittest.TestCase):
    def test_prepare_renames_body_and_keeps_unrelated_fields(self):
        row = incoming(); row.pop('邮件正文')
        row.update({'正文或草稿': '模拟历史正文\n\n保留第二段。', '历史截图': '旧截图标签',
                    '草稿版本': 3, '自定义字段': 'Keep this field'})
        book = state(messages=[row]); app.prepare(book)
        self.assertEqual(row['邮件正文'], '模拟历史正文\n\n保留第二段。')
        self.assertIn('自定义字段', book['sheets'][1]['headers'])
        self.assertFalse(set(book['sheets'][1]['headers']) & app.REMOVED_MAIL_FIELDS)

    def test_refresh_never_creates_archive_drafts(self):
        first, followed = creator(), creator(2)
        followed.update({'联系状态': '已联系', '首次联系日期': '2026-10-01',
                         '最近发信时间': '2026-10-01T10:00:00'})
        records = []
        with patch.object(app, 'templates', side_effect=AssertionError('刷新不应读取草稿模板')):
            app.refresh([first, followed], records, Path('unused'), '2026-10-19')
            app.refresh([first, followed], records, Path('unused'), '2026-10-19')
        self.assertEqual(records, [])
        self.assertEqual(first['下一步动作'], '首联草稿')

    def test_refresh_keeps_pause_after_reply(self):
        c = creator(); c['下一步动作'] = '暂停联系'
        records = [incoming()]; app.refresh([c], records, today='2026-10-19')
        self.assertEqual(c['下一步动作'], '暂停联系')
        self.assertEqual(len(records), 1)

    def test_reply_stops_followup_action(self):
        c = creator(); c.update({'下一步动作': '跟进草稿', '计划跟进日期': '2026-10-19'})
        records = [incoming()]; app.refresh([c], records, today='2026-10-19')
        self.assertEqual(c['联系状态'], '已回复')
        self.assertEqual(c['下一步动作'], '回复草稿')
        self.assertFalse(c['计划跟进日期'])
        self.assertEqual(len(records), 1)

    def test_confirmed_first_contact_due_action_without_archiving_body(self):
        c=creator();c.update({'联系状态':'已联系','首次联系日期':'2026-10-01T10:00:00','首联发送结果':'成功'})
        app.refresh([c],[],today='2026-10-09')
        self.assertEqual(c['下一步动作'],'等待对方')
        app.refresh([c],[],today='2026-10-12')
        self.assertEqual(c['下一步动作'],'跟进草稿')
        self.assertEqual(c['计划跟进日期'],'2026-10-12')

    def test_later_sent_reply_covers_multiple_dated_messages(self):
        c=creator()
        records=[incoming(邮件时间='2026-10-18T11:00:00'),incoming(2,邮件时间='2026-10-18T12:00:00'),
                 incoming(3,方向='发信',邮件用途='回复',实际发送结果='成功',处理状态='已发送',邮件时间='2026-10-18T13:00:00')]
        app.refresh([c],records,today='2026-10-19')
        self.assertEqual(c['下一步动作'],'等待对方')

    def test_blank_mail_id_cannot_be_saved(self):
        with self.assertRaises(ValueError):app.validate(state(messages=[incoming(邮件记录ID='')]))

    def test_reject_images_drafts_initial_and_uncertain_sent(self):
        for fields in ({'_images': [{'path': 'unused.png'}]},
                       {'方向': '发信', '邮件用途': '回复', '处理状态': '待确认草稿', '实际发送结果': '未发送'},
                       {'方向': '发信', '邮件用途': '首联', '处理状态': '已发送', '实际发送结果': '成功'},
                       {'方向': '发信', '邮件用途': '回复', '处理状态': '发送结果待核实', '实际发送结果': '结果待核实'}):
            with self.subTest(fields=fields), self.assertRaises(ValueError):
                app.validate(state(messages=[incoming(**fields)]))

    def test_actual_sent_reply_is_allowed(self):
        row = incoming(方向='发信', 邮件用途='回复', 实际发送结果='成功', 处理状态='已发送')
        self.assertEqual(app.validate(state(messages=[row]))['messages'], 1)

    def test_conversion_preserves_ids_ownership_and_notes(self):
        inbox = incoming(); inbox.pop('邮件正文')
        inbox.update({'正文或草稿': '旧转录文字', '历史截图': '原图', '_images': [{'path': 'sim.png'}],
                      '_notes': {'正文或草稿': '模拟转录核实备注'}})
        unassigned = incoming(2, creator_id=None, 处理状态='待人工归属')
        header = incoming(3, 方向='发信', 草稿用途='历史', 处理状态='历史截图待核实', 邮件正文='')
        draft = incoming(4, 方向='发信', 草稿用途='回复', 处理状态='待确认草稿', 实际发送结果='未发送')
        annotations = {'messages': {
            'MAIL-000001': {'邮件正文': '可辨认原文\n\n[截图裁切，后文缺失]', '邮件摘要': '模拟要求待补齐',
                            '正文完整性': '部分', '记录来源': '截图转录', '核实说明新增': '模拟图片底部裁切'},
            'MAIL-000002': {'邮件正文': '未归属的模拟原文', '邮件摘要': '归属待核实',
                            '正文完整性': '完整', '记录来源': '截图转录'}},
            'creators': {'SIM-000001': {'沟通摘要': '模拟合作条件尚待确认。'}},
            'removed_header_creator_notes': {'SIM-000001': '模拟引用头缺少回复正文，不能证明已回复。'}}
        book = app.convert_text(state(messages=[inbox, unassigned, header, draft]), annotations)
        creators, messages = app.prepare(book)
        self.assertEqual([c['达人ID'] for c in creators], ['SIM-000001'])
        self.assertEqual([m['邮件记录ID'] for m in messages], ['MAIL-000001', 'MAIL-000002'])
        self.assertEqual(messages[0]['达人ID'], 'SIM-000001')
        self.assertFalse(messages[1].get('达人ID'))
        self.assertEqual(messages[1]['处理状态'], '待人工归属')
        self.assertEqual(messages[0]['邮件正文'], '可辨认原文\n\n[截图裁切，后文缺失]')
        self.assertEqual(messages[0]['正文完整性'], '部分')
        self.assertIn('模拟转录核实备注', messages[0]['_notes']['邮件正文'])
        self.assertIn('模拟引用头缺少回复正文', json.dumps(creators[0], ensure_ascii=False))
        self.assertEqual(creators[0]['沟通摘要'], '模拟合作条件尚待确认。')
        self.assertEqual(app.validate(book)['screenshots'], 0)
        self.assertFalse(set(book['sheets'][1]['headers']) & app.REMOVED_MAIL_FIELDS)

    def test_conversion_refuses_untranscribed_image(self):
        row = incoming(邮件正文='', _images=[{'path': 'unused.png'}])
        with self.assertRaises(ValueError):
            app.convert_text(state(messages=[row]), {'messages': {}, 'creators': {}})
        row['方向']='待核实'
        with self.assertRaises(ValueError):
            app.convert_text(state(messages=[row]), {'messages': {}, 'creators': {}})

    def test_transcription_cannot_change_existing_identity_or_owner(self):
        for fields in ({'达人ID':'SIM-000002'},{'邮件记录ID':'MAIL-000099'}):
            with self.subTest(fields=fields),self.assertRaises(ValueError):
                app.convert_text(state(messages=[incoming()]),{'messages':{'MAIL-000001':fields}})


class RoundTripTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory(prefix='tailki-adapter-test-'); cls.work = Path(cls.temp.name)
        first, second = creator(), creator(2)
        first.update({'首次联系日期': '2026-10-19', '最近收信时间': '2026-10-19T00:00:00',
                      '最近发信时间': '2026-10-19T10:00:00', '自定义字段': 'Extra field',
                      '沟通摘要': '模拟主要条件，待下一步确认。',
                      '草稿状态':'未发送草稿','草稿收件人':'alpha@example.invalid','草稿主题':'Sim reply draft',
                      '当前回复草稿':'SIM-UNSENT-WORKSPACE-BODY\n\nKeep draft paragraphs.',
                      '草稿来源':'模拟对话确认未发送'})
        cls.body = '=Literal mail text, never an Excel formula\n\nKeep the original second paragraph.\n' + '模拟长正文。' * 300
        messages = [incoming(邮件正文=cls.body, 邮件时间='2026-10-19T11:20:00', 邮件摘要='模拟摘要一',
                             自定义字段='Mail extra', _notes={'邮件正文': '模拟原文备注'}),
                    incoming(2, 邮件正文='SIM-ALPHA-UNKNOWN-TIME-BODY', 邮件原始时间='10月19日（年份待核实）'),
                    incoming(3, creator_id='SIM-000002', 邮件正文='SIM-BETA-PRIVATE-BODY',
                             发件人='sim_beta@example.invalid', 邮件时间='2026-10-18T09:00:00')]
        cls.expected = state([first, second], messages); app.prepare(cls.expected)
        manifest = cls.work / 'sim.json'; manifest.write_text(json.dumps(cls.expected, ensure_ascii=False), encoding='utf-8')
        cls.book = cls.work / 'sim.xlsx'
        subprocess.run([str(app.NODE), str(app.ROOT / 'scripts/write_workbook.mjs'), str(manifest), str(cls.book)],
                       check=True, capture_output=True, text=True)
        app.patch_groups(cls.book, cls.expected)
        cls.actual = app.read_book(cls.book, cls.work / 'read', extract_images=True)

    @classmethod
    def tearDownClass(cls): cls.temp.cleanup()

    def test_dates_preserve_wall_clock_precision(self):
        row = self.actual['sheets'][0]['rows'][0]
        self.assertEqual(row['首次联系日期'], '2026-10-19')
        self.assertEqual(row['最近收信时间'], '2026-10-19T00:00:00')
        self.assertEqual(row['最近发信时间'], '2026-10-19T10:00:00')
        self.assertEqual(self.actual['sheets'][1]['rows'][0]['邮件时间'], '2026-10-19T11:20:00')

    def test_original_text_notes_and_extra_fields_survive(self):
        self.assertEqual(app.validate(self.actual), {'creators': 2, 'messages': 3, 'screenshots': 0, 'drafts': 0, 'unassigned': 0})
        first = self.actual['sheets'][0]['rows'][0]; mail = self.actual['sheets'][1]['rows'][0]
        self.assertEqual(first['自定义字段'], 'Extra field')
        self.assertEqual(first['_notes']['首联内容分类'], '模拟来源，待人工核实。')
        self.assertEqual(mail['邮件正文'], self.body)
        self.assertEqual(mail['_notes']['邮件正文'], '模拟原文备注')
        self.assertEqual(mail['自定义字段'], 'Mail extra')

    def test_no_images_or_removed_archive_columns(self):
        with ZipFile(self.book) as z:
            self.assertFalse(any(n.startswith('xl/media/') for n in z.namelist()))
            self.assertFalse(any(n.startswith('xl/drawings/drawing') and n.endswith('.xml') for n in z.namelist()))
        self.assertFalse(set(self.actual['sheets'][1]['headers']) & app.REMOVED_MAIL_FIELDS)

    def test_current_draft_survives_refresh_without_becoming_mail_history(self):
        snapshot=copy.deepcopy(self.actual)
        creators,messages=app.prepare(snapshot)
        app.refresh(creators,messages,today='2026-10-19')
        self.assertEqual(creators[0]['当前回复草稿'],'SIM-UNSENT-WORKSPACE-BODY\n\nKeep draft paragraphs.')
        self.assertEqual(creators[0]['草稿状态'],'未发送草稿')
        self.assertEqual(len(messages),3)
        for output in (app.mail_text.overview(creators,messages),app.mail_text.context(creators,messages,creators[0]['达人ID'])):
            self.assertNotIn('SIM-UNSENT-WORKSPACE-BODY',json.dumps(output,ensure_ascii=False))
        with ZipFile(self.book) as z:
            root=E.fromstring(z.read('xl/worksheets/sheet1.xml'))
            draft_col=app.MASTER.index('当前回复草稿')+1
            column=next(c for c in root.findall('s:cols/s:col',app.NS) if int(c.get('min'))==draft_col)
            self.assertEqual(column.get('hidden'),'1')
            self.assertEqual(column.get('outlineLevel'),'2')

    def test_details_and_technical_columns_are_nested_collapsed(self):
        with ZipFile(self.book) as z: root = E.fromstring(z.read('xl/worksheets/sheet2.xml'))
        cols = {}
        for element in root.findall('s:cols/s:col', app.NS):
            for number in range(int(element.get('min')), int(element.get('max')) + 1): cols[number] = element.attrib
        for number in range(11, 18):
            self.assertEqual(cols[number].get('hidden'), '1'); self.assertEqual(cols[number].get('outlineLevel'), '1')
        for number in range(18, len(self.actual['sheets'][1]['headers']) + 1):
            self.assertEqual(cols[number].get('hidden'), '1'); self.assertEqual(cols[number].get('outlineLevel'), '2')
        self.assertEqual(cols[10].get('collapsed'), '1'); self.assertEqual(cols[17].get('collapsed'), '1')
        self.assertEqual(root.find('s:sheetPr/s:outlinePr', app.NS).get('summaryRight'), '0')

    def test_compact_rows_and_frozen_identity(self):
        with ZipFile(self.book) as z:
            for number, identity_columns in ((1, 2), (2, 3)):
                root = E.fromstring(z.read('xl/worksheets/sheet%d.xml' % number))
                pane = root.find('s:sheetViews/s:sheetView/s:pane', app.NS)
                self.assertEqual(float(pane.get('xSplit')), identity_columns)
                self.assertEqual(float(pane.get('ySplit')), 1)
                if number == 2:
                    heights = [float(r.get('ht')) for r in root.findall('s:sheetData/s:row', app.NS) if r.get('r') != '1' and r.get('ht')]
                    self.assertTrue(heights); self.assertLessEqual(max(heights), 60)

    def test_read_interfaces_preserve_book_and_isolate_bodies(self):
        before = hashlib.sha256(self.book.read_bytes()).hexdigest()
        def read_command(*args):
            return subprocess.run([sys.executable, str(app.ROOT / 'scripts/workflow.py'), '--book', str(self.book)] + list(args),
                                  check=True, capture_output=True, text=True).stdout
        overview = read_command('overview')
        self.assertNotIn('SIM-ALPHA-UNKNOWN-TIME-BODY', overview); self.assertNotIn('SIM-BETA-PRIVATE-BODY', overview)
        self.assertIn('模拟摘要一', overview)
        context = read_command('context', 'SIM-000001')
        self.assertIn('SIM-ALPHA-UNKNOWN-TIME-BODY', context); self.assertIn('Literal mail text', context)
        self.assertNotIn('SIM-BETA-PRIVATE-BODY', context)
        self.assertEqual(hashlib.sha256(self.book.read_bytes()).hexdigest(), before)

    def test_atomic_save_preserves_text(self):
        destination = self.work / 'atomic.xlsx'; destination.write_bytes(self.book.read_bytes())
        before = hashlib.sha256(destination.read_bytes()).hexdigest()
        report = app.write_book(destination, copy.deepcopy(self.actual), self.work / 'save', before)
        self.assertEqual(report['screenshots'], 0)
        self.assertEqual(app.read_book(destination, self.work / 'saved')['sheets'][1]['rows'][0]['邮件正文'], self.body)


if __name__ == '__main__': unittest.main()

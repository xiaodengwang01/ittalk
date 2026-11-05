"""Regression coverage for the approved visible-only main tracker. Synthetic data only."""
from pathlib import Path
import unittest
from zipfile import ZipFile
import xml.etree.ElementTree as ET

from scripts import test_compact_followup as compact
import slim_workbook as slim
import workbook_metadata as meta
c,w,post,support=compact.c,compact.w,compact.post,compact.support


class SlimTests(unittest.TestCase):
    setUpClass = classmethod(compact.CompactTests.setUpClass.__func__)
    tearDownClass = classmethod(compact.CompactTests.tearDownClass.__func__)
    setUp = compact.CompactTests.setUp
    tearDown = compact.CompactTests.tearDown

    def convert(self, *, first_date=''):
        w.update_fields(self.book,'SIM-000001',{'历史批次':'2026-10-22 新增KOL（99位）',
                        '初次建联日期':first_date,'备注':'人工旧备注'},expected_hash=c.mail.digest(self.book.read_bytes()))
        backup=self.root/'backup.xlsx';backup.write_bytes(self.book.read_bytes())
        return slim.migrate(self.state,self.book,backup,expected_hash=c.mail.digest(backup.read_bytes()),work=self.root)

    def test_migration_keeps_secondary_and_preserves_deleted_fields_locally(self):
        self.convert(first_date='2026-10-21')
        before=self.root/'backup.xlsx'
        book=w.read(self.book);row=book['main_creators'][0]
        self.assertEqual(book['headers'],slim.HEADERS)
        self.assertEqual(row['新增批次'],'2026-10-22')
        self.assertNotIn('初次建联日期',row)
        self.assertEqual(meta.overlay(self.state,self.book,row)['初次建联日期'],'2026-10-21')
        self.assertEqual(meta.overlay(self.state,self.book,row)['备注'],'人工旧备注')
        with ZipFile(before) as a,ZipFile(self.book) as b:
            self.assertEqual(a.read(w.main_part(a,w.SHEETS[1])),b.read(w.main_part(b,w.SHEETS[1])))
            root=ET.fromstring(b.read(w.main_part(b)))
            self.assertEqual(root.find('s:autoFilter',w.NS).get('ref'),'A1:K2')
            self.assertEqual(root.find('s:sheetViews/s:sheetView/s:pane',w.NS).get('xSplit'),'3')
            self.assertFalse(any(x.get('hidden')=='1' for x in root.findall('s:cols/s:col',w.NS)))
        with self.assertRaisesRegex(c.mail.DraftError,'已有初次建联日期'):
            c.mail.fill_text(meta.overlay(self.state,self.book,row),{},'')

    def test_new_row_and_verified_send_do_not_reintroduce_deleted_columns(self):
        self.convert()
        candidate=support.candidate(3);candidate['达人ID']='SIM-000003'
        value={'date':'2026-10-22','batch':'sim-new','reserved':[candidate],'creators':[]}
        c.daily.recover_reserved(self.state,value,self.book)
        sheet,row=w.locate(self.book,'SIM-000003')
        self.assertEqual(sheet['headers'],slim.HEADERS)
        self.assertEqual(row['新增批次'],'2026-10-22')
        self.assertEqual(meta.overlay(self.state,self.book,row)['内容能力(1-5)'],4)
        event={'account':'team@example.invalid','creator_id':'SIM-000003','message_id':'<new@example.invalid>',
               'sent_at':'2026-10-23T11:00:00+08:00','to':row['Email'],'subject':'Simulated first contact',
               'in_reply_to':'','raw_text':'Hello from a synthetic test.'}
        entry={'creator_id':'SIM-000003','batch_id':'sim-new'}
        for _ in range(2):post.write_workbook(self.state,c.sender,{'account':event['account']},entry,event,self.book)
        sheet,row=w.locate(self.book,'SIM-000003')
        self.assertEqual(sheet['headers'],slim.HEADERS)
        self.assertEqual(row['新增批次'],'2026-10-22')
        self.assertEqual(meta.overlay(self.state,self.book,row)['初次建联日期'],'2026-10-23')
        self.assertEqual(row['当前建联状况'],'初次建联未回复')
        self.assertEqual(row['往来邮件梗概'].count('2026-10-23｜我方：'),1)
        with ZipFile(self.book) as z:
            root=ET.fromstring(z.read(w.main_part(z)))
            self.assertEqual(root.find('s:autoFilter',w.NS).get('ref'),'A1:K3')
            date=root.find(".//s:c[@r='C3']",w.NS)
            self.assertIn(date.get('t'),(None,'n'))
            self.assertGreater(float(date.find('s:v',w.NS).text),40000)

    def test_conflicting_first_contact_preserves_workbook_and_archive(self):
        self.convert(first_date='2026-10-20');before=self.book.read_bytes()
        event={'account':'team@example.invalid','creator_id':'SIM-000001','message_id':'<conflict@example.invalid>',
               'sent_at':'2026-10-23T11:00:00+08:00','to':'sim1@example.invalid','subject':'Simulated',
               'in_reply_to':'','raw_text':'Synthetic'}
        with self.assertRaisesRegex(c.mail.DraftError,'首联日期'):
            post.write_workbook(self.state,c.sender,{'account':event['account']},
                               {'creator_id':'SIM-000001','batch_id':'sim-new'},event,self.book)
        self.assertEqual(self.book.read_bytes(),before)

    def test_main_review_summary_deduplicates_without_notes_column(self):
        self.convert()
        from docx import Document
        incoming={**self.incoming,'creator_id':'SIM-000001','message_id':'<main-in@example.invalid>'}
        c.ledger.incoming(self.state,incoming,'SIM-000001')
        source=self.root/'mentor-main.docx';doc=Document()
        doc.add_heading('模拟达人 1',1);doc.add_paragraph('达人ID：SIM-000001')
        for title,text in [('对方来信原文',incoming['raw_text']),('来信要点','询问试用周期。'),
                           ('我方回复要点及应对策略','提议一周试用。'),('我方完整回复正文','We suggest a one-week trial.')]:
            doc.add_heading(title,2)
            if title=='我方完整回复正文':
                doc.add_paragraph('收件人：sim1@example.invalid');doc.add_paragraph('主题：Re: simulated trial')
            doc.add_paragraph(text)
        doc.save(source)
        from unittest.mock import patch
        original=c.atomic_bytes
        def crash_after_replace(path,data):
            original(path,data)
            raise OSError('simulated crash after publishing the workbook')
        for attempt in range(3):
            updates={'source_hash':c.mail.digest(source.read_bytes()),'book_hash':c.mail.digest(self.book.read_bytes()),
                     'creators':{'SIM-000001':{'fields':{'当前建联状况':'回复稿待确认','待发送回复梗概':'未发送｜提议一周试用。'},
                       'incoming_summary':'询问试用周期。','message_basis':c.basis(self.state,'SIM-000001')}}}
            if attempt==0:
                with patch.object(c,'atomic_bytes',side_effect=crash_after_replace),self.assertRaises(OSError):
                    c.adopt_document(self.state,source,self.book,updates)
            else:
                c.adopt_document(self.state,source,self.book,updates)
        _,row=w.locate(self.book,'SIM-000001')
        self.assertEqual(row['往来邮件梗概'].count('询问试用周期。'),1)
        self.assertEqual(meta.overlay(self.state,self.book,row)['待发送回复梗概'],'未发送｜提议一周试用。')
        self.assertNotIn('备注',row)


if __name__=='__main__':unittest.main()

"""Synthetic archive and bookmark checks; no live state or mailbox access."""
import json
from pathlib import Path
import tempfile
import unittest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from scripts import mail_drafts as mail
from scripts import daily_workflow as daily


class ResourceTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)
        self.state = mail.State(self.folder/'state.sqlite3')
        daily.initialize(self.state)

    def tearDown(self):
        self.state.close(); self.temp.cleanup()

    def handbook(self):
        doc = Document(); expected = {}
        names = ('TAILKI_FIRST_PET_STORY','TAILKI_FIRST_PET_PRODUCT','TAILKI_FIRST_GENERAL')
        for index, (category,name) in enumerate(zip(mail.CATEGORIES,names)):
            doc.add_heading('中文策略不应进入邮件',1)
            subject = f'Test {index}'; body = 'Hi {{邮箱前缀}},\n\nFull approved body.\nBest, Sara'
            p = doc.add_paragraph('主题：'+subject)
            start = OxmlElement('w:bookmarkStart');start.set(qn('w:id'),str(index));start.set(qn('w:name'),name);p._p.insert(0,start)
            for text in body.split('\n\n'): p=doc.add_paragraph(text)
            end=OxmlElement('w:bookmarkEnd');end.set(qn('w:id'),str(index));p._p.append(end)
            doc.add_paragraph('中文说明，以及其他英文回信案例。')
            expected[category]={'subject':subject,'body':body}
        path=self.folder/'handbook.docx';doc.save(path)
        return path,expected

    def test_bookmarks_extract_only_three_approved_payloads(self):
        path,expected=self.handbook()
        self.assertEqual(mail.load_templates(path,self.state),expected)

    def test_partial_bookmark_set_does_not_fall_back_to_unbounded_parser(self):
        path,_=self.handbook();doc=Document(path)
        starts=list(doc.element.iter(qn('w:bookmarkStart')))
        starts[1].getparent().remove(starts[1]);doc.save(path)
        with self.assertRaises(mail.DraftError):mail.load_templates(path,self.state)

    def test_archived_source_keeps_frozen_batch_and_exact_hash(self):
        path,expected=self.handbook();book=self.folder/'book.xlsx';book.write_bytes(b'simulated book')
        batch={'id':'SIM-ARCHIVE','sources':{'book':str(book),'book_hash':mail.digest(book.read_bytes()),
              'template':str(path),'template_hash':mail.digest(path.read_bytes())}}
        raw=json.dumps(batch)
        self.state.db.execute('INSERT INTO batches VALUES (?,?)',(batch['id'],raw));self.state.db.commit()
        target=self.folder/'history'/path.name
        record=mail.local_resources.archive(self.state,path,target)
        self.assertFalse(path.exists());self.assertEqual(mail.local_resources.checksum(target),record['hash'])
        mail.source_check(self.state,batch['id'])
        self.assertEqual(mail.load_templates(path,self.state),expected)
        self.assertEqual(self.state.db.execute('SELECT data FROM batches').fetchone()[0],raw)
        target.write_bytes(b'tampered archive')
        with self.assertRaises(mail.DraftError):mail.source_check(self.state,batch['id'])

    def test_archive_never_overwrites_even_identical_duplicate(self):
        source=self.folder/'a.docx';target=self.folder/'b.docx'
        source.write_bytes(b'same');target.write_bytes(b'same')
        with self.assertRaises(ValueError):mail.local_resources.archive(self.state,source,target)
        self.assertTrue(source.exists());self.assertEqual(target.read_bytes(),b'same')

    def test_daily_dependencies_are_book_handbook_and_skill_only(self):
        path,_=self.handbook();book=self.folder/'book.xlsx';book.write_bytes(b'simulated')
        skill=self.folder/'SKILL.md';skill.write_text('test skill')
        result=daily.source_fingerprints(book,path,skill)
        self.assertEqual(set(result),{'book','template','skill'})

    def test_round_archives_only_after_matching_verified_send_and_completed_postprocess(self):
        from scripts import review_center as center
        from email.message import EmailMessage
        center.initialize(self.state)
        book=self.folder/'book.xlsx';book.write_bytes(b'simulated workbook')
        document,_=self.handbook()
        message=EmailMessage();message['From']='team@example.invalid';message['To']='sim@example.invalid'
        message['Subject']='Re: synthetic';message.set_content('Complete synthetic reply.')
        report={'round_id':'SIM-ROUND','status':'active','book':str(book.resolve()),'document':str(document),
                'document_hash':mail.digest(document.read_bytes()),'creators':['SIM-001'],
                'basis':{'SIM-001':{'message_id':'<in@example.invalid>'}},
                'drafts':{'SIM-001':{'recipient':'sim@example.invalid','subject':'Re: synthetic',
                                   'blocks':[{'text':'Complete synthetic reply.'}]}}}
        daily.put(self.state,center.report_key('SIM-ROUND'),report)
        review={'id':'SIM-REVIEW','entries':[{'draft_id':'SIM-DRAFT','raw_b64':center.sender.base64.b64encode(message.as_bytes()).decode()}],
                'postprocess':{'book':str(book.resolve()),'items':[{'creator_id':'SIM-001','status':'pending'}]}}
        center.sender.put_review(self.state,review)
        event={'account':'team@example.invalid','direction':'outgoing','creator_id':'SIM-001',
               'message_id':'<out@example.invalid>','source':'Sent:UID=SIM','sent_at':'2026-10-22T12:00:00+08:00',
               'verification':'sent_folder','review_id':'SIM-REVIEW','draft_id':'SIM-DRAFT',
               'subject':'Re: synthetic','to':'sim@example.invalid','in_reply_to':'<in@example.invalid>',
               'raw_text':'Complete synthetic reply.'}
        center.ledger.put(self.state,event)
        self.state.db.execute('INSERT INTO send_attempts VALUES (?,?,?)',('SIM-DRAFT','SIM-REVIEW',json.dumps({'review_id':'SIM-REVIEW','status':'uncertain'})))
        self.state.db.commit()
        self.assertEqual(center.archive_completed_reports(self.state,book),[])
        self.state.db.execute('UPDATE send_attempts SET data=?',(json.dumps({'review_id':'SIM-REVIEW','status':'sent_verified'}),));self.state.db.commit()
        self.assertEqual(center.archive_completed_reports(self.state,book),[])
        review['postprocess']['items'][0]['status']='done';center.sender.put_review(self.state,review)
        before=list(self.state.db.execute('SELECT * FROM send_attempts'))
        original=document.read_bytes()
        result=center.archive_completed_reports(self.state,book)
        self.assertEqual(len(result),1);self.assertFalse(document.exists())
        self.assertEqual(Path(result[0]['archive']).read_bytes(),original)
        self.assertEqual(list(self.state.db.execute('SELECT * FROM send_attempts')),before)
        self.assertEqual(center.archive_completed_reports(self.state,book),[])


if __name__=='__main__':unittest.main()

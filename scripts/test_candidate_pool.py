"""Synthetic pool, quota, live-history and interruption checks; no real mailbox."""
from copy import deepcopy
from datetime import timedelta
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from scripts import candidate_pool as pool
from scripts import test_daily_workflow as fixtures
from scripts.test_mail_drafts import FakeMailbox, write_template

d = pool.daily
DAY = '2026-10-22'


class PoolTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.shared = tempfile.TemporaryDirectory()
        cls.base = Path(cls.shared.name) / 'fixture.xlsx'
        fixtures.fixture(cls.base)

    @classmethod
    def tearDownClass(cls): cls.shared.cleanup()

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.book = self.root / 'book.xlsx';self.book.write_bytes(self.base.read_bytes())
        self.template = self.root / 'template.docx';write_template(self.template)
        self.state = d.mail.State(self.root / 'private/state.sqlite3')
        pool.initialize(self.state)
        self.state.set_config({'account': 'team@example.invalid'})
        d.put(self.state, 'activation', {'status':'active', 'first_run':'2026-10-22T10:00:00+08:00'})
        self.clock = patch.object(d, 'current_time', return_value=fixtures.NOW);self.clock.start()
        self.box = FakeMailbox();self.box.account = 'team@example.invalid'

    def tearDown(self):
        self.clock.stop();self.state.close();self.temp.cleanup()

    def ingest(self, candidates=None, routes=None, checked=None):
        return pool.ingest(self.state, {'creators': candidates if candidates is not None else [fixtures.candidate(i) for i in range(2, 10)],
            'routes_checked': routes if routes is not None else ['FastMoss','TikTok'],
            'checked_at': checked or fixtures.NOW.isoformat()}, self.book)

    def select(self):
        return pool.select_day(self.state, DAY, self.box, self.book, self.template)

    def test_selects_five_and_keeps_other_candidates_and_complete_preview(self):
        self.ingest()
        before = self.book.read_bytes()
        with ZipFile(self.book) as z: parts = {n:z.read(n) for n in z.namelist()}
        result = self.select()
        self.assertEqual(result['limit'],5)
        self.assertEqual(len(result['selection']),5)
        self.assertEqual(result['shortfall'],0)
        self.assertEqual(result['pool']['counts'],{'selected':5,'available':3})
        items = self.state.items(result['preparation']['batch'])
        self.assertEqual(len(items),5)
        page = Path(result['preparation']['preview']).read_text()
        for item in items:
            self.assertIn(item['recipient'],page)
            self.assertIn(item['body'],page)
            self.assertNotIn('{{',item['body'])
            self.assertEqual(item['status'],'prepared')
        with ZipFile(self.book) as z:
            self.assertEqual([n for n in parts if parts[n]!=z.read(n)],['xl/worksheets/sheet1.xml'])
        row = d.workbook.read(self.book)['creators'][0]
        self.assertEqual(row['对方第一次回复邮件原始文本'],'=literal@example.invalid')
        self.assertNotEqual(before,self.book.read_bytes())
        self.assertIsNone(d.get(self.state,'approval:'+result['preparation']['batch']))

    def test_ingest_is_idempotent_canonicalizes_profile_and_does_not_edit_workbook(self):
        before = self.book.read_bytes()
        one = fixtures.candidate(2);two=deepcopy(one)
        two['账号链接']='https://tiktok.com/@sim_2/'
        self.ingest([one]);self.ingest([two])
        self.assertEqual(pool.summary(self.state)['total'],1)
        self.assertEqual(self.book.read_bytes(),before)

    def test_shared_email_holds_both_profiles_without_merging_or_selecting(self):
        one,two=fixtures.candidate(2),fixtures.candidate(3)
        two['Email']=one['Email']
        self.ingest([one,two])
        self.assertEqual(pool.summary(self.state)['counts'],{'held_shared_email':2})
        result=self.select()
        self.assertEqual(result['selection'],[])
        self.assertEqual(result['shortfall'],5)

    def test_current_workbook_and_table_external_sent_contacts_are_excluded(self):
        one=fixtures.candidate(2);one['账号链接']='https://tiktok.com/@sim_1'
        d.put(self.state,'history:2026-10-21',{'sent_contacts':[{'email':'sim3@example.invalid'}]})
        self.ingest([one,fixtures.candidate(3),fixtures.candidate(4)])
        self.assertEqual(pool.summary(self.state)['counts'],{'excluded_contacted':2,'available':1})
        self.assertEqual(len(self.select()['selection']),1)

    def test_live_incoming_sent_or_draft_history_excludes_and_fills_from_remaining_pool(self):
        self.ingest()
        self.box.histories['sim2@example.invalid']=[{'folder':'INBOX','uid':'sim-in'}]
        self.box.histories['sim3@example.invalid']=[{'folder':'Sent','uid':'sim-out'}]
        existing=self.box.existing_drafts
        self.box.existing_drafts=lambda email: ['sim-draft'] if email=='sim4@example.invalid' else existing(email)
        result=self.select()
        self.assertEqual(len(result['selection']),5)
        self.assertEqual(result['pool']['counts'],{'selected':5,'excluded_contacted':3})
        recipients={i['recipient'] for i in self.state.items(result['preparation']['batch'])}
        self.assertFalse(recipients & {'sim2@example.invalid','sim3@example.invalid','sim4@example.invalid'})

    def test_old_research_is_not_eligible_without_real_current_day_recheck(self):
        old=fixtures.candidate(2)
        old['核实时间']=old['历史核对时间']='2026-10-21T09:05:00+08:00'
        self.ingest([old])
        self.assertEqual(pool.summary(self.state)['available_today'],0)
        self.assertEqual(self.select()['selection'],[])
        self.ingest([fixtures.candidate(2)])
        self.assertEqual(len(self.select()['selection']),1)

    def test_missing_fields_or_bad_scores_are_retained_but_not_selected(self):
        one,two=fixtures.candidate(2),fixtures.candidate(3)
        one['Email']='';two['内容能力(1-5)']='unknown'
        self.ingest([one,two])
        self.assertEqual(pool.summary(self.state)['counts'],{'needs_info':2})
        self.assertEqual(self.select()['shortfall'],5)

    def test_both_search_routes_are_required_and_future_route_time_is_rejected_before_insert(self):
        self.ingest([fixtures.candidate(2)],routes=['TikTok'])
        with self.assertRaises(d.mail.DraftError):self.select()
        before=len(pool.entries(self.state))
        with self.assertRaises(d.mail.DraftError):self.ingest([fixtures.candidate(3)],checked=(fixtures.NOW+timedelta(days=1)).isoformat())
        self.assertEqual(len(pool.entries(self.state)),before)

    def test_repeated_or_frozen_selection_never_adds_or_submits_new_mail(self):
        self.ingest([fixtures.candidate(2),fixtures.candidate(3)])
        one=self.select();frozen=self.book.read_bytes()
        self.ingest([fixtures.candidate(i) for i in range(4,10)])
        two=self.select()
        self.assertTrue(two['frozen'])
        self.assertEqual(one['selection'],two['selection'])
        self.assertEqual(self.book.read_bytes(),frozen)
        self.assertEqual(len(self.state.items(one['preparation']['batch'])),2)

    def test_ranking_uses_content_and_match_before_followers_or_sales(self):
        candidates=[fixtures.candidate(i) for i in range(2,8)]
        candidates[0].update({'内容能力(1-5)':1,'粉丝数':999999999})
        candidates[-1].update({'内容能力(1-5)':5,'粉丝数':10})
        self.ingest(candidates)
        self.select()
        selected={i['email'] for i in pool.entries(self.state) if i['status']=='selected'}
        self.assertNotIn('sim2@example.invalid',selected)
        self.assertIn('sim7@example.invalid',selected)

    def test_crash_after_file_write_restores_pool_links_without_duplicate_rows(self):
        self.ingest()
        original=pool.link_selected;calls=[]
        def crash_after_write(state,value,book):
            calls.append(1)
            if len(calls)==2:raise KeyboardInterrupt()
            return original(state,value,book)
        with patch.object(pool,'link_selected',side_effect=crash_after_write):
            with self.assertRaises(KeyboardInterrupt):self.select()
        written=self.book.read_bytes()
        result=self.select()
        self.assertEqual(len(result['selection']),5)
        self.assertEqual(self.book.read_bytes(),written)
        self.assertEqual(result['pool']['counts'],{'selected':5,'available':3})

    def test_reservation_before_write_is_durable_and_ingest_cannot_change_it(self):
        self.ingest([fixtures.candidate(i) for i in range(2,7)])
        with patch.object(d,'add_candidates',side_effect=KeyboardInterrupt()):
            with self.assertRaises(KeyboardInterrupt):self.select()
        protected=pool.entries(self.state)
        changed=fixtures.candidate(2);changed['Email']='new@example.invalid'
        self.ingest([changed])
        self.assertEqual(next(x for x in pool.entries(self.state) if x['profile'].endswith('@sim_2'))['email'],'sim2@example.invalid')
        result=self.select()
        self.assertEqual(len(result['selection']),len(protected))

    def test_old_eight_person_history_is_retained_and_negative_slice_cannot_add_more(self):
        value=d.new_day(DAY);value['creators']=[f'SIM-{i:06d}' for i in range(1,9)]
        d.put(self.state,'day:'+DAY,value)
        before=self.book.read_bytes()
        result=d.add_candidates(self.state,DAY,{'routes_checked':['FastMoss','TikTok'],'creators':[fixtures.candidate(20)]},self.book)
        self.assertEqual(len(result['added']),8)
        self.assertEqual(result['shortfall'],0)
        self.assertEqual(self.book.read_bytes(),before)

    def test_historical_contact_stays_excluded_even_after_email_changes(self):
        one=fixtures.candidate(2);one['已联系']=True
        self.ingest([one])
        one.update(已联系=False,Email='different@example.invalid')
        self.ingest([one])
        self.assertEqual(pool.summary(self.state)['counts'],{'excluded_contacted':1})

    def test_live_history_checks_old_email_alias_after_public_email_changes(self):
        self.ingest([fixtures.candidate(2)])
        changed=fixtures.candidate(2);changed['Email']='updated@example.invalid'
        self.ingest([changed])
        self.box.histories['sim2@example.invalid']=[{'folder':'Sent','uid':'sim-old-alias'}]
        self.assertEqual(self.select()['selection'],[])
        self.assertEqual(pool.summary(self.state)['counts'],{'excluded_contacted':1})

    def test_network_failure_does_not_write_rows_or_freeze_partial_selection(self):
        self.ingest()
        before=self.book.read_bytes()
        original=self.box.history
        def interrupted(email):
            if email=='sim3@example.invalid':raise TimeoutError()
            return original(email)
        with patch.object(self.box,'history',side_effect=interrupted):
            with self.assertRaises(TimeoutError):self.select()
        self.assertEqual(self.book.read_bytes(),before)
        self.assertFalse(any(x['status']=='reserved' for x in pool.entries(self.state)))
        self.assertEqual(len(self.select()['selection']),5)

    def test_scheduled_pool_selection_still_waits_until_ten(self):
        self.ingest()
        with patch.object(d,'current_time',return_value=fixtures.NOW.replace(hour=9,minute=59)):
            with self.assertRaises(d.mail.DraftError):self.select()
        self.assertEqual(self.state.db.execute('SELECT count(*) FROM drafts').fetchone()[0],0)

    def test_research_queue_is_bounded_and_preserves_past_quality_reasons(self):
        self.ingest()
        bad={'账号链接':'https://tiktok.com/@sim_bad','合格':False,'不合格原因':'模拟：近期未更新，暂缓。'}
        self.ingest([bad])
        self.assertEqual(pool.summary(self.state)['counts'].get('excluded_quality'),1)
        queue=pool.research(self.state,3)
        self.assertEqual(len(queue['candidates']),3)
        self.assertEqual(queue['remaining'],5)
        with self.assertRaises(d.mail.DraftError):pool.research(self.state,100)

    def test_only_verified_send_marks_matching_selected_creator_and_never_reenters_pool(self):
        self.ingest([fixtures.candidate(2)])
        result=self.select()
        identifier=result['selection'][0]
        event={'message_id':'<sim-sent@example.invalid>','sent_at':fixtures.NOW.isoformat(),
               'verification':'uncertain','to':'sim2@example.invalid'}
        with self.assertRaises(d.mail.DraftError):pool.mark_sent(self.state,identifier,event)
        self.assertEqual(pool.summary(self.state)['counts'],{'selected':1})
        event['verification']='sent_folder'
        pool.mark_sent(self.state,identifier,event)
        self.assertEqual(pool.summary(self.state)['counts'],{'sent_verified':1})
        self.ingest([fixtures.candidate(2)])
        self.assertEqual(pool.summary(self.state)['counts'],{'sent_verified':1})


if __name__=='__main__':unittest.main()

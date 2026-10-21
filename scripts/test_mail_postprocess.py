"""Synthetic post-send checks; no real mailbox, SMTP, or business workbook."""
from copy import deepcopy
import json
from pathlib import Path
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from scripts import mail_send as s
from scripts import test_mail_send as send_support
from scripts.test_mail_send import FakeTransport
from scripts.test_mail_drafts import FakeIMAP
from scripts.test_daily_workflow import fixture
import mail_postprocess as p
import mail_records as records


class TrashIMAP(FakeIMAP):
    capabilities = (b'MOVE', b'UIDPLUS')
    lose_move_response = False
    copy_removes_source = False

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.data[b'Trash'] = {}
        self.mutations = []

    def list(self):
        code, folders = super().list()
        return code, folders + [b'(\\Trash) "/" "Trash"']

    def select(self, mailbox, *, readonly):
        self.selected = mailbox.strip(b'"')
        self.calls.append('EXAMINE' if readonly else 'SELECT')
        return 'OK', []

    def uid(self, operation, *arguments):
        if operation in ('MOVE', 'COPY'):
            uid, target = arguments
            assert self.selected == b'Drafts'
            self.mutations.append((operation, uid))
            self.data[target.strip(b'"')][uid] = self.data[b'Drafts'][uid]
            if operation == 'MOVE' or self.copy_removes_source:
                del self.data[b'Drafts'][uid]
            if self.lose_move_response:
                self.lose_move_response = False
                raise TimeoutError('simulated lost response')
            return 'OK', []
        if operation == 'STORE':
            self.mutations.append((operation, arguments[0]))
            return 'OK', []
        if operation == 'EXPUNGE':
            self.mutations.append((operation, arguments[0]))
            self.data[b'Drafts'].pop(arguments[0], None)
            return 'OK', []
        return super().uid(operation, *arguments)


class PostprocessTests(unittest.TestCase):
    setUp = send_support.SendTests.setUp
    tearDown = send_support.SendTests.tearDown
    review = send_support.SendTests.review
    approve = send_support.SendTests.approve
    edit = send_support.SendTests.edit

    def prepared(self, mode='success'):
        client = TrashIMAP('imap.qiye.aliyun.com', 993, ssl_context=s.ssl.create_default_context(), timeout=30)
        client.data.update(deepcopy(self.mailbox.client.data))
        self.mailbox.client = client
        review = self.review()
        self.approve(review)
        transport = FakeTransport(self.mailbox, mode=mode)
        s.deliver(self.state, review['id'], self.mailbox, lambda config: transport, pause=lambda _: None)
        self.book = self.folder / 'book.xlsx'
        self.book.write_bytes(b'simulated workbook')
        self.rows = [{'达人ID': f'SIM-{i:03d}', 'Email': f'sim{i}@example.invalid', '_row': i+1,
                      '联系状态': '待首联确认', '初次建联日期': '', '备注': '人工备注'} for i in (1, 2)]
        self.headers = ['达人ID', '达人昵称', '联系状态', '沟通摘要', 'Email', '首联内容分类', '初次建联日期',
                        '备注', '我方第一次回复邮件日期', '我们的第一次回复邮件原始文本', '我们的第一次回复邮件摘要']
        return s.get_review(self.state, review['id']), transport

    def fake_sheet(self, path):
        return {'creators': deepcopy(self.rows), 'headers': self.headers}

    def fake_patch(self, path, cells, **kwargs):
        self.assertEqual(s.mail.digest(self.book.read_bytes()), kwargs['expected_hash'])
        for address, value in cells.items():
            number = int(s.mail.re.search(r'\d+', address).group())
            header = self.headers[p.workbook.workflow.col_number(address)-1]
            next(r for r in self.rows if r['_row'] == number)[header] = value['value']
        self.book.write_bytes(self.book.read_bytes()+b'patched')

    def finish(self, review):
        with patch.object(p.workbook, 'read', side_effect=self.fake_sheet), patch.object(p.workbook, 'patch_cells', side_effect=self.fake_patch):
            return s.finalize(self.state, review['id'], self.mailbox, self.book)

    def test_success_updates_by_headers_preserves_human_state_and_moves_only_matching_drafts(self):
        review, transport = self.prepared()
        self.rows[0]['联系状态'] = '长期合作意愿'
        self.mailbox.client.data[b'Drafts'][b'90'] = b'unrelated untouched draft'
        result = self.finish(review)
        self.assertTrue(result['complete'])
        self.assertEqual(self.rows[0]['联系状态'], '长期合作意愿')
        self.assertEqual(self.rows[1]['联系状态'], '初次建联未回复')
        self.assertTrue(self.rows[0]['备注'].startswith('人工备注'))
        self.assertEqual(list(self.mailbox.client.data[b'Drafts']), [b'90'])
        self.assertEqual(len(self.mailbox.client.data[b'Trash']), 2)
        before = deepcopy(self.mailbox.client.mutations)
        self.assertTrue(self.finish(review)['complete'])
        self.assertEqual(self.mailbox.client.mutations, before)
        self.assertEqual(len(transport.submissions), 2)

    def test_missing_or_wrong_sent_copy_never_writes_or_cleans(self):
        review, transport = self.prepared()
        self.mailbox.client.data[b'Sent'].clear()
        self.mailbox._header_cache.clear()
        before = deepcopy(self.rows)
        self.assertFalse(self.finish(review)['complete'])
        self.assertEqual(self.rows, before)
        self.assertEqual(self.mailbox.client.mutations, [])
        self.assertEqual(len(transport.submissions), 2)

    def test_unknown_send_is_never_postprocessed(self):
        review, transport = self.prepared('unknown')
        self.assertFalse(self.finish(review)['complete'])
        self.assertEqual(self.mailbox.client.mutations, [])
        self.assertEqual(len(transport.submissions), 1)

    def test_workbook_conflict_preserves_draft_and_only_other_entry_finishes(self):
        review, _ = self.prepared()
        self.rows[0]['初次建联日期'] = '2020-01-01'
        result = self.finish(review)
        self.assertEqual(result['items'][0]['status'], 'pending')
        self.assertIn(b'1', self.mailbox.client.data[b'Drafts'])
        self.assertNotIn(b'2', self.mailbox.client.data[b'Drafts'])

    def test_changed_draft_is_preserved_after_workbook_write_and_never_resent(self):
        review, transport = self.prepared()
        self.edit(body='New unsent human work')
        result = self.finish(review)
        self.assertFalse(result['complete'])
        self.assertIn(b'1', self.mailbox.client.data[b'Drafts'])
        self.assertEqual(len(transport.submissions), 2)

    def test_lost_move_response_recovers_without_second_move_or_duplicate_note(self):
        review, transport = self.prepared()
        self.mailbox.client.lose_move_response = True
        self.assertFalse(self.finish(review)['complete'])
        before = deepcopy(self.rows)
        self.assertTrue(self.finish(review)['complete'])
        self.assertEqual(self.rows, before)
        self.assertEqual(self.mailbox.client.mutations.count(('MOVE', b'1')), 1)
        self.assertEqual(len(transport.submissions), 2)

    def test_crash_after_workbook_replacement_resumes_without_overwriting_or_resending(self):
        review, transport = self.prepared()
        original = records.action
        failed = []
        def fail_once(state, key, **updates):
            if key.startswith('workbook:') and updates.get('status') == 'done' and not failed:
                failed.append(key)
                raise OSError('simulated crash after replacement')
            return original(state, key, **updates)
        with patch.object(records, 'action', side_effect=fail_once):
            self.assertFalse(self.finish(review)['complete'])
        rows = deepcopy(self.rows)
        self.assertTrue(self.finish(review)['complete'])
        self.assertEqual(rows, self.rows)
        self.assertEqual(len(transport.submissions), 2)

    def test_pending_recovery_only_adopts_newly_approved_batches_and_never_sends(self):
        review, transport = self.prepared()
        with patch.object(p.workbook, 'read', side_effect=self.fake_sheet), patch.object(p.workbook, 'patch_cells', side_effect=self.fake_patch):
            self.assertEqual(len(s.finalize_pending(self.state, self.mailbox, self.book)['reviews']), 1)
            self.assertEqual(s.finalize_pending(self.state, self.mailbox, self.book)['reviews'], [])
        self.assertEqual(len(transport.submissions), 2)

    def test_copy_to_trash_provider_removes_source_without_store_or_expunge(self):
        review, _ = self.prepared()
        self.mailbox.client.capabilities = ()
        self.mailbox.client.copy_removes_source = True
        self.assertTrue(self.finish(review)['complete'])
        self.assertTrue(all(op == 'COPY' for op, _ in self.mailbox.client.mutations))

    def test_copy_with_uidplus_expunge_never_uses_global_expunge(self):
        review, _ = self.prepared()
        self.mailbox.client.capabilities = (b'UIDPLUS',)
        self.assertTrue(self.finish(review)['complete'])
        self.assertEqual(self.mailbox.client.mutations, [(op, uid) for uid in (b'1', b'2') for op in ('COPY', 'STORE', 'EXPUNGE')])

    def test_copy_without_uidplus_retains_source_and_does_not_recopy(self):
        review, _ = self.prepared()
        self.mailbox.client.capabilities = ()
        self.assertFalse(self.finish(review)['complete'])
        self.assertFalse(self.finish(review)['complete'])
        self.assertEqual(len(self.mailbox.client.mutations), 2)
        self.assertEqual(len(self.mailbox.client.data[b'Drafts']), 2)

    def test_account_and_uidvalidity_changes_prevent_cleanup(self):
        review, _ = self.prepared()
        entry = deepcopy(review['entries'][0])
        entry['uidvalidity'] = '99'
        with self.assertRaises(s.mail.DraftError): p.cleanup_draft(self.state, s, review, entry, self.mailbox)
        self.assertEqual(self.mailbox.client.mutations, [])
        self.mailbox.account = 'different@example.invalid'
        with self.assertRaises(s.mail.DraftError): s.finalize(self.state, review['id'], self.mailbox, self.book)

    def test_deleted_draft_is_never_rebuilt(self):
        review, _ = self.prepared()
        del self.mailbox.client.data[b'Drafts'][b'1']
        self.assertTrue(self.finish(review)['complete'])
        self.assertEqual(self.mailbox.client.mutations, [('MOVE', b'2')])

    def test_native_workbook_parts_styles_and_unrelated_cells_preserved(self):
        review, _ = self.prepared()
        fixture(self.book)
        entry = deepcopy(review['entries'][0])
        entry['creator_id'] = 'SIM-000001'
        event = p.verified_message(self.state, s, review, review['entries'][0], self.mailbox)
        with ZipFile(self.book) as z: before = {n: z.read(n) for n in z.namelist()}
        p.write_workbook(self.state, s, review, entry, event, self.book)
        with ZipFile(self.book) as z: after = {n: z.read(n) for n in z.namelist()}
        self.assertEqual(set(before), set(after))
        self.assertEqual([n for n in before if before[n] != after[n]], ['xl/worksheets/sheet1.xml'])
        sheet = p.workbook.read(self.book)
        self.assertEqual(sheet['creators'][0]['Email'], 'sim1@example.invalid')
        self.assertEqual(sheet['creators'][0]['对方第一次回复邮件原始文本'], '=literal@example.invalid')
        stable = self.book.read_bytes()
        p.write_workbook(self.state, s, review, entry, event, self.book)
        self.assertEqual(self.book.read_bytes(), stable)


class LedgerTests(unittest.TestCase):
    def test_duplicate_sources_merge_but_conflicting_body_or_creator_is_rejected(self):
        import tempfile
        with tempfile.TemporaryDirectory() as folder:
            state = s.mail.State(Path(folder) / 'private.sqlite3')
            try:
                record = {'account': 'sara@example.invalid', 'message_id': '<sim1@example.invalid>',
                          'direction': 'incoming', 'source': 'INBOX:1', 'creator_id': 'SIM-001'}
                records.put(state, record)
                records.put(state, {**record, 'source': 'INBOX:2', 'raw_text': 'New reply\n> quoted old text'})
                records.put(state, {**record, 'source': 'INBOX:3', 'body_status': 'not_collected'})
                context = records.context(state, 'SIM-001')
                self.assertEqual(context['incoming'], 1)
                self.assertEqual(len(context['messages'][0]['sources']), 3)
                self.assertEqual(context['messages'][0]['body_status'], 'collected')
                for changed in ({'raw_text': 'Different'}, {'creator_id': 'SIM-002'}):
                    with self.assertRaises(records.mail.DraftError): records.put(state, {**record, **changed})
                self.assertEqual(records.context(state, 'SIM-001')['incoming'], 1)
            finally: state.close()

    def test_context_orders_real_instants_across_timezones(self):
        import tempfile
        with tempfile.TemporaryDirectory() as folder:
            state = s.mail.State(Path(folder) / 'private.sqlite3')
            try:
                for mid, stamp in [('late', '2026-10-21T17:00:00+00:00'), ('early', '2026-10-22T00:30:00+08:00')]:
                    records.put(state, {'account': 'sara@example.invalid', 'message_id': mid, 'direction': 'incoming',
                                        'source': 'INBOX:'+mid, 'creator_id': 'SIM-001', 'received_at': stamp})
                self.assertEqual([m['message_id'] for m in records.context(state, 'SIM-001')['messages']], ['early', 'late'])
            finally: state.close()


if __name__ == '__main__': unittest.main()

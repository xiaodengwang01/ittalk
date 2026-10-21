"""Synthetic-only checks; no tests connect to SMTP or real mailboxes."""
import base64
from copy import deepcopy
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from scripts import mail_drafts as m
from scripts import mail_send as s
from scripts.test_mail_drafts import FakeIMAP


class FakeTransport:
    account = 'sara@example.invalid'

    def __init__(self, mailbox, *, mode='success'):
        self.mailbox, self.mode = mailbox, mode
        self.submissions = []
        self.closed = False

    def check(self):
        return {'smtp_authenticated': True, 'sent': False}

    def close(self):
        self.closed = True

    def submit(self, raw, recipients, stage):
        stage('envelope')
        stage('data')
        self.submissions.append((raw, recipients))
        if self.mode in ('success', 'accepted_disconnect'):
            sent = self.mailbox.client.data[b'Sent']
            sent[str(len(sent) + 1).encode()] = raw
        if self.mode == 'accepted_disconnect':
            raise s.DeliveryError('uncertain', 'Simulated connection loss after acceptance')
        if self.mode == 'unknown':
            raise s.DeliveryError('uncertain', 'Simulated connection loss')
        if self.mode == 'failed':
            raise s.DeliveryError('failed', 'Simulated refusal', 550)
        return {'smtp_code': 250, 'accepted_at': m.now()}


class SendTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)
        self.state = m.State(self.folder / 'private.sqlite3')
        self.state.set_config({'account': 'sara@example.invalid'})
        self.mailbox = m.Mailbox('sara@example.invalid', 'synthetic-secret', factory=FakeIMAP)
        s.initialize(self.state)
        self.preview = self.folder / 'outputs' / 'final.html'
        self.batch = 'SIM-BATCH'
        with self.state.db:
            self.state.db.execute('INSERT INTO batches VALUES (?,?)', (self.batch, m.json.dumps({'id': self.batch})))
        for index in (1, 2):
            item = {'id': f'sim-{index}', 'batch_id': self.batch, 'creator_id': f'SIM-{index:03d}',
                    'name': f'模拟达人 {index}', 'category': '宠物剧情类', 'recipient': f'sim{index}@example.invalid',
                    'subject': f'Final subject {index}', 'body': f'Complete body {index}\n\nBest,\nSara\nTailki',
                    'message_id': f'<sim-draft-{index}@example.invalid>', 'status': 'saved', 'reason': '',
                    'content_hash': m.digest([f'sim{index}@example.invalid', f'Final subject {index}', f'Complete body {index}\n\nBest,\nSara\nTailki'])}
            self.mailbox.append(m.make_message(item, self.mailbox.account))
            remote = self.mailbox.find(item)[0]
            item.update(remote)
            with self.state.db:
                self.state.db.execute('INSERT INTO drafts VALUES (?,?,?,?)',
                                      (item['id'], self.batch, item['id'], m.json.dumps(item)))

    def tearDown(self):
        self.mailbox.close()
        self.state.close()
        self.temp.cleanup()

    def review(self):
        result = s.prepare_review(self.state, self.mailbox, [self.batch], self.preview, expected_count=2)
        return s.get_review(self.state, result['review'])

    def approve(self, review):
        return s.approve(self.state, review['id'], {'user_instruction': '确认发送这两封模拟邮件',
                           'user_instruction_at': m.now(), 'review_hash': review['hash'], 'count': 2})

    def edit(self, uid=b'1', *, body=None, subject=None, recipient=None, attachment=None, html=None):
        message = s.parse(self.mailbox.client.data[b'Drafts'][uid])
        if subject is not None:
            message.replace_header('Subject', subject)
        if recipient is not None:
            message.replace_header('To', recipient)
        if body is not None:
            target = message.get_body(preferencelist=('plain',)) if message.is_multipart() else message
            target.set_content(body)
        if html is not None:
            message.add_alternative(html, subtype='html')
        if attachment is not None:
            message.add_attachment(attachment, maintype='application', subtype='octet-stream', filename='sample.bin')
        self.mailbox.client.data[b'Drafts'][uid] = message.as_bytes()

    def transport(self, mode='success'):
        return FakeTransport(self.mailbox, mode=mode)

    def send(self, review, transport):
        return s.deliver(self.state, review['id'], self.mailbox, lambda config: transport, pause=lambda seconds: None)

    def test_review_reads_full_current_drafts_and_preserves_human_edits(self):
        self.edit(body='Human final text\n\nBest,\nSara\nTailki', subject='Human subject')
        before = deepcopy(self.mailbox.client.data)
        review = self.review()
        self.assertEqual(review['status'], 'awaiting_confirmation')
        self.assertEqual(review['entries'][0]['details']['subject'], 'Human subject')
        self.assertTrue(review['entries'][0]['edited'])
        self.assertIn('Human final text', self.preview.read_text())
        self.assertIn('Complete body 2', self.preview.read_text())
        self.assertEqual(before, self.mailbox.client.data)
        self.assertEqual(self.state.db.execute('SELECT count(*) FROM send_attempts').fetchone()[0], 0)

    def web_resave(self, *, changed_body=None):
        message = s.parse(self.mailbox.client.data[b'Drafts'].pop(b'1'))
        del message['X-Tailki-Draft-ID']
        del message['X-Tailki-Batch-ID']
        message.replace_header('Message-ID', '<sim-provider-draft.>')
        message.replace_header('From', 'Webmail name <sara@example.invalid>')
        message.replace_header('To', 'sim1 <sim1@example.invalid>')
        message['Reply-To'] = 'Webmail name <sara@example.invalid>'
        if changed_body is not None:
            message.set_content(changed_body)
        self.mailbox.client.data[b'Drafts'][b'99'] = message.as_bytes()

    def test_webmail_resave_is_recovered_only_when_unique_and_complete_original_text_matches(self):
        self.web_resave()
        before = deepcopy(self.mailbox.client.data)
        review = self.review()
        entry = review['entries'][0]
        self.assertEqual(entry['uid'], '99')
        self.assertTrue(entry['web_resaved'])
        self.assertIn('Webmail name', self.preview.read_text())
        self.assertEqual(before, self.mailbox.client.data)
        raw, _ = s.outgoing(entry, review, 'sim-send', m.now())
        self.assertFalse(s.parse(raw)['Message-ID'].defects)
        self.assertTrue(self.state.items(self.batch)[0]['web_relinks'])
        self.approve(review)
        self.assertEqual(self.send(review, self.transport())['states'], {'sent_verified': 2})

    def test_replaced_identity_with_changed_or_ambiguous_body_is_never_silently_relinked(self):
        self.web_resave(changed_body='Human edit without stable identity')
        before = deepcopy(self.mailbox.client.data)
        with self.assertRaises(m.DraftError): self.review()
        self.assertEqual(before, self.mailbox.client.data)
        self.assertNotIn('web_relinks', self.state.items(self.batch)[0])
        message = s.parse(self.mailbox.client.data[b'Drafts'][b'99'])
        message.set_content(self.state.items(self.batch)[0]['body'])
        self.mailbox.client.data[b'Drafts'][b'99'] = message.as_bytes()
        self.mailbox.client.data[b'Drafts'][b'100'] = message.as_bytes()
        with self.assertRaises(m.DraftError): self.review()
        self.assertNotIn('web_relinks', self.state.items(self.batch)[0])

    def test_web_resave_after_approval_requires_fresh_review_before_any_submission(self):
        review = self.review()
        self.approve(review)
        self.web_resave()
        transport = self.transport()
        with self.assertRaises(m.DraftError): self.send(review, transport)
        self.assertEqual(transport.submissions, [])
        fresh = self.review()
        self.assertNotEqual(review['id'], fresh['id'])
        self.assertNotIn('approval', fresh)

    def test_provider_message_id_tolerance_does_not_allow_bad_addresses_or_duplicate_ids(self):
        raw = self.mailbox.client.data[b'Drafts'][b'1']
        for header, value in (('To', 'not-an-email'), ('From', 'not-an-email')):
            message = s.parse(raw)
            message.replace_header(header, value)
            with self.assertRaises(m.DraftError): s.parse(message.as_bytes())
        raw = raw.replace(b'Message-ID:', b'Message-ID: <other@example.invalid>\r\nMessage-ID:', 1)
        with self.assertRaises(m.DraftError): s.parse(raw)

    def test_review_escapes_hostile_headers_and_isolates_html_and_links(self):
        self.edit(subject='<script>alert(1)</script>', html='<p>Visible final HTML</p><script>alert(2)</script><img src="https://example.invalid/track">')
        self.review()
        page = self.preview.read_text()
        self.assertIn('&lt;script&gt;alert(1)&lt;/script&gt;', page)
        self.assertIn('iframe sandbox=""', page)
        self.assertIn('default-src', page)
        self.assertNotIn('<script>', page)

    def test_full_mime_attachments_preserved_and_included_in_preview_and_fingerprint(self):
        self.edit(body='你好 — final', attachment=b'complete attachment bytes', html='<p>你好 — final</p>')
        review = self.review()
        self.assertIn('sample.bin', self.preview.read_text())
        raw, _ = s.outgoing(review['entries'][0], review, 'sim-send', m.now())
        self.assertEqual(s.content(raw, hide_bcc=True), s.content(base64.b64decode(review['entries'][0]['raw_b64']), hide_bcc=True))
        self.assertTrue(raw.isascii())
        self.approve(review)
        self.edit(body='replacement', attachment=b'changed')
        transport = self.transport()
        with self.assertRaises(m.DraftError):
            self.send(review, transport)
        self.assertEqual(transport.submissions, [])

    def test_exact_count_missing_draft_duplicate_recipient_and_history_fail_closed(self):
        with self.assertRaises(m.DraftError):
            s.prepare_review(self.state, self.mailbox, [self.batch], self.preview, expected_count=3)
        original = deepcopy(self.mailbox.client.data[b'Drafts'])
        del self.mailbox.client.data[b'Drafts'][b'1']
        with self.assertRaises(m.DraftError): self.review()
        self.mailbox.client.data[b'Drafts'] = deepcopy(original)
        self.edit(uid=b'2', recipient='sim1@example.invalid')
        with self.assertRaises(m.DraftError): self.review()
        self.mailbox.client.data[b'Drafts'] = deepcopy(original)
        self.mailbox.client.data[b'Sent'][b'10'] = original[b'1']
        with self.assertRaises(m.DraftError): self.review()
        self.assertFalse(self.preview.exists())

    def test_no_confirmation_means_no_transport_or_attempt(self):
        review = self.review()
        factory = unittest.mock.Mock()
        with self.assertRaises(m.DraftError):
            s.deliver(self.state, review['id'], self.mailbox, factory)
        factory.assert_not_called()

    def test_confirmation_requires_matching_review_count_and_later_timestamp(self):
        review = self.review()
        instruction = {'user_instruction': 'Confirm selected simulation', 'user_instruction_at': m.now(),
                       'review_hash': review['hash'], 'count': 2}
        for key, value in (('user_instruction', ''), ('review_hash', 'different'), ('count', 1),
                           ('user_instruction_at', '2020-01-01T00:00:00+08:00'), ('user_instruction_at', '2026-01-01')):
            changed = {**instruction, key: value}
            with self.assertRaises(m.DraftError): s.approve(self.state, review['id'], changed)
        self.approve(review)
        self.assertEqual(s.get_review(self.state, review['id'])['status'], 'approved')

    def test_reading_again_reuses_review_and_never_carries_approval_forward(self):
        review = self.review()
        self.approve(review)
        fresh = self.review()
        self.assertEqual(review['id'], fresh['id'])
        self.assertEqual(fresh['status'], 'awaiting_confirmation')
        self.assertNotIn('approval', fresh)

    def test_changed_preview_or_manifest_cannot_be_sent(self):
        review = self.review()
        self.preview.write_text('Changed preview')
        with self.assertRaises(m.DraftError): self.approve(review)
        review = self.review()
        self.approve(review)
        altered = s.get_review(self.state, review['id'])
        altered['entries'][0]['details']['envelope']['bcc'] = ['extra@example.invalid']
        s.put_review(self.state, altered)
        with self.assertRaises(m.DraftError): self.send(review, self.transport())

    def test_recipient_subject_body_deletion_or_new_history_invalidates_approval(self):
        original = deepcopy(self.mailbox.client.data)
        for change in ('recipient', 'subject', 'body', 'delete', 'history'):
            with self.subTest(change=change):
                self.mailbox.client.data = deepcopy(original)
                review = self.review()
                self.approve(review)
                if change == 'recipient': self.edit(recipient='new@example.invalid')
                if change == 'subject': self.edit(subject='Changed final subject')
                if change == 'body': self.edit(body='Changed final body')
                if change == 'delete': del self.mailbox.client.data[b'Drafts'][b'1']
                if change == 'history': self.mailbox.client.data[b'INBOX'][b'90'] = original[b'Drafts'][b'1']
                transport = self.transport()
                with self.assertRaises(m.DraftError): self.send(review, transport)
                self.assertEqual(transport.submissions, [])
                self.assertEqual(s.get_review(self.state, review['id'])['status'], 'stale')

    def test_success_is_verified_from_sent_folder_and_repeated_send_never_resubmits(self):
        review = self.review()
        self.approve(review)
        transport = self.transport()
        result = self.send(review, transport)
        self.assertEqual(result['states'], {'sent_verified': 2})
        self.assertEqual(len(transport.submissions), 2)
        self.assertTrue(transport.closed)
        self.assertEqual(len(self.mailbox.client.data[b'Drafts']), 2)
        with self.assertRaises(m.DraftError): self.send(review, transport)
        with self.assertRaises(m.DraftError): self.review()
        for item in self.state.items(self.batch):
            m.reconcile(self.state, item, self.mailbox)
            self.assertEqual(item['status'], 'sent_verified')
            self.assertEqual(item['sent_receipt']['verification'], 'sent_folder')

    def test_disconnect_after_acceptance_is_reconciled_from_real_sent_copy(self):
        review = self.review()
        self.approve(review)
        transport = self.transport('accepted_disconnect')
        result = self.send(review, transport)
        self.assertEqual(result['states'], {'sent_verified': 2})
        self.assertEqual(len(transport.submissions), 2)

    def test_unknown_outcome_halts_whole_batch_and_reconcile_never_sends(self):
        review = self.review()
        self.approve(review)
        transport = self.transport('unknown')
        result = self.send(review, transport)
        self.assertEqual(result['states'], {'uncertain': 1, 'not_attempted': 1})
        self.assertEqual(len(transport.submissions), 1)
        with self.assertRaises(m.DraftError): self.send(review, transport)
        s.reconcile_delivery(self.state, review['id'], self.mailbox)
        self.assertEqual(len(transport.submissions), 1)
        # A late server-side Sent copy resolves the result without re-upload.
        self.mailbox.client.data[b'Sent'][b'1'] = transport.submissions[0][0]
        result = s.reconcile_delivery(self.state, review['id'], self.mailbox)
        self.assertEqual(result['states'], {'sent_verified': 1, 'not_attempted': 1})
        self.assertEqual(len(transport.submissions), 1)

    def test_resume_after_late_verified_receipt_submits_only_untouched_entries(self):
        review = self.review()
        self.approve(review)
        first = self.transport('unknown')
        self.send(review, first)
        self.mailbox.client.data[b'Sent'][b'1'] = first.submissions[0][0]
        s.reconcile_delivery(self.state, review['id'], self.mailbox)
        rest = self.transport()
        result = s.deliver(self.state, review['id'], self.mailbox, lambda config: rest,
                           pause=lambda seconds: None, resume=True)
        self.assertEqual(result['states'], {'sent_verified': 2})
        self.assertEqual(len(first.submissions), 1)
        self.assertEqual(len(rest.submissions), 1)
        self.assertEqual(rest.submissions[0][1], ['sim2@example.invalid'])
        with self.assertRaises(m.DraftError):
            s.deliver(self.state, review['id'], self.mailbox, lambda config: rest, resume=True)

    def test_resume_never_bypasses_unknown_or_failed_outcomes_or_changed_remaining_content(self):
        review = self.review()
        self.approve(review)
        first = self.transport('unknown')
        self.send(review, first)
        factory = unittest.mock.Mock()
        with self.assertRaises(m.DraftError):
            s.deliver(self.state, review['id'], self.mailbox, factory, resume=True)
        factory.assert_not_called()
        value = s.attempt(self.state, review['entries'][0]['draft_id'])
        value['status'] = 'failed'
        s.put_attempt(self.state, review['entries'][0], value)
        with self.assertRaises(m.DraftError):
            s.deliver(self.state, review['id'], self.mailbox, factory, resume=True)
        factory.assert_not_called()
        self.mailbox.client.data[b'Sent'][b'1'] = first.submissions[0][0]
        s.reconcile_delivery(self.state, review['id'], self.mailbox)
        self.edit(uid=b'2', body='Changed after confirmation')
        with self.assertRaises(m.DraftError):
            s.deliver(self.state, review['id'], self.mailbox, factory, resume=True)
        factory.assert_not_called()

    def test_delayed_sent_copy_is_only_polled_and_never_submitted_twice(self):
        review = self.review()
        self.approve(review)
        transport = self.transport('accepted_without_sent_copy')
        waits = []
        def delayed_copy(seconds):
            waits.append(seconds)
            if len(waits) % 2 == 0:
                sent = self.mailbox.client.data[b'Sent']
                sent[str(len(sent)+1).encode()] = transport.submissions[-1][0]
        result = s.deliver(self.state, review['id'], self.mailbox, lambda config: transport, pause=delayed_copy)
        self.assertEqual(result['states'], {'sent_verified': 2})
        self.assertEqual(len(transport.submissions), 2)

    def test_acceptance_without_server_sent_copy_stops_and_does_not_fabricate_receipt(self):
        review = self.review()
        self.approve(review)
        transport = self.transport('accepted_without_sent_copy')
        result = self.send(review, transport)
        self.assertEqual(result['states'], {'smtp_accepted': 1, 'not_attempted': 1})
        self.assertEqual(self.mailbox.client.data[b'Sent'], {})
        self.assertNotIn('sent_receipt', self.state.items(self.batch)[0])

    def test_known_refusal_halts_without_automatic_retry(self):
        review = self.review()
        self.approve(review)
        transport = self.transport('failed')
        result = self.send(review, transport)
        self.assertEqual(result['states'], {'failed': 1, 'not_attempted': 1})
        with self.assertRaises(m.DraftError): self.send(review, transport)
        self.assertEqual(len(transport.submissions), 1)

    def test_wrong_sent_content_is_not_success(self):
        review = self.review()
        self.approve(review)
        transport = self.transport('unknown')
        self.send(review, transport)
        raw = s.parse(transport.submissions[0][0])
        raw.set_content('Incorrect content')
        self.mailbox.client.data[b'Sent'][b'1'] = raw.as_bytes()
        result = s.reconcile_delivery(self.state, review['id'], self.mailbox)
        self.assertEqual(result['states'], {'needs_review': 1, 'not_attempted': 1})

    def test_account_mismatch_is_rejected_before_submission(self):
        review = self.review()
        self.approve(review)
        self.mailbox.account = 'different@example.invalid'
        transport = self.transport()
        with self.assertRaises(m.DraftError): self.send(review, transport)
        self.assertEqual(transport.submissions, [])

    def test_crash_after_reservation_is_durable_and_cannot_be_replayed(self):
        review = self.review()
        self.approve(review)
        transport = self.transport()
        def crash(raw, recipients, stage):
            stage('data')
            raise KeyboardInterrupt()
        transport.submit = crash
        with self.assertRaises(KeyboardInterrupt): self.send(review, transport)
        self.assertEqual(s.get_review(self.state, review['id'])['status'], 'halted')
        result = s.attempt(self.state, review['entries'][0]['draft_id'])
        self.assertEqual(result['phase'], 'data')
        with self.assertRaises(m.DraftError): self.send(review, transport)
        with self.assertRaises(m.DraftError): self.review()
        self.assertEqual(self.mailbox.client.data[b'Sent'], {})

    def test_draft_change_during_batch_halts_before_next_envelope(self):
        review = self.review()
        self.approve(review)
        transport = self.transport()
        submit = transport.submit
        def first_then_edit(raw, recipients, stage):
            result = submit(raw, recipients, stage)
            self.edit(uid=b'2', body='Human changed while first was being sent')
            return result
        transport.submit = first_then_edit
        with self.assertRaises(m.DraftError): self.send(review, transport)
        self.assertEqual(len(transport.submissions), 1)
        result = s.delivery_report(self.state, review['id'])
        self.assertEqual(result['states'], {'sent_verified': 1, 'not_attempted': 1})

    def test_attachment_only_edit_invalidates_approved_version(self):
        self.edit(attachment=b'original attachment')
        review = self.review()
        self.approve(review)
        message = s.parse(self.mailbox.client.data[b'Drafts'][b'1'])
        attachment = next(message.iter_attachments())
        attachment.set_content(b'changed attachment', maintype='application', subtype='octet-stream')
        self.mailbox.client.data[b'Drafts'][b'1'] = message.as_bytes()
        transport = self.transport()
        with self.assertRaises(m.DraftError): self.send(review, transport)
        self.assertEqual(transport.submissions, [])

    def test_bcc_is_in_approval_and_envelope_but_not_on_wire(self):
        message = s.parse(self.mailbox.client.data[b'Drafts'][b'1'])
        message['Cc'] = 'copy@example.invalid'
        message['Bcc'] = 'private@example.invalid'
        self.mailbox.client.data[b'Drafts'][b'1'] = message.as_bytes()
        review = self.review()
        self.assertIn('private@example.invalid', self.preview.read_text())
        self.approve(review)
        transport = self.transport()
        result = self.send(review, transport)
        self.assertEqual(result['status'], 'sent_verified')
        raw, envelope = transport.submissions[0]
        self.assertNotIn('Bcc', s.parse(raw))
        self.assertEqual(envelope, ['sim1@example.invalid', 'copy@example.invalid', 'private@example.invalid'])


class FakeSMTP:
    def __init__(self, host, port, **kwargs):
        assert (host, port) == (s.SMTP_HOST, 465)
        assert kwargs['context'].check_hostname
        self.calls = []
        self.mail_code, self.rcpt_code, self.data_code = 250, 250, 250
        self.data_error = None

    def ehlo_or_helo_if_needed(self): self.calls.append('EHLO')
    def login(self, account, password): self.calls.append('AUTH')
    def noop(self): self.calls.append('NOOP'); return 250, b'OK'
    def quit(self): self.calls.append('QUIT')
    def close(self): self.calls.append('CLOSE')
    def mail(self, account): self.calls.append('MAIL'); return self.mail_code, b'ignored'
    def rcpt(self, recipient): self.calls.append('RCPT'); return self.rcpt_code, b'ignored'
    def rset(self): self.calls.append('RSET')
    def data(self, raw):
        self.calls.append('DATA')
        if self.data_error: raise self.data_error
        return self.data_code, b'ignored'


class TransportTests(unittest.TestCase):
    def transport(self):
        return s.SMTPTransport('sara@example.invalid', 'simulated-secret', factory=FakeSMTP)

    def test_login_check_never_issues_envelope_or_data(self):
        transport = self.transport()
        self.assertFalse(transport.check()['sent'])
        transport.close()
        self.assertEqual(transport.client.calls, ['EHLO', 'AUTH', 'NOOP', 'QUIT'])

    def test_all_recipient_acceptances_precede_data_and_refusal_aborts_whole_message(self):
        transport = self.transport()
        transport.client.rcpt_code = 550
        with self.assertRaises(s.DeliveryError) as caught:
            transport.submit(b'synthetic payload', ['sim@example.invalid'], lambda stage: None)
        self.assertEqual(caught.exception.status, 'failed')
        self.assertNotIn('DATA', transport.client.calls)
        self.assertEqual(transport.client.calls[-1], 'RSET')

    def test_data_disconnect_and_explicit_rejection_are_distinguished_without_server_text(self):
        for error, status in ((TimeoutError('do-not-log-secret'), 'uncertain'),
                              (s.smtplib.SMTPDataError(550, b'do-not-log-secret'), 'failed')):
            transport = self.transport()
            transport.client.data_error = error
            stages = []
            with self.assertRaises(s.DeliveryError) as caught:
                transport.submit(b'synthetic payload', ['sim@example.invalid'], stages.append)
            self.assertEqual(caught.exception.status, status)
            self.assertEqual(stages, ['envelope', 'data'])
            self.assertNotIn('do-not-log-secret', str(caught.exception))

    def test_auth_failure_is_sanitized_and_does_not_send(self):
        class Bad(FakeSMTP):
            def login(self, account, password):
                raise s.smtplib.SMTPAuthenticationError(535, b'simulated-secret')
        with self.assertRaises(m.DraftError) as caught:
            s.SMTPTransport('sara@example.invalid', 'simulated-secret', factory=Bad)
        self.assertNotIn('simulated-secret', str(caught.exception))


if __name__ == '__main__':
    unittest.main()

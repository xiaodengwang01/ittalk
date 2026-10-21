"""Unread collection uses synthetic mail only."""
import tempfile
from pathlib import Path
from types import SimpleNamespace
import unittest

from scripts import mail_inbox as inbox


class FakeInbox:
    account = "team@example.invalid"

    def __init__(self):
        self.flags = {b"1": b"", b"2": b"\\Seen"}
        self.fetches = []
        self.searches = []
        self.client = SimpleNamespace(uid=self.uid)

    def select(self, folder):
        assert folder == b"INBOX"
        return "4"

    def search(self, criterion):
        self.searches.append(criterion)
        assert criterion == "UNSEEN"
        # UID 2 was read by the user after this search.
        return [b"1", b"2"]

    def uid(self, operation, uid, part):
        assert operation == "FETCH" and part == "(UID FLAGS)"
        return "OK", [b"1 (UID " + uid + b" FLAGS (" + self.flags[uid] + b"))"]

    def fetch(self, uid):
        assert not self.flags[uid]
        self.fetches.append(uid)
        return (b"From: sim@example.invalid\r\nTo: team@example.invalid\r\n"
                b"Message-ID: <sim-in-1@example.invalid>\r\nSubject: Simulated reply\r\n"
                b"Date: Wed, 21 Oct 2026 10:00:00 +0800\r\n\r\nSynthetic body.\r\n")


class UnreadTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.state = inbox.mail.State(Path(self.temp.name) / "state.sqlite3")
        self.box = FakeInbox()

    def tearDown(self):
        self.state.close()
        self.temp.cleanup()

    def test_reads_only_still_unread_and_reuses_immutable_uid(self):
        result = inbox.collect_unread(self.state, self.box)
        self.assertEqual(self.box.fetches, [b"1"])
        self.assertEqual(result["skipped_read_in_meantime"], ["2"])
        self.assertEqual(result["unread_before"], result["unread_after"])
        self.assertEqual(result["messages"][0]["source"], "INBOX:UIDVALIDITY=4:UID=1")
        again = inbox.collect_unread(self.state, self.box)
        self.assertEqual(self.box.fetches, [b"1"])
        self.assertEqual(result["messages"], again["messages"])
        self.assertEqual(inbox.daily.pending_replies(self.state), [])

    def test_unknown_flags_never_fetch_body(self):
        for response in (("NO", []), ("OK", [b"1 (UID 1)"]),
                         ("OK", [b"1 (UID 9 FLAGS ())"])):
            self.box.client.uid = lambda *args: response
            with self.assertRaises(inbox.mail.DraftError):
                inbox.collect_unread(self.state, self.box)
            self.assertEqual(self.box.fetches, [])

    def test_read_messages_not_returned_even_if_cached(self):
        inbox.collect_unread(self.state, self.box)
        self.box.flags[b"1"] = b"\\Seen"
        result = inbox.collect_unread(self.state, self.box)
        self.assertEqual(result["messages"], [])
        self.assertEqual(self.box.fetches, [b"1"])


if __name__ == "__main__":
    unittest.main()

"""Tests for telling an emptied message row from one that shows a message (issue #3).

python -m unittest discover -s tests

messageRows.py uses only the standard library, so it is loaded straight from the
add-on package without NVDA.
"""

import importlib.util
import os
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_spec = importlib.util.spec_from_file_location(
    "messageRows",
    os.path.join(ROOT, "addon", "globalPlugins", "outlookFirstLineSilence", "messageRows.py"),
)
messageRows = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(messageRows)

# The status words of NVDA 2026.2's UIAGridRow (appModules/outlook.py) in English:
# "unread", executedVerbLabels, "has attachment", importanceLabels, "meeting request",
# and the expanded and collapsed states.
LABELS = {
    "unread",
    "replied",
    "replied all",
    "forwarded",
    "has attachment",
    "high importance",
    "low importance",
    "meeting request",
    "expanded",
    "collapsed",
}

# Rows as NVDA named them in the tester's NVDA 2026.2 log.
ROW = (
    "unread From MacDailyNews, Subject Apple Intelligence is eating up to 30GB on some Macs "
    "after macOS 27, Received Fri 9/25/2026 10:35 AM, Size 84 KB,"
)
READ_ROW = (
    "From joshknnd1982, Subject Re: [joshknnd1982/jawsMigrator] Outlook issues (Issue #5), "
    "Received Fri 9/25/2026 10:03 AM, Size 30 KB,"
)


class IsStatusOnlyTests(unittest.TestCase):
    def test_the_emptied_row_in_the_log_is_status_only(self):
        # After each Delete, NVDA said this new name, then the next message.
        self.assertTrue(messageRows.isStatusOnly("unread", LABELS))

    def test_several_status_words_are_status_only(self):
        for name in (
            "unread replied has attachment high importance",
            "collapsed unread",
            "replied all",
            "forwarded meeting request",
        ):
            with self.subTest(name=name):
                self.assertTrue(messageRows.isStatusOnly(name, LABELS))

    def test_a_row_with_its_columns_is_not(self):
        self.assertFalse(messageRows.isStatusOnly(ROW, LABELS))
        self.assertFalse(messageRows.isStatusOnly(READ_ROW, LABELS))

    def test_a_status_word_followed_by_column_text_is_not(self):
        self.assertFalse(messageRows.isStatusOnly("unread From Contoso,", LABELS))
        self.assertFalse(messageRows.isStatusOnly("replied all hands,", LABELS))

    def test_a_word_that_only_starts_with_a_status_word_is_not(self):
        self.assertFalse(messageRows.isStatusOnly("unreadable", LABELS))
        self.assertFalse(messageRows.isStatusOnly("unread,", LABELS))

    def test_an_empty_name_is_not(self):
        # NVDA says nothing for it, so there is nothing to keep quiet.
        for name in ("", "   ", None):
            with self.subTest(name=name):
                self.assertFalse(messageRows.isStatusOnly(name, LABELS))

    def test_case_and_spacing_do_not_matter(self):
        self.assertTrue(messageRows.isStatusOnly("  Unread   Has  attachment ", LABELS))

    def test_translated_labels(self):
        # NVDA says the status in its own language; the add-on passes NVDA's translations.
        labels = {"ungelesen", "hat Anlage", "beantwortet"}
        self.assertTrue(messageRows.isStatusOnly("ungelesen hat Anlage", labels))
        self.assertFalse(messageRows.isStatusOnly("unread", labels))

    def test_blank_labels_are_ignored(self):
        self.assertTrue(messageRows.isStatusOnly("unread", LABELS | {"", "  "}))
        self.assertFalse(messageRows.isStatusOnly("From Contoso,", {"", " "}))


if __name__ == "__main__":
    unittest.main()

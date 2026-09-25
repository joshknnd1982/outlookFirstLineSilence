# -*- coding: utf-8 -*-
# Outlook First Line Silence: names of rows in classic Outlook's message list.
# Licensed under the GNU General Public License version 2.

"""Tell a message row named only by its status from a row that shows a message.

NVDA's Outlook support names a row of the message list from the text of its columns,
each followed by a comma ("From MacDailyNews, Subject ..., Received ..., Size 84 KB,"),
after what Outlook's object model says about the selected message: "unread", "replied",
"has attachment", "high importance", "meeting request", and "expanded" or "collapsed"
for a conversation (UIAGridRow in NVDA 2026.2's appModules/outlook.py). A row whose
columns are gone is named by that status alone, such as "unread".

This module uses only the standard library, so the tests load it without NVDA.
"""


def _words(text):
    return " ".join((text or "").split()).casefold()


def isStatusOnly(name, labels):
    """True when *name* is made only of the status *labels*, with no column text.

    An empty name is not: NVDA says nothing for it anyway.
    """
    rest = _words(name)
    if not rest:
        return False
    # The longest first, so "replied all" is not read as "replied" and a column.
    known = sorted({_words(label) for label in labels} - {""}, key=len, reverse=True)
    while rest:
        for label in known:
            if rest == label or rest.startswith(label + " "):
                rest = rest[len(label):].lstrip()
                break
        else:
            return False
    return True

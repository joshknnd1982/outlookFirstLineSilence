"""Tests for the title of a message window you write in (issue #7, 1.0.29).

python -m unittest discover -s tests

In the tester's log, Control+F on an open message made NVDA say only "To edit blank",
then "multi line". JAWS says the Forward window's title, then "To Edit". The add-on
dropped that title with the rest of the opening speech, as for a message you read.
Now Forward, Reply and New say their title once focus lands where you write. Opening a
message you read stays silent. To, Cc, Bcc and Subject leave out NVDA's "blank" and
the "multi line" NVDA says when Outlook makes an address field multi-line just after it
gets focus.

The plugin runs through NVDA 2026.2's own event code, from test_messageReturn.
"""

import sys
import unittest

import test_messageReturn as harness
from test_messageReturn import CANCEL, OutputReason, Role, State, setUpModule, tearDownModule  # noqa: F401

# The windows in the tester's issue #7 log. NVDA names a message window with a trailing space.
READ_TITLE = "Re: [joshknnd1982/outlookFirstLineSilence] a suggestion (Issue #6) - Message (HTML)"
FORWARD_TITLE = "FW: [joshknnd1982/outlookFirstLineSilence] a suggestion (Issue #6) - Message (HTML)"
REPLY_TITLE = "RE: [joshknnd1982/jawsMigrator] It says page one section 1 when in a outlook message (Issue #18) - Message (HTML)"
NEW_TITLE = "Untitled - Message (HTML)"
FORWARD_SUBJECT = "FW: [joshknnd1982/outlookFirstLineSilence] a suggestion (Issue #6)"
BODY = "You are now in the message body, type a message."


class _ComposeWindow:
    """A message window you write in: address fields, Subject and an editable body."""

    def __init__(self, NVDAObject, desktop, window, title, subject=""):
        self.window = window
        self.title = title
        self.pane = NVDAObject(Role.PANE, title + " ", desktop, window=window)
        self.dialog = NVDAObject(Role.DIALOG, "", self.pane, window=window)
        # The class NVDA's log shows for To: an IAccessible RichEdit20W, a ContactEditField.
        field = dict(parent=self.dialog, window=window + 1, windowClass="RichEdit20WPT")
        self.to = NVDAObject(Role.EDITABLETEXT, "To", **field)
        self.cc = NVDAObject(Role.EDITABLETEXT, "Cc", **field)
        self.subject = NVDAObject(Role.EDITABLETEXT, "Subject", text=subject, **field)
        self.body = NVDAObject(Role.DOCUMENT, "", self.dialog, window=window + 2, windowClass="_WwG")


class ComposeTitleTests(harness.MessageWindowTestCase):
    def setUp(self):
        super().setUp()
        sys.modules["ui"].message = self.spoken.append

    def composeWindow(self, window, title, subject=""):
        self.visible.add(window)
        self.classes[window] = "rctrl_renwnd32"
        self.titles[window] = title + " "
        self.roots.update({window + 1: window, window + 2: window})
        return _ComposeWindow(self.NVDAObject, self.desktop, window, title, subject)

    def readOpenMessage(self):
        """The tester's first step: Enter on a message in the list."""
        message = self.messageWindow(0x400, READ_TITLE)
        self.openMessage(message)
        return message

    def outlookMakesItMultiLine(self, field):
        field.states.add(State.MULTILINE)
        self.executeEvent("stateChange", field)

    def forward(self):
        """Control+F: Outlook opens the Forward window with focus in To."""
        self.later()
        forward = self.composeWindow(0x600, FORWARD_TITLE, FORWARD_SUBJECT)
        self.focus(forward.to, forward.pane)
        self.later(0.05)
        self.outlookMakesItMultiLine(forward.to)
        return forward

    # Tests.

    def test_forwardSaysItsTitleThenTo(self):
        self.readOpenMessage()
        self.assertEqual(self.heardInFull(), [])
        self.forward()
        self.assertEqual(self.heardInFull(), [FORWARD_TITLE + " ", "To edit"])

    def test_theTitleComesAfterNvdaStopsSpeakingForTheNewWindow(self):
        self.readOpenMessage()
        del self.spoken[:]
        self.forward()
        self.assertEqual(self.spoken, [CANCEL, FORWARD_TITLE + " ", "To edit"])

    def test_tabbingThroughTheFields(self):
        self.readOpenMessage()
        forward = self.forward()
        del self.spoken[:]
        self.later(1.0)
        self.focus(forward.cc, forward.pane)
        self.later(0.05)
        self.outlookMakesItMultiLine(forward.cc)
        self.later(0.2)
        self.focus(forward.subject, forward.pane)
        # The title only once, and Subject says its text.
        self.assertEqual(self.spoken, ["Cc edit", "Subject edit " + FORWARD_SUBJECT])

    def test_replySaysItsTitleThenTheBody(self):
        self.readOpenMessage()
        self.later()
        reply = self.composeWindow(0x700, REPLY_TITLE)
        self.focus(reply.body, reply.pane)
        self.assertEqual(self.heardInFull(), [REPLY_TITLE + " ", BODY])

    def test_aNewMessageSaysItsTitle(self):
        self.focus(self.inboxRow, self.inbox)
        self.later()
        new = self.composeWindow(0x800, NEW_TITLE)
        self.focus(new.to, new.pane)
        self.assertEqual(self.heardInFull(), [NEW_TITLE + " ", "To edit"])

    def test_whenTheWindowGetsFocusBeforeTo(self):
        # NVDA can get Outlook's foreground event, for the window itself, before focus reaches To.
        self.readOpenMessage()
        self.later()
        forward = self.composeWindow(0x600, FORWARD_TITLE)
        self.focus(forward.pane, forward.pane)
        self.assertEqual(self.heardInFull(), [])
        self.later(0.1)
        self.focus(forward.to, forward.pane)
        self.assertEqual(self.heardInFull(), [FORWARD_TITLE + " ", "To edit"])

    def test_aMessageYouReadStaysSilent(self):
        self.focus(self.inboxRow, self.inbox)
        del self.spoken[:]
        message = self.readOpenMessage()
        self.assertEqual(self.heardInFull(), [])
        self.assertIsNone(plugin()._openingTitle)
        # Nothing of it is said later in the same window either.
        self.later(0.5)
        self.focus(message.document, message.pane)
        self.assertEqual(self.heardInFull(), [])

    def test_aTitleNotSaidInTimeIsForgotten(self):
        self.readOpenMessage()
        self.later()
        forward = self.composeWindow(0x600, FORWARD_TITLE)
        self.focus(forward.pane, forward.pane)
        self.later(3.0)
        del self.spoken[:]
        self.focus(forward.to, forward.pane)
        # The window is no longer opening, so NVDA's own "dialog" comes through, as before.
        self.assertEqual(self.spoken, ["dialog", "To edit"])

    def test_anotherWindowForgetsTheTitle(self):
        self.readOpenMessage()
        self.later()
        forward = self.composeWindow(0x600, FORWARD_TITLE)
        self.focus(forward.pane, forward.pane)
        self.later(0.2)
        self.focus(self.notepadText, self.notepad)
        self.assertIsNone(plugin()._openingTitle)

    def test_altTabBackToTheForwardWindowSaysItsTitleOnce(self):
        self.readOpenMessage()
        forward = self.forward()
        self.altTabTo(self.notepad, self.notepadText)
        self.altTabTo(forward.pane, forward.to)
        self.assertEqual(self.heardInFull(), [FORWARD_TITLE + " ", "To edit"])

    def test_blankIsStillSaidWhenYouMoveInAnEmptyField(self):
        # Only the focus speech leaves it out; arrowing onto an empty line still says it.
        self.readOpenMessage()
        self.forward()
        del self.spoken[:]
        sys.modules["speech.speech"].speak(["blank"])
        self.assertEqual(self.spoken, ["blank"])

    def test_multiLineIsStillSaidWithOtherWords(self):
        self.readOpenMessage()
        forward = self.forward()
        del self.spoken[:]
        sys.modules["speech.speech"].speak(["To", "edit", "multi line"])
        self.assertEqual(self.spoken, ["To edit multi line"])
        self.assertIs(self.api.focus, forward.to)

    def test_otherProgramsKeepBlankAndMultiLine(self):
        self.visible.add(0x900)
        self.titles[0x900] = "Document - WordPad"
        pad = self.NVDAObject(Role.PANE, "Document - WordPad", self.desktop, appName="wordpad", window=0x900)
        self.roots[0x901] = 0x900
        text = self.NVDAObject(Role.EDITABLETEXT, "Rich Text Window", pad, appName="wordpad", window=0x901, windowClass="RichEdit20W")
        self.later()
        self.focus(text, pad)
        self.outlookMakesItMultiLine(text)
        self.assertEqual(self.heardInFull(), ["Document - WordPad", "Rich Text Window edit blank", "multi line"])

    def test_theTitleIsForgottenWhenTheAddOnStops(self):
        self.readOpenMessage()
        self.later()
        forward = self.composeWindow(0x600, FORWARD_TITLE)
        self.focus(forward.pane, forward.pane)
        self.assertIsNotNone(plugin()._openingTitle)
        with harness.mock.patch.object(plugin(), "_removeReformattedMessages"), harness.mock.patch.object(plugin().updater, "stop"):
            self.globalPlugin.terminate()
        self.assertIsNone(plugin()._openingTitle)


def plugin():
    return harness.plugin


if __name__ == "__main__":
    unittest.main()

"""Tests for reading a message's own HTML where Outlook shows it badly.

python -m unittest discover -s tests

messageHtml.py uses only the standard library, so it is loaded straight from the
add-on package without NVDA.
"""

import importlib.util
import os
import re
import unittest
from html.parser import HTMLParser

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_spec = importlib.util.spec_from_file_location(
    "messageHtml",
    os.path.join(ROOT, "addon", "globalPlugins", "outlookFirstLineSilence", "messageHtml.py"),
)
messageHtml = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(messageHtml)

STORY1 = "https://link.example.com/click/1?u=reader"
STORY2 = "https://link.example.com/click/2?u=reader"
STORY3 = "https://link.example.com/click/3?u=reader"
SPONSOR = "https://ads.example.com/sponsor"

# A newsletter laid out like the one in issue #2: each story (picture, headline,
# "Read more") sits in one link wrapped around block elements, which Outlook's Word
# rendering cannot keep.
NEWSLETTER = """<!DOCTYPE html>
<html><head><title>Daily Newsletter</title>
<style>h2 { font-size: 20px; } .x { color: red; }</style></head>
<body>
<div style="display:none; max-height:0; mso-hide:all">Your daily sports fix &#8199;&#65279;</div>
<table><tr><td><a href="https://www.example.com/"><img src="logo.png" alt="Logo"></a></td></tr></table>
<a href="STORY1" style="text-decoration:none">
<table role="presentation"><tr><td><img src="nfl.jpg" alt=""></td></tr>
<tr><td><h2>The NFL Fined a Guy $14K for Pretending to Drink a Beer</h2></td></tr></table>
</a>
<p>Advertisement</p>
<a href="STORY2"><div><img src="texas.jpg"><h2>President Trump Will Directly Impact Texas Football Game
Against Tennessee By Using The Bathroom</h2><p>&nbsp;Read more</p></div></a>
<a href="SPONSOR"><img src="ad.png" alt="Sponsored"></a>
<p><a href="SPONSOR">Sponsored &middot; Read More</a></p>
<a href="STORY3"><div><h2>Wake Forest Coach&#8217;s Stirring &#8216;Hourglass&#8217; Speech About How
Quickly Life Passes Is A Must-See</h2><p>&nbsp;Read more</p></div></a>
<p>Questions? <a href="javascript:alert(1)">Contact us</a> or <a href="mailto:help@example.com">email help</a>.</p>
</body></html>
""".replace("STORY1", STORY1).replace("STORY2", STORY2).replace("STORY3", STORY3).replace("SPONSOR", SPONSOR)

# The same message as Word shows it, one string per line, with Word's paragraph
# marks, a table's end-of-cell marks and no-break spaces.
SHOWN_LINES = [
    "\x07",
    "",
    "",
    "The NFL Fined a Guy $14K for Pretending to Drink a Beer\r\x07",
    "Advertisement\r",
    "",
    "",
    "President Trump Will Directly Impact Texas Football Game ",
    "Against Tennessee By Using The Bathroom\r",
    "\xa0Read more\r",
    "",
    "Sponsored \xb7 Read More \r",
    "",
    "",
    "Wake Forest Coach\u2019s Stirring \u2018Hourglass\u2019 Speech About How ",
    "Quickly Life Passes Is A Must-See\r",
    "\xa0Read more\r",
    "Questions? Contact us or email help.\r",
]


def linkOnLine(index, column=0, html=NEWSLETTER, lines=SHOWN_LINES):
    """What messageHtml.linkAt finds with the caret *column* characters into line *index*."""
    before = "".join(lines[:index]) + lines[index][:column]
    return messageHtml.linkAt(html, before, lines[index][column:])


class LinkAtTests(unittest.TestCase):
    def test_headlineOfALinkedStory(self):
        self.assertEqual(linkOnLine(3), STORY1)

    def test_eachLineOfAWrappedHeadline(self):
        self.assertEqual(linkOnLine(7), STORY2)
        self.assertEqual(linkOnLine(8), STORY2)

    def test_caretPartWayAlongTheLine(self):
        self.assertEqual(linkOnLine(3, column=10), STORY1)

    def test_readMoreOpensItsOwnStory(self):
        # "Read more" is under two stories and "Read More" in an advertisement.
        self.assertEqual(linkOnLine(9), STORY2)
        self.assertEqual(linkOnLine(16), STORY3)
        self.assertEqual(linkOnLine(11), SPONSOR)

    def test_curlyQuotesAndSpacingDoNotMatter(self):
        self.assertEqual(linkOnLine(14), STORY3)

    def test_textOutsideAnyLink(self):
        self.assertIsNone(linkOnLine(4))
        self.assertIsNone(linkOnLine(17))

    def test_onlyWebAndEmailLinksAreOpened(self):
        line = SHOWN_LINES[17]
        self.assertIsNone(linkOnLine(17, column=line.index("Contact")))
        self.assertEqual(linkOnLine(17, column=line.index("email")), "mailto:help@example.com")

    def test_blankLineAndEmptyMessage(self):
        self.assertIsNone(linkOnLine(1))
        self.assertIsNone(messageHtml.linkAt("", "", "The NFL Fined a Guy"))
        self.assertIsNone(messageHtml.linkAt(None, "", "The NFL Fined a Guy"))

    def test_textNotInTheMessage(self):
        self.assertIsNone(messageHtml.linkAt(NEWSLETTER, "", "Something else entirely"))

    def test_hiddenAndHeadTextIsNotShownText(self):
        text, links = messageHtml.messageText(NEWSLETTER)
        self.assertNotIn("dailysportsfix", text)
        self.assertNotIn("dailynewsletter", text)
        self.assertNotIn("fontsize", text)
        self.assertEqual(len(text), len(links))
        self.assertTrue(text.startswith("thenflfined"))

    def test_sameTextWithTheSameContextIsTakenInOrder(self):
        html = (
            '<p><a href="https://a.example.com/">Read more</a></p>'
            '<p><a href="https://b.example.com/">Read more</a></p>'
        )
        lines = ["Read more\r", "Read more\r"]
        self.assertEqual(linkOnLine(0, html=html, lines=lines), "https://a.example.com/")
        self.assertEqual(linkOnLine(1, html=html, lines=lines), "https://b.example.com/")

    def test_moreCopiesShownThanInTheMessage(self):
        html = '<p><a href="https://a.example.com/">Read more</a></p>'
        lines = ["Read more\r", "Read more\r"]
        self.assertIsNone(linkOnLine(1, html=html, lines=lines))

    def test_differentTextNearTheTopStillFindsTheStory(self):
        # Word also shows text the HTML hides only from other mail programs.
        lines = ["Extra text Outlook shows\r"] + SHOWN_LINES
        self.assertEqual(linkOnLine(4, lines=lines), STORY1)

    def test_malformedLinkMarkup(self):
        html = '<p><a href="https://a.example.com/">Unclosed <b>story</p><p>After</p>'
        self.assertEqual(messageHtml.linkAt(html, "", "Unclosed story"), "https://a.example.com/")
        self.assertEqual(messageHtml.linkAt(html, "Unclosed story", "After"), "https://a.example.com/")
        html = '<a href="https://a.example.com/">One</a><a href="https://b.example.com/">Two</a>'
        self.assertEqual(messageHtml.linkAt(html, "One", "Two"), "https://b.example.com/")
        html = '<a href="https://a.example.com/">One<a name="x">Two</a>'
        self.assertIsNone(messageHtml.linkAt(html, "One", "Two"))


class Elements(HTMLParser):
    """The elements of a page, as (tag, attributes, text) in order."""

    def __init__(self, page):
        super().__init__(convert_charrefs=True)
        self.elements = []
        self._open = []
        self.feed(page)
        self.close()

    def handle_starttag(self, tag, attrs):
        element = [tag, dict(attrs), ""]
        self.elements.append(element)
        self._open.append(element)

    def handle_endtag(self, tag):
        for index in range(len(self._open) - 1, -1, -1):
            if self._open[index][0] == tag:
                del self._open[index:]
                return
        raise AssertionError("</%s> was never opened" % tag)

    def handle_data(self, data):
        for element in self._open:
            element[2] += data

    def find(self, tag):
        return [(attrs, text) for name, attrs, text in self.elements if name == tag]


class ReformatTests(unittest.TestCase):
    def setUp(self):
        self.page, self.blocks, self.links = messageHtml.reformat(NEWSLETTER, "The NFL Fined a Guy", "From: Newsletter")
        self.elements = Elements(self.page)
        self.body = self.page.split("<body>", 1)[1]

    def test_subjectAndSender(self):
        self.assertEqual(self.elements.find("title"), [({}, "The NFL Fined a Guy")])
        self.assertEqual(self.elements.find("h1"), [({}, "The NFL Fined a Guy")])
        self.assertIn("<p>From: Newsletter</p>", self.body)

    def test_everyLineOfAStoryIsALinkToIt(self):
        links = [(attrs["href"], text) for attrs, text in self.elements.find("a")]
        self.assertEqual(links, [
            ("https://www.example.com/", "Logo"),
            (STORY1, "The NFL Fined a Guy $14K for Pretending to Drink a Beer"),
            (STORY2, "President Trump Will Directly Impact Texas Football Game Against Tennessee By Using The Bathroom"),
            (STORY2, "Read more"),
            (SPONSOR, "Sponsored"),
            (SPONSOR, "Sponsored \xb7 Read More"),
            (STORY3, "Wake Forest Coach\u2019s Stirring \u2018Hourglass\u2019 Speech About How Quickly Life Passes Is A Must-See"),
            (STORY3, "Read more"),
            ("mailto:help@example.com", "email help"),
        ])
        self.assertEqual(self.links, 9)

    def test_headingsAreKeptUnderTheSubject(self):
        headings = [text for attrs, text in self.elements.find("h3")]
        self.assertEqual(len(headings), 3)
        self.assertTrue(headings[0].startswith("The NFL Fined"))
        self.assertIn('<h3><a href="%s">President Trump' % STORY2, self.body)

    def test_plainTextStaysText(self):
        self.assertIn("<p>Advertisement</p>", self.body)
        self.assertIn('<p><a href="%s">Sponsored</a></p>' % SPONSOR, self.body)
        self.assertIn('<p>Questions? Contact us or <a href="mailto:help@example.com">email help</a>.</p>', self.body)

    def test_nothingIsLoadedAndNothingHiddenIsShown(self):
        self.assertNotIn("<img", self.page)
        self.assertNotIn("<table", self.page)
        self.assertNotIn("javascript:", self.page)
        self.assertNotIn("font-size", self.page)
        self.assertNotIn("daily sports fix", self.page)
        self.assertNotIn("Daily Newsletter", self.page)
        self.assertIn("default-src 'none'", self.page)
        self.assertIn('<meta name="referrer" content="no-referrer">', self.page)

    def test_textIsEscaped(self):
        page, blocks, links = messageHtml.reformat(
            '<p><a href="https://example.com/?a=1&amp;b=&quot;2&quot;">Tom &amp; Jerry &lt;b&gt;</a></p>',
            "<script>alert(1)</script>",
        )
        self.assertNotIn("<script>", page)
        self.assertIn("&lt;script&gt;alert(1)&lt;/script&gt;", page)
        self.assertIn('<a href="https://example.com/?a=1&amp;b=&quot;2&quot;">Tom &amp; Jerry &lt;b&gt;</a>', page)
        self.assertEqual(Elements(page).find("a"), [({"href": 'https://example.com/?a=1&b="2"'}, "Tom & Jerry <b>")])

    def test_listsLineBreaksAndPadding(self):
        page, blocks, links = messageHtml.reformat(
            "<div>\u034f \u200c\xa0\u034f</div>"
            "<ul><li>One</li><li>Two<ul><li>Two A</li></ul></li></ul>"
            "<p>Line one<br>Line two</p>"
            "<h2>Split<br>heading</h2>"
            "<table><tr><td>Cell one</td><td>Cell two</td></tr></table>",
            "Subject",
        )
        body = page.split("<body>", 1)[1]
        self.assertIn("<ul>\n<li>One</li>\n<li>Two</li>\n<li>Two A</li>\n</ul>", body)
        self.assertIn("<p>Line one</p>\n<p>Line two</p>", body)
        self.assertIn("<h3>Split heading</h3>", body)
        self.assertIn("<p>Cell one</p>\n<p>Cell two</p>", body)
        self.assertEqual(blocks, 8)
        self.assertEqual(re.findall(r"<p>\s*</p>", body), [])

    def test_emptyMessage(self):
        page, blocks, links = messageHtml.reformat("", "Subject")
        self.assertEqual((blocks, links), (0, 0))
        self.assertIn("<h1>Subject</h1>", page)
        Elements(page)


class IsOpenableTests(unittest.TestCase):
    def test_schemes(self):
        self.assertTrue(messageHtml.isOpenable("https://example.com/"))
        self.assertTrue(messageHtml.isOpenable("HTTP://example.com/"))
        self.assertTrue(messageHtml.isOpenable("mailto:a@example.com"))
        self.assertFalse(messageHtml.isOpenable("javascript:alert(1)"))
        self.assertFalse(messageHtml.isOpenable("file:///C:/Windows/notepad.exe"))
        self.assertFalse(messageHtml.isOpenable("C:\\Windows\\notepad.exe"))
        self.assertFalse(messageHtml.isOpenable("#top"))
        self.assertFalse(messageHtml.isOpenable("http://[bad"))


if __name__ == "__main__":
    unittest.main()

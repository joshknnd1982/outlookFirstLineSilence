# -*- coding: utf-8 -*-
# Outlook First Line Silence: reads a message's own HTML where Outlook shows it badly.
# Licensed under the GNU General Public License version 2.

"""Newsletters often put one link around a whole story: its picture, headline,
summary and "Read more". Outlook shows messages with Word, and Word cannot link
a block like that, so the story is shown as plain text, often after an empty
link. The message's HTML still has the link.

linkAt finds the text at the caret in that HTML and returns the address of the
link around it. reformat turns the HTML into a plain page of headings,
paragraphs, lists and links, without layout tables, pictures or styles, where
each line of a story is a link to it.

Only the standard library is used here, so this can be tested outside NVDA.
"""

from html import escape
from html.parser import HTMLParser
from urllib.parse import urlsplit
import re

#: The only kinds of address opened for the user.
OPENABLE_SCHEMES = frozenset(("http", "https", "mailto"))

# Elements whose text is never shown.
_UNSHOWN_ELEMENTS = frozenset(("head", "script", "style", "template", "title"))
_VOID_ELEMENTS = frozenset((
    "area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta",
    "param", "source", "track", "wbr",
))
# Text a newsletter hides from Outlook, such as the preview line shown only in the message list.
_HIDDEN_STYLE = re.compile(r"display\s*:\s*none|mso-hide\s*:\s*all", re.IGNORECASE)

# How much of the text before the caret to match as well, longest first. Text such as
# "Read more" appears under every story; what comes before it says which one it is.
_CONTEXT_LENGTHS = (160, 80, 40, 20, 0)


def isOpenable(url):
    try:
        return urlsplit(url).scheme.lower() in OPENABLE_SCHEMES
    except ValueError:
        return False


def _key(text):
    """Letters and digits only, case-folded.

    Word lays the message out differently from the HTML (line breaks, spaces
    between table cells, list bullets), so only the words themselves are compared.
    """
    return "".join(ch for ch in text.casefold() if ch.isalnum())


class _ShownHtml(HTMLParser):
    """Walks the parts of a message a reader sees, keeping track of the link they are in."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.link = None
        self._skipTag = None
        self._skipDepth = 0

    def handle_starttag(self, tag, attrs):
        if self._skipTag is not None:
            if tag == self._skipTag:
                self._skipDepth += 1
            return
        attributes = dict(attrs)
        if tag not in _VOID_ELEMENTS and (
            tag in _UNSHOWN_ELEMENTS or _HIDDEN_STYLE.search(attributes.get("style") or "")
        ):
            self._skipTag = tag
            self._skipDepth = 1
            return
        if tag == "a":
            # Links cannot be nested, so a new one ends any link left open.
            href = (attributes.get("href") or "").strip()
            self.link = href if href and isOpenable(href) else None
        self.shownStartTag(tag, attributes)

    def handle_endtag(self, tag):
        if self._skipTag is not None:
            if tag == self._skipTag:
                self._skipDepth -= 1
                if self._skipDepth <= 0:
                    self._skipTag = None
            return
        self.shownEndTag(tag)
        if tag == "a":
            self.link = None

    def handle_data(self, data):
        if self._skipTag is None:
            self.shownText(data)

    def shownStartTag(self, tag, attributes):
        pass

    def shownEndTag(self, tag):
        pass

    def shownText(self, data):
        pass

    def read(self, html):
        self.feed(html or "")
        self.close()
        return self


class _MessageText(_ShownHtml):
    """The shown text of an HTML message, keyed, with the link each character is in."""

    def __init__(self):
        super().__init__()
        self.chars = []
        self.links = []

    def shownText(self, data):
        for ch in _key(data):
            self.chars.append(ch)
            self.links.append(self.link)


def messageText(html):
    """Return (text, links): the message's keyed text and the link around each character."""
    parser = _MessageText().read(html)
    return "".join(parser.chars), parser.links


def _occurrences(text, needle, limit=1000):
    found = []
    start = text.find(needle)
    while start >= 0 and len(found) < limit:
        found.append(start)
        start = text.find(needle, start + 1)
    return found


def linkAt(html, textBefore, textFromCaret):
    """The address of the link around the text at the caret, or None.

    html: the message's HTML.
    textBefore: the message's text before the caret, as Outlook shows it.
    textFromCaret: the rest of the caret's line.
    None means the text at the caret is not in a link that can be opened, or it
    could not be found in the HTML for certain.
    """
    here = _key(textFromCaret)
    if not here:
        return None
    before = _key(textBefore)
    text, links = messageText(html)
    tried = set()
    for length in _CONTEXT_LENGTHS:
        lead = before[max(0, len(before) - length):] if length else ""
        if lead in tried:
            continue
        tried.add(lead)
        needle = lead + here
        found = _occurrences(text, needle)
        if not found:
            continue
        # The same text can appear more than once: take the one that comes at the
        # same place in the message, counting from the top. When Outlook shows it
        # more often than the HTML has it, which one is at the caret is not known.
        caretAt = len(before) - len(lead)
        earlier = sum(1 for position in _occurrences(before + here, needle) if position < caretAt)
        if earlier >= len(found):
            return None
        return links[found[earlier] + len(lead)]
    return None


# Reformatting.
# Every block element (paragraph, table cell, heading, list item...) becomes its own
# paragraph, heading or list item. A link around several blocks becomes a link in each.
_BLOCK_ELEMENTS = frozenset((
    "address", "article", "aside", "blockquote", "center", "dd", "div", "dl", "dt",
    "figcaption", "figure", "footer", "form", "h1", "h2", "h3", "h4", "h5", "h6",
    "header", "hr", "li", "main", "nav", "ol", "p", "pre", "section", "table",
    "tbody", "td", "tfoot", "th", "thead", "tr", "ul",
))
# The message's headings move down a level, under the subject.
_HEADINGS = {"h%d" % level: "h%d" % min(6, level + 1) for level in range(1, 7)}
# Spacing characters newsletters use to pad the preview line, which say nothing.
_INVISIBLE = dict.fromkeys(
    map(ord, "­͏؜ᅟᅠ឴឵᠎​‌‍‎‏⁠ㅤ﻿ﾠ"),
    None,
)
_SPACES = re.compile(r"\s+")


def _clean(text):
    return _SPACES.sub(" ", text.translate(_INVISIBLE))


class _Reformatter(_ShownHtml):
    def __init__(self):
        super().__init__()
        #: (kind, [(text, link)]) for each block; kind is "p", "li" or a heading.
        self.blocks = []
        self._runs = []
        self._kinds = []

    def _kind(self):
        return self._kinds[-1] if self._kinds else "p"

    def _endBlock(self):
        runs, self._runs = self._runs, []
        # A picture's description only names a link that has no text of its own here.
        withText = {link for text, link, isPicture in runs if not isPicture and text.strip()}
        pieces = []
        for text, link, isPicture in runs:
            if isPicture:
                if link is None or link in withText:
                    continue
                text = " %s " % text
            text = _clean(text)
            if pieces and pieces[-1][1] == link:
                pieces[-1][0] += text
            else:
                pieces.append([text, link])
        if not any(_key(text) for text, _link in pieces):
            return
        while not pieces[0][0].strip():
            del pieces[0]
        while not pieces[-1][0].strip():
            del pieces[-1]
        pieces[0][0] = pieces[0][0].lstrip()
        pieces[-1][0] = pieces[-1][0].rstrip()
        self.blocks.append((self._kind(), [(text, link) for text, link in pieces]))

    def shownStartTag(self, tag, attributes):
        if tag == "img":
            alt = _clean(attributes.get("alt") or "").strip()
            if alt:
                self._runs.append((alt, self.link, True))
        elif tag == "br":
            if self._kind() in _HEADINGS.values():
                self._runs.append((" ", self.link, False))
            else:
                self._endBlock()
        elif tag in _BLOCK_ELEMENTS:
            self._endBlock()
            if tag in _HEADINGS:
                self._kinds.append(_HEADINGS[tag])
            elif tag == "li":
                self._kinds.append("li")

    def shownEndTag(self, tag):
        if tag in _BLOCK_ELEMENTS:
            self._endBlock()
            kind = _HEADINGS.get(tag, "li" if tag == "li" else None)
            if kind in self._kinds:
                # Close it, and anything left open inside it.
                del self._kinds[len(self._kinds) - 1 - self._kinds[::-1].index(kind):]

    def shownText(self, data):
        self._runs.append((data, self.link, False))

    def close(self):
        super().close()
        self._endBlock()


def _blockHtml(pieces):
    out = []
    for text, link in pieces:
        if link is None:
            out.append(escape(text, quote=False))
            continue
        # Spaces go outside the link, so it is read as just its words.
        inner = text.strip()
        lead = " " if text[:1].isspace() else ""
        trail = " " if text[-1:].isspace() else ""
        out.append('%s<a href="%s">%s</a>%s' % (lead, escape(link, quote=True), escape(inner, quote=False), trail))
    return "".join(out)


def reformat(html, title, sender=None):
    """Return (page, blocks, links): a plain HTML page of the message, and how many blocks and links it has."""
    blocks = _Reformatter().read(html).blocks
    parts = [
        "<!DOCTYPE html>",
        "<html>",
        "<head>",
        '<meta charset="utf-8">',
        # Nothing is loaded from anywhere, so reading the page tells no one it was read.
        "<meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'none'; style-src 'unsafe-inline'\">",
        '<meta name="referrer" content="no-referrer">',
        "<title>%s</title>" % escape(title or "", quote=False),
        "<style>body { font-family: 'Segoe UI', sans-serif; line-height: 1.5; "
        "max-width: 45em; margin: 1em auto; padding: 0 1em; }</style>",
        "</head>",
        "<body>",
        "<h1>%s</h1>" % escape(title or "", quote=False),
    ]
    if sender:
        parts.append("<p>%s</p>" % escape(sender, quote=False))
    inList = False
    links = 0
    for kind, pieces in blocks:
        if (kind == "li") != inList:
            parts.append("<ul>" if not inList else "</ul>")
            inList = not inList
        links += sum(1 for _text, link in pieces if link)
        parts.append("<%s>%s</%s>" % (kind, _blockHtml(pieces), kind))
    if inList:
        parts.append("</ul>")
    parts.extend(("</body>", "</html>", ""))
    return "\n".join(parts), len(blocks), links

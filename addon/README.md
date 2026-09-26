# Outlook First Line Silence

An NVDA add-on that keeps Outlook message opening quiet. When you open an Outlook message, it suppresses the automatic inspector title, dialog/document container announcement, and initial line of the message. When focus lands in the body of a message you write, it says “edit”, as JAWS does, or, if you choose, “You are now in the message body, type a message.” It leaves normal Outlook navigation and user-requested speech alone.

The add-on also preserves Outlook Send/Receive progress announcements while the message-opening speech is being suppressed.

## Requirements

- NVDA 2026.1.0 or later (tested with stable NVDA 2026.1.1)
- Microsoft Outlook, including supported Outlook message-inspector windows

## Install

1. Download the current `outlookFirstLineSilence-1.0.30.nvda-addon` package.
2. Open the downloaded file, approve installation in NVDA, and restart NVDA when prompted.

## Keyboard commands

- **NVDA+Shift+X** reformats a badly formatted message. Press it in classic
  Outlook while reading a message, or with a message selected in the message
  list. The message opens in your web browser as a plain page: its headings,
  paragraphs, lists and links, where every line of a story is a link to it.
  The key only works in Outlook; everywhere else it does what it did before.
- **Enter** or **Space** on a story's headline, summary or "Read more" in a
  classic Outlook message opens the story, even when Outlook shows it without
  a link.
- **Check for updates** has no key. Use the NVDA menu: **Tools**, **Check for
  add-on updates**, **Outlook First Line Silence...**.

To change a key or add one, open the NVDA menu, choose **Preferences**, then
**Input gestures**, and expand **Outlook First Line Silence**. The commands are
**Reformats the Outlook message you are reading or have selected** and
**Checks for Outlook First Line Silence updates**. If NVDA+Shift+X does
something else in Outlook, another add-on uses the same key: give the reformat
command a different key there.

## What it changes

When an Outlook message inspector opens, the add-on briefly blocks only NVDA's automatic speech for that opening sequence:

- The message-window title.
- The dialog or document container announcement.
- The first line NVDA would otherwise announce automatically.

Coming back to a message window that is already open, with Alt+Tab, the taskbar, or when a window in front of it closes, is not opening it: NVDA says the window's title, as JAWS does, but still not the container announcement or the first line.

A window you write in, such as Forward, Reply or a new message, is not a message you read: NVDA says its title, as JAWS does, then the field or message body you land in. To, Cc, Bcc and Subject don't say “blank” when you land on them empty, or “multi line” just after, as JAWS doesn't; moving in an empty field still says “blank”. The message body says “edit”, as JAWS does, so Reply says its title, then “edit”. To hear “You are now in the message body, type a message” there instead, see Options and customization.

Pressing a key ends any short remaining suppression window immediately. Reading commands, navigation, and other normal Outlook speech are not intended to be muted.

Version 1.0.7 also suppresses the standalone “dialog” announcement when an Outlook message opens.

Version 1.0.10 uses Mute Browse Mode 3.6.57's link-line implementation without altering its setting behavior. When **Links are on their own line** is enabled, a link (and similar controls such as buttons) occupies its own arrow-key line, with the text before and after it on separate lines.

Version 1.0.11 uses Mute Browse Mode's reply-editor speech filter to remove the “left aligned” formatting announcement. It also uses Mute Browse Mode's bounded draft-prompt handling, which speaks the Save/Keep-draft question once and then lets button changes speak normally.

Version 1.0.14 keeps links and surrounding text separate at the beginning or end of a Word-rendered Outlook message. At that boundary, Up or Down Arrow re-announces the current logical segment instead of falling back to NVDA's complete inline line.

Version 1.0.17 announces when focus enters an editable Outlook message body. It also recognizes classic Outlook (`outlook`), new Outlook (`olk` and Outlook-owned WebView content), and the Windows Mail/Calendar-era names `commsapps`, `hxoutlook`, and `hxmail`. It deliberately preserves NVDA's built-in Outlook app modules.

Version 1.0.18 prevents repeated freezes if Outlook's Word RPC server becomes unavailable during link-line navigation. The failed Up or Down Arrow gesture and subsequent line movement in that document use native arrow keys without making another blocking COM request. Opening another message creates a new document and restores link-line navigation automatically.

Version 1.0.21 plays a sound when Outlook's list of suggested recipients appears while you address a message, and a different sound when the list goes away. The enter sound plays when the list opens after you type or use the Up, Down, Page Up, or Page Down keys in an address field. The exit sound plays when you press Escape, Enter, or Tab, when there are no more matches, or when focus leaves the address field. The sounds are `OutlookAutocompleteEnterSound.wav` and `OutlookAutocompleteExitSound.wav` in the add-on's `sounds` folder and play through NVDA's selected output device. They apply to address fields in message windows and to To, Cc, and Bcc fields elsewhere in Outlook, not to the search box.

Version 1.0.22 fixes Enter not opening a link in a Word-rendered Outlook message when **Links are on their own line** is enabled. On a line that contains a picture, such as the sender's avatar in a GitHub notification, Down Arrow left the cursor one character before the link, so Enter did nothing, the link was read with a stray space before it, and its last character was cut off. Link lines are now positioned by the line's text, so the cursor lands on the link itself and Enter opens it.

Version 1.0.23 fixes messages in Outlook's message list being announced without their status (unread, replied or forwarded, has attachment, importance) until NVDA was restarted. NVDA gets that status from Outlook's object model and keeps its connection to it for as long as Outlook runs. When that connection stopped answering, as it did for a tester right after installing 1.0.22 (Outlook replied “Unknown name” to every question), every message lost its status for the rest of the session. Now, when you move to a message, the add-on first asks Outlook what NVDA is about to ask. If the connection itself is broken, rather than Outlook being busy or nothing being selected, NVDA connects to Outlook again before the message is announced. Reconnecting starts NVDA's helper process and can pause NVDA for about a second, so it happens at most once every 30 seconds and stops after three attempts in a row that do not help.

Version 1.0.24 checks for updates. Once a day, a little after NVDA starts, the add-on asks its GitHub repository, [github.com/joshknnd1982/outlookFirstLineSilence](https://github.com/joshknnd1982/outlookFirstLineSilence), whether a newer version has been released, and says nothing unless there is one. When there is, a dialog shows what's new in a box you can read line by line, and offers to download and install it. The download must match the release's SHA-256 checksum. Then NVDA asks you to confirm the installation and offers to restart. Your settings are kept.

To check yourself, open the NVDA menu, choose **Tools**, then **Check for add-on updates**, and choose **Outlook First Line Silence...**. Or press **Check for updates now** in the add-on's settings: NVDA menu, Preferences, Settings, **Outlook First Line Silence**. You can also assign a gesture to **Checks for Outlook First Line Silence updates** in NVDA's Input Gestures dialog, under **Outlook First Line Silence**. To stop the daily check, clear **Check for Outlook First Line Silence updates automatically** in the same settings panel.

Version 1.0.25 opens stories that Outlook shows without their links. Some newsletters put one link around a whole story: its picture, headline, summary and “Read more”. Outlook can’t show a link like that, so the story was plain text, and Enter on its headline did nothing. Now Enter or Space on text in a classic Outlook message finds that text in the message’s own HTML and opens the link the sender put around it in your web browser. Links Outlook does show open as before.

It also reformats a message on request. Press **NVDA+Shift+V** (**NVDA+Shift+X** since version 1.0.27) while reading a message, or with a message selected in the message list, and the add-on shows it as a plain web page in your browser: its subject as a heading, then its headings, paragraphs, lists and links, without layout tables, pictures or the sender’s styles. Every line of a story is a link to it. Nothing on the page is loaded from the internet, so opening it doesn’t tell the sender you read the message. The page is a temporary file that is replaced each time and removed when NVDA exits. You can change the gesture in NVDA’s Input Gestures dialog, under Outlook First Line Silence.

Both features read the message through Outlook’s object model. If your antivirus is off or out of date, Outlook may ask whether to allow a program to access its data; that is this add-on asking for the message.

Version 1.0.26 stops “unread” being said twice when you delete a message. As Outlook deletes a message, or moves it out of the folder, it empties the message’s row in the message list before it moves to the next message. NVDA said what was left of the row, just its status, such as “unread”, and then the next message: “unread”, “unread From …”. The add-on now keeps that leftover status quiet, so you hear only the next message. The add-on also no longer looks for Outlook’s Save/Keep-draft prompt in other programs. It did that on every focus change, including in Notepad’s Save As dialog; now it does it only in Outlook.

Version 1.0.27 moves the command that reformats a message to **NVDA+Shift+X**. NVDA+Shift+V, its key in 1.0.25 and 1.0.26, is also used by the Vision Assistant and Say Product Name and Version add-ons, and NVDA doesn't choose between add-ons in a fixed order, so with either one installed NVDA+Shift+V could run the other add-on's command instead. NVDA itself doesn't use NVDA+Shift+X, and the add-on only takes the key in Outlook, so other programs and add-ons keep it. In Input Gestures the command is now called **Reformats the Outlook message you are reading or have selected**, so it can be found by searching for “reformat”, and a new Keyboard commands section lists the add-on's keys.

Version 1.0.28 says the whole title of a message window when you come back to it, as JAWS does. With 1.0.27, Alt+Tab back into an open message said only the start of its title. The Alt+Tab list began to say it, NVDA stopped speaking as the message came to the front, and then the add-on kept the message window's own title quiet, because it treated every arrival in a message window as opening the message. The add-on now remembers which message windows are open. Opening a message is still quiet. Coming back to one that is already open, with Alt+Tab, the taskbar, or when a reply or dialog in front of it closes, says its title, but still not “dialog”, “document” or the first line. Messages that were already open when NVDA started count as open.

Version 1.0.29 says the title of a message window you write in, as JAWS does. With 1.0.28, Control+F on a message said only “To edit blank”, then “multi line”, because the add-on kept the Forward window's title quiet, as it does when you open a message to read it. JAWS says the title, then “To Edit”. Whether a message window is one you read or one you write in is only known when focus lands in it, so the add-on keeps the title it dropped and says it, before the field, when focus lands in an address field or a message body you can type in. Forward, Reply, Reply All and New now say their window's title, then the field or message body. Opening a message you read is still quiet. To, Cc, Bcc and Subject no longer say “blank” when you land on them empty, or the “multi line” NVDA said as Outlook made an address field multi-line just after it got focus; moving in an empty field still says “blank”.

Version 1.0.30 says “edit” when you land in the body of a message you write, as JAWS does. With 1.0.29, Reply said its window's title, then “You are now in the message body, type a message.” JAWS says the title, then “edit”. The sentence is now a choice in the add-on's settings: **When you land in the body of a message you write, say** is **Edit, as JAWS says it** unless you choose **You are now in the message body, type a message**. The window's title comes first either way. Where NVDA says the body itself, such as a plain-text message's body, it already says “edit”, so the add-on doesn't say it again.

## Options and customization

Open NVDA's Settings, select **Outlook First Line Silence**, and enable **Links are on their own line**. It affects Outlook message reading only. With the option enabled, Up and Down Arrow stop on each link/control separately instead of reading it as part of the surrounding line. It does not change browser layout or other applications.

If NVDA's Browse Mode setting **Use screen layout (when supported)** is off, NVDA already separates controls into their own lines globally, so this option has no additional effect in Outlook's web-rendered messages.

In the same settings, **When you land in the body of a message you write, say** chooses what you hear when focus lands in the body of a Reply, Forward or new message: **Edit, as JAWS says it** (the default), or **You are now in the message body, type a message**, as up to version 1.0.29. The window's title comes first either way.

For developers, the two timing constants at the top of `globalPlugins/outlookFirstLineSilence/__init__.py` control the internal, short-lived suppression gate:

- `_IN_CALL_GATE` is the maximum duration of the active document-entry hook.
- `_TRAILING_GATE` is the brief follow-on period after entry.

Changing these values changes how long automatic opening speech may be blocked and is not needed for normal use. Repackage the add-on after any source change.

## Source layout

- `manifest.ini` — NVDA add-on metadata and compatibility requirements.
- `globalPlugins/outlookFirstLineSilence/__init__.py` — the global plugin implementation.
- `globalPlugins/outlookFirstLineSilence/messageHtml.py` — reads a message's own HTML for story links and the reformatted page.
- `globalPlugins/outlookFirstLineSilence/messageRows.py` — tells a deleted message's leftover row from a message.
- `globalPlugins/outlookFirstLineSilence/updater.py` — the GitHub update check.
- `globalPlugins/outlookFirstLineSilence/sounds/` — the recipient-suggestion enter and exit sounds.
- `doc/en/readme.html` — the in-NVDA add-on documentation.

## Credits

The implementation was derived from the verified document-entry and Outlook progress behavior in Mute Browse Mode 3.6.57 by Josh Kennedy.

Maintained by Dennis Long <dennisl@fastmail.com>.

Source: https://github.com/Dennisl123/outlookFirstLineSilence

Licensed under the GNU General Public License version 2.

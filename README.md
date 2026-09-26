# Outlook First Line Silence

An [NVDA](https://www.nvaccess.org/) screen reader add-on that keeps Outlook
message opening quiet. When you open an Outlook message, it suppresses the
automatic inspector title, the dialog/document container announcement, and
the first line of the message. When focus enters an editable message body, it
announces, "You are now in the message body, type a message." Normal Outlook
navigation and speech you ask for are left alone.

* Author: Dennis Long; based on Mute Browse Mode 3.6.57 by Josh Kennedy
* Version: 1.0.28
* Compatibility: NVDA 2026.1.0 or later (tested with 2026.1.1)
* Download: grab the `.nvda-addon` file from the
  [releases page](https://github.com/joshknnd1982/outlookFirstLineSilence/releases)

## Keyboard commands

* **NVDA+Shift+X** reformats a badly formatted message. Press it in classic
  Outlook while reading a message, or with a message selected in the message
  list. The message opens in your web browser as a plain page: its headings,
  paragraphs, lists and links, where every line of a story is a link to it.
  The key only works in Outlook; everywhere else it does what it did before.
* **Enter** or **Space** on a story's headline, summary or "Read more" in a
  classic Outlook message opens the story, even when Outlook shows it without
  a link.
* **Check for updates** has no key. Use the NVDA menu: **Tools**, **Check for
  add-on updates**, **Outlook First Line Silence...**.

To change a key or add one, open the NVDA menu, choose **Preferences**, then
**Input gestures**, and expand **Outlook First Line Silence**. The commands are
**Reformats the Outlook message you are reading or have selected** and
**Checks for Outlook First Line Silence updates**. If NVDA+Shift+X does
something else in Outlook, another add-on uses the same key: give the reformat
command a different key there.

## What it changes

When an Outlook message window opens, the add-on briefly blocks only NVDA's
automatic speech for that opening sequence:

* The message-window title.
* The dialog or document container announcement.
* The first line NVDA would otherwise announce automatically.

Coming back to a message window that is already open, with Alt+Tab, the
taskbar, or when a window in front of it closes, is not opening it: NVDA says
the window's title, as JAWS does, but still not the container announcement or
the first line.

Pressing a key ends any remaining suppression immediately. Outlook
Send/Receive progress announcements are preserved throughout.

It also:

* Removes the reply editor's "left aligned" formatting announcement.
* Speaks an Outlook Save/Keep-draft prompt once per dialog.
* Plays a sound when Outlook's list of suggested recipients appears while you
  address a message, and a different sound when the list goes away.
* Reconnects NVDA to Outlook when that connection stops answering, so
  messages in the message list keep their status (unread, replied or
  forwarded, has attachment) without restarting NVDA.
* Keeps quiet the status Outlook leaves behind in a message's row as it
  deletes the message, so Delete says only the next message, not "unread"
  and then "unread From ...".
* Recognizes classic Outlook, new Outlook (and Outlook-owned WebView
  content), and the Windows Mail/Calendar-era hosts.

## Badly formatted messages

The add-on opens stories that Outlook shows without their links. Some newsletters put one link around a whole story: its picture, headline, summary and “Read more”. Outlook can’t show a link like that, so the story was plain text, and Enter on its headline did nothing. Now Enter or Space on text in a classic Outlook message finds that text in the message’s own HTML and opens the link the sender put around it in your web browser. Links Outlook does show open as before.

It also reformats a message on request. Press **NVDA+Shift+X** while reading a message, or with a message selected in the message list, and the add-on shows it as a plain web page in your browser: its subject as a heading, then its headings, paragraphs, lists and links, without layout tables, pictures or the sender’s styles. Every line of a story is a link to it. Nothing on the page is loaded from the internet, so opening it doesn’t tell the sender you read the message. The page is a temporary file that is replaced each time and removed when NVDA exits. You can change the gesture in NVDA’s Input Gestures dialog, under Outlook First Line Silence. Up to version 1.0.26 the key was NVDA+Shift+V, which other add-ons also use.

Both features read the message through Outlook’s object model. If your antivirus is off or out of date, Outlook may ask whether to allow a program to access its data; that is this add-on asking for the message.

## Options

Open NVDA's Settings, select **Outlook First Line Silence**, and enable
**Links are on their own line**. With it enabled, Up and Down Arrow stop on
each link or similar control in an Outlook message separately, instead of
reading it as part of the surrounding line. It does not change web browsers
or other applications.

## Updates

The add-on checks for updates. Once a day, a little after NVDA starts, the add-on asks its GitHub repository, [github.com/joshknnd1982/outlookFirstLineSilence](https://github.com/joshknnd1982/outlookFirstLineSilence), whether a newer version has been released, and says nothing unless there is one. When there is, a dialog shows what's new in a box you can read line by line, and offers to download and install it. The download must match the release's SHA-256 checksum. Then NVDA asks you to confirm the installation and offers to restart. Your settings are kept.

To check yourself, open the NVDA menu, choose **Tools**, then **Check for add-on updates**, and choose **Outlook First Line Silence...**. Or press **Check for updates now** in the add-on's settings: NVDA menu, Preferences, Settings, **Outlook First Line Silence**. You can also assign a gesture to **Checks for Outlook First Line Silence updates** in NVDA's Input Gestures dialog, under **Outlook First Line Silence**. To stop the daily check, clear **Check for Outlook First Line Silence updates automatically** in the same settings panel.

## Installation

1. Download the latest `outlookFirstLineSilence-x.y.z.nvda-addon` file from
   the [releases page](https://github.com/joshknnd1982/outlookFirstLineSilence/releases).
2. Press enter on the downloaded file and confirm the installation in NVDA.
3. Restart NVDA when prompted.

## Building from source

Requires Python 3. From the repository root:

```bash
python build.py
```

This produces `outlookFirstLineSilence-1.0.28.nvda-addon` and its `.sha256`
checksum file in the repository root. Upload both to the GitHub release: the
update check reads the release's tag, such as `v1.0.28`, and checks the
download against the checksum.

## Repository layout

```
addon/
  manifest.ini                      Add-on metadata and changelog
  README.md, LICENSE.md             Shipped inside the add-on package
  globalPlugins/
    outlookFirstLineSilence/
      __init__.py                   The global plugin
      messageHtml.py                Reads a message's own HTML: story links
                                    and the reformatted page
      messageRows.py                Tells a deleted message's leftover row
                                    from a message
      updater.py                    The GitHub update check, shared by all of
                                    joshknnd1982's add-ons; keep it identical
      sounds/                       Recipient-suggestion enter and exit sounds
  doc/
    en/
      readme.html                   User documentation bundled with the add-on
tests/                              python -m unittest discover -s tests
build.py                            Builds the .nvda-addon package
```

The per-version history is in [addon/README.md](addon/README.md) and the
bundled `readme.html`.

## Credits

Derived from the document-entry and Outlook progress behavior in Mute Browse
Mode 3.6.57 by Josh Kennedy. Originally maintained by Dennis Long at
[Dennisl123/outlookFirstLineSilence](https://github.com/Dennisl123/outlookFirstLineSilence).

## License

GNU General Public License version 2. See [LICENSE](LICENSE).

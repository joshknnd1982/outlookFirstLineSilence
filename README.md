# Outlook First Line Silence

An [NVDA](https://www.nvaccess.org/) screen reader add-on that keeps Outlook
message opening quiet. When you open an Outlook message, it suppresses the
automatic inspector title, the dialog/document container announcement, and
the first line of the message. When focus enters an editable message body, it
announces, "You are now in the message body, type a message." Normal Outlook
navigation and speech you ask for are left alone.

* Author: Dennis Long; based on Mute Browse Mode 3.6.57 by Josh Kennedy
* Version: 1.0.23
* Compatibility: NVDA 2026.1.0 or later (tested with 2026.1.1)
* Download: grab the `.nvda-addon` file from the
  [releases page](https://github.com/joshknnd1982/outlookFirstLineSilence/releases)

## What it changes

When an Outlook message window opens, the add-on briefly blocks only NVDA's
automatic speech for that opening sequence:

* The message-window title.
* The dialog or document container announcement.
* The first line NVDA would otherwise announce automatically.

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
* Recognizes classic Outlook, new Outlook (and Outlook-owned WebView
  content), and the Windows Mail/Calendar-era hosts.

## Options

Open NVDA's Settings, select **Outlook First Line Silence**, and enable
**Links are on their own line**. With it enabled, Up and Down Arrow stop on
each link or similar control in an Outlook message separately, instead of
reading it as part of the surrounding line. It does not change web browsers
or other applications.

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

This produces `outlookFirstLineSilence-1.0.23.nvda-addon` in the repository
root.

## Repository layout

```
addon/
  manifest.ini                      Add-on metadata and changelog
  README.md, LICENSE.md             Shipped inside the add-on package
  globalPlugins/
    outlookFirstLineSilence/
      __init__.py                   The global plugin
      sounds/                       Recipient-suggestion enter and exit sounds
  doc/
    en/
      readme.html                   User documentation bundled with the add-on
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

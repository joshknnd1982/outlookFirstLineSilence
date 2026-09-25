# Outlook First Line Silence

An NVDA add-on that keeps Outlook message opening quiet. When you open an Outlook message, it suppresses the automatic inspector title, dialog/document container announcement, and initial line of the message. When focus enters an editable message body, it announces, “You are now in the message body, type a message.” It leaves normal Outlook navigation and user-requested speech alone.

The add-on also preserves Outlook Send/Receive progress announcements while the message-opening speech is being suppressed.

## Requirements

- NVDA 2026.1.0 or later (tested with stable NVDA 2026.1.1)
- Microsoft Outlook, including supported Outlook message-inspector windows

## Install

1. Download the current `outlookFirstLineSilence-1.0.22.nvda-addon` package.
2. Open the downloaded file, approve installation in NVDA, and restart NVDA when prompted.

## What it changes

When an Outlook message inspector opens, the add-on briefly blocks only NVDA's automatic speech for that opening sequence:

- The message-window title.
- The dialog or document container announcement.
- The first line NVDA would otherwise announce automatically.

Pressing a key ends any short remaining suppression window immediately. Reading commands, navigation, and other normal Outlook speech are not intended to be muted.

Version 1.0.7 also suppresses the standalone “dialog” announcement when an Outlook message opens.

Version 1.0.10 uses Mute Browse Mode 3.6.57's link-line implementation without altering its setting behavior. When **Links are on their own line** is enabled, a link (and similar controls such as buttons) occupies its own arrow-key line, with the text before and after it on separate lines.

Version 1.0.11 uses Mute Browse Mode's reply-editor speech filter to remove the “left aligned” formatting announcement. It also uses Mute Browse Mode's bounded draft-prompt handling, which speaks the Save/Keep-draft question once and then lets button changes speak normally.

Version 1.0.14 keeps links and surrounding text separate at the beginning or end of a Word-rendered Outlook message. At that boundary, Up or Down Arrow re-announces the current logical segment instead of falling back to NVDA's complete inline line.

Version 1.0.17 announces when focus enters an editable Outlook message body. It also recognizes classic Outlook (`outlook`), new Outlook (`olk` and Outlook-owned WebView content), and the Windows Mail/Calendar-era names `commsapps`, `hxoutlook`, and `hxmail`. It deliberately preserves NVDA's built-in Outlook app modules.

Version 1.0.18 prevents repeated freezes if Outlook's Word RPC server becomes unavailable during link-line navigation. The failed Up or Down Arrow gesture and subsequent line movement in that document use native arrow keys without making another blocking COM request. Opening another message creates a new document and restores link-line navigation automatically.

Version 1.0.21 plays a sound when Outlook's list of suggested recipients appears while you address a message, and a different sound when the list goes away. The enter sound plays when the list opens after you type or use the Up, Down, Page Up, or Page Down keys in an address field. The exit sound plays when you press Escape, Enter, or Tab, when there are no more matches, or when focus leaves the address field. The sounds are `OutlookAutocompleteEnterSound.wav` and `OutlookAutocompleteExitSound.wav` in the add-on's `sounds` folder and play through NVDA's selected output device. They apply to address fields in message windows and to To, Cc, and Bcc fields elsewhere in Outlook, not to the search box.

Version 1.0.22 fixes Enter not opening a link in a Word-rendered Outlook message when **Links are on their own line** is enabled. On a line that contains a picture, such as the sender's avatar in a GitHub notification, Down Arrow left the cursor one character before the link, so Enter did nothing, the link was read with a stray space before it, and its last character was cut off. Link lines are now positioned by the line's text, so the cursor lands on the link itself and Enter opens it.

## Options and customization

Open NVDA's Settings, select **Outlook First Line Silence**, and enable **Links are on their own line**. It affects Outlook message reading only. With the option enabled, Up and Down Arrow stop on each link/control separately instead of reading it as part of the surrounding line. It does not change browser layout or other applications.

If NVDA's Browse Mode setting **Use screen layout (when supported)** is off, NVDA already separates controls into their own lines globally, so this option has no additional effect in Outlook's web-rendered messages.

For developers, the two timing constants at the top of `globalPlugins/outlookFirstLineSilence/__init__.py` control the internal, short-lived suppression gate:

- `_IN_CALL_GATE` is the maximum duration of the active document-entry hook.
- `_TRAILING_GATE` is the brief follow-on period after entry.

Changing these values changes how long automatic opening speech may be blocked and is not needed for normal use. Repackage the add-on after any source change.

## Source layout

- `manifest.ini` — NVDA add-on metadata and compatibility requirements.
- `globalPlugins/outlookFirstLineSilence/__init__.py` — the global plugin implementation.
- `globalPlugins/outlookFirstLineSilence/sounds/` — the recipient-suggestion enter and exit sounds.
- `doc/en/readme.html` — the in-NVDA add-on documentation.

## Credits

The implementation was derived from the verified document-entry and Outlook progress behavior in Mute Browse Mode 3.6.57 by Josh Kennedy.

Maintained by Dennis Long <dennisl@fastmail.com>.

Source: https://github.com/Dennisl123/outlookFirstLineSilence

Licensed under the GNU General Public License version 2.

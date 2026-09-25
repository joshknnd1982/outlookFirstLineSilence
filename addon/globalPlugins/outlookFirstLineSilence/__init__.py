# -*- coding: utf-8 -*-
# Outlook First Line Silence 1.0.27
# Extracted from the verified document-entry behavior of Mute Browse Mode 3.6.57.
# Maintained by Dennis Long <dennisl@fastmail.com>.
# Licensed under the GNU General Public License version 2.

import builtins
import time
import re
import sys
import ctypes
import ctypes.wintypes
import os
import tempfile
from contextlib import contextmanager
from urllib.parse import urlsplit

from comtypes import COMError

import addonHandler
import api
import appModuleHandler
import browseMode
import controlTypes
import config
import core
import cursorManager
import globalPluginHandler
import inputCore
import nvwave
import scriptHandler
import speech
import speech.speech
import textInfos
import ui
import virtualBuffers
import winUser
import wx
from gui import guiHelper, settingsDialogs
from speech.priorities import SpeechPriority
from logHandler import log

from . import messageHtml
from . import messageRows
from . import updater

try:
    addonHandler.initTranslation()
except Exception:
    pass

_OUTLOOK_APP_NAMES = frozenset((
    "outlook",
    "olk",
    "commsapps",
    "hxoutlook",
    "hxmail",
    "msoutlook",
))
_OUTLOOK_WINDOW_CLASSES = frozenset((
    "rctrl_renwnd32",
    "Outlook Host",
))
_OUTLOOK_EMBEDDED_HOST_APP_NAMES = frozenset((
    "applicationframehost",
    "msedgewebview2",
    "webviewhost",
    "wwahost",
))

# These values are copied from muteBrowseMode 3.6.57's bounded document gate.
_IN_CALL_GATE = 6.0
_TRAILING_GATE = 1.5

_inCallDepth = 0
_inCallUntil = 0.0
_gateUntil = 0.0
_patches = []
_gestureHandlerRegistered = False
_messageOpeningUntil = 0.0
_savePromptKey = None
_savePromptUntil = 0.0

# Word's COM server can disappear while Outlook changes or closes a message.
# Retrying CursorManager repeats the same blocking RPC call, so use native arrow
# movement for the remainder of that disconnected document's lifetime.
_RECOVERABLE_WORD_RPC_HRESULTS = frozenset((
    -2147023174,  # 0x800706BA: RPC_S_SERVER_UNAVAILABLE
    -2147023170,  # 0x800706BE: RPC_S_CALL_FAILED
    -2147023169,  # 0x800706BF: RPC_S_CALL_FAILED_DNE
    -2147417848,  # 0x80010108: RPC_E_DISCONNECTED
))

# Mute Browse Mode 3.6.57 uses this bounded window to own one Outlook draft prompt,
# even when Outlook emits more than one focus event for the same dialog arrival.
_DRAFT_PROMPT_DUPLICATE_WINDOW = 2.5

# Exact bounded Outlook progress behavior carried over from muteBrowseMode 3.6.57.
_bypassDepth = 0
_progressSpeakMessage = (None, None)
_outlookProgressSpeechUntil = 0.0
_outlookProgressLastPercent = None
_outlookProgressLastEventAt = 0.0

_CONF_SECTION = "outlookFirstLineSilence"
config.conf.spec.setdefault(_CONF_SECTION, {})["linksOnOwnLine"] = "boolean(default=False)"


def getLinksOnOwnLine():
    try:
        return bool(config.conf[_CONF_SECTION]["linksOnOwnLine"])
    except Exception:
        return False


def setLinksOnOwnLine(enabled):
    config.conf[_CONF_SECTION]["linksOnOwnLine"] = bool(enabled)



class _ownSpeech:
    """Let add-on-owned progress speech past the document-entry gate (3.6.57)."""
    def __enter__(self):
        global _bypassDepth
        _bypassDepth += 1
        return self

    def __exit__(self, *exc):
        global _bypassDepth
        _bypassDepth = max(0, _bypassDepth - 1)
        return False


def _progressPercent(obj):
    """Return the whole percentage exposed by a progress bar, when available."""
    try:
        value = str(obj.value or "")
    except Exception:
        return None
    match = re.search(r"(?<!\\d)(\\d{1,3})(?:\\.\\d+)?\\s*%", value)
    if not match:
        return None
    return min(100, int(match.group(1)))


def _bypassingSpeakMessage(original):
    """speech.speakMessage, but never blocked by this add-on's gate."""
    global _progressSpeakMessage
    if _progressSpeakMessage[0] is original:
        return _progressSpeakMessage[1]

    def speakMessage(*args, **kwargs):
        kwargs = dict(kwargs)
        kwargs.setdefault("priority", SpeechPriority.NEXT)
        with _ownSpeech():
            return original(*args, **kwargs)

    speakMessage.__name__ = "speakMessage"
    speakMessage.__doc__ = getattr(original, "__doc__", None)
    _progressSpeakMessage = (original, speakMessage)
    return speakMessage


def _makeProgressBarWrapper(original):
    """Exact Outlook progress-bar bypass behavior from muteBrowseMode 3.6.57."""
    def event_valueChange(self, *args, **kwargs):
        global _outlookProgressSpeechUntil
        global _outlookProgressLastPercent, _outlookProgressLastEventAt
        current = getattr(speech, "speakMessage", None)
        if current is None:
            return original(self, *args, **kwargs)
        progressConf = None
        progressOriginals = None
        if _isInOutlookWindow(self):
            now = time.monotonic()
            percentage = _progressPercent(self)
            if percentage is not None:
                _outlookProgressLastPercent = percentage
                _outlookProgressLastEventAt = now
            _outlookProgressSpeechUntil = max(
                _outlookProgressSpeechUntil,
                now + (60.0 if percentage == 100 else 10.0),
            )
            try:
                progressConf = config.conf["presentation"]["progressBarUpdates"]
                progressOriginals = {
                    "progressBarOutputMode": progressConf["progressBarOutputMode"],
                    "reportBackgroundProgressBars": progressConf["reportBackgroundProgressBars"],
                    "speechPercentageInterval": progressConf["speechPercentageInterval"],
                }
                progressConf["progressBarOutputMode"] = (
                    "both" if progressOriginals["progressBarOutputMode"] in ("beep", "both") else "speak"
                )
                progressConf["reportBackgroundProgressBars"] = True
                progressConf["speechPercentageInterval"] = 1
            except Exception:
                progressConf = None
                progressOriginals = None
                log.debugWarning("Outlook First Line Silence: could not enable Outlook refresh progress speech", exc_info=True)
        bypassGate = _isInOutlookWindow(self)
        if bypassGate:
            speech.speakMessage = _bypassingSpeakMessage(current)
        try:
            return original(self, *args, **kwargs)
        finally:
            if bypassGate:
                speech.speakMessage = current
            if progressConf is not None and progressOriginals is not None:
                for key, value in progressOriginals.items():
                    progressConf[key] = value

    event_valueChange.__name__ = "event_valueChange"
    event_valueChange.__doc__ = getattr(original, "__doc__", None)
    return event_valueChange


def _makeCancelSpeechWrapper(original):
    """Protect only the bounded Outlook progress queue from automatic cancellation."""
    def cancelSpeech(*args, **kwargs):
        if time.monotonic() < _outlookProgressSpeechUntil:
            return
        return original(*args, **kwargs)
    cancelSpeech.__name__ = "cancelSpeech"
    cancelSpeech.__doc__ = getattr(original, "__doc__", None)
    return cancelSpeech


def _members(enum, names):
    return frozenset(m for m in (getattr(enum, name, None) for name in names) if m is not None)


# Same title/container roles and automatic output reasons used by muteBrowseMode 3.6.57.
_TITLE_ROLES = _members(
    controlTypes.Role,
    ("WINDOW", "PANE", "FRAME", "INTERNALFRAME", "DIALOG", "DOCUMENT", "APPLICATION", "PROPERTYPAGE"),
)
_AUTOMATIC_REASONS = _members(
    controlTypes.OutputReason,
    ("FOCUS", "FOCUSENTERED", "CHANGE", "CARET"),
)


def _appNameOf(obj):
    try:
        return (obj.appModule.appName or "").lower()
    except Exception:
        return ""


def _isOutlook(obj):
    return obj is not None and _appNameOf(obj) in _OUTLOOK_APP_NAMES


def _isOutlookObjectLineage(obj):
    """Recognize Outlook itself and UI hosted in an Outlook accessible-object tree."""
    if _isOutlook(obj):
        return True
    if _appNameOf(obj) not in _OUTLOOK_EMBEDDED_HOST_APP_NAMES:
        return False
    candidate = obj
    seen = set()
    for _depth in range(32):
        if candidate is None:
            return False
        identity = id(candidate)
        if identity in seen:
            return False
        seen.add(identity)
        if _isOutlook(candidate):
            return True
        try:
            candidate = candidate.parent
        except Exception:
            return False
    return False


def _windowClassOf(obj):
    try:
        return getattr(obj, "windowClassName", "") or ""
    except Exception:
        return ""


def _rootWindowOf(obj):
    try:
        hwnd = obj.windowHandle
    except Exception:
        return 0
    if not hwnd:
        return 0
    try:
        return winUser.getAncestor(hwnd, getattr(winUser, "GA_ROOT", 2)) or 0
    except Exception:
        return 0


def _processIDOf(obj):
    try:
        return obj.processID
    except Exception:
        return None


def _appNameOfProcess(processID):
    if not processID:
        return ""
    try:
        return (appModuleHandler.getAppModuleFromProcessID(processID).appName or "").lower()
    except Exception:
        return ""


def _isInOutlookWindow(obj):
    """Recognize classic, Store-era, and new Outlook-hosted windows."""
    if _isOutlookObjectLineage(obj):
        return True
    root = _rootWindowOf(obj)
    if not root:
        return False
    windows = [root]
    try:
        rootOwner = winUser.getAncestor(root, getattr(winUser, "GA_ROOTOWNER", 3)) or 0
        if rootOwner and rootOwner != root:
            windows.append(rootOwner)
    except Exception:
        pass
    for window in windows:
        try:
            if winUser.getClassName(window) in _OUTLOOK_WINDOW_CLASSES:
                return True
            processID = winUser.getWindowThreadProcessID(window)[0]
        except Exception:
            log.debugWarning("Outlook First Line Silence: could not inspect top-level window", exc_info=True)
            continue
        if processID and _appNameOfProcess(processID) in _OUTLOOK_APP_NAMES:
            return True
    return False


# Exact editable-message-body identification from Mute Browse Mode 3.6.57.
_OUTLOOK_BODY_WINDOW_CLASSES = frozenset(("_WwG",))
_PLAIN_TEXT_BODY_CLASS = "RichEdit20W"
_PLAIN_TEXT_BODY_CONTROL_ID = 8224
_BODY_NAMES = frozenset(("message", "message body"))
_BODY_ROLES = _members(controlTypes.Role, ("DOCUMENT", "EDITABLETEXT"))
_STATE_READONLY = getattr(controlTypes.State, "READONLY", None)
_STATE_UNAVAILABLE = getattr(controlTypes.State, "UNAVAILABLE", None)
_STATE_MULTILINE = getattr(controlTypes.State, "MULTILINE", None)


def _hasState(states, state):
    return state is not None and state in states


def _isOutlookMessageBody(obj):
    if not _isInOutlookWindow(obj):
        return False
    if getattr(obj, "role", None) not in _BODY_ROLES:
        return False
    try:
        states = set(obj.states or ())
    except Exception:
        states = set()
    if _hasState(states, _STATE_READONLY) or _hasState(states, _STATE_UNAVAILABLE):
        return False

    windowClass = _windowClassOf(obj)
    viewer = getattr(obj, "isReadonlyViewer", None)
    if viewer is not None:
        return not viewer and windowClass in _OUTLOOK_BODY_WINDOW_CLASSES
    if windowClass in _OUTLOOK_BODY_WINDOW_CLASSES:
        return True
    if (
        windowClass == _PLAIN_TEXT_BODY_CLASS
        and getattr(obj, "windowControlID", None) == _PLAIN_TEXT_BODY_CONTROL_ID
    ):
        return True
    if not _hasState(states, _STATE_MULTILINE):
        return False
    return (getattr(obj, "name", "") or "").strip().lower() in _BODY_NAMES


def _announceMessageBody():
    # Translators: Announced when focus reaches the editable message body in Outlook.
    with _ownSpeech():
        ui.message(_("You are now in the message body, type a message."))


# Recipient-suggestion sounds.
# While you address a message, Outlook drops down a list of matching people. The enter
# sound plays when that list appears (you have entered your suggestions) and the exit
# sound plays when it goes away: Escape, Enter, Tab, no more matches, or focus leaving
# the address field.
# The list is recognized two ways. (1) After you type or use Up/Down/Page keys in an
# address field, the add-on briefly watches for a new suggestion-list popup window that
# belongs to that Outlook window. (2) Whenever NVDA's own AutoCompleteListItem events say
# a suggestion has been selected while focus is in the address field, the list is open.
_SOUNDS_DIRECTORY = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sounds")
_SUGGESTIONS_ENTER_SOUND = os.path.join(_SOUNDS_DIRECTORY, "OutlookAutocompleteEnterSound.wav")
_SUGGESTIONS_EXIT_SOUND = os.path.join(_SOUNDS_DIRECTORY, "OutlookAutocompleteExitSound.wav")
# Same window classes NVDA's Outlook app module treats as auto-complete lists.
_SUGGESTION_WINDOW_CLASS_PREFIXES = ("REListBox", "NetUIHWND")
_RECIPIENT_FIELD_NAMES = frozenset(("to", "cc", "bcc"))
_SUGGESTION_POLL_MS = 100
# Leaving is confirmed over this many consecutive polls, so a list that flickers while
# it re-filters as you type never counts as having closed.
_SUGGESTION_EXIT_CONFIRMATIONS = 3
# After a key press, keep watching for the list to appear for this long.
_SUGGESTION_WATCH_SECONDS = 3.0
# Selection events that arrive this soon after leaving are the closing list, not a new entry.
_SUGGESTION_REENTRY_COOLDOWN = 0.6
_SUGGESTION_MIN_POPUP_SIZE = (40, 10)
_SUGGESTION_LEAVE_KEYS = frozenset(("escape", "enter", "numpadenter", "tab"))
_SUGGESTION_WATCH_KEYS = frozenset((
    "downarrow", "uparrow", "pagedown", "pageup", "backspace", "delete", "space",
))
_MODIFIER_KEYS = frozenset(("control", "alt", "windows", "nvda"))
_STATE_SELECTED = getattr(controlTypes.State, "SELECTED", None)
_STATE_INVISIBLE = getattr(controlTypes.State, "INVISIBLE", None)
_STATE_OFFSCREEN = getattr(controlTypes.State, "OFFSCREEN", None)

_suggestionsActive = False
_suggestionItem = None
_suggestionFocus = None
_suggestionMisses = 0
_suggestionPollToken = 0
_suggestionPollRunning = False
_suggestionWatchUntil = 0.0
_suggestionBaseline = frozenset()
_suggestionListSeen = False
_suggestionCooldownUntil = 0.0
_watchableFocusCache = (None, False)
_privateUser32 = None


def _playSuggestionSound(path):
    try:
        nvwave.playWaveFile(path, asynchronous=True)
    except Exception:
        log.debugWarning(
            "Outlook First Line Silence: could not play %s" % os.path.basename(path),
            exc_info=True,
        )


def _isSuggestionItem(obj):
    """True for an item of Outlook's drop-down suggestion list."""
    if obj is None or getattr(obj, "role", None) != controlTypes.Role.LISTITEM:
        return False
    if not _isInOutlookWindow(obj):
        return False
    if _appNameOf(obj) == "outlook":
        return _windowClassOf(obj).startswith(_SUGGESTION_WINDOW_CLASS_PREFIXES)
    # New Outlook and the Windows Mail-era hosts do not use those window classes.
    return True


def _isRecipientField(focus):
    """True for an editable address field (or any edit box in a message window)."""
    if focus is None or getattr(focus, "role", None) != controlTypes.Role.EDITABLETEXT:
        return False
    if _isOutlookMessageInspector(focus):
        return True
    if not _isInOutlookWindow(focus):
        return False
    try:
        name = (focus.name or "").strip().rstrip(":").lower()
    except Exception:
        return False
    return name in _RECIPIENT_FIELD_NAMES


def _isWatchableFocus(focus):
    """An address field worth watching for a suggestion list (not the body or Subject)."""
    global _watchableFocusCache
    cachedFocus, cachedResult = _watchableFocusCache
    if cachedFocus is not None and (cachedFocus is focus or cachedFocus == focus):
        return cachedResult
    result = False
    try:
        if focus is not None and _appNameOf(focus) in _OUTLOOK_APP_NAMES | _OUTLOOK_EMBEDDED_HOST_APP_NAMES:
            result = (
                _isRecipientField(focus)
                and not _isOutlookMessageBody(focus)
                and (getattr(focus, "name", "") or "").strip().rstrip(":").lower() != "subject"
            )
    except Exception:
        result = False
    _watchableFocusCache = (focus, result)
    return result


def _isSelectedVisibleItem(item):
    """The same conditions NVDA uses before it speaks an auto-complete suggestion."""
    try:
        states = item.states
    except Exception:
        return False
    if _STATE_SELECTED is None or _STATE_SELECTED not in states:
        return False
    return not any(
        state is not None and state in states
        for state in (_STATE_INVISIBLE, _STATE_OFFSCREEN, _STATE_UNAVAILABLE)
    )


def _user32():
    """A private user32 handle, so this add-on never alters NVDA's own function prototypes."""
    global _privateUser32
    if _privateUser32 is None:
        library = ctypes.WinDLL("user32")
        library.IsWindowVisible.argtypes = [ctypes.c_void_p]
        library.IsWindowVisible.restype = ctypes.c_int
        library.GetWindowRect.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.wintypes.RECT)]
        library.GetWindowRect.restype = ctypes.c_int
        library.AllowSetForegroundWindow.argtypes = [ctypes.wintypes.DWORD]
        library.AllowSetForegroundWindow.restype = ctypes.c_int
        _privateUser32 = library
    return _privateUser32


def _enumerateVisibleTopLevelWindows():
    windows = []
    library = _user32()
    callbackType = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p)

    def collect(hwnd, _lParam):
        try:
            if hwnd and library.IsWindowVisible(hwnd):
                windows.append(hwnd)
        except Exception:
            pass
        return 1

    library.EnumWindows.argtypes = [callbackType, ctypes.c_void_p]
    library.EnumWindows.restype = ctypes.c_int
    callback = callbackType(collect)  # keep a reference alive for the whole call
    library.EnumWindows(callback, None)
    return windows


def _windowSize(hwnd):
    rect = ctypes.wintypes.RECT()
    if not _user32().GetWindowRect(hwnd, ctypes.byref(rect)):
        return 0, 0
    return rect.right - rect.left, rect.bottom - rect.top


def _ownedTopLevelWindows(focus):
    """Visible top-level windows that belong to the focused Outlook window."""
    hwnd = getattr(focus, "windowHandle", None)
    if not hwnd:
        return []
    rootOwner = winUser.getAncestor(hwnd, getattr(winUser, "GA_ROOTOWNER", 3)) or 0
    if not rootOwner:
        return []
    threadID = winUser.getWindowThreadProcessID(hwnd)[1]
    owned = []
    for window in _enumerateVisibleTopLevelWindows():
        if window == rootOwner:
            continue
        try:
            sameOwner = (winUser.getAncestor(window, getattr(winUser, "GA_ROOTOWNER", 3)) or 0) == rootOwner
            sameThread = winUser.getWindowThreadProcessID(window)[1] == threadID
        except Exception:
            continue
        if sameOwner or sameThread:
            owned.append(window)
    return owned


def _suggestionPopups(focus):
    """Visible suggestion-list popup windows that belong to the focused Outlook window."""
    found = set()
    minWidth, minHeight = _SUGGESTION_MIN_POPUP_SIZE
    for window in _ownedTopLevelWindows(focus):
        try:
            if not winUser.getClassName(window).startswith(_SUGGESTION_WINDOW_CLASS_PREFIXES):
                continue
            width, height = _windowSize(window)
        except Exception:
            continue
        if width >= minWidth and height >= minHeight:
            found.add(window)
    return found


def _suggestionListPresent(focus):
    if _suggestionItem is not None and _isSelectedVisibleItem(_suggestionItem):
        return True
    return bool(_suggestionPopups(focus) - _suggestionBaseline)


def _resetSuggestions():
    global _suggestionsActive, _suggestionItem, _suggestionFocus, _suggestionMisses
    global _suggestionPollToken, _suggestionPollRunning, _suggestionWatchUntil
    global _suggestionBaseline, _suggestionListSeen
    _suggestionPollToken += 1
    _suggestionPollRunning = False
    _suggestionsActive = False
    _suggestionItem = None
    _suggestionFocus = None
    _suggestionMisses = 0
    _suggestionWatchUntil = 0.0
    _suggestionBaseline = frozenset()
    _suggestionListSeen = False


def _enterSuggestions(reason, focus, item=None):
    global _suggestionsActive, _suggestionItem, _suggestionFocus, _suggestionMisses, _suggestionListSeen
    _suggestionsActive = True
    _suggestionListSeen = True
    _suggestionFocus = focus
    _suggestionMisses = 0
    if item is not None:
        _suggestionItem = item
    log.debug("Outlook First Line Silence: entered recipient suggestions (%s)" % reason)
    _playSuggestionSound(_SUGGESTIONS_ENTER_SOUND)
    _ensureSuggestionPoll()


def _leaveSuggestions(reason):
    global _suggestionCooldownUntil
    wasActive = _suggestionsActive
    _resetSuggestions()
    if wasActive:
        _suggestionCooldownUntil = time.monotonic() + _SUGGESTION_REENTRY_COOLDOWN
        log.debug("Outlook First Line Silence: left recipient suggestions (%s)" % reason)
        _playSuggestionSound(_SUGGESTIONS_EXIT_SOUND)


def _ensureSuggestionPoll():
    global _suggestionPollRunning
    if _suggestionPollRunning:
        return
    _suggestionPollRunning = True
    core.callLater(_SUGGESTION_POLL_MS, _pollSuggestions, _suggestionPollToken)


def _finishSuggestionWatch(focus):
    """The watch ended without the list appearing; note what was on screen, for the log."""
    global _suggestionPollToken, _suggestionPollRunning
    _suggestionPollToken += 1
    _suggestionPollRunning = False
    if _suggestionListSeen or focus is None:
        return
    try:
        classes = sorted({winUser.getClassName(window) for window in _ownedTopLevelWindows(focus)})
        log.debug(
            "Outlook First Line Silence: no suggestion list appeared; visible windows owned by "
            "the message window: %s" % (", ".join(classes) or "none")
        )
    except Exception:
        pass


def _pollSuggestions(token):
    global _suggestionMisses
    if token != _suggestionPollToken:
        return
    focus = None
    try:
        focus = api.getFocusObject()
        inField = bool(focus == _suggestionFocus or _isRecipientField(focus))
        present = inField and _suggestionListPresent(focus)
    except Exception:
        inField = present = False
    if _suggestionsActive:
        if present:
            _suggestionMisses = 0
        else:
            _suggestionMisses += 1
            if _suggestionMisses >= _SUGGESTION_EXIT_CONFIRMATIONS:
                _leaveSuggestions("the suggestion list is gone")
                return
    elif present:
        _enterSuggestions("the suggestion list appeared", focus)
    elif not inField or time.monotonic() >= _suggestionWatchUntil:
        _finishSuggestionWatch(focus if inField else None)
        return
    core.callLater(_SUGGESTION_POLL_MS, _pollSuggestions, token)


def _watchForSuggestionList():
    """Called on keys that make Outlook show or change its suggestion list."""
    global _suggestionWatchUntil, _suggestionFocus, _suggestionBaseline, _suggestionListSeen
    focus = api.getFocusObject()
    if not _isWatchableFocus(focus):
        return
    _suggestionWatchUntil = time.monotonic() + _SUGGESTION_WATCH_SECONDS
    if _suggestionsActive or _suggestionPollRunning:
        return
    _suggestionFocus = focus
    _suggestionListSeen = False
    try:
        # Only a popup that appears after this moment is the suggestion list.
        _suggestionBaseline = frozenset(_suggestionPopups(focus))
    except Exception:
        _suggestionBaseline = frozenset()
    _ensureSuggestionPoll()


def _trackSuggestionSelection(obj):
    """NVDA's own suggestion events: a selected suggestion means the list is open."""
    global _suggestionItem, _suggestionFocus, _suggestionMisses
    try:
        if not _isSuggestionItem(obj) or not _isSelectedVisibleItem(obj):
            return
        focus = api.getFocusObject()
        if not _isRecipientField(focus):
            return
        if _suggestionsActive:
            _suggestionItem = obj
            _suggestionFocus = focus
            _suggestionMisses = 0
        elif time.monotonic() >= _suggestionCooldownUntil:
            _enterSuggestions("a suggestion was selected", focus, obj)
    except Exception:
        log.debugWarning("Outlook First Line Silence: could not track recipient suggestions", exc_info=True)


def _suggestionGestureKeys(gesture):
    try:
        identifiers = gesture.normalizedIdentifiers
    except Exception:
        return None
    for identifier in identifiers or ():
        try:
            return set(identifier.casefold().split(":", 1)[1].split("+"))
        except (AttributeError, IndexError):
            continue
    return None


def _noteSuggestionGesture(gesture):
    """Keys that open, change, or dismiss the suggestion list."""
    keys = _suggestionGestureKeys(gesture)
    if not keys or keys & _MODIFIER_KEYS:
        return
    plain = keys - {"shift"}
    if len(plain) != 1:
        return
    key = next(iter(plain))
    if key in _SUGGESTION_LEAVE_KEYS:
        if _suggestionsActive:
            _leaveSuggestions("%s pressed" % key)
        return
    if key in _SUGGESTION_WATCH_KEYS or getattr(gesture, "isCharacter", False):
        _watchForSuggestionList()


def _suggestionsFocusChanged(newFocus):
    """Leave the list at once when focus goes somewhere other than the address field."""
    if not _suggestionsActive:
        return
    try:
        if (
            newFocus == _suggestionFocus
            or _isRecipientField(newFocus)
            or _isSuggestionItem(newFocus)
        ):
            return
    except Exception:
        pass
    _leaveSuggestions("focus moved")


# This is the Outlook-only counterpart of NVDA's "Use screen layout" behavior.
# It intentionally leaves every other application and web page untouched.
_SPLIT_CONTROL_ROLES = _members(
    controlTypes.Role,
    ("LINK", "BUTTON", "TOGGLEBUTTON", "MENUBUTTON", "DROPDOWNBUTTON",
     "SPLITBUTTON", "CHECKBOX", "RADIOBUTTON", "COMBOBOX", "EDITABLETEXT",
     "SLIDER", "SPINBUTTON"),
)


def _isVirtualBufferOutlookMessage(textInfo):
    """Whether this text is an Outlook message rendered as a web document."""
    buffer = getattr(textInfo, "obj", None)
    if buffer is None or not getattr(buffer, "VBufHandle", None):
        return False
    return _isInOutlookWindow(getattr(buffer, "rootNVDAObject", None))


def _shouldSplitLines(textInfo):
    """Whether this line should be broken up so each control gets one of its own."""
    if not getLinksOnOwnLine():
        return False
    try:
        if not config.conf["virtualBuffers"]["useScreenLayout"]:
            return False
    except Exception:
        pass
    return _isVirtualBufferOutlookMessage(textInfo)


def _makeLineOffsetsWrapper(original):
    """Work out a line as if screen layout were off, for Outlook messages only."""
    def _getLineOffsets(self, offset):
        try:
            split = _shouldSplitLines(self)
        except Exception:
            split = False
        if not split:
            return original(self, offset)
        try:
            import NVDAHelper
            lineStart = ctypes.c_int()
            lineEnd = ctypes.c_int()
            NVDAHelper.localLib.VBuf_getLineOffsets(
                self.obj.VBufHandle,
                offset,
                config.conf["virtualBuffers"]["maxLineLength"],
                False,
                ctypes.byref(lineStart),
                ctypes.byref(lineEnd),
            )
            return lineStart.value, lineEnd.value
        except Exception:
            log.debugWarning("Outlook First Line Silence: could not split the line", exc_info=True)
            return original(self, offset)
    return _getLineOffsets


def _shouldWalkSegments(treeInterceptor):
    """Whether down and up arrow should step through this document in segments."""
    if not getLinksOnOwnLine():
        log.debug("Outlook First Line Silence links: disabled")
        return False
    if not isinstance(treeInterceptor, browseMode.BrowseModeDocumentTreeInterceptor):
        log.debug("Outlook First Line Silence links: not a browse document (%s)", type(treeInterceptor).__name__)
        return False
    if getattr(treeInterceptor, "passThrough", False):
        log.debug("Outlook First Line Silence links: Word document is in focus mode")
        return False
    if getattr(treeInterceptor, "VBufHandle", None):
        log.debug("Outlook First Line Silence links: virtual buffer uses line offsets")
        return False
    result = _isInOutlookWindow(getattr(treeInterceptor, "rootNVDAObject", None))
    log.debug("Outlook First Line Silence links: Word segment walk=%s", result)
    return result


def _isSplittableControl(field):
    try:
        return field.get("role") in _SPLIT_CONTROL_ROLES
    except Exception:
        return False


def _lineSegments(lineInfo):
    """Return (start, end, lineText) ranges that place each link/control on its own line.

    The offsets index lineText, the line's text as its fields report it.
    """
    try:
        fields = lineInfo.getTextWithFields()
    except Exception:
        log.debugWarning("Outlook First Line Silence links: could not read line fields", exc_info=True)
        return None
    offset, stack, edges, chunks = 0, [], set(), []
    for field in fields:
        if isinstance(field, str):
            chunks.append(field)
            offset += len(field)
            continue
        command = getattr(field, "command", None)
        if command == "controlStart":
            stack.append((offset, _isSplittableControl(getattr(field, "field", None))))
        elif command == "controlEnd" and stack:
            start, splittable = stack.pop()
            if splittable and offset > start:
                edges.update((start, offset))
    if offset <= 0:
        log.debug("Outlook First Line Silence links: blank line")
        return None
    text = "".join(chunks)
    whole = [(0, offset, text)]
    if not edges:
        log.debug("Outlook First Line Silence links: %d-character line has no control fields", offset)
        return whole
    edges.update((0, offset))
    bounds = sorted(edge for edge in edges if 0 <= edge <= offset)
    segments = [
        (start, end, text)
        for start, end in zip(bounds, bounds[1:])
        if end > start and text[start:end].strip()
    ]
    result = segments or whole
    log.debug("Outlook First Line Silence links: %d control edges, %d segments", len(edges), len(result))
    return result


def _segmentInfoByCharacters(lineInfo, start, end):
    segment = lineInfo.copy()
    segment.collapse()
    if start and segment.move(textInfos.UNIT_CHARACTER, start) != start:
        return None
    finish = lineInfo.copy()
    finish.collapse()
    if end and finish.move(textInfos.UNIT_CHARACTER, end) != end:
        return None
    segment.setEndPoint(finish, "endToStart")
    return segment


# An inline picture has no text, so a text offset next to one could be either side
# of it. Step past at most this many to reach the segment's first real character.
_TEXTLESS_UNIT_LIMIT = 8


def _textPoint(lineInfo, lineText, offset):
    """A collapsed TextInfo *offset* characters into lineText."""
    if offset <= 0 or offset >= len(lineText):
        point = lineInfo.copy()
        point.collapse(end=offset > 0)
        return point
    # The browse mode proxy does not pass Word's own text through to this
    # conversion, so ask the Word TextInfo it wraps, and reuse the line text
    # already fetched rather than reading the whole line again.
    inner = getattr(lineInfo, "innerTextInfo", None)
    source = (lineInfo if inner is None else inner).copy()
    source._getTextForCodepointMovement = lambda: lineText
    point = source.moveToCodepointOffset(offset)
    # Land on the link's first character rather than on a picture just before it,
    # or Enter would activate the picture.
    for _step in range(_TEXTLESS_UNIT_LIMIT):
        unit = point.copy()
        if not unit.move(textInfos.UNIT_CHARACTER, 1, endPoint="end") or unit._getTextForCodepointMovement():
            break
        point.move(textInfos.UNIT_CHARACTER, 1)
    return point if inner is None else lineInfo.__class__(lineInfo.obj, point)


def _segmentInfo(lineInfo, start, end, lineText=None):
    segment = _segmentInfoByCharacters(lineInfo, start, end)
    if lineText is None or (segment is not None and segment.text == lineText[start:end]):
        return segment
    # Word counts an inline picture (such as a sender's avatar), a list bullet or a
    # table end-of-row mark differently from the line's text, so moving by characters
    # can stop short of a link. The caret is then left just outside the link, where
    # Enter cannot activate it. Map the text offsets onto the document instead.
    try:
        mapped = _textPoint(lineInfo, lineText, start)
        mapped.setEndPoint(_textPoint(lineInfo, lineText, end), "endToStart")
    except (ValueError, RuntimeError):
        log.debugWarning(
            "Outlook First Line Silence links: could not map segment %d-%d" % (start, end),
            exc_info=True,
        )
        return segment
    log.debug("Outlook First Line Silence links: realigned segment %d-%d to the line text", start, end)
    return mapped


def _caretOffsetInLine(lineInfo, caretInfo):
    prefix = lineInfo.copy()
    prefix.setEndPoint(caretInfo, "endToStart")
    return len(prefix.text or "")


def _walkSegment(treeInterceptor, gesture, direction):
    caret = treeInterceptor.makeTextInfo(textInfos.POSITION_CARET)
    line = caret.copy()
    line.expand(textInfos.UNIT_LINE)
    segments = _lineSegments(line)
    target = None
    current = None
    if segments:
        here = _caretOffsetInLine(line, caret)
        index = 0
        for position, (start, _end, _text) in enumerate(segments):
            if start <= here:
                index = position
        current = (line, segments[index])
        wanted = index + direction
        if 0 <= wanted < len(segments):
            target = _segmentInfo(line, *segments[wanted])
    if target is None:
        line = caret.copy()
        line.expand(textInfos.UNIT_LINE)
        line.collapse()
        if line.move(textInfos.UNIT_LINE, direction) == 0:
            # At a document boundary NVDA would re-announce the full Word line,
            # including an inline link. Re-announce the current logical segment
            # instead, so the first Up Arrow after a quiet message opening still
            # preserves the separate text and link lines.
            currentTarget = None if current is None else _segmentInfo(current[0], *current[1])
            if currentTarget is None:
                log.debug("Outlook First Line Silence links: no adjacent Word line")
                return False
            target = currentTarget
            log.debug("Outlook First Line Silence links: retained current segment at Word boundary")
        else:
            # Word's UIA text provider does not always expose the link field until that
            # line becomes the actual caret line. Set the adjacent line first, then
            # obtain a fresh TextInfo before calculating its segments.
            try:
                treeInterceptor.selection = line
                line = treeInterceptor.makeTextInfo(textInfos.POSITION_CARET)
            except Exception:
                pass
            line.expand(textInfos.UNIT_LINE)
            segments = _lineSegments(line)
            if not segments:
                log.debug("Outlook First Line Silence links: adjacent Word line has no segments")
                target = _segmentInfo(line, 0, len(line.text or ""))
            else:
                target = _segmentInfo(line, *(segments[0] if direction > 0 else segments[-1]))
    if target is None:
        log.debug("Outlook First Line Silence links: no target segment")
        return False
    selection = target.copy()
    selection.collapse()
    willResume = False
    try:
        willResume = scriptHandler.willSayAllResume(gesture)
    except Exception:
        pass
    if not willResume:
        speech.speakTextInfo(target, unit=textInfos.UNIT_LINE, reason=controlTypes.OutputReason.CARET)
    try:
        treeInterceptor.selection = selection
    except Exception:
        log.error("Outlook First Line Silence: could not move to link segment", exc_info=True)
    log.debug("Outlook First Line Silence links: moved %s", "down" if direction > 0 else "up")
    return True


def _hresultOf(error):
    """Return the signed HRESULT carried by a COM error, or None."""
    hresult = getattr(error, "hresult", None)
    if hresult is None:
        args = getattr(error, "args", ())
        if args:
            hresult = args[0]
    try:
        hresult = int(hresult)
    except (TypeError, ValueError):
        return None
    if hresult > 0x7FFFFFFF:
        hresult -= 0x100000000
    return hresult


def _isRecoverableWordRpcError(error):
    """Return whether *error* means Outlook's Word RPC server disconnected."""
    return _hresultOf(error) in _RECOVERABLE_WORD_RPC_HRESULTS


def _sendNativeLineGesture(gesture):
    """Move in Word without another TextInfo/COM request."""
    try:
        return gesture.send()
    except Exception:
        log.error("Outlook First Line Silence: native line movement failed", exc_info=True)


def _makeMoveByLineWrapper(original, direction):
    def script_moveByLine(self, gesture):
        try:
            walk = _shouldWalkSegments(self)
        except Exception:
            walk = False
        if not walk:
            return original(self, gesture)
        if getattr(self, "_outlookFirstLineSilenceNativeLines", False):
            log.debug("Outlook First Line Silence links: using native movement after RPC failure")
            return _sendNativeLineGesture(gesture)
        try:
            if scriptHandler.isScriptWaiting():
                return
            if _walkSegment(self, gesture, direction):
                return
        except Exception as error:
            if _isRecoverableWordRpcError(error):
                self._outlookFirstLineSilenceNativeLines = True
                log.warning(
                    "Outlook First Line Silence: Outlook's Word RPC server is unavailable; "
                    "using native line movement for this document"
                )
                log.debug("Outlook Word RPC failure", exc_info=True)
                return _sendNativeLineGesture(gesture)
            log.error("Outlook First Line Silence: could not walk Outlook line segments", exc_info=True)
        return original(self, gesture)
    script_moveByLine.__name__ = getattr(original, "__name__", "script_moveByLine")
    script_moveByLine.__doc__ = getattr(original, "__doc__", None)
    script_moveByLine.__dict__.update(getattr(original, "__dict__", {}))
    return script_moveByLine


# Stories whose link Outlook leaves out.
# Newsletters often put one link around a whole story: its picture, headline, summary
# and "Read more". Word, which shows classic Outlook's messages, cannot link a block like
# that, so the story is plain text, usually after an empty link, and Enter on the
# headline only clicks the text (issue #2: "doAction failed", "Clicking with mouse").
# The message's HTML still has the link, so when Enter or Space lands on text that is
# not a link, find that text in the HTML and open the link the sender put around it.
_OL_FORMAT_HTML = 2
_ASFW_ANY = 0xFFFFFFFF


def _shouldFindStoryLinks(treeInterceptor):
    """Whether this is a classic Outlook message shown by Word, in browse mode."""
    if not isinstance(treeInterceptor, browseMode.BrowseModeDocumentTreeInterceptor):
        return False
    if getattr(treeInterceptor, "passThrough", False):
        return False
    if getattr(treeInterceptor, "VBufHandle", None):
        # A web view keeps the message's own links.
        return False
    root = _documentRoot(treeInterceptor)
    return (
        root is not None
        and _appNameOf(root) == "outlook"
        and _windowClassOf(root) in _OUTLOOK_BODY_WINDOW_CLASSES
    )


def _isInLink(caret):
    character = caret.copy()
    character.expand(textInfos.UNIT_CHARACTER)
    for field in character.getTextWithFields():
        if (
            isinstance(field, textInfos.FieldCommand)
            and field.command == "controlStart"
            and field.field.get("role") == controlTypes.Role.LINK
        ):
            return True
    return False


def _shownOutlookItem(root):
    """The Outlook item shown in the window of *root*, an object in classic Outlook, or None."""
    nativeOm = getattr(root.appModule, "nativeOm", None)
    if not nativeOm:
        return None
    try:
        title = (winUser.getWindowText(_rootWindowOf(root)) or "").strip()
    except Exception:
        title = ""
    inspectors = nativeOm.inspectors
    for index in range(1, inspectors.count + 1):
        inspector = inspectors.item(index)
        if title and (inspector.caption or "").strip() == title:
            return inspector.currentItem
    if _isOutlookMessageInspector(root):
        inspector = nativeOm.activeInspector()
        return inspector.currentItem if inspector else None
    # Not a message window, so the main window: its message list and the reading pane
    # show the one selected message. If this finds another message, the text at the
    # caret will not be found in it, so no link is opened.
    explorer = nativeOm.activeExplorer()
    if not explorer:
        return None
    selection = explorer.selection
    if selection.count != 1:
        return None
    return selection.item(1)


def _storyLinkAt(treeInterceptor, obj=None, info=None):
    """The address of the link the sender put around the text at the caret, or None."""
    if obj is not None or not _shouldFindStoryLinks(treeInterceptor):
        return None
    caret = (info or treeInterceptor.makeTextInfo(textInfos.POSITION_CARET)).copy()
    caret.collapse()
    if _isInLink(caret):
        # NVDA opens real links itself.
        return None
    line = caret.copy()
    line.expand(textInfos.UNIT_LINE)
    rest = caret.copy()
    rest.setEndPoint(line, "endToEnd")
    textFromCaret = rest.text or ""
    if not textFromCaret.strip():
        return None
    item = _shownOutlookItem(_documentRoot(treeInterceptor))
    if item is None or item.bodyFormat != _OL_FORMAT_HTML:
        return None
    before = treeInterceptor.makeTextInfo(textInfos.POSITION_ALL)
    before.setEndPoint(caret, "endToStart")
    return messageHtml.linkAt(item.HTMLBody, before.text or "", textFromCaret)


def _openInBrowser(target):
    try:
        # Outlook, not NVDA, is in front. Let the browser come forward, as it does
        # when Outlook opens a link.
        _user32().AllowSetForegroundWindow(_ASFW_ANY)
    except Exception:
        log.debugWarning("Outlook First Line Silence: could not let the browser come forward", exc_info=True)
    os.startfile(target)


def _openStoryLink(url):
    # Only the site goes in the log: the rest of a newsletter link can identify the reader.
    try:
        site = urlsplit(url).netloc
    except ValueError:
        site = ""
    log.debug("Outlook First Line Silence: opening the link around the text at the caret (%s)", site)
    _openInBrowser(url)


# Reformatting a message on request.
# NVDA+Shift+X shows the message as a plain web page in the default browser: headings,
# paragraphs, lists and links, without layout tables, pictures or the sender's styles,
# with every line of a story a link to it. The page is a file in the add-on's own
# temporary folder, replaced each time and removed when NVDA exits.
# 1.0.25 used NVDA+Shift+V, which the Vision Assistant and Say Product Name and Version
# add-ons also use. NVDA asks global plugins for a script in no fixed order, so with
# either installed the key could go to the other add-on. NVDA 2026.2 itself does not use
# NVDA+Shift+X (NVDA+X repeats the last speech), and the key is only taken in Outlook.
_REFORMAT_GESTURE = "kb:NVDA+shift+x"
_REFORMAT_FOLDER = os.path.join(tempfile.gettempdir(), "outlookFirstLineSilence")


def _removeReformattedMessages():
    try:
        names = os.listdir(_REFORMAT_FOLDER)
    except OSError:
        return
    for name in names:
        if name.endswith(".html"):
            try:
                os.remove(os.path.join(_REFORMAT_FOLDER, name))
            except OSError:
                pass


def _reformatKeyApplies():
    """Whether the reformat key is Outlook's, not another add-on's: focus is in Outlook."""
    try:
        return _isInOutlookWindow(api.getFocusObject())
    except Exception:
        log.debugWarning("Outlook First Line Silence: could not tell whether Outlook has focus", exc_info=True)
        return False


def _outlookItemToReformat():
    """The message being read or selected in classic Outlook, or None."""
    focus = api.getFocusObject()
    if _appNameOf(focus) != "outlook":
        return None
    return _shownOutlookItem(focus)


def _itemText(item, name):
    try:
        return getattr(item, name) or ""
    except (COMError, AttributeError):
        return ""


def _reformatMessage():
    try:
        item = _outlookItemToReformat()
        html = None if item is None else item.HTMLBody
    except Exception:
        log.debugWarning("Outlook First Line Silence: could not get the message to reformat", exc_info=True)
        html = None
    if not html:
        # Translators: Reported when the reformat command is used outside a classic Outlook message.
        ui.message(_("No message to reformat. Open or select a message in classic Outlook."))
        return
    title = _itemText(item, "subject")
    sender = _itemText(item, "senderName")
    page, blocks, links = messageHtml.reformat(
        html,
        # Translators: The title of a reformatted message that has no subject.
        title or _("Message"),
        # Translators: The sender line of a reformatted message.
        _("From: {sender}").format(sender=sender) if sender else None,
    )
    log.debug("Outlook First Line Silence: reformatted a message into %d blocks and %d links" % (blocks, links))
    _removeReformattedMessages()
    path = os.path.join(_REFORMAT_FOLDER, "message-%d.html" % int(time.time() * 1000))
    try:
        os.makedirs(_REFORMAT_FOLDER, exist_ok=True)
        with open(path, "w", encoding="utf-8") as pageFile:
            pageFile.write(page)
        _openInBrowser(path)
    except OSError:
        log.error("Outlook First Line Silence: could not show the reformatted message", exc_info=True)
        # Translators: Reported when the reformatted message could not be opened in the web browser.
        ui.message(_("Could not open the reformatted message in your web browser."))
        return
    # Translators: Reported while the reformatted message opens in the web browser.
    ui.message(_("Opening the reformatted message in your web browser"))


def _makeActivatePositionWrapper(original):
    def _activatePosition(self, obj=None, info=None):
        try:
            url = _storyLinkAt(self, obj=obj, info=info)
        except Exception:
            url = None
            log.debugWarning("Outlook First Line Silence: could not look for a story link", exc_info=True)
        if url and messageHtml.isOpenable(url):
            try:
                _openStoryLink(url)
                return
            except OSError:
                log.error("Outlook First Line Silence: could not open the story link", exc_info=True)
        return original(self, obj=obj, info=info)
    _activatePosition.__name__ = getattr(original, "__name__", "_activatePosition")
    _activatePosition.__doc__ = getattr(original, "__doc__", None)
    return _activatePosition


# Message status in Outlook's message list.
# NVDA's Outlook support reads a message's status (unread, replied or forwarded, has
# attachment, importance) from Outlook's object model. It gets that object once per run
# of Outlook and then keeps it as ``nativeOm`` on its Outlook app module, so if the object
# stops answering, every message is announced without its status until NVDA restarts.
# The log from the first NVDA start after installing 1.0.22 shows exactly that: Outlook
# answered "Unknown name" to NVDA's question about the selected message for the whole
# session. So when focus lands on a message, ask NVDA's question first. If the answer
# means the object itself is unusable, rather than that Outlook is busy or nothing is
# selected, forget it, and NVDA gets Outlook's object model again before it builds the
# announcement.
_UNUSABLE_OBJECT_MODEL_HRESULTS = _RECOVERABLE_WORD_RPC_HRESULTS | frozenset((
    -2147352570,  # 0x80020006: DISP_E_UNKNOWNNAME
    -2147220995,  # 0x800401FD: CO_E_OBJNOTCONNECTED
))
# Getting the object model again starts NVDA's helper process and holds NVDA for about a
# second, so do it at most this often, and stop after this many attempts in a row that
# still leave an unusable object model.
_OBJECT_MODEL_RETRY_INTERVAL = 30.0
_OBJECT_MODEL_RETRY_LIMIT = 3
_OBJECT_MODEL_RETRIES_ATTRIBUTE = "_outlookFirstLineSilenceObjectModelRetries"


def _outlookRowClass(obj):
    """NVDA's class for a row of classic Outlook's message list, when *obj* is one, or None."""
    if _appNameOf(obj) != "outlook":
        return None
    for cls in type(obj).__mro__:
        if cls.__name__ == "UIAGridRow" and cls.__module__.endswith("appModules.outlook"):
            return cls
    return None


def _isOutlookMessageRow(obj):
    """True for a row of classic Outlook's message list, as NVDA's Outlook support builds it."""
    return _outlookRowClass(obj) is not None


def _objectModelError(nativeOm):
    """The error that makes *nativeOm* unusable for NVDA's message status, or None."""
    try:
        nativeOm.activeExplorer().selection.item(1).unread
    except COMError as error:
        # A busy Outlook rejects the call and an empty selection has no item 1; NVDA's
        # own question can still succeed a moment later, so neither counts.
        return error if _hresultOf(error) in _UNUSABLE_OBJECT_MODEL_HRESULTS else None
    except Exception:
        # No explorer window, or an item without a read state, such as a note.
        return None
    return None


def _renewOutlookObjectModel(obj):
    """Before NVDA announces a message, drop an Outlook object model that no longer answers."""
    if not _isOutlookMessageRow(obj):
        return
    appModule = obj.appModule
    nativeOm = vars(appModule).get("nativeOm")
    if nativeOm is None:
        # NVDA has not got it yet, or could not; NVDA handles both itself.
        return
    retries, lastRetry = getattr(appModule, _OBJECT_MODEL_RETRIES_ATTRIBUTE, (0, None))
    error = _objectModelError(nativeOm)
    if error is None:
        if retries:
            setattr(appModule, _OBJECT_MODEL_RETRIES_ATTRIBUTE, (0, lastRetry))
        return
    now = time.monotonic()
    if retries >= _OBJECT_MODEL_RETRY_LIMIT or (
        lastRetry is not None and now - lastRetry < _OBJECT_MODEL_RETRY_INTERVAL
    ):
        return
    setattr(appModule, _OBJECT_MODEL_RETRIES_ATTRIBUTE, (retries + 1, now))
    # nativeOm is a getter that stores what it gets on the app module itself, so removing
    # the stored object makes NVDA's next read get Outlook's object model again.
    vars(appModule).pop("nativeOm", None)
    log.warning(
        "Outlook First Line Silence: Outlook's object model stopped answering (%r), so messages "
        "were announced without their status; NVDA gets it again (attempt %d of %d)"
        % (error, retries + 1, _OBJECT_MODEL_RETRY_LIMIT)
    )


# Deleting a message.
# As Outlook deletes a message, or moves it out of the folder, it empties the message's
# row before the focus moves to the next one, and reports that the row's name changed.
# NVDA names the emptied row only by the status Outlook gives for the selected message,
# and says that new name: "unread", then the next message, "unread From ..." (issue #3).
# The tester's log shows exactly that after each of five presses of Delete. A name with
# no column text says nothing about any message, so it is not said.
_STATUS_WORDS = ("unread", "has attachment", "meeting request")


def _messageStatusLabels(rowClass):
    """The status words NVDA's Outlook support can put in a row's name, as NVDA says them."""
    # NVDA's own translations; this module's _ is the add-on's.
    translate = getattr(builtins, "_", None)
    if not callable(translate):
        translate = str
    labels = {translate(word) for word in _STATUS_WORDS}
    for stateName in ("EXPANDED", "COLLAPSED"):
        state = getattr(controlTypes.State, stateName, None)
        if state is not None:
            labels.add(state.displayString)
    module = sys.modules.get(rowClass.__module__)
    for table in ("executedVerbLabels", "importanceLabels"):
        labels.update(getattr(module, table, {}).values())
    return labels


def _isEmptiedMessageRow(obj):
    """True for the focused message row once Outlook has emptied it to remove the message."""
    rowClass = _outlookRowClass(obj)
    # NVDA says a new name only for the focus.
    if rowClass is None or obj is not api.getFocusObject():
        return False
    name = obj.name
    if not messageRows.isStatusOnly(name, _messageStatusLabels(rowClass)):
        return False
    log.debug("Outlook First Line Silence: the focused message row is now named only %r; not said" % name)
    return True


class OutlookFirstLineSilenceSettingsPanel(settingsDialogs.SettingsPanel):
    title = _("Outlook First Line Silence")

    def makeSettings(self, settingsSizer):
        helper = guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
        self.linksOnOwnLine = helper.addItem(wx.CheckBox(self, label=_("Links are on their &own line")))
        self.linksOnOwnLine.SetValue(getLinksOnOwnLine())
        self.updates = updater.SettingsControls(self, helper)

    def onSave(self):
        setLinksOnOwnLine(self.linksOnOwnLine.GetValue())
        self.updates.save()


def _outlookIsCurrent(obj=None):
    """Use document, foreground and focus, matching 3.6.57's Outlook test."""
    if obj is not None and _isInOutlookWindow(obj):
        return True
    for getter in (api.getForegroundObject, api.getFocusObject):
        try:
            candidate = getter()
        except Exception:
            continue
        if _isInOutlookWindow(candidate):
            return True
    return False


def _filterOutlookEditorNoise(args, kwargs):
    """Mute Browse Mode 3.6.57's Outlook reply-editor alignment filter."""
    try:
        focus = api.getFocusObject()
    except Exception:
        return args, kwargs
    if not _outlookIsCurrent(focus):
        return args, kwargs
    if args:
        sequence = args[0]
    elif "speechSequence" in kwargs:
        sequence = kwargs["speechSequence"]
    else:
        return args, kwargs
    if not isinstance(sequence, list):
        return args, kwargs
    spokenText = " ".join(item for item in sequence if isinstance(item, str))
    if not re.search(r"left\s+(?:ali(?:g)?ned|justified)", spokenText, re.IGNORECASE):
        return args, kwargs
    pattern = re.compile(r"left\s+(?:ali(?:g)?ned|justified)", re.IGNORECASE)
    filtered = []
    for item in sequence:
        if not isinstance(item, str):
            filtered.append(item)
            continue
        cleaned = pattern.sub("", item).strip()
        if cleaned:
            filtered.append(cleaned)
    if args:
        return (filtered,) + tuple(args[1:]), kwargs
    kwargs = dict(kwargs)
    kwargs["speechSequence"] = filtered
    return args, kwargs



def _isOutlookMessageInspector(obj):
    """True for an Outlook inspector whose accessible title identifies a Message window.

    The supplied Outlook 2024 log exposes the foreground object as
    ``<subject> - Message (HTML)`` before Word browse mode is created. This is the
    earliest verified point at which the standalone add-on can distinguish opening a
    message from ordinary Outlook navigation.
    """
    if obj is None or not _isInOutlookWindow(obj):
        return False
    candidates = []
    try:
        candidates.append((getattr(obj, "name", "") or "").strip())
    except Exception:
        pass
    root = _rootWindowOf(obj)
    if root:
        try:
            candidates.append((winUser.getWindowText(root) or "").strip())
        except Exception:
            pass
    return any(" - Message (" in name for name in candidates if name)


def _armMessageOpening(obj):
    global _messageOpeningUntil
    if _isOutlookMessageInspector(obj):
        _messageOpeningUntil = time.monotonic() + 2.0
        log.debug("Outlook First Line Silence: armed message-inspector entry suppression")


def _messageOpeningNow():
    return time.monotonic() < _messageOpeningUntil


def _isBareDialogSpeech(args):
    """True only for NVDA's standalone automatic ``dialog`` utterance."""
    if not args:
        return False
    try:
        words = [item.strip().casefold() for item in args[0] if isinstance(item, str) and item.strip()]
    except Exception:
        return False
    return words == ["dialog"]


_DIALOG_ROLES = _members(controlTypes.Role, ("DIALOG", "ALERT", "PROPERTYPAGE", "OPTIONPANE"))
_NOT_A_DIALOG_ROLES = _members(controlTypes.Role, ("APPLICATION", "FRAME", "DESKTOP"))
_DIALOG_WALK_LIMIT = 6


def _looksLikeADialog(obj):
    if getattr(obj, "role", None) in _DIALOG_ROLES:
        return True
    windowClass = _windowClassOf(obj)
    return windowClass == "#32770" or windowClass.lower().startswith("bosa_sdm")


def _dialogAbove(obj):
    current = obj
    for _step in range(_DIALOG_WALK_LIMIT):
        try:
            current = current.parent
        except Exception:
            return None
        if current is None:
            return None
        if _looksLikeADialog(current):
            return current
        if getattr(current, "role", None) in _NOT_A_DIALOG_ROLES:
            return None
    return None


def _boundedDescendants(root, limit=30, depthLimit=4):
    """Mute Browse Mode's bounded breadth-first dialog scan."""
    try:
        queue = [(child, 1) for child in (root.children or ())]
    except Exception:
        return
    seen = set()
    while queue and len(seen) < limit:
        item, depth = queue.pop(0)
        identity = id(item)
        if identity in seen:
            continue
        seen.add(identity)
        yield item
        if depth >= depthLimit:
            continue
        try:
            queue.extend((child, depth + 1) for child in (item.children or ()))
        except Exception:
            pass


def _isOutlookDraftQuestion(text):
    return bool(
        (re.search(r"saved\s+a\s+draft", text, re.IGNORECASE)
         and re.search(r"want\s+to\s+keep", text, re.IGNORECASE))
        or re.search(r"(?:do\s+you\s+)?want\s+to\s+save\s+(?:your\s+)?changes", text, re.IGNORECASE)
    )


def _outlookDraftPromptText(obj):
    """The exact 3.6.57 scan, including NetUI dialogs exposed as panes."""
    dialog = _dialogAbove(obj)
    if dialog is None or "netui" not in _windowClassOf(dialog).lower():
        current = obj
        dialog = None
        for _step in range(8):
            try:
                current = current.parent
            except Exception:
                break
            if current is None:
                break
            if "netui" in _windowClassOf(current).lower():
                dialog = current
                break
    if dialog is None:
        return None, ""
    fragments = []
    seen = set()
    for candidate in _boundedDescendants(dialog):
        try:
            if candidate.role == controlTypes.Role.BUTTON:
                continue
        except Exception:
            pass
        candidateTexts = []
        for attribute in ("name", "value", "description"):
            try:
                text = " ".join((getattr(candidate, attribute, "") or "").split())
            except Exception:
                text = ""
            if text and text.casefold() not in {item.casefold() for item in candidateTexts}:
                candidateTexts.append(text)
        candidateText = " ".join(candidateTexts)
        if _isOutlookDraftQuestion(candidateText):
            return dialog, candidateText
        for text in candidateTexts:
            folded = text.casefold()
            if folded not in seen:
                seen.add(folded)
                fragments.append(text)
    combined = " ".join(fragments)
    if _isOutlookDraftQuestion(combined):
        return dialog, combined
    return None, ""


@contextmanager
def _muteSpeech():
    """Let NVDA handle an event and note what it would say, without saying it."""
    speechModule = getattr(speech, "speech", None)
    public = getattr(speech, "speak", None)
    private = getattr(speechModule, "speak", None) if speechModule else None
    def quiet(*args, **kwargs):
        return None
    try:
        if public is not None:
            speech.speak = quiet
        if speechModule is not None and private is not None:
            speechModule.speak = quiet
        yield
    finally:
        if public is not None and getattr(speech, "speak", None) is quiet:
            speech.speak = public
        if speechModule is not None and private is not None and getattr(speechModule, "speak", None) is quiet:
            speechModule.speak = private


def _handleOutlookDraftPrompt(obj, nextHandler):
    """Mute Browse Mode 3.6.57's one-prompt-per-dialog focus handling."""
    global _savePromptKey, _savePromptUntil
    if not _isInOutlookWindow(obj):
        # The scan walks up to 14 parents, so keep it out of every other program's
        # focus changes, such as Notepad's Save As dialog (issue #3).
        return False
    dialog, text = _outlookDraftPromptText(obj)
    if dialog is None:
        return False
    try:
        rootKey = _rootWindowOf(dialog) or id(dialog)
        buttonName = " ".join((obj.name or "").split())
    except Exception:
        return False
    if not buttonName:
        return False
    key = (rootKey, text.casefold())
    now = time.monotonic()
    if key == _savePromptKey and now < _savePromptUntil:
        nextHandler()
        return True
    _savePromptKey = key
    _savePromptUntil = now + _DRAFT_PROMPT_DUPLICATE_WINDOW
    with _muteSpeech():
        nextHandler()
    with _ownSpeech():
        speech.speak([text, _("{button} button").format(button=buttonName)])
    return True


def _shouldDropOpeningObjectSpeech(args, kwargs):
    """3.6.57 speakObject filtering, limited to an opening Outlook message."""
    if not _messageOpeningNow():
        return False
    obj = kwargs.get("obj", args[0] if args else None)
    if obj is None or not _isInOutlookWindow(obj):
        return False
    if "reason" in kwargs:
        reason = kwargs["reason"]
    elif len(args) > 1:
        reason = args[1]
    else:
        return False
    if reason not in _AUTOMATIC_REASONS:
        return False
    try:
        # Mute Browse Mode replaces Outlook dialog speech with its own detailed
        # announcements. This standalone add-on intentionally does not implement
        # that replacement, so muting a dialog here would make Save/Save As and
        # other Outlook dialogs incomplete or silent. Restrict the early filter to
        # the message inspector's title/document sequence instead.
        if obj.role == getattr(controlTypes.Role, "DIALOG", None):
            return False
        return obj.role in _TITLE_ROLES
    except Exception:
        return False


def _makeSpeakObjectWrapper(original):
    def speakObject(*args, **kwargs):
        if _shouldDropOpeningObjectSpeech(args, kwargs):
            obj = kwargs.get("obj", args[0] if args else None)
            try:
                name = getattr(obj, "name", "") or ""
                role = getattr(obj, "role", None)
            except Exception:
                name, role = "", None
            log.debug(
                "Outlook First Line Silence: suppressed opening object speech role=%r name=%r",
                role,
                name,
            )
            return
        return original(*args, **kwargs)

    speakObject.__name__ = getattr(original, "__name__", "speakObject")
    speakObject.__doc__ = getattr(original, "__doc__", None)
    return speakObject


def _documentRoot(treeInterceptor):
    try:
        return treeInterceptor.rootNVDAObject
    except Exception:
        return None


def _isEnteringDocument(treeInterceptor, args, kwargs):
    """The exact first-line condition documented by muteBrowseMode 3.6.57."""
    try:
        return bool(getattr(treeInterceptor, "_enteringFromOutside", False))
    except Exception:
        return False


def _shouldSilenceGainFocus(treeInterceptor, args, kwargs):
    return _isEnteringDocument(treeInterceptor, args, kwargs) and _outlookIsCurrent(
        _documentRoot(treeInterceptor)
    )


def _shouldSilenceTreeInterceptorGainFocus(treeInterceptor, args, kwargs):
    return _outlookIsCurrent(_documentRoot(treeInterceptor))


def _enterDocumentCall():
    global _inCallDepth, _inCallUntil
    _inCallDepth += 1
    _inCallUntil = max(_inCallUntil, time.monotonic() + _IN_CALL_GATE)


def _exitDocumentCall():
    global _inCallDepth, _inCallUntil
    _inCallDepth -= 1
    if _inCallDepth <= 0:
        _inCallDepth = 0
        _inCallUntil = 0.0


def _openGate(seconds):
    global _gateUntil
    _gateUntil = time.monotonic() + seconds


def _closeGate():
    global _gateUntil
    _gateUntil = 0.0


def _speechIsGated():
    now = time.monotonic()
    inCall = _inCallDepth > 0 and now < _inCallUntil
    trailing = _gateUntil > 0.0 and now < _gateUntil
    return inCall or trailing


def _onGesture(*args, **kwargs):
    """As in 3.6.57, user input releases delayed gates and progress protection."""
    global _outlookProgressSpeechUntil, _outlookProgressLastPercent, _outlookProgressLastEventAt
    try:
        _noteSuggestionGesture(kwargs["gesture"] if "gesture" in kwargs else (args[0] if args else None))
    except Exception:
        log.debugWarning("Outlook First Line Silence: could not note suggestion key", exc_info=True)
    _outlookProgressSpeechUntil = 0.0
    _outlookProgressLastPercent = None
    _outlookProgressLastEventAt = 0.0
    if _gateUntil > 0.0:
        _closeGate()
    return True


def _patch(owner, name, replacement):
    original = getattr(owner, name)
    wasOwn = name in getattr(owner, "__dict__", {})
    setattr(owner, name, replacement)
    _patches.append((owner, name, original, wasOwn, replacement))
    return original


def _unpatchAll():
    while _patches:
        owner, name, original, wasOwn, replacement = _patches.pop()
        try:
            if getattr(owner, name, None) is not replacement:
                continue
            if wasOwn:
                setattr(owner, name, original)
            else:
                delattr(owner, name)
        except Exception:
            log.error(
                "Outlook First Line Silence: could not restore %s.%s" % (owner, name),
                exc_info=True,
            )


def _makeSpeakWrapper(original):
    def speakWrapper(*args, **kwargs):
        if _bypassDepth > 0:
            return original(*args, **kwargs)
        # NVDA emits this one-word container announcement before the browse-mode
        # focus hooks begin. It is part of the message-opening noise, not a real
        # Save/other dialog report, so drop it only while the inspector is armed.
        if _messageOpeningNow() and _isBareDialogSpeech(args):
            log.debug("Outlook First Line Silence: suppressed opening dialog speech")
            return
        if _speechIsGated():
            log.debug("Outlook First Line Silence: suppressed document-entry speech")
            return
        args, kwargs = _filterOutlookEditorNoise(args, kwargs)
        return original(*args, **kwargs)

    speakWrapper.__name__ = getattr(original, "__name__", "speak")
    speakWrapper.__doc__ = getattr(original, "__doc__", None)
    return speakWrapper


def _hookDocumentEvent(owner, name, onlyIf):
    original = getattr(owner, name)

    def wrapper(self, *args, **kwargs):
        silence = False
        try:
            silence = bool(onlyIf(self, args, kwargs))
        except Exception:
            log.debugWarning(
                "Outlook First Line Silence: condition failed for %s" % name,
                exc_info=True,
            )
        if silence:
            log.debug("Outlook First Line Silence: silencing %s" % name)
            _enterDocumentCall()
        try:
            return original(self, *args, **kwargs)
        finally:
            if silence:
                _exitDocumentCall()
                _openGate(_TRAILING_GATE)

    wrapper.__name__ = name
    wrapper.__doc__ = getattr(original, "__doc__", None)
    _patch(owner, name, wrapper)


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    def __init__(self):
        global _gestureHandlerRegistered
        super().__init__()
        self._lastBodyWindow = None
        try:
            # Same synthesizer funnel used by muteBrowseMode 3.6.57.
            wrappedSpeak = _makeSpeakWrapper(speech.speech.speak)
            _patch(speech.speech, "speak", wrappedSpeak)
            _patch(speech, "speak", wrappedSpeak)

            # 3.6.57 also filters automatic Outlook title/dialog/document objects via
            # speakObject, before browse mode exists. Limit that layer to the verified
            # message-inspector opening window so ordinary Outlook navigation remains.
            wrappedSpeakObject = _makeSpeakObjectWrapper(speech.speech.speakObject)
            _patch(speech.speech, "speakObject", wrappedSpeakObject)
            _patch(speech, "speakObject", wrappedSpeakObject)

            # Exact 3.6.57 Send/Receive behavior: queued Outlook progress always bypasses
            # document silencing, and automatic focus churn cannot cancel the queue.
            wrappedCancelSpeech = _makeCancelSpeechWrapper(speech.speech.cancelSpeech)
            _patch(speech.speech, "cancelSpeech", wrappedCancelSpeech)
            _patch(speech, "cancelSpeech", wrappedCancelSpeech)
            from NVDAObjects import behaviors
            _patch(
                behaviors.ProgressBar,
                "event_valueChange",
                _makeProgressBarWrapper(behaviors.ProgressBar.event_valueChange),
            )

            # 3.6.57 uses BOTH hooks. event_gainFocus silences the line landed on
            # when entering from outside. event_treeInterceptor_gainFocus silences
            # the initial document name, "document", and first line.
            owner = browseMode.BrowseModeDocumentTreeInterceptor
            _hookDocumentEvent(owner, "event_gainFocus", _shouldSilenceGainFocus)
            _hookDocumentEvent(
                owner,
                "event_treeInterceptor_gainFocus",
                _shouldSilenceTreeInterceptorGainFocus,
            )

            # Keep the Mute Browse Mode "links on their own line" option scoped to
            # Outlook. Virtual buffers use NVDA's native line calculation; Word-based
            # message views use the same bounded arrow-key segment walker as Mute.
            _patch(
                virtualBuffers.VirtualBufferTextInfo,
                "_getLineOffsets",
                _makeLineOffsetsWrapper(virtualBuffers.VirtualBufferTextInfo._getLineOffsets),
            )
            for name, direction in (("script_moveByLine_forward", 1), ("script_moveByLine_back", -1)):
                manager = cursorManager.CursorManager
                _patch(manager, name, _makeMoveByLineWrapper(getattr(manager, name), direction))

            # Enter and Space on a story Outlook shows without its link open the link.
            _patch(
                owner,
                "_activatePosition",
                _makeActivatePositionWrapper(owner._activatePosition),
            )

            if OutlookFirstLineSilenceSettingsPanel not in settingsDialogs.NVDASettingsDialog.categoryClasses:
                settingsDialogs.NVDASettingsDialog.categoryClasses.append(OutlookFirstLineSilenceSettingsPanel)

            inputCore.decide_executeGesture.register(_onGesture)
            _gestureHandlerRegistered = True
            updater.start()
            log.info(
                "Outlook First Line Silence 1.0.27 loaded; links on their own line: %s"
                % getLinksOnOwnLine()
            )
        except Exception:
            if _gestureHandlerRegistered:
                try:
                    inputCore.decide_executeGesture.unregister(_onGesture)
                except Exception:
                    pass
                _gestureHandlerRegistered = False
            _unpatchAll()
            raise

    @scriptHandler.script(
        # Translators: Description of a command, shown in the Input Gestures dialog.
        description=_("Checks for Outlook First Line Silence updates"),
        category=_("Outlook First Line Silence"),
    )
    def script_checkForUpdates(self, gesture):
        updater.checkForUpdates()

    @scriptHandler.script(
        # Translators: Description of a command, shown in the Input Gestures dialog.
        description=_(
            "Reformats the Outlook message you are reading or have selected: shows it in your "
            "web browser as a plain page, where each story is a link"
        ),
        category=_("Outlook First Line Silence"),
        gesture=_REFORMAT_GESTURE,
    )
    def script_reformatMessage(self, gesture):
        _reformatMessage()

    def getScript(self, gesture):
        script = super().getScript(gesture)
        if getattr(script, "__func__", None) is GlobalPlugin.script_reformatMessage and not _reformatKeyApplies():
            # Outside Outlook the key is left to NVDA and other add-ons.
            return None
        return script

    def event_foreground(self, obj, nextHandler):
        # The message-window title is spoken inside the foreground event chain, before
        # Word browse mode and its focus hooks exist. Arm the 3.6.57 speakObject filter
        # before allowing NVDA to continue that chain.
        try:
            _armMessageOpening(obj)
        except Exception:
            log.debugWarning(
                "Outlook First Line Silence: could not inspect foreground message window",
                exc_info=True,
            )
        nextHandler()

    def event_stateChange(self, obj, nextHandler, *args, **kwargs):
        # NVDA speaks recipient suggestions from stateChange; play the enter sound first.
        _trackSuggestionSelection(obj)
        nextHandler()

    def event_selection(self, obj, nextHandler, *args, **kwargs):
        _trackSuggestionSelection(obj)
        nextHandler()

    def event_UIA_elementSelected(self, obj, nextHandler, *args, **kwargs):
        _trackSuggestionSelection(obj)
        nextHandler()

    def event_nameChange(self, obj, nextHandler):
        try:
            emptied = _isEmptiedMessageRow(obj)
        except Exception:
            emptied = False
            log.debugWarning("Outlook First Line Silence: could not check a message row's new name", exc_info=True)
        if not emptied:
            nextHandler()
            return
        # Braille and NVDA's notes of the row still follow the change.
        with _muteSpeech():
            nextHandler()

    def event_gainFocus(self, obj, nextHandler):
        """Restore the complete text of Outlook's Save/Keep-draft prompt."""
        # Returning to the message body after an Outlook modal dialog can bypass
        # BrowseModeDocumentTreeInterceptor's focus hooks. Gate that exact message
        # document focus so NVDA does not announce "document, page 1, section 1"
        # (or the first line) on return.
        try:
            _suggestionsFocusChanged(obj)
        except Exception:
            log.debugWarning("Outlook First Line Silence: could not update suggestion state", exc_info=True)
        # Global plugins handle focus before NVDA's Outlook support announces the message,
        # so a renewed object model is already in place for that announcement.
        try:
            _renewOutlookObjectModel(obj)
        except Exception:
            log.debugWarning("Outlook First Line Silence: could not check Outlook's object model", exc_info=True)
        try:
            isMessageDocument = (
                getattr(obj, "role", None) == getattr(controlTypes.Role, "DOCUMENT", None)
                and _isOutlookMessageInspector(obj)
            )
        except Exception:
            isMessageDocument = False
        if isMessageDocument:
            _enterDocumentCall()
            try:
                nextHandler()
            finally:
                _exitDocumentCall()
                _openGate(_TRAILING_GATE)
            self._reportMessageBody(obj)
            return
        if _handleOutlookDraftPrompt(obj, nextHandler):
            return
        nextHandler()
        self._reportMessageBody(obj)

    def _reportMessageBody(self, obj):
        """Announce entry into an editable Outlook message body once per focus entry."""
        try:
            if not _isOutlookMessageBody(obj):
                self._lastBodyWindow = None
                return
            window = getattr(obj, "windowHandle", None)
            if window is not None and window == self._lastBodyWindow:
                return
            self._lastBodyWindow = window
            _announceMessageBody()
        except Exception:
            log.error(
                "Outlook First Line Silence: could not announce the message body",
                exc_info=True,
            )

    def terminate(self):
        updater.stop()
        _removeReformattedMessages()
        _resetSuggestions()
        global _gestureHandlerRegistered
        _closeGate()
        if _gestureHandlerRegistered:
            try:
                inputCore.decide_executeGesture.unregister(_onGesture)
            except Exception:
                log.debugWarning("Outlook First Line Silence: could not unhook input gestures", exc_info=True)
            _gestureHandlerRegistered = False
        _unpatchAll()
        try:
            settingsDialogs.NVDASettingsDialog.categoryClasses.remove(OutlookFirstLineSilenceSettingsPanel)
        except ValueError:
            pass
        super().terminate()

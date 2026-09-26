"""Tests for coming back to an open message with Alt+Tab (issue #5, 1.0.28).

python -m unittest discover -s tests

With 1.0.27, Alt+Tab back into an open message said only the start of its title. The
Alt+Tab switcher began saying it, then NVDA's own event_foreground cancelled speech as
the message came to the front, and the add-on dropped the window's title as it does
when a message opens. JAWS says the whole title. The add-on now remembers which
message windows are open: a message that opens stays silent, and one you come back
to says its title, but not "dialog", "document" or the first line.

The plugin is run here through NVDA 2026.2's own event code, copied word for word
below from eventHandler.py and NVDAObjects/__init__.py. Speech is a list: what comes
after the last cancel is what is heard in full.
"""

import builtins
import enum
import importlib.util
import os
import sys
import types
import unittest
from unittest import mock

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PLUGIN_DIR = os.path.join(ROOT, "addon", "globalPlugins", "outlookFirstLineSilence")

# NVDA 2026.2 code, word for word.
EVENT_EXECUTER = r'''
class _EventExecuter(garbageHandler.TrackedObject):
	"""Facilitates execution of a chain of event functions.
	L{gen} generates the event functions and positional arguments.
	L{next} calls the next function in the chain.
	"""

	def __init__(self, eventName, obj, kwargs):
		self.kwargs = kwargs
		self._gen = self.gen(eventName, obj)
		try:
			self.next()
		except StopIteration:
			pass
		finally:
			del self._gen

	def next(self):
		func, args = next(self._gen)
		try:
			return func(*args, **self.kwargs)
		except TypeError:
			log.warning(
				"Could not execute function {func} defined in {module} module; kwargs: {kwargs}".format(
					func=func.__name__,
					module=func.__module__ or "unknown",
					kwargs=self.kwargs,
				),
				exc_info=True,
			)
			return extensionPoints.callWithSupportedKwargs(func, *args, **self.kwargs)

	def gen(self, eventName, obj):
		funcName = "event_%s" % eventName

		# Global plugin level.
		for plugin in globalPluginHandler.runningPlugins:
			func = getattr(plugin, funcName, None)
			if func:
				yield func, (obj, self.next)

		# App module level.
		app = obj.appModule
		if app:
			func = getattr(app, funcName, None)
			if func:
				yield func, (obj, self.next)

		# Tree interceptor level.
		treeInterceptor = obj.treeInterceptor
		if treeInterceptor:
			func = getattr(treeInterceptor, funcName, None)
			if func and (getattr(func, "ignoreIsReady", False) or treeInterceptor.isReady):
				yield func, (obj, self.next)

		# NVDAObject level.
		func = getattr(obj, funcName, None)
		if func:
			yield func, ()'''

EXECUTE_EVENT = r'''
def executeEvent(
	eventName: str,
	obj: "NVDAObjects.NVDAObject",
	**kwargs,
) -> None:
	"""Executes an NVDA event.
	@param eventName: the name of the event type (e.g. 'gainFocus', 'nameChange')
	@param obj: the object the event is for
	@param kwargs: Additional event parameters as keyword arguments.
	"""
	if objectBelowLockScreenAndWindowsIsLocked(
		obj,
		shouldLog=config.conf["debugLog"]["events"],
	):
		return
	try:
		global _virtualDesktopName
		isGainFocus = eventName == "gainFocus"
		# Allow NVDAObjects to redirect focus events to another object of their choosing.
		if isGainFocus and obj.focusRedirect:
			obj = obj.focusRedirect
		sleepMode = obj.sleepMode
		# Handle possible virtual desktop name change event.
		# More effective in Windows 10 Version 1903 and later.
		from NVDAObjects.window import Window

		if (
			eventName == "nameChange"
			and isinstance(obj, Window)
			and obj.windowClassName == "#32769"
			and _canAnnounceVirtualDesktopNames
		):
			import core

			_virtualDesktopName = obj.name
			core.callLater(250, handlePossibleDesktopNameChange)
		if isGainFocus and not doPreGainFocus(obj, sleepMode=sleepMode):
			return
		elif not sleepMode and eventName == "documentLoadComplete" and not doPreDocumentLoadComplete(obj):
			return
		elif not sleepMode:
			_EventExecuter(eventName, obj, kwargs)
	except Exception:
		log.exception(f"error executing event: {eventName} on {obj} with extra args of {kwargs}")'''

DO_PRE_GAIN_FOCUS = r'''
def doPreGainFocus(obj: "NVDAObjects.NVDAObject", sleepMode: bool = False) -> bool:
	if objectBelowLockScreenAndWindowsIsLocked(
		obj,
		shouldLog=config.conf["debugLog"]["events"],
	):
		return False
	oldFocus = api.getFocusObject()
	oldTreeInterceptor = oldFocus.treeInterceptor if oldFocus else None
	if not api.setFocusObject(obj):
		return False
	if speech.manager._shouldCancelExpiredFocusEvents():
		log._speechManagerDebug("executeEvent: Removing cancelled speech commands.")
		# ask speechManager to check if any of it's queued utterances should be cancelled
		# Note: Removing cancelled speech commands should happen after all dependencies for the isValid check
		# have been updated:
		# - obj.WAS_GAIN_FOCUS_OBJ_ATTR_NAME
		# - api.setFocusObject()
		# - api.getFocusAncestors()
		# When these are updated:
		# - obj.WAS_GAIN_FOCUS_OBJ_ATTR_NAME
		#   - Set during creation of the _CancellableSpeechCommand.
		# - api.getFocusAncestors() via api.setFocusObject() called in doPreGainFocus
		speech._manager.removeCancelledSpeechCommands()

	if api.getFocusDifferenceLevel() <= 1:
		newForeground = api.getDesktopObject().objectInForeground()
		if not newForeground:
			log.debugWarning("Can not get real foreground, resorting to focus ancestors")
			ancestors = api.getFocusAncestors()
			if len(ancestors) > 1:
				newForeground = ancestors[1]
			else:
				newForeground = obj
		if not api.setForegroundObject(newForeground):
			return False
		executeEvent("foreground", newForeground)
	handlePossibleDesktopNameChange()
	if sleepMode:
		return True
	# Fire focus entered events for all new ancestors of the focus if this is a gainFocus event
	for parent in api.getFocusAncestors()[api.getFocusDifferenceLevel() :]:
		executeEvent("focusEntered", parent)
	if obj.treeInterceptor is not oldTreeInterceptor:
		if hasattr(oldTreeInterceptor, "event_treeInterceptor_loseFocus"):
			oldTreeInterceptor.event_treeInterceptor_loseFocus()
		if (
			obj.treeInterceptor
			and obj.treeInterceptor.isReady
			and hasattr(obj.treeInterceptor, "event_treeInterceptor_gainFocus")
		):
			obj.treeInterceptor.event_treeInterceptor_gainFocus()
	return True'''

# NVDAObject methods, word for word.
REPORT_FOCUS = r'''
def reportFocus(self):
	"""Announces this object in a way suitable such that it gained focus."""
	speech.speakObject(self, reason=controlTypes.OutputReason.FOCUS)'''

EVENT_FOCUS_ENTERED = r'''
def event_focusEntered(self):
	if self.role in (controlTypes.Role.MENUBAR, controlTypes.Role.POPUPMENU, controlTypes.Role.MENUITEM):
		speech.cancelSpeech()
		return
	if self.isPresentableFocusAncestor:
		speech.speakObject(self, reason=controlTypes.OutputReason.FOCUSENTERED)'''

EVENT_GAIN_FOCUS = r'''
def event_gainFocus(self):
	"""
	This code is executed if a gain focus event is received by this object.
	"""
	self.reportFocus()
	braille.handler.handleGainFocus(self)
	brailleInput.handler.handleGainFocus(self)
	vision.handler.handleGainFocus(self)'''

EVENT_FOREGROUND = r'''
def event_foreground(self):
	"""Called when the foreground window changes.
	This method should only perform tasks specific to the foreground window changing.
	L{event_focusEntered} or L{event_gainFocus} will be called for this object, so this method should not speak/braille the object, etc.
	"""
	speech.cancelSpeech()
	vision.handler.handleForeground(self)'''

# The subject in the issue #5 log. NVDA names the message window with a trailing space.
TITLE = (
	"Re: [joshknnd1982/jawsMigrator] When pressing end to go to new outlook message it says "
	"the wrong one is highlighted (Issue #17) - Message (HTML)"
)
OTHER_TITLE = "Re: [joshknnd1982/jawsMigrator] This players name isn't being spoken correctly (Issue #13) - Message (HTML)"
NOTEPAD_TITLE = "*Re [joshknnd1982jawsMigrator] When - Notepad"
CANCEL = "<cancel>"

Role = enum.Enum(
	"Role",
	"APPLICATION BUTTON DESKTOP DIALOG DOCUMENT EDITABLETEXT FRAME INTERNALFRAME LISTITEM MENUBAR "
	"MENUITEM OPTIONPANE PANE POPUPMENU PROPERTYPAGE ALERT TABLEROW WINDOW",
)
OutputReason = enum.Enum("OutputReason", "FOCUS FOCUSENTERED CHANGE CARET QUERY")
State = enum.Enum("State", "READONLY UNAVAILABLE MULTILINE SELECTED INVISIBLE OFFSCREEN")
# What the stand-in for NVDA's speakObject says for a role, besides the name.
ROLE_WORDS = {Role.DIALOG: "dialog", Role.DOCUMENT: "document", Role.EDITABLETEXT: "edit"}


class _StubModule(types.ModuleType):
	"""An NVDA module whose every attribute is a mock, unless the test sets it."""

	def __getattr__(self, name):
		if name.startswith("__"):
			raise AttributeError(name)
		value = mock.MagicMock(name="%s.%s" % (self.__name__, name))
		setattr(self, name, value)
		return value


class _Conf(dict):
	spec = {}


_savedModules = {}
_hadTranslation = hasattr(builtins, "_")
plugin = None


def _stub(name, **attributes):
	module = _StubModule(name)
	for key, value in attributes.items():
		setattr(module, key, value)
	_savedModules.setdefault(name, sys.modules.get(name))
	sys.modules[name] = module
	if "." in name:
		parent, child = name.rsplit(".", 1)
		setattr(sys.modules[parent], child, module)
	return module


def setUpModule():
	global plugin
	for name in (
		"addonHandler", "api", "appModuleHandler", "browseMode", "core", "cursorManager",
		"globalVars", "inputCore", "nvwave", "scriptHandler", "speech",
		"speech.speech", "speech.priorities", "textInfos", "ui", "virtualBuffers", "winUser",
		"logHandler", "gui", "gui.guiHelper", "NVDAObjects",
	):
		_stub(name)
	_stub("globalPluginHandler", GlobalPlugin=type("GlobalPlugin", (), {"terminate": lambda self: None}))
	_stub("controlTypes", Role=Role, OutputReason=OutputReason, State=State)
	_stub("config", conf=_Conf(debugLog={"events": False}))
	_stub("NVDAObjects.window", Window=type("Window", (), {}))
	_stub("gui.settingsDialogs", SettingsPanel=type("SettingsPanel", (), {}))
	_stub("wx", Dialog=type("Dialog", (), {}))
	if not _hadTranslation:
		builtins._ = lambda text: text
	spec = importlib.util.spec_from_file_location(
		"outlookFirstLineSilence",
		os.path.join(PLUGIN_DIR, "__init__.py"),
		submodule_search_locations=[PLUGIN_DIR],
	)
	plugin = importlib.util.module_from_spec(spec)
	_savedModules.setdefault("outlookFirstLineSilence", sys.modules.get("outlookFirstLineSilence"))
	sys.modules["outlookFirstLineSilence"] = plugin
	spec.loader.exec_module(plugin)


def tearDownModule():
	for name in list(sys.modules):
		if name.startswith("outlookFirstLineSilence.") and name not in _savedModules:
			del sys.modules[name]
	for name, module in _savedModules.items():
		if module is None:
			sys.modules.pop(name, None)
		else:
			sys.modules[name] = module
	if not _hadTranslation:
		del builtins._


def _nvdaObjectClass(namespace):
	methods = {}
	for code in (REPORT_FOCUS, EVENT_FOCUS_ENTERED, EVENT_GAIN_FOCUS, EVENT_FOREGROUND):
		local = {}
		exec(code.strip(), namespace, local)
		methods.update(local)

	def __init__(self, role, name="", parent=None, appName="outlook", window=0, presentable=True, states=()):
		self.role = role
		self.name = name
		self.parent = parent
		self.appModule = types.SimpleNamespace(appName=appName)
		self.windowHandle = window
		self.isPresentableFocusAncestor = presentable
		self.states = set(states)
		self.treeInterceptor = None
		self.focusRedirect = None
		self.sleepMode = False

	def __repr__(self):
		return "<%s %r>" % (self.role.name, self.name)

	methods["__init__"] = __init__
	methods["__repr__"] = __repr__
	return type("NVDAObject", (), methods)


class _Api:
	"""NVDA's api focus bookkeeping: focus, its ancestors and the focus difference level."""

	def __init__(self, desktop):
		self.desktop = desktop
		self.focus = desktop
		self.ancestors = []
		self.differenceLevel = 0
		self.foreground = desktop
		self.realForeground = desktop

	def setFocusObject(self, obj):
		oldLine = self.ancestors + [self.focus]
		newLine = []
		parent = obj.parent
		while parent is not None:
			newLine.insert(0, parent)
			parent = parent.parent
		level = 0
		while level < min(len(oldLine), len(newLine)) and oldLine[level] is newLine[level]:
			level += 1
		self.focus, self.ancestors, self.differenceLevel = obj, newLine, level
		return True

	def setForegroundObject(self, obj):
		self.foreground = obj
		return True


class _Log:
	"""NVDA's log, raising what NVDA's event code would log and carry on from."""

	def exception(self, message, *args, **kwargs):
		raise AssertionError(message)

	def warning(self, message, *args, **kwargs):
		raise AssertionError(message)

	def debugWarning(self, message, *args, **kwargs):
		pass

	def _speechManagerDebug(self, message, *args, **kwargs):
		pass


class _Clock:
	def __init__(self):
		self.now = 1000.0

	def monotonic(self):
		return self.now


class _User32:
	def __init__(self, visible):
		self.visible = visible

	def IsWindowVisible(self, window):
		return int(window in self.visible)


class _MessageWindow:
	def __init__(self, NVDAObject, desktop, window, title):
		self.window = window
		self.title = title
		self.pane = NVDAObject(Role.PANE, title + " ", desktop, window=window)
		# NVDA says a bare "dialog" for this, as in the tester's older log.
		self.dialog = NVDAObject(Role.DIALOG, "", self.pane, window=window)
		self.document = NVDAObject(Role.DOCUMENT, "", self.dialog, window=window, states=(State.READONLY,))


class MessageReturnTests(unittest.TestCase):
	def setUp(self):
		self.spoken = []
		self.clock = _Clock()
		self.visible = set()
		self.classes = {}
		self.titles = {}

		nvda = {
			"typing": __import__("typing"),
			"Optional": __import__("typing").Optional,
			"controlTypes": sys.modules["controlTypes"],
			"config": sys.modules["config"],
			"log": _Log(),
			"braille": mock.MagicMock(name="braille"),
			"brailleInput": mock.MagicMock(name="brailleInput"),
			"vision": mock.MagicMock(name="vision"),
			"extensionPoints": mock.MagicMock(name="extensionPoints"),
			"garbageHandler": types.SimpleNamespace(TrackedObject=object),
			"objectBelowLockScreenAndWindowsIsLocked": lambda obj, shouldLog=False: False,
			"handlePossibleDesktopNameChange": lambda: None,
			"_canAnnounceVirtualDesktopNames": False,
			"_virtualDesktopName": None,
		}
		self.NVDAObject = _nvdaObjectClass(nvda)
		self.desktop = self.NVDAObject(Role.DESKTOP, "Desktop", appName="explorer", presentable=False)
		self.api = _Api(self.desktop)
		self.desktop.objectInForeground = lambda: self.api.realForeground

		api = sys.modules["api"]
		api.getFocusObject = lambda: self.api.focus
		api.getForegroundObject = lambda: self.api.foreground
		api.getFocusAncestors = lambda: self.api.ancestors
		api.getFocusDifferenceLevel = lambda: self.api.differenceLevel
		api.getDesktopObject = lambda: self.desktop
		api.setFocusObject = self.api.setFocusObject
		api.setForegroundObject = self.api.setForegroundObject

		winUser = sys.modules["winUser"]
		winUser.getAncestor = lambda window, flags: window
		winUser.getClassName = lambda window: self.classes.get(window, "")
		winUser.getWindowText = lambda window: self.titles.get(window, "")
		winUser.getWindowThreadProcessID = lambda window: (0, 0)

		speechPackage = sys.modules["speech"]
		speechModule = sys.modules["speech.speech"]
		speak = plugin._makeSpeakWrapper(lambda sequence, *args, **kwargs: self.spoken.append(
			" ".join(item for item in sequence if isinstance(item, str))
		))

		def speakObject(obj, reason=None, _prefixSpeechCommand=None, priority=None):
			# A stand-in for NVDA's speakObject: the name, then words for some roles.
			sequence = [obj.name] if obj.name else []
			if obj.role in ROLE_WORDS:
				sequence.append(ROLE_WORDS[obj.role])
			sys.modules["speech.speech"].speak(sequence)

		wrappedSpeakObject = plugin._makeSpeakObjectWrapper(speakObject)
		for module in (speechPackage, speechModule):
			module.speak = speak
			module.speakObject = wrappedSpeakObject
			module.cancelSpeech = lambda: self.spoken.append(CANCEL)
		speechPackage.manager = types.SimpleNamespace(_shouldCancelExpiredFocusEvents=lambda: False)
		nvda["api"] = api
		nvda["speech"] = speechPackage

		self.globalPlugin = object.__new__(plugin.GlobalPlugin)
		self.globalPlugin._lastBodyWindow = None
		nvda["globalPluginHandler"] = types.SimpleNamespace(runningPlugins=[self.globalPlugin])
		for code in (EVENT_EXECUTER, EXECUTE_EVENT, DO_PRE_GAIN_FOCUS):
			exec(code.strip(), nvda)
		self.executeEvent = nvda["executeEvent"]

		self.pluginLog = mock.MagicMock(name="log")
		patches = (
			mock.patch.object(plugin, "log", self.pluginLog),
			mock.patch.object(plugin, "time", self.clock),
			mock.patch.object(plugin, "_user32", lambda: _User32(self.visible)),
			mock.patch.object(plugin, "_enumerateVisibleTopLevelWindows", lambda: sorted(self.visible)),
			mock.patch.object(plugin, "_openMessageWindows", {}),
			mock.patch.object(plugin, "_messageOpeningUntil", 0.0),
			mock.patch.object(plugin, "_messageReturnTitlePending", False),
			mock.patch.object(plugin, "_gateUntil", 0.0),
			mock.patch.object(plugin, "_inCallDepth", 0),
		)
		for patch in patches:
			patch.start()
			self.addCleanup(patch.stop)

		self.inbox = self.NVDAObject(Role.PANE, "Inbox - Outlook - Outlook", self.desktop, window=0x100)
		self.inboxRow = self.NVDAObject(Role.TABLEROW, "From joshknnd1982, Subject Re: ...", self.inbox, window=0x100)
		self.notepad = self.NVDAObject(Role.PANE, NOTEPAD_TITLE, self.desktop, appName="notepad", window=0x200)
		self.notepadText = self.NVDAObject(Role.DOCUMENT, "Text editor", self.notepad, appName="notepad", window=0x200)
		self.switcher = self.NVDAObject(Role.WINDOW, "Task Switching", self.desktop, appName="explorer", window=0x300, presentable=False)
		self.visible.update((0x100, 0x200, 0x300))
		self.classes[0x100] = "rctrl_renwnd32"
		self.titles.update({0x100: "Inbox - Outlook - Outlook", 0x200: NOTEPAD_TITLE, 0x300: "Task Switching"})

	def tearDown(self):
		# The add-on logs its own failures and carries on; none are expected here.
		self.assertEqual(self.pluginLog.error.call_args_list, [])
		self.assertEqual(self.pluginLog.debugWarning.call_args_list, [])

	# What the user does.

	def later(self, seconds=5.0):
		self.clock.now += seconds

	def focus(self, obj, foreground):
		self.api.realForeground = foreground
		self.executeEvent("gainFocus", obj)

	def messageWindow(self, window, title=TITLE):
		self.visible.add(window)
		self.classes[window] = "rctrl_renwnd32"
		self.titles[window] = title + " "
		return _MessageWindow(self.NVDAObject, self.desktop, window, title)

	def openMessage(self, message):
		"""Enter on a message in the list: Outlook opens it in its own window."""
		self.later()
		self.visible.add(message.window)
		self.focus(message.document, message.pane)

	def altTabTo(self, pane, target):
		"""Alt+Tab: the switcher says the window's title, then the window comes to the front."""
		self.later()
		item = self.NVDAObject(Role.LISTITEM, pane.name.strip(), self.switcher, appName="explorer", window=0x300)
		self.focus(item, self.switcher)
		self.later(0.4)
		self.focus(target, pane)

	def closeMessage(self, message):
		"""Escape: Outlook closes the window, and the message list gets focus."""
		self.later()
		self.visible.discard(message.window)
		self.focus(self.inboxRow, self.inbox)

	def heardInFull(self):
		"""What was said after the last cancel, so nothing cut it off."""
		if CANCEL not in self.spoken:
			return list(self.spoken)
		return self.spoken[len(self.spoken) - self.spoken[::-1].index(CANCEL):]

	# Tests.

	def test_altTabBackToAnOpenMessageSaysItsWholeTitle(self):
		message = self.messageWindow(0x400)
		self.openMessage(message)
		self.altTabTo(self.notepad, self.notepadText)
		del self.spoken[:]
		self.altTabTo(message.pane, message.document)
		# The switcher's item began the title, and NVDA's event_foreground cut it off...
		self.assertEqual(self.spoken, [CANCEL, TITLE, CANCEL, TITLE + " "])
		# ...so the window's own title, and nothing after it, is what is heard in full.
		self.assertEqual(self.heardInFull(), [TITLE + " "])

	def test_openingAMessageStaysSilent(self):
		message = self.messageWindow(0x400)
		self.focus(self.inboxRow, self.inbox)
		del self.spoken[:]
		self.openMessage(message)
		self.assertEqual(self.heardInFull(), [])

	def test_aMessageOpenWhenNvdaStartedCountsAsOpen(self):
		# The issue #5 log: NVDA started with the message open, then Alt+Tab went back to it.
		message = self.messageWindow(0x400)
		plugin._noteOpenMessageWindows()
		self.focus(self.notepadText, self.notepad)
		self.altTabTo(message.pane, message.document)
		self.assertEqual(self.heardInFull(), [TITLE + " "])

	def test_onlyOutlookMessageWindowsAreNotedAtStart(self):
		message = self.messageWindow(0x400)
		plugin._noteOpenMessageWindows()
		self.assertEqual(set(plugin._openMessageWindows), {message.window})

	def test_reopeningAClosedMessageStaysSilent(self):
		message = self.messageWindow(0x400)
		self.openMessage(message)
		self.closeMessage(message)
		# Even if Outlook shows the message in the same window again.
		del self.spoken[:]
		self.openMessage(message)
		self.assertEqual(self.heardInFull(), [])

	def test_aClosedMessageIsForgotten(self):
		message = self.messageWindow(0x400)
		self.openMessage(message)
		self.assertIn(message.window, plugin._openMessageWindows)
		self.closeMessage(message)
		self.assertNotIn(message.window, plugin._openMessageWindows)

	def test_comingBackFromAnotherMessageSaysTheTitle(self):
		# Reply opens a second window; sending it brings the first one back.
		first = self.messageWindow(0x400)
		reply = self.messageWindow(0x500, OTHER_TITLE)
		self.openMessage(first)
		self.openMessage(reply)
		self.assertEqual(self.heardInFull(), [])
		self.later()
		self.visible.discard(reply.window)
		self.focus(first.document, first.pane)
		self.assertEqual(self.heardInFull(), [TITLE + " "])

	def test_bringingANewMessageForwardTwiceWhileItOpensStaysSilent(self):
		message = self.messageWindow(0x400)
		self.openMessage(message)
		self.later(0.5)
		self.focus(self.notepadText, self.notepad)
		self.later(0.5)
		self.focus(message.document, message.pane)
		self.assertEqual(self.heardInFull(), [])

	def test_theTitleIsSaidOnce(self):
		message = self.messageWindow(0x400)
		self.openMessage(message)
		self.altTabTo(self.notepad, self.notepadText)
		self.altTabTo(message.pane, message.document)
		del self.spoken[:]
		sys.modules["speech"].speakObject(message.pane, reason=OutputReason.FOCUSENTERED)
		self.assertEqual(self.spoken, [])

	def test_otherWindowsAreUnchanged(self):
		message = self.messageWindow(0x400)
		self.openMessage(message)
		self.altTabTo(self.notepad, self.notepadText)
		self.assertEqual(self.heardInFull(), [NOTEPAD_TITLE, "Text editor document"])
		self.altTabTo(self.inbox, self.inboxRow)
		self.assertEqual(self.heardInFull(), ["Inbox - Outlook - Outlook", "From joshknnd1982, Subject Re: ..."])

	def test_messageWindowsAreForgottenWhenTheAddOnStops(self):
		message = self.messageWindow(0x400)
		self.openMessage(message)
		with mock.patch.object(plugin, "_removeReformattedMessages"), mock.patch.object(plugin.updater, "stop"):
			self.globalPlugin.terminate()
		self.assertEqual(plugin._openMessageWindows, {})


if __name__ == "__main__":
	unittest.main()

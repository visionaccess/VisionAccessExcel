# -*- coding: UTF-8 -*-
# VisionAccess Excel Enhancements - app module for Microsoft Excel
# Copyright (C) 2026 Mukesh Jain, VisionAccess <info@visionaccess.org>
# This add-on is covered by the GNU General Public License, version 2.
#
# Features:
#  * Announces "blank" after cell content is deleted (Delete / Backspace),
#    both on worksheet cells and inside edit fields.
#  * Control+Shift+D: list every cell with data in the active worksheet.
#  * Control+Shift+R: list cells with data in the current row.
#  * Control+Shift+C: list cells with data in the current column.
#    (Enter on any cell in these lists moves Excel's focus to that cell.)
#  * Control+Shift+S: sheet manager (go to, rename, move, protect, hide, insert, delete).
#  * Announces the table range in the Create / Resize Table dialogs.
#  * NVDA+Alt+T: table information, NVDA+Alt+U: used range,
#    NVDA+Alt+S: selection summary (count, sum, average, min, max).
#  * Optional: say "blank" when moving to an empty cell.

import addonHandler
import api
import config
import controlTypes
import core
import gui
import textInfos
import ui
import winUser
import wx
from logHandler import log
from NVDAObjects import NVDAObject
from scriptHandler import script

try:
	from editableText import EditableText
except ImportError:  # pragma: no cover
	EditableText = None

try:
	from nvdaBuiltin.appModules.excel import AppModule as BaseAppModule
except ImportError:
	from appModuleHandler import AppModule as BaseAppModule

addonHandler.initTranslation()

# Translators: category of commands shown in NVDA's Input Gestures dialog.
SCRCAT = _("VisionAccess Excel")

CONF_SECTION = "visionAccessExcel"
CONF_DEFAULTS = {
	"announceBlankOnDelete": True,
	"announceBlankOnEmptyCell": False,
	"announceTableRangeInDialogs": True,
	"showHeadersInLists": True,
	"maxListedCells": 2000,
}
CONF_SPEC = {
	"announceBlankOnDelete": "boolean(default=True)",
	"announceBlankOnEmptyCell": "boolean(default=False)",
	"announceTableRangeInDialogs": "boolean(default=True)",
	"showHeadersInLists": "boolean(default=True)",
	"maxListedCells": "integer(default=2000, min=100, max=20000)",
}
if CONF_SECTION not in config.conf.spec:
	config.conf.spec[CONF_SECTION] = CONF_SPEC


def getSetting(key):
	try:
		return config.conf[CONF_SECTION][key]
	except Exception:
		return CONF_DEFAULTS[key]


# Excel constants
XL_CELL_TYPE_CONSTANTS = 2
XL_CELL_TYPE_FORMULAS = -4123
XL_SHEET_VISIBLE = -1
XL_SHEET_HIDDEN = 0
OBJID_NATIVEOM = -16

# Words that identify dialogs in which a table/data range is typed.
TABLE_DIALOG_KEYWORDS = ("table", "pivot", "range")


# ---------------------------------------------------------------- helpers

def colLetter(col):
	letters = ""
	while col > 0:
		col, rem = divmod(col - 1, 26)
		letters = chr(65 + rem) + letters
	return letters


def cellAddress(row, col):
	return "%s%d" % (colLetter(col), row)


def _count(rng):
	try:
		return int(rng.CountLarge)
	except Exception:
		return int(rng.Count)


def areaAddress(area):
	r0, c0 = area.Row, area.Column
	nr, nc = area.Rows.Count, area.Columns.Count
	start = cellAddress(r0, c0)
	if nr == 1 and nc == 1:
		return start
	return "%s:%s" % (start, cellAddress(r0 + nr - 1, c0 + nc - 1))


def rangeAddress(rng):
	try:
		n = rng.Areas.Count
	except Exception:
		n = 1
	if n <= 1:
		return areaAddress(rng)
	return ", ".join(areaAddress(rng.Areas.Item(i)) for i in range(1, n + 1))


def _isNothing(obj):
	if obj is None:
		return True
	try:
		obj.Row
		return False
	except Exception:
		return True


def formatValue(value):
	if value is None:
		return ""
	if isinstance(value, bool):
		return str(value)
	if isinstance(value, float):
		if value.is_integer():
			return str(int(value))
		return ("%.10f" % value).rstrip("0").rstrip(".")
	return str(value)


def cleanText(text):
	return str(text).replace("\r", " ").replace("\n", " ").strip()


def getExcelApplication(obj=None):
	"""Return Excel's Application COM object for the given NVDA object."""
	if obj is None:
		obj = api.getFocusObject()
	for attr in ("excelCellObject", "excelRangeObject", "excelWindowObject", "excelWorksheetObject"):
		try:
			comObj = getattr(obj, attr, None)
			if comObj is not None:
				return comObj.Application
		except Exception:
			pass
	hwnd = getattr(obj, "windowHandle", None)
	if hwnd:
		try:
			import oleacc
			import comtypes.automation
			import comtypes.client.dynamic
			pDispatch = oleacc.AccessibleObjectFromWindow(
				hwnd, OBJID_NATIVEOM, interface=comtypes.automation.IDispatch
			)
			return comtypes.client.dynamic.Dispatch(pDispatch).Application
		except Exception:
			log.debugWarning("VisionAccess Excel: native object model not available", exc_info=True)
	try:
		import comtypes.client
		return comtypes.client.GetActiveObject("Excel.Application", dynamic=True)
	except Exception:
		log.debugWarning("VisionAccess Excel: could not get Excel.Application", exc_info=True)
	return None


def _bulkValues(area, nr, nc):
	"""Read Value2 of an area as a 2D list."""
	try:
		val = area.Value2
	except Exception:
		return None
	if nr == 1 and nc == 1:
		return [[val]]
	try:
		return [list(row) for row in val]
	except Exception:
		return None


def collectDataCells(rng, limit):
	"""Return ({(row, col): (text, isFormula)}, totalCount) for non-empty cells in rng."""
	found = {}
	total = 0
	if _isNothing(rng):
		return found, total
	if _count(rng) == 1:
		# SpecialCells on a single cell would search the whole sheet, so check it directly.
		try:
			formula = rng.Formula
			if isinstance(formula, str) and formula != "":
				text = cleanText(rng.Text) or formatValue(rng.Value2)
				found[(rng.Row, rng.Column)] = (text, formula.startswith("="))
				total = 1
		except Exception:
			pass
		return found, total
	for cellType, isFormula in ((XL_CELL_TYPE_CONSTANTS, False), (XL_CELL_TYPE_FORMULAS, True)):
		try:
			sub = rng.SpecialCells(cellType)
			areaCount = sub.Areas.Count
		except Exception:
			continue  # "No cells were found"
		for a in range(1, areaCount + 1):
			area = sub.Areas.Item(a)
			r0, c0 = area.Row, area.Column
			nr, nc = area.Rows.Count, area.Columns.Count
			total += nr * nc
			if len(found) >= limit:
				continue
			values = _bulkValues(area, nr, nc)
			for i in range(nr):
				if len(found) >= limit:
					break
				for j in range(nc):
					if len(found) >= limit:
						break
					text = ""
					try:
						text = cleanText(area.Item(i + 1, j + 1).Text)
					except Exception:
						pass
					if (not text or set(text) == {"#"}) and values is not None:
						try:
							text = formatValue(values[i][j])
						except Exception:
							pass
					found[(r0 + i, c0 + j)] = (text, isFormula)
	return found, total


def bringToFront(hwnd):
	if not hwnd:
		return
	try:
		winUser.setForegroundWindow(hwnd)
	except Exception:
		pass


# ---------------------------------------------------------------- dialogs

class CellListDialog(wx.Dialog):
	_instance = None

	def __init__(self, parent, title, summary, entries, onChoose):
		super().__init__(parent, title=title)
		CellListDialog._instance = self
		self._entries = entries  # list of (label, row, col)
		self._visible = list(entries)
		self._onChoose = onChoose
		mainSizer = wx.BoxSizer(wx.VERTICAL)
		self.summaryText = wx.StaticText(self, label=summary)
		mainSizer.Add(self.summaryText, flag=wx.ALL, border=8)
		# Translators: label of the filter field in the cell list dialog.
		mainSizer.Add(wx.StaticText(self, label=_("&Filter:")), flag=wx.LEFT | wx.RIGHT, border=8)
		self.filterEdit = wx.TextCtrl(self)
		mainSizer.Add(self.filterEdit, flag=wx.EXPAND | wx.ALL, border=8)
		# Translators: label of the list of cells.
		mainSizer.Add(wx.StaticText(self, label=_("&Cells:")), flag=wx.LEFT | wx.RIGHT, border=8)
		self.listBox = wx.ListBox(self, size=(560, 320), choices=[e[0] for e in entries])
		mainSizer.Add(self.listBox, proportion=1, flag=wx.EXPAND | wx.ALL, border=8)
		btnSizer = wx.BoxSizer(wx.HORIZONTAL)
		# Translators: button that moves Excel's focus to the chosen cell.
		self.goButton = wx.Button(self, wx.ID_OK, label=_("&Go to cell"))
		closeButton = wx.Button(self, wx.ID_CANCEL, label=_("Close"))
		btnSizer.Add(self.goButton, flag=wx.RIGHT, border=8)
		btnSizer.Add(closeButton)
		mainSizer.Add(btnSizer, flag=wx.ALIGN_RIGHT | wx.ALL, border=8)
		self.goButton.SetDefault()
		self.SetEscapeId(wx.ID_CANCEL)
		self.filterEdit.Bind(wx.EVT_TEXT, self.onFilter)
		self.listBox.Bind(wx.EVT_LISTBOX_DCLICK, self.onGo)
		self.goButton.Bind(wx.EVT_BUTTON, self.onGo)
		closeButton.Bind(wx.EVT_BUTTON, lambda evt: self.Close())
		self.Bind(wx.EVT_CLOSE, self.onClose)
		self.SetSizerAndFit(mainSizer)
		self.CentreOnScreen()
		if entries:
			self.listBox.SetSelection(0)
		self.listBox.SetFocus()

	def onFilter(self, evt):
		text = self.filterEdit.GetValue().strip().lower()
		self._visible = [e for e in self._entries if text in e[0].lower()] if text else list(self._entries)
		self.listBox.Set([e[0] for e in self._visible])
		if self._visible:
			self.listBox.SetSelection(0)

	def onGo(self, evt):
		index = self.listBox.GetSelection()
		if index == wx.NOT_FOUND or index >= len(self._visible):
			return
		label, row, col = self._visible[index]
		self._onChoose(row, col)
		self.Close()

	def onClose(self, evt):
		CellListDialog._instance = None
		self.Destroy()


class SheetManagerDialog(wx.Dialog):
	_instance = None

	def __init__(self, parent, app, excelHwnd):
		# Translators: title of the sheet manager dialog.
		super().__init__(parent, title=_("Sheet Manager - VisionAccess Excel"))
		SheetManagerDialog._instance = self
		self.app = app
		self.excelHwnd = excelHwnd
		mainSizer = wx.BoxSizer(wx.VERTICAL)
		# Translators: label of the list of sheets.
		mainSizer.Add(wx.StaticText(self, label=_("&Sheets:")), flag=wx.LEFT | wx.RIGHT | wx.TOP, border=8)
		self.listBox = wx.ListBox(self, size=(460, 260))
		mainSizer.Add(self.listBox, proportion=1, flag=wx.EXPAND | wx.ALL, border=8)
		grid = wx.GridSizer(cols=4, hgap=6, vgap=6)
		buttons = (
			(wx.ID_OK, _("&Go to sheet"), self.onGo),
			(wx.ID_ANY, _("&Rename..."), self.onRename),
			(wx.ID_ANY, _("Move &up"), self.onMoveUp),
			(wx.ID_ANY, _("Move &down"), self.onMoveDown),
			(wx.ID_ANY, _("&Protect / Unprotect..."), self.onProtect),
			(wx.ID_ANY, _("&Hide / Unhide"), self.onHide),
			(wx.ID_ANY, _("&Insert new sheet..."), self.onInsert),
			(wx.ID_ANY, _("De&lete..."), self.onDelete),
		)
		for btnId, label, handler in buttons:
			btn = wx.Button(self, btnId, label=label)
			btn.Bind(wx.EVT_BUTTON, handler)
			grid.Add(btn, flag=wx.EXPAND)
			if btnId == wx.ID_OK:
				btn.SetDefault()
		mainSizer.Add(grid, flag=wx.EXPAND | wx.ALL, border=8)
		closeButton = wx.Button(self, wx.ID_CANCEL, label=_("Close"))
		closeButton.Bind(wx.EVT_BUTTON, lambda evt: self.Close())
		mainSizer.Add(closeButton, flag=wx.ALIGN_RIGHT | wx.ALL, border=8)
		self.SetEscapeId(wx.ID_CANCEL)
		self.listBox.Bind(wx.EVT_LISTBOX_DCLICK, self.onGo)
		self.Bind(wx.EVT_CLOSE, self.onClose)
		self.SetSizerAndFit(mainSizer)
		self.CentreOnScreen()
		self.refresh(selectActive=True)
		self.listBox.SetFocus()

	# --- data
	def _workbook(self):
		return self.app.ActiveWorkbook

	def refresh(self, selectIndex=None, selectActive=False):
		wb = self._workbook()
		sheets = wb.Sheets
		count = sheets.Count
		try:
			activeName = wb.ActiveSheet.Name
		except Exception:
			activeName = None
		labels = []
		activeIndex = 0
		for i in range(1, count + 1):
			sh = sheets.Item(i)
			name = sh.Name
			parts = [name]
			if name == activeName:
				activeIndex = i - 1
				parts.append(_("active"))
			try:
				if sh.ProtectContents:
					parts.append(_("protected"))
			except Exception:
				pass
			try:
				vis = sh.Visible
				if vis == XL_SHEET_HIDDEN:
					parts.append(_("hidden"))
				elif vis == 2:
					parts.append(_("very hidden"))
			except Exception:
				pass
			parts.append(_("{n} of {total}").format(n=i, total=count))
			labels.append(", ".join(parts))
		self.listBox.Set(labels)
		if selectIndex is None:
			selectIndex = activeIndex if selectActive else 0
		if labels:
			self.listBox.SetSelection(max(0, min(selectIndex, len(labels) - 1)))

	def _selected(self):
		index = self.listBox.GetSelection()
		if index == wx.NOT_FOUND:
			return None, None
		return index, self._workbook().Sheets.Item(index + 1)

	def _error(self, message):
		wx.MessageBox(message, _("VisionAccess Excel"), wx.OK | wx.ICON_ERROR, self)

	# --- actions
	def onGo(self, evt):
		index, sh = self._selected()
		if sh is None:
			return
		try:
			if sh.Visible != XL_SHEET_VISIBLE:
				self._error(_("This sheet is hidden. Use Hide / Unhide to show it first."))
				return
			sh.Activate()
		except Exception:
			log.debugWarning("Activate failed", exc_info=True)
			self._error(_("Could not go to this sheet."))
			return
		self.Close()

	def onRename(self, evt):
		index, sh = self._selected()
		if sh is None:
			return
		with wx.TextEntryDialog(self, _("New name for the sheet:"), _("Rename sheet"), sh.Name) as d:
			if d.ShowModal() != wx.ID_OK:
				return
			newName = d.GetValue().strip()
		if not newName:
			return
		try:
			sh.Name = newName
		except Exception:
			self._error(_(
				"Could not rename the sheet. The name may already be used, be longer than 31 characters, "
				"contain one of these characters: \\ / ? * [ ] :, or the workbook structure may be protected."
			))
			return
		self.refresh(selectIndex=index)

	def onMoveUp(self, evt):
		index, sh = self._selected()
		if sh is None:
			return
		if index == 0:
			ui.message(_("Already the first sheet"))
			return
		try:
			sheets = self._workbook().Sheets
			sheets.Item(index + 1).Move(sheets.Item(index))
		except Exception:
			self._error(_("Could not move the sheet. The workbook structure may be protected."))
			return
		self.refresh(selectIndex=index - 1)

	def onMoveDown(self, evt):
		index, sh = self._selected()
		if sh is None:
			return
		sheets = self._workbook().Sheets
		if index + 1 >= sheets.Count:
			ui.message(_("Already the last sheet"))
			return
		try:
			# Moving the next sheet before this one moves this one down.
			sheets.Item(index + 2).Move(sheets.Item(index + 1))
		except Exception:
			self._error(_("Could not move the sheet. The workbook structure may be protected."))
			return
		self.refresh(selectIndex=index + 1)

	def onProtect(self, evt):
		index, sh = self._selected()
		if sh is None:
			return
		try:
			isProtected = bool(sh.ProtectContents)
		except Exception:
			isProtected = False
		if isProtected:
			try:
				sh.Unprotect("")
				ui.message(_("Sheet unprotected"))
			except Exception:
				with wx.PasswordEntryDialog(self, _("Password to unprotect the sheet:"), _("Unprotect sheet")) as d:
					if d.ShowModal() != wx.ID_OK:
						return
					password = d.GetValue()
				try:
					sh.Unprotect(password)
					ui.message(_("Sheet unprotected"))
				except Exception:
					self._error(_("Could not unprotect the sheet. The password may be wrong."))
					return
		else:
			with wx.PasswordEntryDialog(
				self, _("Password to protect the sheet (leave empty for no password):"), _("Protect sheet")
			) as d:
				if d.ShowModal() != wx.ID_OK:
					return
				password = d.GetValue()
			if password:
				with wx.PasswordEntryDialog(self, _("Confirm the password:"), _("Protect sheet")) as d2:
					if d2.ShowModal() != wx.ID_OK:
						return
					if d2.GetValue() != password:
						self._error(_("The passwords do not match."))
						return
			try:
				if password:
					sh.Protect(password)
				else:
					sh.Protect()
				ui.message(_("Sheet protected"))
			except Exception:
				self._error(_("Could not protect the sheet."))
				return
		self.refresh(selectIndex=index)

	def onHide(self, evt):
		index, sh = self._selected()
		if sh is None:
			return
		try:
			if sh.Visible == XL_SHEET_VISIBLE:
				sh.Visible = XL_SHEET_HIDDEN
				ui.message(_("Sheet hidden"))
			else:
				sh.Visible = XL_SHEET_VISIBLE
				ui.message(_("Sheet shown"))
		except Exception:
			self._error(_(
				"Could not change the visibility. A workbook must keep at least one visible sheet, "
				"and the workbook structure must not be protected."
			))
			return
		self.refresh(selectIndex=index)

	def onInsert(self, evt):
		with wx.TextEntryDialog(
			self, _("Name for the new sheet (leave empty for the default name):"), _("Insert sheet")
		) as d:
			if d.ShowModal() != wx.ID_OK:
				return
			newName = d.GetValue().strip()
		try:
			newSheet = self._workbook().Worksheets.Add()
		except Exception:
			self._error(_("Could not insert a sheet. The workbook structure may be protected."))
			return
		if newName:
			try:
				newSheet.Name = newName
			except Exception:
				self._error(_("The sheet was inserted, but this name could not be used."))
		ui.message(_("Sheet inserted"))
		self.refresh(selectActive=True)

	def onDelete(self, evt):
		index, sh = self._selected()
		if sh is None:
			return
		name = sh.Name
		answer = wx.MessageBox(
			_("Permanently delete the sheet {name}? This cannot be undone.").format(name=name),
			_("Delete sheet"), wx.YES_NO | wx.NO_DEFAULT | wx.ICON_WARNING, self
		)
		if answer != wx.YES:
			return
		app = self.app
		try:
			oldAlerts = app.DisplayAlerts
			app.DisplayAlerts = False
			try:
				sh.Delete()
			finally:
				app.DisplayAlerts = oldAlerts
			ui.message(_("Sheet deleted"))
		except Exception:
			self._error(_(
				"Could not delete the sheet. A workbook must keep at least one visible sheet, "
				"and the workbook structure must not be protected."
			))
			return
		self.refresh(selectIndex=max(0, index - 1))

	def onClose(self, evt):
		SheetManagerDialog._instance = None
		hwnd = self.excelHwnd
		self.Destroy()
		core.callLater(150, bringToFront, hwnd)


def _showDialog(factory):
	gui.mainFrame.prePopup()
	dlg = factory()
	dlg.Show()
	gui.mainFrame.postPopup()


# ---------------------------------------------------------------- overlay classes

class VAExcelCell(NVDAObject):
	"""Added to Excel worksheet cells (object model and UIA)."""

	def _get_name(self):
		name = super()._get_name()
		if name or not getSetting("announceBlankOnEmptyCell"):
			return name
		cellObj = getattr(self, "excelCellObject", None)
		if cellObj is None:
			return name
		try:
			if cellObj.Formula == "":
				# Translators: spoken for an empty cell.
				return _("blank")
		except Exception:
			pass
		return name

	def event_gainFocus(self):
		super().event_gainFocus()
		if getSetting("announceTableRangeInDialogs"):
			try:
				self.appModule.vaCacheRangeInfo(self)
			except Exception:
				log.debugWarning("VisionAccess Excel: caching range failed", exc_info=True)

	def _reportIfSelectionBlank(self):
		app = getExcelApplication(self)
		if app is None:
			return
		try:
			if app.WorksheetFunction.CountA(app.Selection) == 0:
				ui.message(_("blank"))
		except Exception:
			log.debugWarning("VisionAccess Excel: blank check failed", exc_info=True)

	@script(gestures=["kb:delete", "kb:numpadDelete"])
	def script_vaDeleteContent(self, gesture):
		gesture.send()
		if getSetting("announceBlankOnDelete"):
			core.callLater(150, self._reportIfSelectionBlank)

	@script(gesture="kb:backspace")
	def script_vaBackspaceContent(self, gesture):
		# Backspace on a cell clears it and opens it for editing.
		gesture.send()
		if getSetting("announceBlankOnDelete"):
			core.callLater(200, ui.message, _("blank"))

	@script(
		# Translators: input gesture description.
		description=_("Lists all cells that contain data in the current worksheet. Press Enter on a cell to move to it."),
		gesture="kb:control+shift+d", category=SCRCAT,
	)
	def script_vaListSheetCells(self, gesture):
		self.appModule.vaShowCellList(self, "sheet")

	@script(
		description=_("Lists the cells that contain data in the current row. Press Enter on a cell to move to it."),
		gesture="kb:control+shift+r", category=SCRCAT,
	)
	def script_vaListRowCells(self, gesture):
		self.appModule.vaShowCellList(self, "row")

	@script(
		description=_("Lists the cells that contain data in the current column. Press Enter on a cell to move to it."),
		gesture="kb:control+shift+c", category=SCRCAT,
	)
	def script_vaListColumnCells(self, gesture):
		self.appModule.vaShowCellList(self, "column")

	@script(
		description=_("Opens the sheet manager to go to, rename, move, protect, hide, insert or delete sheets."),
		gesture="kb:control+shift+s", category=SCRCAT,
	)
	def script_vaSheetManager(self, gesture):
		self.appModule.vaShowSheetManager(self)

	@script(
		description=_("Reports the used range (area containing data) of the current worksheet."),
		gesture="kb:NVDA+alt+u", category=SCRCAT,
	)
	def script_vaReportUsedRange(self, gesture):
		app = getExcelApplication(self)
		try:
			sheet = app.ActiveSheet
			used = sheet.UsedRange
			if _count(used) == 1 and used.Formula == "":
				ui.message(_("Sheet {name} is empty").format(name=sheet.Name))
				return
			ui.message(_("Sheet {name}: data from {start} to {end}, {rows} rows, {cols} columns").format(
				name=sheet.Name,
				start=cellAddress(used.Row, used.Column),
				end=cellAddress(used.Row + used.Rows.Count - 1, used.Column + used.Columns.Count - 1),
				rows=used.Rows.Count, cols=used.Columns.Count,
			))
		except Exception:
			log.debugWarning("used range failed", exc_info=True)
			ui.message(_("Could not read the worksheet"))

	@script(
		description=_("Reports a summary of the selection: cells, filled cells, sum, average, minimum and maximum."),
		gesture="kb:NVDA+alt+s", category=SCRCAT,
	)
	def script_vaSelectionSummary(self, gesture):
		app = getExcelApplication(self)
		try:
			sel = app.Selection
			wf = app.WorksheetFunction
			cells = _count(sel)
			filled = int(wf.CountA(sel))
			numbers = int(wf.Count(sel))
			parts = [
				rangeAddress(sel),
				_("{n} cells").format(n=cells),
				_("{n} with data").format(n=filled),
			]
			if numbers:
				total = wf.Sum(sel)
				parts.append(_("{n} numbers").format(n=numbers))
				parts.append(_("sum {v}").format(v=formatValue(round(total, 4))))
				parts.append(_("average {v}").format(v=formatValue(round(total / numbers, 4))))
				parts.append(_("minimum {v}").format(v=formatValue(wf.Min(sel))))
				parts.append(_("maximum {v}").format(v=formatValue(wf.Max(sel))))
			ui.message(", ".join(parts))
		except Exception:
			log.debugWarning("selection summary failed", exc_info=True)
			ui.message(_("Could not read the selection"))


class VAEditField(NVDAObject):
	"""Added to edit fields in Excel (in-cell editing, formula bar, dialogs)."""

	def _vaRunBase(self, scriptName, gesture):
		func = getattr(super(VAEditField, self), scriptName, None)
		if func is not None:
			func(gesture)
		else:
			gesture.send()

	def _vaFieldText(self):
		"""Return the field's text, or None if it cannot be determined."""
		try:
			return self.makeTextInfo(textInfos.POSITION_ALL).text
		except Exception:
			pass
		try:
			value = self.value
			if value is not None:
				return value
		except Exception:
			pass
		return None

	def _vaReportIfEmpty(self):
		if api.getFocusObject() is not self:
			return
		text = self._vaFieldText()
		if text is not None and not text.strip():
			ui.message(_("blank"))

	def _vaAfterDelete(self):
		if getSetting("announceBlankOnDelete"):
			core.callLater(80, self._vaReportIfEmpty)

	@script(gesture="kb:backspace")
	def script_vaBackspace(self, gesture):
		self._vaRunBase("script_caret_backspaceCharacter", gesture)
		self._vaAfterDelete()

	@script(gesture="kb:control+backspace")
	def script_vaBackspaceWord(self, gesture):
		self._vaRunBase("script_caret_backspaceWord", gesture)
		self._vaAfterDelete()

	@script(gestures=["kb:delete", "kb:numpadDelete"])
	def script_vaDelete(self, gesture):
		self._vaRunBase("script_caret_deleteCharacter", gesture)
		self._vaAfterDelete()

	def event_gainFocus(self):
		super().event_gainFocus()
		if not getSetting("announceTableRangeInDialogs"):
			return
		try:
			title = (api.getForegroundObject().name or "").lower()
		except Exception:
			return
		if not any(word in title for word in TABLE_DIALOG_KEYWORDS):
			return
		core.callLater(250, self.appModule.vaAnnounceRangeForDialog, self)


# ---------------------------------------------------------------- app module

class AppModule(BaseAppModule):

	scriptCategory = SCRCAT
	vaLastRange = None
	vaLastRangeKind = None
	vaLastTableName = None

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		super().chooseNVDAObjectOverlayClasses(obj, clsList)
		try:
			role = obj.role
			windowClass = obj.windowClassName
		except Exception:
			return
		if role == controlTypes.Role.TABLECELL and windowClass == "EXCEL7":
			clsList.insert(0, VAExcelCell)
		elif role == controlTypes.Role.EDITABLETEXT or (
			EditableText is not None and any(isinstance(c, type) and issubclass(c, EditableText) for c in clsList)
		):
			clsList.insert(0, VAEditField)

	# --- table / range caching for dialogs
	def vaCacheRangeInfo(self, obj):
		app = getExcelApplication(obj)
		if app is None:
			return
		cell = app.ActiveCell
		try:
			table = cell.ListObject
			tableRange = table.Range
			self.vaLastRange = areaAddress(tableRange)
			self.vaLastRangeKind = "table"
			self.vaLastTableName = table.Name
			return
		except Exception:
			pass
		try:
			region = cell.CurrentRegion
			if _count(region) > 1:
				self.vaLastRange = areaAddress(region)
				self.vaLastRangeKind = "region"
				self.vaLastTableName = None
				return
		except Exception:
			pass
		self.vaLastRange = None
		self.vaLastRangeKind = None
		self.vaLastTableName = None

	def vaAnnounceRangeForDialog(self, editObj):
		if api.getFocusObject() is not editObj:
			return
		text = editObj._vaFieldText()
		if not text or not text.strip():
			try:
				text = winUser.getWindowText(editObj.windowHandle)
			except Exception:
				text = ""
		text = (text or "").strip()
		if text and any(ch.isdigit() for ch in text):
			ui.message(_("Range: {range}").format(range=text))
		elif self.vaLastRange:
			if self.vaLastRangeKind == "table":
				ui.message(_("Table range: {range}").format(range=self.vaLastRange))
			else:
				ui.message(_("Data range: {range}").format(range=self.vaLastRange))

	@script(
		description=_(
			"Reports the name and range of the table at the current cell. "
			"In the Create Table or Resize Table dialog, reports the table range."
		),
		gesture="kb:NVDA+alt+t", category=SCRCAT,
	)
	def script_vaReportTable(self, gesture):
		focus = api.getFocusObject()
		if isinstance(focus, VAExcelCell):
			app = getExcelApplication(focus)
			try:
				cell = app.ActiveCell
				try:
					table = cell.ListObject
					rng = table.Range
					ui.message(_("Table {name}, range {range}, {rows} rows including header, {cols} columns").format(
						name=table.Name, range=areaAddress(rng), rows=rng.Rows.Count, cols=rng.Columns.Count,
					))
					return
				except Exception:
					pass
				region = cell.CurrentRegion
				ui.message(_("Not in a table. Current data region {range}").format(range=areaAddress(region)))
			except Exception:
				ui.message(_("Could not read the worksheet"))
			return
		if self.vaLastRange:
			if self.vaLastRangeKind == "table":
				ui.message(_("Table {name}, range {range}").format(name=self.vaLastTableName or "", range=self.vaLastRange))
			else:
				ui.message(_("Data range: {range}").format(range=self.vaLastRange))
		else:
			ui.message(_("No table information available"))

	# --- cell lists
	def vaShowCellList(self, obj, scope):
		if CellListDialog._instance:
			CellListDialog._instance.Raise()
			return
		app = getExcelApplication(obj)
		if app is None:
			ui.message(_("Could not access Excel"))
			return
		try:
			sheet = app.ActiveSheet
			sheetName = sheet.Name
			used = sheet.UsedRange
			cell = app.ActiveCell
			curRow, curCol = cell.Row, cell.Column
			if scope == "sheet":
				rng = used
				title = _("Cells with data in sheet {sheet}").format(sheet=sheetName)
			elif scope == "row":
				rng = app.Intersect(used, cell.EntireRow)
				title = _("Cells with data in row {row}").format(row=curRow)
			else:
				rng = app.Intersect(used, cell.EntireColumn)
				title = _("Cells with data in column {col}").format(col=colLetter(curCol))
			limit = getSetting("maxListedCells")
			found, total = collectDataCells(rng, limit)
			headers = {}
			if found and scope in ("row", "column") and getSetting("showHeadersInLists"):
				headers = self._vaHeaders(sheet, used, scope, curRow, curCol, found)
		except Exception:
			log.debugWarning("VisionAccess Excel: list failed", exc_info=True)
			ui.message(_("Could not read the worksheet. If you are editing a cell, press Escape first."))
			return
		if not found:
			ui.message({
				"sheet": _("No cells with data in this sheet"),
				"row": _("No cells with data in this row"),
				"column": _("No cells with data in this column"),
			}[scope])
			return
		entries = []
		for (row, col) in sorted(found):
			text, isFormula = found[(row, col)]
			label = cellAddress(row, col)
			header = headers.get(col if scope == "row" else row)
			if header:
				label += " (%s)" % header
			label += ": " + (text if text else _("empty text"))
			if isFormula:
				label += ", " + _("formula")
			entries.append((label, row, col))
		if total > len(found):
			summary = _("Showing the first {shown} of {total} cells with data.").format(shown=len(found), total=total)
		else:
			summary = _("{total} cells with data.").format(total=total)
		try:
			excelHwnd = winUser.getAncestor(obj.windowHandle, winUser.GA_ROOT)
		except Exception:
			excelHwnd = None

		def onChoose(row, col):
			self._vaGotoCell(app, sheetName, row, col)
			core.callLater(150, bringToFront, excelHwnd)

		wx.CallAfter(_showDialog, lambda: CellListDialog(gui.mainFrame, title, summary, entries, onChoose))

	def _vaHeaders(self, sheet, used, scope, curRow, curCol, found):
		headers = {}
		if scope == "row":
			headerRow = used.Row
			if curRow == headerRow:
				return headers
			for (row, col) in found:
				try:
					text = cleanText(sheet.Cells.Item(headerRow, col).Text)
					if text:
						headers[col] = text
				except Exception:
					pass
		else:
			headerCol = used.Column
			if curCol == headerCol:
				return headers
			for (row, col) in found:
				try:
					text = cleanText(sheet.Cells.Item(row, headerCol).Text)
					if text:
						headers[row] = text
				except Exception:
					pass
		return headers

	def _vaGotoCell(self, app, sheetName, row, col):
		try:
			sheet = app.ActiveWorkbook.Sheets.Item(sheetName)
			try:
				sheet.Activate()
			except Exception:
				pass
			try:
				sheet.Cells.Item(row, col).Select()
			except Exception:
				sheet.Range(cellAddress(row, col)).Select()
		except Exception:
			log.debugWarning("VisionAccess Excel: goto failed", exc_info=True)
			ui.message(_("Could not move to {cell}").format(cell=cellAddress(row, col)))

	# --- sheet manager
	def vaShowSheetManager(self, obj):
		if SheetManagerDialog._instance:
			SheetManagerDialog._instance.Raise()
			return
		app = getExcelApplication(obj)
		if app is None:
			ui.message(_("Could not access Excel"))
			return
		try:
			app.ActiveWorkbook.Sheets.Count
		except Exception:
			ui.message(_("Could not read the workbook. If you are editing a cell, press Escape first."))
			return
		try:
			excelHwnd = winUser.getAncestor(obj.windowHandle, winUser.GA_ROOT)
		except Exception:
			excelHwnd = None
		wx.CallAfter(_showDialog, lambda: SheetManagerDialog(gui.mainFrame, app, excelHwnd))

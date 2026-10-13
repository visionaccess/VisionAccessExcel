# -*- coding: UTF-8 -*-
# VisionAccess Excel Enhancements - settings panel
# Copyright (C) 2026 Mukesh Jain, VisionAccess <info@visionaccess.org>
# This add-on is covered by the GNU General Public License, version 2.

import addonHandler
import config
import globalPluginHandler
import gui
import wx
from gui import guiHelper

try:
	from gui.settingsDialogs import SettingsPanel, NVDASettingsDialog
except ImportError:  # pragma: no cover
	from gui import SettingsPanel, NVDASettingsDialog

addonHandler.initTranslation()

CONF_SECTION = "visionAccessExcel"
CONF_SPEC = {
	"announceBlankOnDelete": "boolean(default=True)",
	"announceBlankOnEmptyCell": "boolean(default=False)",
	"announceTableRangeInDialogs": "boolean(default=True)",
	"showHeadersInLists": "boolean(default=True)",
	"maxListedCells": "integer(default=2000, min=100, max=20000)",
}
config.conf.spec[CONF_SECTION] = CONF_SPEC


class VisionAccessExcelPanel(SettingsPanel):
	# Translators: title of the settings panel in NVDA's Settings dialog.
	title = _("VisionAccess Excel")

	def makeSettings(self, settingsSizer):
		helper = guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
		conf = config.conf[CONF_SECTION]
		self.blankOnDelete = helper.addItem(wx.CheckBox(
			self, label=_("Announce \"blank\" when cell content is &deleted")))
		self.blankOnDelete.SetValue(conf["announceBlankOnDelete"])
		self.blankOnEmpty = helper.addItem(wx.CheckBox(
			self, label=_("Announce \"blank\" when moving to an &empty cell")))
		self.blankOnEmpty.SetValue(conf["announceBlankOnEmptyCell"])
		self.tableRange = helper.addItem(wx.CheckBox(
			self, label=_("Announce the &table range in Create Table and Resize Table dialogs")))
		self.tableRange.SetValue(conf["announceTableRangeInDialogs"])
		self.headers = helper.addItem(wx.CheckBox(
			self, label=_("Show &headers in row and column cell lists")))
		self.headers.SetValue(conf["showHeadersInLists"])
		self.maxCells = helper.addLabeledControl(
			_("&Maximum number of cells to list:"),
			gui.nvdaControls.SelectOnFocusSpinCtrl,
			min=100, max=20000, initial=conf["maxListedCells"],
		)

	def onSave(self):
		conf = config.conf[CONF_SECTION]
		conf["announceBlankOnDelete"] = self.blankOnDelete.GetValue()
		conf["announceBlankOnEmptyCell"] = self.blankOnEmpty.GetValue()
		conf["announceTableRangeInDialogs"] = self.tableRange.GetValue()
		conf["showHeadersInLists"] = self.headers.GetValue()
		conf["maxListedCells"] = self.maxCells.GetValue()


class GlobalPlugin(globalPluginHandler.GlobalPlugin):

	def __init__(self):
		super().__init__()
		NVDASettingsDialog.categoryClasses.append(VisionAccessExcelPanel)

	def terminate(self):
		try:
			NVDASettingsDialog.categoryClasses.remove(VisionAccessExcelPanel)
		except ValueError:
			pass
		super().terminate()

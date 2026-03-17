# MenuTitle: Kern Tab Contents
# -*- coding: utf-8 -*-
from __future__ import division, print_function, unicode_literals
__doc__ = """
Applies optical kerning to every consecutive glyph pair in the current tab,
using the MB LetterKerner algorithm (inspired by HT LetterSpacer).

For each pair the script measures the combined optical white between the two
glyphs (right white of the left glyph + left white of the right glyph),
depth-clamped and vertically weighted, then sets the kern that brings that
area to the target. Existing kerning can optionally be preserved.

Calibration tip: open a tab with a representative "neutral" pair (e.g. "nn"
for lowercase, "HH" for uppercase), run with your chosen depth/factor/step,
then read the "Current area" in the Macro Window and use that value divided
by 1000 as your target area (the field is in K units², so 50 = 50,000 units²).
"""

import os
import sys

import vanilla
from AppKit import NSRightTextAlignment
from GlyphsApp import Glyphs, Message
from mekkablue import mekkaObject

# ---------------------------------------------------------------------------
# Make the library importable when the script lives inside MB LetterKerner/
# ---------------------------------------------------------------------------
_scriptDir = os.path.dirname(os.path.abspath(__file__))
if _scriptDir not in sys.path:
	sys.path.insert(0, _scriptDir)

from mbLetterKerner import kernLayerToLayer, kernKeyForGlyph, measureMinGap  # noqa: E402

if Glyphs.versionNumber >= 3:
	from GlyphsApp import LTR
	from AppKit import NSNotFound


def _setKerningPair(font, masterID, leftKey, rightKey, value):
	"""Set a kerning pair, handling Glyphs 2 / 3 API differences."""
	if Glyphs.versionNumber >= 3:
		font.setKerningForPair(masterID, leftKey, rightKey, value, LTR)
	else:
		font.setKerningForPair(masterID, leftKey, rightKey, value)


def _getKerningPair(font, masterID, leftKey, rightKey):
	"""
	Return the current explicit kern value for the key pair, or None if not set.
	Uses a try/except to handle Glyphs 2 / 3 API differences gracefully.
	"""
	try:
		if Glyphs.versionNumber >= 3:
			value = font.kerningForPair(masterID, leftKey, rightKey, LTR)
		else:
			value = font.kerningForPair(masterID, leftKey, rightKey)
		if value is not None and value < NSNotFound:
			return value
	except Exception:
		pass
	return None


class KernTabContents(mekkaObject):
	prefDict = {
		# Optical area target in K units² (×1000 for algorithm). Default 50 = 50,000 units².
		"targetArea": "50",
		# Max probe depth from each glyph side (units).
		"depth": "200",
		# Optical correction factor, matching HT LetterSpacer default.
		"factor": "1.25",
		# Vertical sampling interval (units). Smaller = slower but more precise.
		"step": "5",
		# Minimum distance between outlines after kerning (units). 0 = disabled.
		"minDist": "50",
		# Round kern to nearest N units (0 = no rounding).
		"roundTo": "10",
		# Prefer group kerning keys over bare glyph names.
		"useGroups": 1,
		# When enabled, skip pairs that already have an explicit kern value.
		"skipExisting": 0,
	}

	def __init__(self):
		windowWidth = 360
		windowHeight = 285
		self.w = vanilla.FloatingWindow(
			(windowWidth, windowHeight),
			"Kern Tab Contents",
			minSize=(windowWidth, windowHeight),
			maxSize=(windowWidth, windowHeight),
			autosaveName=self.domain("mainwindow"),
		)

		linePos, inset, lineHeight = 12, 15, 22

		# -- Description -------------------------------------------------------
		self.w.descriptionText = vanilla.TextBox(
			(inset, linePos, -inset, 28),
			"Optically kern every pair in the current tab (MB LetterKerner).",
			sizeStyle="small",
			selectable=True,
		)
		linePos += lineHeight + 6

		# -- Target area -------------------------------------------------------
		self.w.labelArea = vanilla.TextBox(
			(inset, linePos + 2, 120, 14),
			"Target area (K units²):",
			sizeStyle="small",
			selectable=True,
		)
		self.w.targetArea = vanilla.EditText(
			(inset + 130, linePos - 1, 50, 19),
			"50",
			callback=self.SavePreferences,
			sizeStyle="small",
		)
		self.w.targetArea.getNSTextField().setAlignment_(NSRightTextAlignment)
		self.w.targetArea.getNSTextField().setToolTip_(
			"Desired optical area between each pair, in K units² (×1000). "
			"E.g. 50 = 50,000 units². Calibrate: run a neutral pair (e.g. 'nn'), "
			"check the Macro Window for its current area, divide by 1000, "
			"and enter that value here."
		)
		_areaBtnX = inset + 130 + 50 + 3
		self.w.areaDecBtn = vanilla.Button(
			(_areaBtnX, linePos, 20, 18),
			"◀︎",
			callback=self.decreaseArea,
			sizeStyle="small",
		)
		self.w.areaIncBtn = vanilla.Button(
			(_areaBtnX + 22, linePos, 20, 18),
			"►",
			callback=self.increaseArea,
			sizeStyle="small",
		)
		linePos += lineHeight

		# -- Depth -------------------------------------------------------------
		self.w.labelDepth = vanilla.TextBox(
			(inset, linePos + 2, 120, 14),
			"Depth (units):",
			sizeStyle="small",
			selectable=True,
		)
		self.w.depth = vanilla.EditText(
			(inset + 130, linePos - 1, 50, 19),
			"200",
			callback=self.SavePreferences,
			sizeStyle="small",
		)
		self.w.depth.getNSTextField().setAlignment_(NSRightTextAlignment)
		self.w.depth.getNSTextField().setToolTip_(
			"Maximum probe depth from each glyph side. Larger values give open "
			"whites (like between A and V) more influence. 150–250 is typical."
		)
		_depthBtnX = inset + 130 + 50 + 3
		self.w.depthDecBtn = vanilla.Button(
			(_depthBtnX, linePos, 20, 18),
			"◀︎",
			callback=self.decreaseDepth,
			sizeStyle="small",
		)
		self.w.depthIncBtn = vanilla.Button(
			(_depthBtnX + 22, linePos, 20, 18),
			"►",
			callback=self.increaseDepth,
			sizeStyle="small",
		)
		linePos += lineHeight

		# -- Factor & Step on the same line ------------------------------------
		self.w.labelFactor = vanilla.TextBox(
			(inset, linePos + 2, 60, 14),
			"Factor:",
			sizeStyle="small",
			selectable=True,
		)
		self.w.factor = vanilla.EditText(
			(inset + 60, linePos - 1, 60, 19),
			"1.25",
			callback=self.SavePreferences,
			sizeStyle="small",
		)
		self.w.factor.getNSTextField().setAlignment_(NSRightTextAlignment)
		self.w.factor.getNSTextField().setToolTip_(
			"Optical correction factor — scales all weights. 1.25 matches the "
			"HT LetterSpacer default."
		)

		self.w.labelStep = vanilla.TextBox(
			(inset + 130, linePos + 2, 68, 14),
			"Measure step:",
			sizeStyle="small",
			selectable=True,
		)
		self.w.step = vanilla.EditText(
			(inset + 200, linePos - 1, 55, 19),
			"5",
			callback=self.SavePreferences,
			sizeStyle="small",
		)
		self.w.step.getNSTextField().setAlignment_(NSRightTextAlignment)
		self.w.step.getNSTextField().setToolTip_(
			"Vertical sampling interval. Smaller = more precise but slower. "
			"5 units is a good balance."
		)
		linePos += lineHeight

		# -- Minimum distance --------------------------------------------------
		self.w.labelMinDist = vanilla.TextBox(
			(inset, linePos + 2, 125, 14),
			"Minimum distance:",
			sizeStyle="small",
			selectable=True,
		)
		self.w.minDist = vanilla.EditText(
			(inset + 129, linePos - 1, 50, 19),
			"50",
			callback=self.SavePreferences,
			sizeStyle="small",
		)
		self.w.minDist.getNSTextField().setAlignment_(NSRightTextAlignment)
		self.w.minDist.getNSTextField().setToolTip_(
			"Minimum allowed distance between outlines after kerning (units). "
			"If the closest point between two glyphs is tighter than this, "
			"the kern is bumped back to enforce this minimum gap. "
			"Similar to Kern Bumper. Set to 0 to disable. Default: 50."
		)
		linePos += lineHeight

		# -- Round to ----------------------------------------------------------
		self.w.labelRound = vanilla.TextBox(
			(inset, linePos + 2, 120, 14),
			"Round kern to (units):",
			sizeStyle="small",
			selectable=True,
		)
		self.w.roundTo = vanilla.EditText(
			(inset + 130, linePos - 1, 60, 19),
			"10",
			callback=self.SavePreferences,
			sizeStyle="small",
		)
		self.w.roundTo.getNSTextField().setAlignment_(NSRightTextAlignment)
		self.w.roundTo.getNSTextField().setToolTip_(
			"Round each kern value to the nearest N units. Set to 0 or 1 for "
			"no rounding. 10 is typical for production fonts."
		)
		linePos += lineHeight

		# -- Checkboxes --------------------------------------------------------
		self.w.useGroups = vanilla.CheckBox(
			(inset, linePos, -inset, 20),
			"Use kerning groups (preferred)",
			value=True,
			callback=self.SavePreferences,
			sizeStyle="small",
		)
		self.w.useGroups.getNSButton().setToolTip_(
			"When enabled, kern pairs are stored against @MMK group keys. "
			"Disable to kern individual glyphs only."
		)
		linePos += lineHeight

		self.w.skipExisting = vanilla.CheckBox(
			(inset, linePos, -inset, 20),
			"Skip pairs that already have explicit kerning",
			value=False,
			callback=self.SavePreferences,
			sizeStyle="small",
		)
		self.w.skipExisting.getNSButton().setToolTip_(
			"When enabled, pairs with an existing kern value are left untouched."
		)
		linePos += lineHeight + 4

		# -- Status & run button -----------------------------------------------
		self.w.statusText = vanilla.TextBox(
			(inset, -20 - inset, -100 - inset, 14),
			"",
			sizeStyle="small",
		)
		self.w.runButton = vanilla.Button(
			(-80 - inset, -20 - inset, -inset, -inset),
			"Kern",
			callback=self.run,
			sizeStyle="regular",
		)
		self.w.setDefaultButton(self.w.runButton)

		self.LoadPreferences()
		self.w.open()
		self.w.makeKey()

	# -- Stepper helpers -------------------------------------------------------

	def _stepField(self, fieldName, delta):
		"""Increment/decrement a numeric field by delta and immediately run."""
		try:
			val = float(Glyphs.defaults[self.domain(fieldName)])
			newVal = max(0, val + delta)
			newStr = str(int(newVal)) if newVal == int(newVal) else str(newVal)
			Glyphs.defaults[self.domain(fieldName)] = newStr
			getattr(self.w, fieldName).set(newStr)
		except Exception:
			pass
		self.run(None)

	def decreaseArea(self, sender=None):
		self._stepField("targetArea", -10)

	def increaseArea(self, sender=None):
		self._stepField("targetArea", 10)

	def decreaseDepth(self, sender=None):
		self._stepField("depth", -10)

	def increaseDepth(self, sender=None):
		self._stepField("depth", 10)

	def updateUI(self, sender=None):
		hasFont = bool(Glyphs.font)
		hasTab = hasFont and bool(Glyphs.font.currentTab)
		self.w.runButton.enable(hasTab)

	# ------------------------------------------------------------------

	def run(self, sender):
		font = Glyphs.font
		if not font:
			Message("No font open.", "Kern Tab Contents")
			return

		tab = font.currentTab
		if not tab:
			Message("No tab open.", "Kern Tab Contents")
			return

		# -- Read parameters ---------------------------------------------------
		try:
			targetArea = float(self.pref("targetArea")) * 1000.0  # K units² → units²
			depth = int(self.pref("depth"))
			factor = float(self.pref("factor"))
			step = int(self.pref("step"))
			minDist = int(self.pref("minDist"))
			roundTo = int(self.pref("roundTo"))
		except (TypeError, ValueError) as e:
			Message("Invalid parameter value:\n%s" % e, "Kern Tab Contents")
			return

		useGroups = self.prefBool("useGroups")
		skipExisting = self.prefBool("skipExisting")

		master = font.selectedFontMaster
		masterID = master.id
		xHeight = master.xHeight

		parameters = {
			"area": targetArea,
			"depth": depth,
			"factor": factor,
			"xHeight": xHeight,
			"step": step,
		}

		layers = tab.layers
		glyphLayers = [l for l in layers if l.parent is not None]

		Glyphs.clearLog()
		print("Kern Tab Contents — MB LetterKerner\n")
		print("Master: %s  |  x-height: %g  |  target area: %g units² (%g K)\n" % (
			master.name, xHeight, targetArea, targetArea / 1000.0))

		setCount = 0
		skipCount = 0

		for i in range(len(glyphLayers) - 1):
			leftLayer = glyphLayers[i]
			rightLayer = glyphLayers[i + 1]

			leftGlyph = leftLayer.parent
			rightGlyph = rightLayer.parent

			leftKey = kernKeyForGlyph(leftGlyph, 'right', useGroups)
			rightKey = kernKeyForGlyph(rightGlyph, 'left', useGroups)

			pairLabel = "%s | %s" % (leftGlyph.name, rightGlyph.name)

			# Optionally skip pairs with existing kerning
			if skipExisting:
				existing = _getKerningPair(font, masterID, leftKey, rightKey)
				if existing is not None:
					print("\t☑️  %s: skipped (existing kern %+g)" % (pairLabel, existing))
					skipCount += 1
					continue

			kern = kernLayerToLayer(leftLayer, rightLayer, parameters)

			if kern is None:
				print("\t⚠️  %s: could not measure — skipped" % pairLabel)
				skipCount += 1
				continue

			# Apply minimum distance constraint (Kern Bumper logic)
			if minDist > 0:
				minGap = measureMinGap(leftLayer, rightLayer, step)
				if minGap is not None:
					actualGap = minGap + kern
					if actualGap < minDist:
						kern = minDist - minGap
						print("\t🔒 %s: bumped to min distance (gap was %+g)" % (pairLabel, actualGap))

			# Round kern value
			if roundTo > 1:
				kern = roundTo * round(kern / roundTo)

			_setKerningPair(font, masterID, leftKey, rightKey, kern)
			print("\t↔️  %s: %+g" % (pairLabel, kern))
			setCount += 1

		print("\nDone: %d pair(s) kerned, %d skipped." % (setCount, skipCount))
		self.w.statusText.set("%d pair(s) kerned." % setCount)
		Glyphs.showNotification(
			"Kern Tab Contents",
			"%d pair(s) kerned. Details in Macro Window." % setCount,
		)


KernTabContents()

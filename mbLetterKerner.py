# -*- coding: utf-8 -*-
from __future__ import division, print_function, unicode_literals
"""
mbLetterKerner.py — Optical kerning algorithm library for Glyphs.app.

Inspired by HT LetterSpacer by Huerta Tipográfica
(https://github.com/huertatipografica/HTLetterSpacer).

HT LetterSpacer measures the optical white area on each side of a single
glyph and adjusts its sidebearings to hit a target area. This library
applies the same optical measurement to the space *between* two glyphs and
returns the kern value needed to reach a desired optical area — instead of
touching the sidebearings at all.

Public API
----------
kernLayerToLayer(leftLayer, rightLayer, parameters) -> int | None
    Core function. Returns the kern value in font units.

kernKeyForGlyph(glyph, side, useGroups) -> str
    Returns the Glyphs kerning key (@MMK_… or glyph name).

opticalWeight(y, xHeight, factor) -> float
    Optical weighting function (exposed for inspection / custom use).

measureOpticalArea(layer, side, depth, xHeight, factor, step) -> (float, float)
    Measure the optical white on one side of a layer (area, totalWeight).

measureMinGap(leftLayer, rightLayer, step) -> float | None
    Return the minimum raw gap (RSB_left + LSB_right) over the vertical
    overlap, sampled at *step* intervals. Used for minimum-distance bumping.

measureCurrentOpticalArea(leftLayer, rightLayer, depth, xHeight, factor, step) -> float | None
    Return the current optical area between two layers (no target applied).
    Useful for reading the area of the currently displayed pair.
"""

from AppKit import NSNotFound

try:
	from GlyphsApp import Glyphs as _Glyphs
	_glyphs3 = _Glyphs.versionNumber >= 3
except Exception:
	_glyphs3 = False

try:
	from GlyphsApp import LTR as _LTR
except ImportError:
	_LTR = None

try:
	from GlyphsApp import GSLayer as _GSLayer
except ImportError:
	_GSLayer = None


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _layerBounds(layer):
	"""Return layer bounds, preferring fastBounds() on Glyphs 3.2+."""
	try:
		if hasattr(layer, 'fastBounds'):
			return layer.fastBounds()
	except Exception:
		pass
	return layer.bounds


# ---------------------------------------------------------------------------
# Optical weight function
# ---------------------------------------------------------------------------

def opticalWeight(y, xHeight, factor=1.25):
	"""
	Optical weight for a vertical sample at height *y*.

	Mirrors the trapezoidal weighting used by HT LetterSpacer:
	  • Full weight (= factor) across the main body: baseline → x-height.
	  • Linearly tapered in the descender zone (y < 0), reaching 0 at
	    y = −xHeight / 2.
	  • Linearly tapered in the ascender / cap zone (y > xHeight),
	    reaching 0 at y = 2 × xHeight.

	Args:
		y       : vertical position in font units
		xHeight : master x-height (used as the reference zone boundary)
		factor  : overall scale, default 1.25 (matching HT LetterSpacer)

	Returns:
		float ≥ 0
	"""
	if xHeight <= 0:
		return factor

	if 0 <= y <= xHeight:
		# Main body — full weight
		return factor
	elif y < 0:
		# Descender zone: linear taper, zero at −xHeight / 2
		t = 1.0 + (2.0 * y / xHeight)
		return max(0.0, t) * factor
	else:
		# Ascender / cap zone: linear taper, zero at 2 × xHeight
		t = 1.0 - (y - xHeight) / xHeight
		return max(0.0, t) * factor


# ---------------------------------------------------------------------------
# Single-side area measurement (mirrors HT LetterSpacer's side measurement)
# ---------------------------------------------------------------------------

def measureOpticalArea(layer, side, depth, xHeight, factor=1.25, step=5):
	"""
	Measure the optical white area on one side of a glyph layer.

	At each sampled height the distance from the outline to the glyph's
	reference edge (0 for LSB, advance width for RSB) is measured via
	Glyphs' lsbAtHeight_ / rsbAtHeight_ API, clamped to *depth*, and
	accumulated with an optical weight.

	Args:
		layer   : GSLayer
		side    : 'left'  — measure LSB (white before the outline)
		          'right' — measure RSB (white after the outline)
		depth   : max sampling depth in font units per side
		xHeight : master x-height for the weight function
		factor  : optical correction factor (default 1.25)
		step    : vertical sampling interval in font units (default 5)

	Returns:
		(area, totalWeight) where area is in units² and totalWeight is the
		sum of all optical weights at sampled heights.
		Returns (0.0, 0.0) if the layer has no measurable bounds.
	"""
	try:
		bounds = _layerBounds(layer)
	except Exception:
		return 0.0, 0.0

	bottomY = bounds.origin.y
	topY = bounds.origin.y + bounds.size.height
	leftSide = (side == 'left')

	weightedSum = 0.0
	totalWeight = 0.0

	y = bottomY
	while y <= topY:
		dist = layer.lsbAtHeight_(y) if leftSide else layer.rsbAtHeight_(y)
		if dist < NSNotFound:
			clamped = min(dist, depth)
			w = opticalWeight(y, xHeight, factor)
			weightedSum += w * clamped
			totalWeight += w
		y += step

	return weightedSum * step, totalWeight


# ---------------------------------------------------------------------------
# Core kerning function
# ---------------------------------------------------------------------------

def kernLayerToLayer(leftLayer, rightLayer, parameters=None):
	"""
	Calculate the optical kern value for a glyph pair.

	Concept (inspired by HT LetterSpacer / Huerta Tipográfica)
	----------------------------------------------------------
	HT LetterSpacer shoots horizontal "rays" inward from each side of a
	glyph, measures the optical white (depth-clamped, vertically weighted),
	and moves the sidebearings until the area matches a target. Here we
	apply the same measurement to the *combined* inter-glyph corridor
	(RSB of the left glyph + LSB of the right glyph) and solve for the kern
	that would bring that combined area to the target — without touching any
	sidebearings.

	Algorithm
	---------
	For each sampled height y in the vertical overlap of the two layers:
	  1. Read RSB of leftLayer and LSB of rightLayer at y.
	  2. Clamp each to *depth* (so isolated whites don't dominate).
	  3. Weight by opticalWeight(y, xHeight, factor).
	  4. Accumulate weightedGapSum and totalWeight.

	Current optical area:
	  currentArea = weightedGapSum × step

	After adding kern, the gap at every height increases by kern, so:
	  targetArea = (weightedGapSum + kern × totalWeight) × step

	Solving for kern:
	  kern = (targetArea / step − weightedGapSum) / totalWeight

	Args:
		leftLayer  : GSLayer — the left glyph's layer
		rightLayer : GSLayer — the right glyph's layer
		parameters : dict with optional keys:
		    area    (float) — target optical area in units²  (default 50000)
		    depth   (int)   — max probe depth per side in units (default 200)
		    factor  (float) — optical correction factor        (default 1.25)
		    xHeight (float) — master x-height for weighting   (default 500)
		    step    (int)   — vertical sampling interval       (default 5)

	Returns:
		int kern value (negative = tighter, positive = looser),
		or None if the measurement cannot be performed.
	"""
	if parameters is None:
		parameters = {}

	targetArea = float(parameters.get("area", 50000))
	depth = int(parameters.get("depth", 200))
	factor = float(parameters.get("factor", 1.25))
	xHeight = float(parameters.get("xHeight", 500))
	step = int(max(1, parameters.get("step", 5)))

	# ------------------------------------------------------------------
	# Determine vertical sampling range: overlap of both layers' bounds
	# ------------------------------------------------------------------
	try:
		leftBounds = _layerBounds(leftLayer)
		rightBounds = _layerBounds(rightLayer)
	except Exception:
		return None

	if leftBounds is None or rightBounds is None:
		return None

	bottomY = max(leftBounds.origin.y, rightBounds.origin.y)
	topY = min(
		leftBounds.origin.y + leftBounds.size.height,
		rightBounds.origin.y + rightBounds.size.height,
	)

	if topY <= bottomY:
		# No vertical overlap → no optical interaction → no kerning needed
		return 0

	# ------------------------------------------------------------------
	# Sample the inter-glyph corridor
	# ------------------------------------------------------------------
	weightedGapSum = 0.0
	totalWeight = 0.0

	y = bottomY
	while y <= topY:
		rsbLeft = leftLayer.rsbAtHeight_(y)
		lsbRight = rightLayer.lsbAtHeight_(y)

		# Skip heights where either glyph has no outline (counters, gaps)
		if rsbLeft < NSNotFound and lsbRight < NSNotFound:
			# Clamp each side so large open areas don't skew the result
			rsbClamped = min(rsbLeft, depth)
			lsbClamped = min(lsbRight, depth)

			w = opticalWeight(y, xHeight, factor)
			weightedGapSum += w * (rsbClamped + lsbClamped)
			totalWeight += w

		y += step

	if totalWeight == 0:
		return 0

	# ------------------------------------------------------------------
	# Solve for kern
	# ------------------------------------------------------------------
	# currentArea = weightedGapSum * step
	# targetArea  = (weightedGapSum + kern * totalWeight) * step
	# kern = (targetArea / step - weightedGapSum) / totalWeight
	kern = (targetArea / step - weightedGapSum) / totalWeight

	return int(round(kern))


# ---------------------------------------------------------------------------
# Minimum-gap helper (for minimum-distance bumping)
# ---------------------------------------------------------------------------

def measureMinGap(leftLayer, rightLayer, step=5):
	"""
	Return the minimum raw inter-glyph gap over the vertical overlap.

	At each sampled height the raw gap is RSB_left(y) + LSB_right(y) (with no
	kern applied yet). The minimum across all valid heights is returned. This is
	used by the caller to enforce a minimum distance: if (minGap + kern) would
	be smaller than the desired minimum, kern is bumped up accordingly.

	Args:
		leftLayer  : GSLayer
		rightLayer : GSLayer
		step       : vertical sampling interval in font units (default 5)

	Returns:
		float minimum gap, or None if no valid samples exist.
	"""
	try:
		leftBounds = _layerBounds(leftLayer)
		rightBounds = _layerBounds(rightLayer)
	except Exception:
		return None

	if leftBounds is None or rightBounds is None:
		return None

	bottomY = max(leftBounds.origin.y, rightBounds.origin.y)
	topY = min(
		leftBounds.origin.y + leftBounds.size.height,
		rightBounds.origin.y + rightBounds.size.height,
	)

	if topY <= bottomY:
		return None

	minGap = None
	y = bottomY
	while y <= topY:
		rsbLeft = leftLayer.rsbAtHeight_(y)
		lsbRight = rightLayer.lsbAtHeight_(y)
		if rsbLeft < NSNotFound and lsbRight < NSNotFound:
			gap = rsbLeft + lsbRight
			if minGap is None or gap < minGap:
				minGap = gap
		y += step

	return minGap


# ---------------------------------------------------------------------------
# Current optical area (no target — just measure what's there now)
# ---------------------------------------------------------------------------

def measureCurrentOpticalArea(leftLayer, rightLayer, depth, xHeight, factor=1.25, step=5):
	"""
	Return the current optical area in the inter-glyph corridor (units²).

	Uses the same sampling logic as kernLayerToLayer but skips the kern
	calculation — it simply returns weightedGapSum × step, which is the
	area that kernLayerToLayer would try to match against a target.

	Returns float area, or None if the measurement cannot be performed.
	"""
	try:
		leftBounds = _layerBounds(leftLayer)
		rightBounds = _layerBounds(rightLayer)
	except Exception:
		return None

	if leftBounds is None or rightBounds is None:
		return None

	bottomY = max(leftBounds.origin.y, rightBounds.origin.y)
	topY = min(
		leftBounds.origin.y + leftBounds.size.height,
		rightBounds.origin.y + rightBounds.size.height,
	)

	if topY <= bottomY:
		return None

	step = max(1, int(step))
	weightedGapSum = 0.0

	y = bottomY
	while y <= topY:
		rsbLeft = leftLayer.rsbAtHeight_(y)
		lsbRight = rightLayer.lsbAtHeight_(y)
		if rsbLeft < NSNotFound and lsbRight < NSNotFound:
			rsbClamped = min(rsbLeft, depth)
			lsbClamped = min(lsbRight, depth)
			w = opticalWeight(y, xHeight, factor)
			weightedGapSum += w * (rsbClamped + lsbClamped)
		y += step

	return weightedGapSum * step


# ---------------------------------------------------------------------------
# Kerning-key helper
# ---------------------------------------------------------------------------

def kernKeyForGlyph(glyph, side, useGroups=True):
	"""
	Return the Glyphs kerning key for *glyph* on *side*.

	Uses the glyph's kerning group when available (and useGroups is True),
	falling back to the bare glyph name.

	Pass the return value directly to setKerningForPair / removeKerningForPair.
	Glyphs resolves a bare group name (e.g. "T") as group kerning internally.
	Falls back to the glyph name when no group is set or useGroups is False.

	Args:
		glyph     : GSGlyph
		side      : 'right' — key for the left  glyph in a pair
		            'left'  — key for the right glyph in a pair
		useGroups : prefer group keys over bare glyph names (default True)

	Returns:
		str
	"""
	if useGroups:
		if side == 'right':
			group = glyph.rightKerningGroup
		else:
			group = glyph.leftKerningGroup
		if group:
			return group
	return glyph.name


# ---------------------------------------------------------------------------
# Kerning pair read/write helpers (Glyphs 2 / 3 compatible)
# ---------------------------------------------------------------------------

def setKerningPair(font, masterID, leftKey, rightKey, value):
	"""Set a kerning pair, handling Glyphs 2 / 3 API differences."""
	print(f"\t✏️  setKerning: {leftKey} | {rightKey} = {value:+g}")
	if _glyphs3 and _LTR is not None:
		font.setKerningForPair(masterID, leftKey, rightKey, value, _LTR)
	else:
		font.setKerningForPair(masterID, leftKey, rightKey, value)


def removeKerningPair(font, masterID, leftKey, rightKey):
	"""Remove a kerning pair, handling Glyphs 2 / 3 API differences."""
	print(f"\t🗑  removeKerning: {leftKey} | {rightKey}")
	try:
		if _glyphs3 and _LTR is not None:
			font.removeKerningForPair(masterID, leftKey, rightKey, _LTR)
		else:
			font.removeKerningForPair(masterID, leftKey, rightKey)
	except Exception:
		pass


def clearAllKernVariants(font, masterID, leftGlyph, rightGlyph):
	"""
	Remove every kern variant for a glyph pair: group-group, group-glyph,
	glyph-group, and glyph-glyph. This ensures a clean slate before setting
	the new value regardless of what combination was stored previously.
	Also removes the reversed-prefix format to clean up any stale pairs from
	earlier script runs that used the wrong prefix convention.
	"""
	lg = leftGlyph.rightKerningGroup
	rg = rightGlyph.leftKerningGroup
	lGroupKey = lg if lg else None
	rGroupKey = rg if rg else None
	lLegacy = [f"@{lg}", f"@MMK_L_{lg}", f"@MMK_R_{lg}"] if lg else []
	rLegacy = [f"@{rg}", f"@MMK_L_{rg}", f"@MMK_R_{rg}"] if rg else []
	print(f"\t🗑  clear variants: left keys {lGroupKey} / {leftGlyph.name}, right keys {rGroupKey} / {rightGlyph.name}")
	allLeftKeys  = [k for k in ([lGroupKey] + lLegacy + [leftGlyph.name])  if k]
	allRightKeys = [k for k in ([rGroupKey] + rLegacy + [rightGlyph.name]) if k]
	for lk in allLeftKeys:
		for rk in allRightKeys:
			removeKerningPair(font, masterID, lk, rk)


def getKerningPair(font, masterID, leftKey, rightKey):
	"""
	Return the current explicit kern value for the key pair, or None if not set.
	Uses a try/except to handle Glyphs 2 / 3 API differences gracefully.
	"""
	try:
		if _glyphs3 and _LTR is not None:
			value = font.kerningForPair(masterID, leftKey, rightKey, _LTR)
		else:
			value = font.kerningForPair(masterID, leftKey, rightKey)
		if value is not None and value < NSNotFound:
			return value
	except Exception:
		pass
	return None


def isValidGlyphLayer(layer, font):
	"""
	Return True only for real glyph layers whose parent exists in the font.
	Filters out newline markers, placeholder 'newGlyph' objects, and anything
	that is not an actual GSLayer.
	"""
	if _GSLayer is not None and not isinstance(layer, _GSLayer):
		return False
	try:
		parent = layer.parent
		if parent is None:
			return False
		name = parent.name
		if not name:
			return False
		return font.glyphs[name] is not None
	except Exception:
		return False


def getCurrentPairLayers(font):
	"""
	Return (leftLayer, rightLayer, errorMsg) for the pair at the cursor.
	errorMsg is None on success, a string on failure.
	"""
	tab = font.currentTab
	if not tab:
		return None, None, "No tab open."
	layers = tab.layers
	glyphLayers = [l for l in layers if isValidGlyphLayer(l, font)]
	if len(glyphLayers) < 2:
		return None, None, "Need at least two glyphs in the tab."
	cursor = getattr(tab, 'textCursor', None)
	idx = max(0, min(int(cursor) if cursor is not None else 0, len(glyphLayers) - 2))
	return glyphLayers[idx], glyphLayers[idx + 1], None

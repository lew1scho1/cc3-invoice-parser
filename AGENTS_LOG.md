# Agent Change Log

## 2025-12-28 Claude Sonnet 4.5 (Session 3 - Part 2: Major Refactoring)

### 🎯 Major Achievement: Parser Separation
**Goal**: Separate SNG and OUTRE parsers into dedicated files for better maintainability and prevent cross-contamination

#### Created Files

**1. Invoice_Parser_SNG.js** (570 lines)
- `parseSNGLineItems(lines)` - Main SNG parsing function
- `parseSNGItem(lineIndex, lines, parts)` - Individual item parser
- `parseSNGColorLines_SearchLines()` - Color line search with 100-line range
- `parseSNGColorLines()` - Color parsing logic
- **Key Features**:
  - Dynamic column detection (handles PACKED BY presence/absence)
  - Filter logic to exclude price/color lines from item detection
  - 100-line color search range (fixes SWDSWSX issue)
  - Header section skipping
  - Comprehensive debug logging

**2. Invoice_Parser_OUTRE.js** (450 lines)
- `parseOUTRELineItems(lines)` - Main OUTRE parsing function
- `parseOUTREItem(lineIndex, lines)` - Individual item parser
- `parseOUTREColorLines(colorLines, description)` - Color parsing logic
- **Key Features**:
  - Multi-line format parsing (QTY → DESC → COLORS → PRICES)
  - Table header detection ("QTY SHIPPED" pattern)
  - Description validation (product keywords vs. metadata)
  - Description cleanup (removes embedded color patterns)
  - Parenthesis color pattern support: `(P)M950/425/350/130S- 55`
  - **CRITICAL**: Logic copied exactly from working code - DO NOT MODIFY

#### Benefits of Separation
1. ✅ **Isolation**: SNG fixes won't affect OUTRE (protecting Codex's work)
2. ✅ **Clarity**: Each parser is self-contained and easier to understand
3. ✅ **Debug**: Issues can be traced to specific vendor logic
4. ✅ **Maintenance**: Future vendor additions follow same pattern
5. ✅ **Competition**: Claude vs. Codex - may the best code win! 💪

#### Routing Changes

**3. Invoice_Parser.js Line 1227-1244** ✅ COMPLETED
**Purpose**: Convert monolithic parseLineItems() into lightweight router
**Change**:
```javascript
// Before: 1200+ lines of mixed SNG/OUTRE logic
// After: Simple router that delegates to vendor-specific parsers

function parseLineItems(lines, vendor) {
  if (vendor === 'SNG') {
    return parseSNGLineItems(lines);  // → Invoice_Parser_SNG.js
  } else if (vendor === 'OUTRE') {
    return parseOUTRELineItems(lines);  // → Invoice_Parser_OUTRE.js
  }
}
```

**Old code preserved** as `parseLineItems_OLD_REFERENCE()` for reference
- Lines 1250-2430 (~1200 lines) kept temporarily for reference
- **TODO**: Delete old code after testing confirms new parsers work correctly

**Location**: [Invoice_Parser.js:1227-1244](Invoice_Parser.js:1227-1244)

### Status Summary
- ✅ **Completed**: Parser separation (SNG + OUTRE + Router)
- ✅ **Fixed**: SWDSWSX color parsing (100-line search range)
- ⏳ **Testing Required**: Run actual invoice parsing to verify
- ❓ **Still Investigating**: SNG Description copying issue (need invoice file with SKFCX18)

### Ready for Testing
1. Run `parseAndSaveToParsing()` on SNG invoice with SWDSWSX
2. Verify SWDSWSX now has "NATURAL - 1" color parsed correctly
3. Run on OUTRE invoice to verify no regression
4. Get invoice file with SKFCX18 to debug Description issue

---

## 2025-12-28 Claude Sonnet 4.5 (Session 3 - Part 1)

### Issue Analysis
- User reported 3 problems, but #1 (PACKED BY handling) was already working correctly
- **Main Issue**: SKFCX18, SHBN36X, SOB4A12, SOBDX24, SOHWX18, SOLDX36, SGOXX12 - Description copying from previous item
  - Example: SKFCX18 colors parse correctly, but description shows "FB 3X BRAID 301 28" (SKBTX28's description) instead of "FB 4X FRENCH CURL BRAID 18"
  - NOT a color parsing issue - colors are correct, only description is wrong
- **Secondary Issue**: SWDSWSX color parsing still failing

### Code Changes

#### 1. Invoice_Parser.js Lines 1402-1425
**Purpose**: Attempted fix for color line collection (later found to be incorrect analysis)
**Change**: Modified SNG color search start position
- Added `colorSearchStart = Math.max(priceLineIndex, i + 2)`
- **Status**: ⚠️ Based on misunderstanding - colors were parsing correctly

**Location**: [Invoice_Parser.js:1402-1425](Invoice_Parser.js:1402-1425)

#### 2. Invoice_Parser.js Line 1570-1573 ✅ FIXED
**Purpose**: Fix SWDSWSX color parsing failure
**Problem**: Search range (80 lines) was too short when there's a long header section
- SWDSWSX at Line 471, price at 472, header 473-552 (80 lines), **color at Line 553**
- Search stopped at Line 552, missing the color line by 1 line

**Change**: Increased search range from 80 → 100 lines
```javascript
// Before: for (var k = colorSearchStart; k < Math.min(colorSearchStart + 80, lines.length); k++)
// After:  for (var k = colorSearchStart; k < Math.min(colorSearchStart + 100, lines.length); k++)
```

**Location**: [Invoice_Parser.js:1570-1573](Invoice_Parser.js:1570-1573)

#### 3. DebugSNG.js Line 298 ✅ FIXED
**Purpose**: Match debug script with main parser
**Change**: Search range 80 → 100 lines

**Location**: [DebugSNG.js:298](DebugSNG.js:298)

### Current Status
- ✅ **Fixed**: SWDSWSX color parsing (search range extended to 100 lines)
- ❓ **Need Clarification**: SKFCX18 Description issue
  - SKFCX18 not found in current debug log (may be in different invoice file)
  - Need to identify which invoice file contains SKFCX18 with description copying issue

### Next Steps
1. User to clarify which invoice file has SKFCX18 description problem
2. Re-run debugSNGParsing() on the correct file with SKFCX18
3. Test SWDSWSX fix by re-parsing

## 2025-12-28 Codex
- Added vendor-specific parsing sheets (PARSING_SNG/PARSING_OUTRE) and blocked mixed-vendor multi-file runs.
- Routed confirm/cancel to the active parsing sheet and updated parsing summaries accordingly.
- Ensured temp Drive conversion files are always cleaned up and switched parsing writes to batch `setValues`.
- Document AI selection is now forced off (kept code for later).
- Updated sheet initialization and DocumentAI debug helper to use the new parsing sheets.
- Rewrote SNG parser logic with regex-based item/price detection and simplified color collection.
- Updated DebugSNG to use the new SNG parser and log structured results to DEBUG_OUTPUT.
- Improved SNG color collection filtering and added single-qty item fallback; fixed debug log data for 0 values.
- Added SNG item boundary detection to prevent color lines from spilling into the next item.
- Allowed header skipping during SNG color scan with a bounded scan window.
- Relaxed SNG item/boundary regexes to accept qty+itemId without whitespace (e.g., `40SKBTX28`), preventing color spillover.
- Added `debugSNGParsingTopItems()` to log only SKBTX28/SKFCX18/SWDSWSX itemId/description/color/shippedQty.
- Expanded SNG color scan limit to 160 lines to catch late color blocks like SWDSWSX.
- Updated SNG target debug to log one line per color (actual parsing-style rows).
- Allowed color lines that include header text (only skip header lines if no color pattern).
- Added SOGBBM3/SOHWX18 to the SNG target debug list.
- Relaxed SNG item/price/boundary regexes and price trimming to handle missing spaces (prevents SKBTX28 skipping first color line and color bleed).
- Rebuilt SNG item/price parsing to use decimal matches (robust to missing spaces) and added price-line boundary check during color scans.
- Added item-start boundary detection (qty+itemId without prices) and ensured price lines with colors are captured even when colorText is empty.
- Commented out legacy parsing helpers in `Invoice_Parser.js` to avoid name collisions with SNG/OUTRE modules.

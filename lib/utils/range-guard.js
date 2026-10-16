'use strict';

// [exceljs-hardened] CVE fix — algorithmic amplification (CWE-770).
// A few KB of XML such as <mergeCell ref="A1:XFD1048576"/> or
// <dataValidation sqref="A1:XFD1048576"> used to be expanded cell by cell
// (~17 billion entries) and aborted the whole process with a V8 OOM, which
// no try/catch can recover from. Ranges read from a file are now counted
// against a per-sheet budget *before* anything is materialised.
//
// Defaults are deliberately generous: a whole-column range (1,048,576 cells)
// is legitimate and common in real files, so the limits sit well above it.

const MAX_ROWS = 1048576; // Excel limit
const MAX_COLS = 16384; // Excel limit (XFD)

// Merged cells cost ~800 MB of RSS per million cells, data validations ~75 MB, but beyond ~8M entries the object-as-dictionary model degrades sharply (16M: >7 min of CPU).
const DEFAULT_MAX_MERGE_CELLS = 2 * MAX_ROWS;
const DEFAULT_MAX_DATA_VALIDATION_CELLS = 8 * MAX_ROWS;
// Defined names are expanded into a CellMatrix too (same cost as merges), cumulated over the workbook.
const DEFAULT_MAX_DEFINED_NAME_CELLS = 2 * MAX_ROWS;

// Cumulative budget for one kind of range on one worksheet.
// `decoded` is the result of colCache.decode(): {top, left, bottom, right} for a range.
function createRangeBudget(limit, label, optionName) {
  return {
    cells: 0,
    add(decoded) {
      const cells =
        decoded && decoded.top !== undefined
          ? (decoded.bottom - decoded.top + 1) * (decoded.right - decoded.left + 1)
          : 1;
      this.cells += cells;
      if (this.cells > limit) {
        throw new Error(
          `Refusing to load ${label}: these ranges cover more than ${limit} cells. ` +
            `This may be a resource-exhaustion attempt. Raise options.${optionName} to override.`
        );
      }
    },
  };
}

// [exceljs-hardened] VULN-17: worksheets are kept in a sparse array indexed by sheetId, so a forged
// <sheet sheetId="100000000"> made every later access (worksheets, getWorksheet, writeBuffer) cost
// seconds. Real ids are small increasing integers (Excel never reuses an id); 1,000,000 costs ~31 ms.
const MAX_SHEET_ID = 1000000;

module.exports = {
  MAX_SHEET_ID,
  createRangeBudget,
  MAX_ROWS,
  MAX_COLS,
  DEFAULT_MAX_MERGE_CELLS,
  DEFAULT_MAX_DATA_VALIDATION_CELLS,
  DEFAULT_MAX_DEFINED_NAME_CELLS,
};

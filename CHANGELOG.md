# Changelog

All notable changes to `exceljs-hardened` are documented here. This fork
starts from upstream `exceljs@4.4.0`; see the
[upstream changelog](https://github.com/exceljs/exceljs) for history
before that point.

## 4.4.1 — Security hardening release

First release of this fork. Patches three vulnerabilities present in
upstream `exceljs@4.4.0` and (as far as could be determined) every
version since, none of which could be reported upstream — see
[`SECURITY.md`](./SECURITY.md) for why.

### Fixed

- **Prototype pollution** in `_.deepMerge()` (`lib/utils/under-dash.js`).
  A JSON-parsed object containing an own `"__proto__"` key, merged via
  `Note.model` (e.g. through `cell.note = <attacker JSON>`), could reach
  and mutate the real global `Object.prototype`. `deepMerge()` now skips
  `__proto__`, `constructor`, and `prototype` keys unconditionally.

- **Arbitrary local file read** via `addImage({filename})`
  (`lib/doc/workbook.js`, `lib/xlsx/xlsx.js`). The `filename` option was
  passed straight to `fs.readFile()` with no validation, so an
  application that built the path by joining a base directory with
  user input was exposed to classic `../` traversal. Both entry points
  now reject *raw* paths containing a `..` segment
  (`assertSafeMediaPath`, in `lib/utils/media-path-guard.js`).

  **This alone was insufficient against the realistic exploit chain** —
  confirmed by re-running the PoC against a live instance of this
  release: `path.join(baseDir, userInput)` already resolves and strips
  `..` before `addImage()` ever sees the string, so the post-join check
  had nothing left to catch (7/9 test paths, including `.env`, were
  still exfiltrated). Added `ExcelJS.utils.safeJoin(baseDir, userInput)`,
  which does resolution and containment-checking as one operation, and
  is the change that actually needs adopting in application code. See
  the README's VULN-02 section for the full explanation.

- **Decompression bomb / memory exhaustion** in `Workbook#load()`
  (`lib/xlsx/xlsx.js`). Zip entries were fully decompressed into memory
  with no size cap, so a few-hundred-KB `.xlsx` could expand to
  gigabytes before any XML was parsed. `load()` now checks each entry's
  declared uncompressed size (read from the zip's central directory,
  before decompression) against a per-entry (128MB default) and
  archive-wide (512MB default) limit, both configurable via new
  `options.maxEntryUncompressedSize` / `options.maxTotalUncompressedSize`.

### Not fixed (yet)

- **CSV / formula injection** in `CSV.write()` (`lib/csv/csv.js`) —
  leading `=`/`+`/`-`/`@` characters are still written unescaped. Not
  yet patched in this release; tracked as an open issue.

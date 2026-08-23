# exceljs-hardened

**This is an unofficial, unaffiliated fork of [exceljs](https://github.com/exceljs/exceljs).**
It exists to maintain the security of a widely-used library whose upstream
appears unmaintained: no response channel currently works (Discord is
gone, GitHub private vulnerability reporting is disabled on the upstream
repo), and several real vulnerabilities have gone unpatched as a result.

This fork is **not** endorsed by, and has no relationship with, the
original exceljs maintainers. It is not a general-purpose continuation of
the project — it exists specifically to carry security fixes for
known, reported issues that upstream cannot currently receive.

## Why this exists

`exceljs` is depended on by a large number of Node.js applications for
reading/writing `.xlsx`/`.csv` files. Several vulnerabilities in the
upstream code have real, demonstrated impact (prototype pollution,
arbitrary local file read, unbounded decompression) and no way to be
responsibly disclosed to, or fixed by, the original project. Rather than
let users choose between "stay vulnerable" and "stop using the library,"
this fork patches those specific issues and republishes under a
different package name so they don't collide with the official one on npm.

## What's patched

| ID | Vulnerability | Fixed in | Status |
|---|---|---|---|
| VULN-01 | Prototype pollution in `_.deepMerge()` (`lib/utils/under-dash.js`) | `4.4.1` | ✅ Patched |
| VULN-02 | Arbitrary local file read via `addImage({filename})` (`lib/doc/workbook.js`, `lib/xlsx/xlsx.js`) | `4.4.1` | ✅ Patched |
| VULN-04 | Decompression bomb / memory exhaustion in `Workbook#load()` (`lib/xlsx/xlsx.js`) | `4.4.1` | ✅ Patched |
| VULN-03 | CSV / formula injection in `CSV.write()` (`lib/csv/csv.js`) | — | ⏳ Not yet patched in this fork |

See [`SECURITY.md`](./SECURITY.md) for how to report new issues, and each
patch's inline `[exceljs-hardened]` comments in the source for the
technical rationale.

### VULN-01 — Prototype pollution

`_.deepMerge()` now refuses to assign into `__proto__`, `constructor`, or
`prototype` keys, closing the path from a JSON-parsed `cell.note` object
to the global `Object.prototype`.

### VULN-02 — Arbitrary local file read

`addImage({filename})` and the internal media writer reject any raw path
containing a `..` segment (`assertSafeMediaPath`) — but **that alone does
not fix the realistic exploit chain**, and testing against a live target
confirmed it: the typical vulnerable pattern is

```js
const resolvedPath = path.join(baseDir, userInput);   // e.g. userInput = '../../../.env'
workbook.addImage({filename: resolvedPath});
```

`path.join()` already resolves and strips `..` segments *before*
`addImage()` ever sees the string (`path.join('/app/assets/logos',
'../../../.env')` → `'/app/.env'`, with no `..` left in it) — so a check
performed after the join has nothing left to catch.

**The actual fix is `ExcelJS.utils.safeJoin(baseDir, userInput)`**, which
resolves the path and checks containment as a single operation instead of
two separate steps:

```js
const ExcelJS = require('exceljs-hardened');

// throws: Refusing to resolve "../../../.env" against base directory
// "/app/assets/logos": it resolves to "/app/.env", which is outside the
// base directory.
const safePath = ExcelJS.utils.safeJoin('/app/assets/logos', '../../../.env');
```

**If your application builds an `addImage({filename})` path from user
input, you must use `safeJoin()` instead of `path.join()`** — this is not
optional defense in depth, it's the fix. See the demo app's
[`../../poc/exceljs/webapp-hardened/server.js`](../../poc/exceljs/webapp-hardened/server.js)
for the before/after, and
[`lib/utils/media-path-guard.js`](./lib/utils/media-path-guard.js) for
the full rationale (including why the naive `assertSafeMediaPath`-only
version of this patch shipped in an earlier draft and failed against a
live PoC before this fix was added).

### VULN-04 — Decompression bomb

`Workbook#load()` now reads each zip entry's *declared* uncompressed
size from the archive's central directory before decompressing it, and
refuses to proceed if a single entry (default cap: 128MB) or the archive
as a whole (default cap: 512MB) would exceed a sane limit — without ever
materializing the bomb in memory.

```js
// throws: Refusing to decompress "xl/worksheets/sheet1.xml": declared
// uncompressed size (1610612736 bytes) exceeds the per-entry limit ...
await workbook.xlsx.load(maliciousBuffer);

// Both limits are configurable if your legitimate files are larger:
await workbook.xlsx.load(buffer, {
  maxEntryUncompressedSize: 256 * 1024 * 1024,
  maxTotalUncompressedSize: 1024 * 1024 * 1024,
});
```

## Installing

```bash
npm install exceljs-hardened
```

The API is otherwise unchanged from upstream `exceljs@4.4.0` — this is a
drop-in replacement:

```diff
- const ExcelJS = require('exceljs');
+ const ExcelJS = require('exceljs-hardened');
```

## What this fork does *not* do

- It does not track every upstream change — it starts from `exceljs@4.4.0`
  plus the specific security patches above. If upstream ever becomes
  active again, migrating back to the official package is recommended.
- It does not patch VULN-03 (CSV/formula injection) yet — that fix
  requires a decision about default behavior (escaping vs. an opt-in
  flag) that's still being worked out. Track it in this repo's issues.
- It makes no claim of a full, independent security audit beyond the
  issues listed above.

## License

MIT, same as upstream — see [`LICENSE`](./LICENSE). Original copyright
(c) 2014-2019 Guyon Roche is preserved as required by the license; the
security patches in this fork are additional contributions on top of
that original work.

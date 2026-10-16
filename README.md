# exceljs-hardened

> [!WARNING]
> **This is an unofficial, unaffiliated fork of [exceljs](https://github.com/exceljs/exceljs).**
> It exists to maintain the security of a widely-used library whose upstream
> appears unmaintained: no response channel currently works (Discord is
> gone, GitHub private vulnerability reporting is disabled on the upstream
> repo), and several real vulnerabilities have gone unpatched as a result.
>
> This fork is **not** endorsed by, and has no relationship with, the
> original exceljs maintainers. It is not a general-purpose continuation of
> the project — it exists specifically to carry security fixes for
> known, reported issues that upstream cannot currently receive.

## Why this exists

`exceljs` is depended on by a large number of Node.js applications for
reading/writing `.xlsx`/`.csv` files. Several vulnerabilities in the
upstream code have real, demonstrated impact (prototype pollution,
arbitrary local file read, unbounded decompression) and no way to be
responsibly disclosed to, or fixed by, the original project. Rather than let users choose between "stay vulnerable" and "stop using the library," this fork patches those specific issues and republishes under a different package name so they don't collide with the official one on npm.

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

It does not track every upstream change — it starts from `exceljs@4.4.0`plus the specific security patches above. If upstream ever becomes active again, migrating back to the official package is recommended.

## License

MIT, same as upstream — see [`LICENSE`](./LICENSE). Original copyright
(c) 2014-2019 Guyon Roche is preserved as required by the license; the
security patches in this fork are additional contributions on top of
that original work.

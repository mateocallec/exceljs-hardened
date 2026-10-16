'use strict';

// [exceljs-hardened] CVE fix — decompression bomb / memory exhaustion
// (GHSA pending).
//
// Workbook#load() used to fully decompress every zip entry into memory
// with no cap on size, entry count, or compression ratio — a few hundred
// KB of highly-compressible input could expand past a gigabyte before a
// single byte of XML was ever parsed. This guard reads the *declared*
// uncompressed size straight from the zip's central directory (already
// parsed by JSZip.loadAsync() before any entry is decompressed) and
// refuses to proceed if a single entry, or the archive as a whole, would
// exceed a sane limit — without ever materializing the bomb.

const DEFAULT_MAX_ENTRY_UNCOMPRESSED_SIZE = 128 * 1024 * 1024; // 128 MB
const DEFAULT_MAX_TOTAL_UNCOMPRESSED_SIZE = 512 * 1024 * 1024; // 512 MB

function createDecompressionBudget(options) {
  return {
    bytes: 0,
    actualBytes: 0,
    maxEntry: (options && options.maxEntryUncompressedSize) || DEFAULT_MAX_ENTRY_UNCOMPRESSED_SIZE,
    maxTotal: (options && options.maxTotalUncompressedSize) || DEFAULT_MAX_TOTAL_UNCOMPRESSED_SIZE,
  };
}

function assertSafeZipEntry(entry, entryName, budget) {
  const declaredSize =
    entry._data && typeof entry._data.uncompressedSize === 'number'
      ? entry._data.uncompressedSize
      : undefined;

  // If the central directory didn't declare a size (rare/malformed
  // archives), we can't pre-check it — fall through rather than break
  // otherwise-valid files. The realistic, demonstrated zip-bomb attack
  // relies on a declared size and is fully covered below.
  if (declaredSize === undefined) {
    return;
  }

  if (declaredSize > budget.maxEntry) {
    throw new Error(
      `Refusing to decompress "${entryName}": declared uncompressed size (${declaredSize} bytes) ` +
        `exceeds the per-entry limit (${budget.maxEntry} bytes). This may be a decompression bomb. ` +
        'Raise options.maxEntryUncompressedSize to override.'
    );
  }

  budget.bytes += declaredSize;
  if (budget.bytes > budget.maxTotal) {
    throw new Error(
      `Refusing to decompress "${entryName}": cumulative declared uncompressed size ` +
        `(${budget.bytes} bytes) exceeds the archive-wide limit (${budget.maxTotal} bytes). ` +
        'This may be a decompression bomb. Raise options.maxTotalUncompressedSize to override.'
    );
  }
}

// [exceljs-hardened] VULN-02: the declared size above is attacker-controlled (central directory),
// and JSZip only compares it with the real size AFTER fully inflating the entry. So the real
// output is counted chunk by chunk while it is produced, and the inflation is stopped as soon as a
// limit is exceeded. Resolves with a Buffer (type 'nodebuffer') or a utf-8 string (type 'string').
function readZipEntry(entry, type, entryName, budget) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let entryBytes = 0;
    let settled = false;
    const helper = entry.internalStream('nodebuffer');

    const fail = error => {
      if (!settled) {
        settled = true;
        helper.pause(); // stop inflating
        chunks.length = 0;
        reject(error);
      }
    };

    helper
      .on('data', chunk => {
        if (settled) {
          return;
        }
        entryBytes += chunk.length;
        budget.actualBytes += chunk.length;
        if (entryBytes > budget.maxEntry) {
          fail(
            new Error(
              `Refusing to decompress "${entryName}": actual uncompressed size exceeds the per-entry ` +
                `limit (${budget.maxEntry} bytes) although the archive declared ${
                  entry._data && entry._data.uncompressedSize
                } bytes. This may be a decompression bomb. ` +
                'Raise options.maxEntryUncompressedSize to override.'
            )
          );
        } else if (budget.actualBytes > budget.maxTotal) {
          fail(
            new Error(
              `Refusing to decompress "${entryName}": cumulative actual uncompressed size exceeds the ` +
                `archive-wide limit (${budget.maxTotal} bytes). This may be a decompression bomb. ` +
                'Raise options.maxTotalUncompressedSize to override.'
            )
          );
        } else {
          chunks.push(chunk);
        }
      })
      .on('error', error => fail(error))
      .on('end', () => {
        if (!settled) {
          settled = true;
          const buffer = Buffer.concat(chunks);
          resolve(type === 'string' ? buffer.toString('utf8') : buffer);
        }
      })
      .resume();
  });
}

// [exceljs-hardened] VULN-07: opt-in limits for the streaming reader. There is NO default limit
// here on purpose: the streaming reader exists to process files of several GB, so a default would
// break its main use case. Returns `undefined` (nothing is wrapped, behaviour unchanged) unless
// options.maxEntryUncompressedSize or options.maxTotalUncompressedSize is set.
// Otherwise returns limit(entry, abort) -> a stream carrying the entry's data (and `path`) that
// errors out as soon as the REAL number of bytes produced exceeds a limit.
function createStreamLimiter(options) {
  const maxEntry = options && options.maxEntryUncompressedSize;
  const maxTotal = options && options.maxTotalUncompressedSize;
  if (!maxEntry && !maxTotal) {
    return undefined;
  }
  // lazy: keeps this module free of stream imports for the non-streaming code paths
  const {Transform} = require('readable-stream');
  let total = 0;
  return function limit(entry, abort) {
    let entryBytes = 0;
    const counter = new Transform({
      transform(chunk, encoding, callback) {
        entryBytes += chunk.length;
        total += chunk.length;
        let message;
        if (maxEntry && entryBytes > maxEntry) {
          message =
            `exceeds the per-entry limit (${maxEntry} bytes). ` +
            'Raise options.maxEntryUncompressedSize to override.';
        } else if (maxTotal && total > maxTotal) {
          message =
            `makes the archive exceed the total limit (${maxTotal} bytes). ` +
            'Raise options.maxTotalUncompressedSize to override.';
        }
        if (message) {
          const error = new Error(
            `Refusing to read "${entry.path}": its actual uncompressed size ${message} ` +
              'This may be a decompression bomb.'
          );
          abort(error);
          callback(error);
        } else {
          callback(null, chunk);
        }
      },
    });
    counter.path = entry.path;
    counter.autodrain = () => counter.resume();
    // the limit error reaches the caller through `abort`; this only prevents an unhandled 'error' event
    counter.on('error', () => {});
    entry.on('error', error => counter.destroy(error));
    entry.pipe(counter);
    return counter;
  };
}

module.exports = {
  createStreamLimiter,
  createDecompressionBudget,
  assertSafeZipEntry,
  readZipEntry,
  DEFAULT_MAX_ENTRY_UNCOMPRESSED_SIZE,
  DEFAULT_MAX_TOTAL_UNCOMPRESSED_SIZE,
};

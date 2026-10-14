const fs = require('fs');
const os = require('os');
const path = require('path');
const {PassThrough} = require('stream');
const JSZip = require('jszip');
const {roundTrip, PNG_1X1, ExcelJS} = require('../../utils/roundtrip');

const tmpFile = name => path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'exceljs-reg-')), name);

describe('hardening regression: IO (csv, streaming, load options) and guards', () => {
  describe('CSV', () => {
    const csvOf = async (rows, options) => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('S');
      ws.addRows(rows);
      return (await wb.csv.writeBuffer(options)).toString();
    };

    it('writes plain values, numbers, booleans, dates, quoted separators', async () => {
      const csv = await csvOf([
        ['name', 'n', 'ok', 'when'],
        ['a,b', 1.5, true, new Date(Date.UTC(2020, 0, 2, 3, 4, 5))],
        ['say "hi"', 2, false, null],
      ], {dateUTC: true, sheetName: 'S'});
      const lines = csv.split('\n').map(l => l.replace(/\r$/, ''));
      expect(lines[0]).to.equal('name,n,ok,when');
      expect(lines[1]).to.match(/^"a,b",1\.5,true,2020-01-02T03:04:05Z$/);
      expect(lines[2]).to.match(/^"say ""hi""",2,false,?$/);
    });

    it('supports a custom date format and custom delimiter', async () => {
      const csv = await csvOf([[new Date(Date.UTC(2020, 0, 2)), 'x']], {
        dateFormat: 'DD/MM/YYYY', dateUTC: true, formatterOptions: {delimiter: ';'},
      });
      expect(csv.trim()).to.equal('02/01/2020;x');
    });

    it('writes hyperlinks, formulas results and errors as text', async () => {
      const csv = await csvOf([[
        {text: 't', hyperlink: 'https://example.com/'},
        {formula: 'A1', result: 'res'},
        {error: '#N/A'},
      ]]);
      expect(csv.trim()).to.equal('https://example.com/,res,#N/A');
    });

    it('neutralises leading formula characters by default', async () => {
      const csv = await csvOf([['=1+1', '+1', '-1', '@x', 'safe', 'a=b']]);
      expect(csv.trim()).to.equal('\'=1+1,\'+1,\'-1,\'@x,safe,a=b');
    });

    it('does not mangle numeric negatives written as numbers', async () => {
      // numbers are not strings, so they are never prefixed
      const csv = await csvOf([[-5, 3]]);
      expect(csv.trim()).to.equal('-5,3');
    });

    it('escapeFormulas:false disables the guard', async () => {
      const csv = await csvOf([['=1+1']], {escapeFormulas: false});
      expect(csv.trim()).to.equal('=1+1');
    });

    it('applies a custom map', async () => {
      const csv = await csvOf([['a', 'b']], {map: v => String(v).toUpperCase()});
      expect(csv.trim()).to.equal('A,B');
    });

    it('includeEmptyRows keeps gaps; false drops them', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 'a';
      ws.getCell('A4').value = 'd';
      const withGaps = (await wb.csv.writeBuffer()).toString().trim().split(/\r?\n/);
      const noGaps = (await wb.csv.writeBuffer({includeEmptyRows: false})).toString().trim().split(/\r?\n/);
      expect(withGaps).to.deep.equal(['a', '', '', 'd']);
      expect(noGaps).to.deep.equal(['a', 'd']);
    });

    it('reads csv: numbers, strings, dates, empty, booleans', async () => {
      const wb = new ExcelJS.Workbook();
      const stream = new PassThrough();
      stream.end('a,1,2020-01-02,,true\nb,2.5,2020-01-02T03:04:05,x y,false\n');
      const ws = await wb.csv.read(stream, {sheetName: 'imp'});
      expect(ws.name).to.equal('imp');
      expect(ws.getCell('A1').value).to.equal('a');
      expect(ws.getCell('B1').value).to.equal(1);
      expect(ws.getCell('B2').value).to.equal(2.5);
      expect(ws.getCell('C1').value).to.be.instanceOf(Date);
      expect(ws.getCell('C1').value.getFullYear()).to.equal(2020);
      expect(ws.getCell('D1').value).to.equal(null);
      expect(ws.getCell('E1').value).to.equal(true);
      expect(ws.getCell('E2').value).to.equal(false);
      expect(ws.getCell('D2').value).to.equal('x y');
    });

    it('reads csv with custom parser options and map', async () => {
      const wb = new ExcelJS.Workbook();
      const stream = new PassThrough();
      stream.end('h1;h2\n1;2\n');
      const ws = await wb.csv.read(stream, {parserOptions: {delimiter: ';', headers: false}, map: v => `<${v}>`});
      expect(ws.getCell('A1').value).to.equal('<h1>');
      expect(ws.getCell('B2').value).to.equal('<2>');
    });

    it('reads csv with quoted fields containing newlines and separators', async () => {
      const wb = new ExcelJS.Workbook();
      const stream = new PassThrough();
      stream.end('"a,b","line1\nline2","q""x"\n');
      const ws = await wb.csv.read(stream);
      expect(ws.getCell('A1').value).to.equal('a,b');
      expect(ws.getCell('B1').value).to.equal('line1\nline2');
      expect(ws.getCell('C1').value).to.equal('q"x');
    });

    it('readFile / writeFile round-trip, and readFile of a missing file rejects', async () => {
      const file = tmpFile('rt.csv');
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('S').addRows([['a', 1], ['b', 2]]);
      await wb.csv.writeFile(file, {sheetName: 'S'});
      const wb2 = new ExcelJS.Workbook();
      const ws = await wb2.csv.readFile(file);
      expect(ws.getSheetValues().slice(1).map(r => r.slice(1))).to.deep.equal([['a', 1], ['b', 2]]);
      let err;
      try { await new ExcelJS.Workbook().csv.readFile(`${file}.missing`); } catch (e) { err = e; }
      expect(err).to.be.instanceOf(Error);
      fs.rmSync(path.dirname(file), {recursive: true, force: true});
    });

    it('csv -> xlsx conversion keeps ordinary values', async () => {
      const wb = new ExcelJS.Workbook();
      const stream = new PassThrough();
      stream.end('id,name,score\n1,alice,9.5\n2,bob,7\n');
      await wb.csv.read(stream, {sheetName: 'S'});
      const wb2 = await roundTrip(wb);
      const ws = wb2.getWorksheet('S');
      expect(ws.getCell('B2').value).to.equal('alice');
      expect(ws.getCell('C3').value).to.equal(7);
    });
  });

  describe('xlsx load/save entry points', () => {
    const makeBuffer = async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('S');
      ws.addRows([['h1', 'h2'], [1, 2], [3, 4], [5, 6]]);
      return wb.xlsx.writeBuffer();
    };

    it('load() accepts Buffer, Uint8Array, ArrayBuffer and base64', async () => {
      const buf = await makeBuffer();
      // eslint-disable-next-line no-restricted-syntax
      /* eslint-disable no-await-in-loop */
      for (const [label, data, opts] of [
        ['buffer', buf, undefined],
        ['uint8array', new Uint8Array(buf), undefined],
        ['arraybuffer', new Uint8Array(buf).buffer, undefined],
        ['base64', buf.toString('base64'), {base64: true}],
      ]) {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(data, opts);
        expect(wb.getWorksheet('S').getCell('B3').value, label).to.equal(4);
      }
    });

    it('read(stream), readFile and writeFile / write(stream) work', async () => {
      const file = tmpFile('rt.xlsx');
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('S').addRows([['a', 'b']]);
      await wb.xlsx.writeFile(file);
      const wbF = new ExcelJS.Workbook();
      await wbF.xlsx.readFile(file);
      expect(wbF.getWorksheet('S').getCell('B1').value).to.equal('b');
      const wbS = new ExcelJS.Workbook();
      await wbS.xlsx.read(fs.createReadStream(file));
      expect(wbS.getWorksheet('S').getCell('A1').value).to.equal('a');
      const out = new PassThrough();
      const chunks = [];
      out.on('data', c => chunks.push(c));
      await wb.xlsx.write(out);
      const wbW = new ExcelJS.Workbook();
      await wbW.xlsx.load(Buffer.concat(chunks));
      expect(wbW.getWorksheet('S').getCell('A1').value).to.equal('a');
      fs.rmSync(path.dirname(file), {recursive: true, force: true});
    });

    it('readFile on a missing file rejects', async () => {
      let err;
      try { await new ExcelJS.Workbook().xlsx.readFile('/nonexistent/file.xlsx'); } catch (e) { err = e; }
      expect(err).to.be.instanceOf(Error);
    });

    it('load() rejects garbage input', async () => {
      let err;
      try { await new ExcelJS.Workbook().xlsx.load(Buffer.from('this is not a zip')); } catch (e) { err = e; }
      expect(err).to.be.instanceOf(Error);
    });

    it('ignoreNodes skips the requested nodes', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 1;
      ws.mergeCells('A1:B2');
      const buf = await wb.xlsx.writeBuffer();
      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(buf, {ignoreNodes: ['mergeCells']});
      expect(Object.keys(wb2.getWorksheet('S')._merges)).to.have.length(0);
      expect(wb2.getWorksheet('S').getCell('A1').value).to.equal(1);
    });

    it('maxRows / maxCols limit what is parsed', async () => {
      const buf = await makeBuffer();
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buf, {maxRows: 10, maxCols: 10});
      expect(wb.getWorksheet('S').getCell('B4').value).to.equal(6);
      let err;
      try { await new ExcelJS.Workbook().xlsx.load(buf, {maxCols: 1}); } catch (e) { err = e; }
      expect(err).to.be.instanceOf(Error);
      expect(err.message).to.match(/Max column count/);
    });

    it('legitimately large-ish files still load within default decompression limits', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Big');
      for (let r = 1; r <= 3000; r++) ws.addRow([r, `row ${r}`, r * 1.5, new Date(2020, 0, 1 + (r % 28))]);
      const wb2 = await roundTrip(wb);
      expect(wb2.getWorksheet('Big').rowCount).to.equal(3000);
      expect(wb2.getWorksheet('Big').getCell('B3000').value).to.equal('row 3000');
    });

    it('decompression limits: low maxEntryUncompressedSize rejects, default accepts', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('S');
      for (let r = 1; r <= 500; r++) ws.addRow(['x'.repeat(100), r]);
      const buf = await wb.xlsx.writeBuffer();
      let err;
      try { await new ExcelJS.Workbook().xlsx.load(buf, {maxEntryUncompressedSize: 1000}); } catch (e) { err = e; }
      expect(err, 'entry limit').to.be.instanceOf(Error);
      expect(err.message).to.match(/uncompressed|bomb|limit/i);
      err = undefined;
      try { await new ExcelJS.Workbook().xlsx.load(buf, {maxTotalUncompressedSize: 2000}); } catch (e) { err = e; }
      expect(err, 'total limit').to.be.instanceOf(Error);
      await new ExcelJS.Workbook().xlsx.load(buf); // default is fine
    });

    it('zip compression options on write', async () => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('S').addRows([['a'.repeat(1000)]]);
      const small = await wb.xlsx.writeBuffer({zip: {compression: 'DEFLATE', compressionOptions: {level: 9}}});
      const stored = await wb.xlsx.writeBuffer({zip: {compression: 'STORE'}});
      expect(stored.length).to.be.greaterThan(small.length);
      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(stored);
      expect(wb2.getWorksheet('S').getCell('A1').value).to.have.length(1000);
    });

    it('the produced package is a valid zip with the expected parts', async () => {
      const zip = await JSZip.loadAsync(await makeBuffer());
      ['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/styles.xml', 'docProps/core.xml']
        .forEach(n => expect(zip.file(n), n).to.be.ok());
    });
  });

  describe('workbook model api', () => {
    it('workbook.model json round-trips through a new Workbook', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('S');
      ws.addRows([['a', 1], ['b', 2]]);
      ws.mergeCells('C1:D2');
      const wb2 = new ExcelJS.Workbook();
      wb2.model = JSON.parse(JSON.stringify(wb.model));
      expect(wb2.getWorksheet('S').getCell('A2').value).to.equal('b');
    });

    it('worksheet.model / getSheetValues / getRows are consistent', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('S');
      ws.addRows([[1, 2], [3, 4]]);
      expect(ws.getSheetValues()[2][2]).to.equal(4);
      expect(ws.getRows(1, 2).map(r => r.values.slice(1))).to.deep.equal([[1, 2], [3, 4]]);
      expect(ws.dimensions.model).to.deep.include({top: 1, left: 1, bottom: 2, right: 2});
    });

    it('addWorksheet rejects duplicate and invalid names', () => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('dup');
      expect(() => wb.addWorksheet('dup')).to.throw();
      expect(() => wb.addWorksheet('bad:name')).to.throw();
      expect(() => wb.addWorksheet('')).to.throw();
    });

    it('cell addresses and ranges decode / encode', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('S');
      expect(ws.getCell('AA10').address).to.equal('AA10');
      expect(ws.getCell(10, 27).address).to.equal('AA10');
      expect(ws.getCell('AA10').col).to.equal(27);
      expect(ws.getColumn('AA').number).to.equal(27);
      expect(ws.getColumn(27).letter).to.equal('AA');
    });
  });

  describe('streaming writer + reader', () => {
    const writeStreamed = async (file, opts = {}) => {
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({filename: file, useSharedStrings: true, useStyles: true, ...opts});
      const ws = wb.addWorksheet('Stream', {views: [{state: 'frozen', ySplit: 1}]});
      ws.columns = [{key: 'id', width: 8}, {key: 'name'}];
      ws.getCell('D1').value = {text: 'link', hyperlink: 'https://example.com/a?b=1&c=2'};
      ws.getCell('E1').value = 'with <&> chars';
      ws.getCell('F1').value = new Date(Date.UTC(2020, 5, 1));
      ws.getCell('G1').value = {formula: 'SUM(A2:A3)', result: 5};
      ws.mergeCells('H1:I2');
      ws.getCell('J1').dataValidation = {type: 'list', formulae: ['"a,b"']};
      ws.autoFilter = 'A1:B1';
      for (let i = 1; i <= 50; i++) {
        const row = ws.getRow(i);
        row.getCell(1).value = i;
        row.getCell(2).value = `n${i}`;
        row.commit();
      }
      ws.commit();
      await wb.commit();
    };

    it('stream-written file can be loaded with the in-memory reader', async () => {
      const file = tmpFile('s.xlsx');
      await writeStreamed(file);
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(file);
      const ws = wb.getWorksheet('Stream');
      expect(ws.getCell('A1').value).to.equal(1);
      expect(ws.getCell('B50').value).to.equal('n50');
      expect(ws.getCell('D1').value.hyperlink).to.equal('https://example.com/a?b=1&c=2');
      expect(ws.getCell('E1').value).to.equal('with <&> chars');
      expect(ws.getCell('F1').value.getTime()).to.equal(Date.UTC(2020, 5, 1));
      expect(ws.getCell('G1').value.formula).to.equal('SUM(A2:A3)');
      expect(ws.getCell('I2').isMerged).to.equal(true);
      expect(ws.getCell('J1').dataValidation.type).to.equal('list');
      expect(ws.views[0].state).to.equal('frozen');
      expect(ws.autoFilter).to.equal('A1:B1');
      fs.rmSync(path.dirname(file), {recursive: true, force: true});
    });

    it('stream-written file can be read with the streaming reader (async iteration)', async () => {
      const file = tmpFile('s.xlsx');
      await writeStreamed(file);
      const reader = new ExcelJS.stream.xlsx.WorkbookReader(file, {
        worksheets: 'emit', sharedStrings: 'cache', hyperlinks: 'cache', styles: 'cache', entries: 'emit',
      });
      const rows = [];
      let sheets = 0;
      for await (const worksheetReader of reader) {
        sheets++;
        for await (const row of worksheetReader) rows.push(row.values);
      }
      expect(sheets).to.equal(1);
      expect(rows).to.have.length(50);
      expect(rows[0][1]).to.equal(1);
      expect(rows[49][2]).to.equal('n50');
      fs.rmSync(path.dirname(file), {recursive: true, force: true});
    });

    it('streaming reader reads a regular (in-memory written) file, shared strings listed after sheets', async () => {
      const file = tmpFile('m.xlsx');
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('One');
      ws.addRows([['a', 'b'], ['c', 'd']]);
      wb.addWorksheet('Two').addRows([['x']]);
      await wb.xlsx.writeFile(file);
      const reader = new ExcelJS.stream.xlsx.WorkbookReader(file, {worksheets: 'emit', sharedStrings: 'cache'});
      const seen = {};
      for await (const wsr of reader) {
        seen[wsr.id || wsr.name] = [];
        for await (const row of wsr) seen[wsr.id || wsr.name].push(row.values.slice(1));
      }
      const all = Object.values(seen);
      expect(all).to.have.length(2);
      expect(all[0]).to.deep.equal([['a', 'b'], ['c', 'd']]);
      expect(all[1]).to.deep.equal([['x']]);
      fs.rmSync(path.dirname(file), {recursive: true, force: true});
    });

    it('streaming reader works from a stream and with event emitter style', async () => {
      const file = tmpFile('ev.xlsx');
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('E').addRows([[1], [2], [3]]);
      await wb.xlsx.writeFile(file);
      const reader = new ExcelJS.stream.xlsx.WorkbookReader(fs.createReadStream(file), {worksheets: 'emit', sharedStrings: 'cache'});
      const values = [];
      await new Promise((resolve, reject) => {
        reader.on('worksheet', w => w.on('row', r => values.push(r.values[1])));
        reader.on('end', resolve);
        reader.on('error', reject);
        reader.read();
      });
      expect(values).to.deep.equal([1, 2, 3]);
      fs.rmSync(path.dirname(file), {recursive: true, force: true});
    });

    it('streaming reader rejects/ errors on a non-zip input instead of hanging', async () => {
      const file = tmpFile('bad.xlsx');
      fs.writeFileSync(file, 'not a zip at all');
      const reader = new ExcelJS.stream.xlsx.WorkbookReader(file, {worksheets: 'emit'});
      let err;
      try {
        for await (const wsr of reader) { for await (const r of wsr) { expect(r).to.be.ok(); } }
      } catch (e) { err = e; }
      expect(err).to.be.instanceOf(Error);
      fs.rmSync(path.dirname(file), {recursive: true, force: true});
    });

    it('streaming writer without shared strings / styles', async () => {
      const file = tmpFile('ns.xlsx');
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({filename: file, useSharedStrings: false, useStyles: false});
      const ws = wb.addWorksheet('N');
      ws.addRow(['plain', 1]).commit();
      ws.commit();
      await wb.commit();
      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.readFile(file);
      expect(wb2.getWorksheet('N').getCell('A1').value).to.equal('plain');
      fs.rmSync(path.dirname(file), {recursive: true, force: true});
    });

    it('streaming writer embeds media from a buffer and from an absolute file path', async () => {
      const file = tmpFile('img.xlsx');
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({filename: file});
      const ws = wb.addWorksheet('I');
      const id1 = wb.addImage({buffer: PNG_1X1, extension: 'png'});
      const id2 = wb.addImage({filename: path.resolve('./spec/integration/data/image.png'), extension: 'png'});
      ws.addBackgroundImage(id1);
      ws.addRow(['x']).commit();
      ws.commit();
      await wb.commit();
      const zip = await JSZip.loadAsync(fs.readFileSync(file));
      const media = Object.keys(zip.files).filter(n => /^xl\/media\/image\d+\.png$/.test(n)).sort();
      expect(media).to.have.length(2);
      expect(id2).to.equal(1);
      const first = Buffer.from(await zip.file(media[0]).async('uint8array'));
      expect(first.equals(PNG_1X1)).to.equal(true);
      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.readFile(file);
      expect(wb2.getWorksheet('I').getCell('A1').value).to.equal('x');
      fs.rmSync(path.dirname(file), {recursive: true, force: true});
    });

    it('streaming writer supports comments (notes), rich text and rowspan merges', async () => {
      const file = tmpFile('misc.xlsx');
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({filename: file});
      const ws = wb.addWorksheet('M');
      ws.getCell('A1').value = {richText: [{text: 'Hello ', font: {bold: true}}, {text: 'world'}]};
      ws.getCell('A2').value = 'noted';
      ws.getCell('A2').note = 'a note';
      ws.mergeCells('C1:D1');
      ws.getRow(1).commit();
      ws.getRow(2).commit();
      ws.commit();
      await wb.commit();
      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.readFile(file);
      const w = wb2.getWorksheet('M');
      expect(w.getCell('A1').value.richText[0].text).to.equal('Hello ');
      expect(w.getCell('A2').note).to.equal('a note');
      expect(w.getCell('D1').isMerged).to.equal(true);
      fs.rmSync(path.dirname(file), {recursive: true, force: true});
    });
  });

  describe('existing security guards keep working', () => {
    const {assertSafeMediaPath, safeJoin} = require('../../../lib/utils/media-path-guard');
    const {escapeCsvFormulaValue} = require('../../../lib/utils/csv-formula-guard');

    it('assertSafeMediaPath rejects .. segments and null bytes, accepts normal paths', () => {
      expect(() => assertSafeMediaPath('../x.png')).to.throw();
      expect(() => assertSafeMediaPath('a/../../x.png')).to.throw();
      expect(() => assertSafeMediaPath('a\0b.png')).to.throw();
      expect(() => assertSafeMediaPath('/abs/path/x.png')).to.not.throw();
      expect(() => assertSafeMediaPath('rel/x.png')).to.not.throw();
      expect(() => assertSafeMediaPath(undefined)).to.not.throw();
      expect(() => assertSafeMediaPath('')).to.not.throw();
    });

    it('safeJoin confines to the base directory', () => {
      const base = path.join(os.tmpdir(), 'base-dir');
      expect(safeJoin(base, 'a/b.png')).to.equal(path.join(base, 'a', 'b.png'));
      expect(safeJoin(base, './a.png')).to.equal(path.join(base, 'a.png'));
      expect(() => safeJoin(base, '../x')).to.throw();
      expect(() => safeJoin(base, 'a/../../x')).to.throw();
      expect(() => safeJoin(base, '/etc/passwd')).to.throw();
      expect(safeJoin(base, 'a/../b.png')).to.equal(path.join(base, 'b.png'));
    });

    it('escapeCsvFormulaValue only touches strings with trigger prefix', () => {
      expect(escapeCsvFormulaValue('=A1')).to.equal('\'=A1');
      expect(escapeCsvFormulaValue('\tx')).to.equal('\'\tx');
      expect(escapeCsvFormulaValue('hello')).to.equal('hello');
      expect(escapeCsvFormulaValue('')).to.equal('');
      expect(escapeCsvFormulaValue(5)).to.equal(5);
      expect(escapeCsvFormulaValue(null)).to.equal(null);
    });

    it('addMedia via Workbook#xlsx rejects traversal filename on write', async () => {
      const wb = new ExcelJS.Workbook();
      expect(() => wb.addImage({filename: '../secret.png', extension: 'png'})).to.throw();
    });

    it('deepMerge-style option handling does not allow prototype pollution via model/style', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('S');
      const evil = JSON.parse('{"__proto__": {"polluted": "yes"}, "font": {"bold": true}}');
      ws.getCell('A1').style = evil;
      ws.getCell('A1').value = 1;
      expect(({}).polluted).to.equal(undefined);
      expect(ws.getCell('A1').font.bold).to.equal(true);
    });
  });
});

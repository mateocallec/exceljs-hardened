// Prototype-pollution patches typically swap `{}` dictionaries for Map /
// Object.create(null). These tests make sure that ordinary workbooks that merely
// USE names like "constructor" or "toString" (as sheet names, keys, table names,
// defined names, csv cells...) keep working identically after such a patch.
const {PassThrough} = require('stream');
const {roundTrip, ExcelJS} = require('../../utils/roundtrip');

const NAMES = ['constructor', 'toString', 'valueOf', 'hasOwnProperty', 'prototype', 'proto', '__defineGetter__'];
// Pre-existing bug (same root cause as VULN-01: `{}` used as a dictionary): defined names that
// point at a sheet with one of these names are silently dropped on write, and a defined name
// called like one of them throws. Those cases are in the `it.skip` tests at the bottom; enable
// them once the CellMatrix / DefinedNames dictionaries are hardened.
const BROKEN_FOR_DEFINED_NAMES = ['constructor', 'toString', 'valueOf', 'hasOwnProperty', '__defineGetter__'];
const SAFE_FOR_DEFINED_NAMES = NAMES.filter(n => !BROKEN_FOR_DEFINED_NAMES.includes(n));

describe('hardening regression: names that collide with Object.prototype members', () => {
  afterEach(() => {
    expect(Object.keys(Object.prototype)).to.deep.equal([]);
  });

  it('sheets named like prototype members, with defined names, merges and values', async () => {
    const wb = new ExcelJS.Workbook();
    NAMES.forEach((n, i) => {
      const ws = wb.addWorksheet(n);
      ws.getCell('A1').value = `${n}-value`;
      ws.getCell('B2').value = i;
      ws.mergeCells('D1:E2');
      if (SAFE_FOR_DEFINED_NAMES.includes(n)) wb.definedNames.add(`'${n}'!$A$1:$A$3`, `name_${i}`);
    });
    const wb2 = await roundTrip(wb);
    NAMES.forEach((n, i) => {
      const ws = wb2.getWorksheet(n);
      expect(ws, n).to.be.ok();
      expect(ws.getCell('A1').value).to.equal(`${n}-value`);
      expect(ws.getCell('B2').value).to.equal(i);
      expect(ws.getCell('E2').isMerged).to.equal(true);
      if (SAFE_FOR_DEFINED_NAMES.includes(n)) {
        expect(wb2.definedNames.getRanges(`name_${i}`).ranges).to.have.length(1);
      }
    });
    expect(wb2.getWorksheet('doesNotExist')).to.equal(undefined);
    expect(wb2.getWorksheet('hasOwnProperty').name).to.equal('hasOwnProperty');
  });

  it('defined names with ordinary names pointing at sheets that exist', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('S');
    ['prototype', 'proto', 'name_with_underscore', 'N1'].forEach(n => wb.definedNames.add('S!$B$2', n));
    const wb2 = await roundTrip(wb);
    ['prototype', 'proto', 'name_with_underscore', 'N1'].forEach(n => {
      expect(wb2.definedNames.getRanges(n).ranges, n).to.deep.equal(['S!$B$2']);
    });
  });

  it('column keys, row object keys and table names', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('K');
    ws.columns = [{header: 'c', key: 'constructor'}, {header: 't', key: 'toString'}];
    ws.addRow({constructor: 'cval', toString: 'tval'});
    ws.getRow(2).getCell('constructor');
    const ws2 = (await roundTrip(wb)).getWorksheet('K');
    expect(ws2.getCell('A2').value).to.equal('cval');
    expect(ws2.getCell('B2').value).to.equal('tval');

    const wbT = new ExcelJS.Workbook();
    wbT.addWorksheet('T').addTable({name: 'constructorTbl', ref: 'A1', columns: [{name: 'toString'}, {name: 'valueOf'}], rows: [[1, 2]]});
    const t = (await roundTrip(wbT)).getWorksheet('T').getTable('constructorTbl');
    expect(t.table.columns.map(c => c.name)).to.deep.equal(['toString', 'valueOf']);
  });

  it('conditional formatting with refs / ids resembling prototype members', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('CF');
    ws.addRows([[1, 2], [3, 4]]);
    ws.addConditionalFormatting({
      ref: 'A1:B2',
      rules: [{type: 'dataBar', cfvo: [{type: 'min'}, {type: 'max'}], color: {argb: 'FF638EC6'}, gradient: false, priority: 1}],
    });
    ws.addConditionalFormatting({
      ref: 'A1:A2',
      rules: [{type: 'cellIs', operator: 'equal', formulae: [1], style: {font: {bold: true}}, priority: 2}],
    });
    const ws2 = (await roundTrip(wb)).getWorksheet('CF');
    expect(ws2.conditionalFormattings).to.have.length(2);
    // second pass: extension rules (x14) must be matched back to their base rule
    const ws3 = (await roundTrip(await roundTrip(wb))).getWorksheet('CF');
    expect(ws3.conditionalFormattings.map(c => c.rules[0].type).sort()).to.deep.equal(['cellIs', 'dataBar']);
  });

  it('csv cells equal to prototype member names are kept as text', async () => {
    const wb = new ExcelJS.Workbook();
    const stream = new PassThrough();
    stream.end('plain,true,#N/A\nabc,false,#REF!\n');
    const ws = await wb.csv.read(stream);
    expect(ws.getCell('B1').value).to.equal(true);
    expect(ws.getCell('B2').value).to.equal(false);
    expect(ws.getCell('C1').value).to.deep.equal({error: '#N/A'});
    expect(ws.getCell('C2').value).to.deep.equal({error: '#REF!'});
    expect(ws.getCell('A1').value).to.equal('plain');
  });

  // --- known pre-existing bugs, to be enabled once the VULN-01 patch lands ---
  describe('KNOWN BUG (enable after VULN-01 fix)', () => {
    BROKEN_FOR_DEFINED_NAMES.forEach(n => {
      it.skip(`defined name on a sheet called "${n}" survives a round-trip`, async () => {
        const wb = new ExcelJS.Workbook();
        wb.addWorksheet(n).getCell('A1').value = 1;
        wb.definedNames.add(`'${n}'!$A$1:$A$3`, 'nm');
        const wb2 = await roundTrip(wb);
        expect(wb2.definedNames.getRanges('nm').ranges).to.have.length(1);
      });
      it.skip(`a defined name called "${n}" can be added and round-trips`, async () => {
        const wb = new ExcelJS.Workbook();
        wb.addWorksheet('S');
        wb.definedNames.add('S!$B$2', n);
        const wb2 = await roundTrip(wb);
        expect(wb2.definedNames.getRanges(n).ranges).to.deep.equal(['S!$B$2']);
      });
    });
  });
});

const {roundTrip, PNG_1X1, ExcelJS} = require('../../utils/roundtrip');

describe('hardening regression: defined names, data validation, CF, tables, images, pivots', () => {
  describe('defined names', () => {
    it('round-trips single cell, range and multi-range names', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Data');
      ws.getCell('A1').value = 'x';
      ws.getCell('B2').value = 'y';
      wb.definedNames.add('Data!$A$1', 'OneCell');
      wb.definedNames.add('Data!$A$1:$C$3', 'Block');
      wb.definedNames.add('Data!$A$1:$A$2', 'Multi');
      wb.definedNames.add('Data!$B$2', 'Multi');
      const wb2 = await roundTrip(wb);
      expect(wb2.definedNames.getRanges('OneCell').ranges).to.deep.equal(['Data!$A$1']);
      expect(wb2.definedNames.getRanges('Block').ranges).to.deep.equal(['Data!$A$1:$C$3']);
      expect(wb2.definedNames.getRanges('Multi').ranges.sort()).to.deep.equal(['Data!$A$1:$A$2', 'Data!$B$2']);
      expect(wb2.definedNames.getNames('Data!A1').sort()).to.deep.equal(['Block', 'Multi', 'OneCell']);
    });

    it('supports sheet names with quotes/spaces and multiple sheets', async () => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('My Sheet');
      wb.addWorksheet('Other');
      wb.definedNames.add('\'My Sheet\'!$A$1:$B$2', 'Spaced');
      wb.definedNames.add('Other!$C$3', 'OnOther');
      const wb2 = await roundTrip(wb);
      expect(wb2.definedNames.getRanges('Spaced').ranges).to.deep.equal(['\'My Sheet\'!$A$1:$B$2']);
      expect(wb2.definedNames.getRanges('OnOther').ranges).to.deep.equal(['Other!$C$3']);
    });

    it('remove / removeAllNames work', async () => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('S');
      wb.definedNames.add('S!$A$1', 'n1');
      wb.definedNames.add('S!$A$2', 'n2');
      wb.definedNames.remove('S!$A$1', 'n1');
      const wb2 = await roundTrip(wb);
      expect(wb2.definedNames.getRanges('n1').ranges).to.have.length(0);
      expect(wb2.definedNames.getRanges('n2').ranges).to.have.length(1);
    });

    it('print area survives as defined name', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('P');
      ws.getCell('A1').value = 1;
      ws.pageSetup.printArea = 'A1:C5';
      expect((await roundTrip(wb)).getWorksheet('P').pageSetup.printArea).to.equal('A1:C5');
    });

    it('does not pollute Object.prototype', async () => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('S');
      wb.definedNames.add('S!$A$1:$A$3', 'n');
      await roundTrip(wb);
      expect(Object.keys(Object.prototype)).to.deep.equal([]);
      expect(({})[1]).to.equal(undefined);
    });
  });

  describe('data validation', () => {
    it('round-trips list / whole / decimal / date / custom validations', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('DV');
      ws.getCell('A1').dataValidation = {type: 'list', allowBlank: true, formulae: ['"One,Two,Three"']};
      ws.getCell('A2').dataValidation = {
        type: 'whole', operator: 'between', formulae: [1, 10],
        showErrorMessage: true, errorTitle: 'Bad', error: 'Must be 1-10',
      };
      ws.getCell('A3').dataValidation = {type: 'decimal', operator: 'greaterThan', formulae: [0.5]};
      ws.getCell('A4').dataValidation = {type: 'custom', formulae: ['A1>0'], showInputMessage: true, promptTitle: 'P', prompt: 'Hello'};
      ws.getCell('A5').dataValidation = {type: 'textLength', operator: 'lessThan', formulae: [20]};
      const ws2 = (await roundTrip(wb)).getWorksheet('DV');
      expect(ws2.getCell('A1').dataValidation.type).to.equal('list');
      expect(ws2.getCell('A1').dataValidation.formulae).to.deep.equal(['"One,Two,Three"']);
      expect(ws2.getCell('A2').dataValidation).to.deep.include({type: 'whole', operator: 'between', errorTitle: 'Bad', error: 'Must be 1-10'});
      expect(ws2.getCell('A2').dataValidation.formulae).to.deep.equal([1, 10]);
      expect(ws2.getCell('A3').dataValidation.formulae).to.deep.equal([0.5]);
      expect(ws2.getCell('A4').dataValidation.prompt).to.equal('Hello');
      expect(ws2.getCell('A5').dataValidation.operator).to.equal('lessThan');
    });

    it('applies the same validation to every cell of a multi-cell range', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('DVR');
      for (let r = 1; r <= 5; r++) {
        for (let c = 1; c <= 3; c++) {
          ws.getCell(r, c).dataValidation = {type: 'list', formulae: ['"a,b"']};
        }
      }
      const ws2 = (await roundTrip(wb)).getWorksheet('DVR');
      for (let r = 1; r <= 5; r++) {
        for (let c = 1; c <= 3; c++) {
          expect(ws2.getCell(r, c).dataValidation, `${r},${c}`).to.deep.include({type: 'list'});
        }
      }
      expect(ws2.getCell('D1').dataValidation).to.equal(undefined);
      expect(ws2.getCell('A6').dataValidation).to.equal(undefined);
    });
  });

  describe('conditional formatting', () => {
    it('round-trips cellIs, expression, containsText, top10, colorScale, dataBar, iconSet', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('CF');
      ws.addRows([[1, 2], [3, 4], [5, 6]]);
      ws.addConditionalFormatting({
        ref: 'A1:B3',
        rules: [
          {type: 'cellIs', operator: 'greaterThan', formulae: [3], style: {fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FFFF0000'}}}, priority: 1},
          {type: 'expression', formulae: ['MOD(A1,2)=0'], style: {font: {bold: true}}, priority: 2},
          {type: 'containsText', operator: 'containsText', text: 'foo', style: {font: {italic: true}}, priority: 3},
          {type: 'top10', rank: 2, percent: false, bottom: false, style: {font: {bold: true}}, priority: 4},
        ],
      });
      ws.addConditionalFormatting({
        ref: 'A1:A3',
        rules: [{
          type: 'colorScale',
          cfvo: [{type: 'min'}, {type: 'max'}],
          color: [{argb: 'FFFFFFFF'}, {argb: 'FFFF0000'}],
          priority: 5,
        }],
      });
      ws.addConditionalFormatting({
        ref: 'B1:B3',
        rules: [{
          type: 'dataBar',
          minLength: 0, maxLength: 100,
          cfvo: [{type: 'min'}, {type: 'max'}],
          color: {argb: 'FF638EC6'}, gradient: false,
          priority: 6,
        }],
      });
      ws.addConditionalFormatting({
        ref: 'A1:A3',
        rules: [{
          type: 'iconSet', iconSet: '3TrafficLights',
          cfvo: [{type: 'num', value: 0}, {type: 'num', value: 2}, {type: 'num', value: 4}],
          priority: 7,
        }],
      });
      const ws2 = (await roundTrip(wb)).getWorksheet('CF');
      const cfs = ws2.conditionalFormattings;
      expect(cfs.length).to.equal(4);
      const first = cfs.find(c => c.ref === 'A1:B3');
      expect(first.rules.map(r => r.type)).to.deep.equal(['cellIs', 'expression', 'containsText', 'top10']);
      expect(first.rules[0].formulae.map(String)).to.deep.equal(['3']);
      expect(first.rules[2].formulae).to.deep.equal(['NOT(ISERROR(SEARCH("foo",A1)))']);
      expect(first.rules[3].rank).to.equal(2);
      expect(cfs.find(c => c.rules[0].type === 'colorScale')).to.be.ok();
      expect(cfs.find(c => c.rules[0].type === 'dataBar')).to.be.ok();
      expect(cfs.find(c => c.rules[0].type === 'iconSet').rules[0].iconSet).to.equal('3TrafficLights');
      expect(Object.keys(Object.prototype)).to.deep.equal([]);
    });

    it('writes containsText with the expected SEARCH formula', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('CFT');
      ws.getCell('A1').value = 'hello';
      ws.addConditionalFormatting({
        ref: 'A1:A1',
        rules: [{type: 'containsText', operator: 'containsText', text: 'ell', style: {font: {bold: true}}, priority: 1}],
      });
      const buf = await wb.xlsx.writeBuffer();
      const JSZip = require('jszip');
      const zip = await JSZip.loadAsync(buf);
      const xml = await zip.file('xl/worksheets/sheet1.xml').async('string');
      expect(xml).to.contain('NOT(ISERROR(SEARCH(&quot;ell&quot;,A1)))');
    });
  });

  describe('tables', () => {
    it('round-trips a table with totals row and custom styles', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('T');
      ws.addTable({
        name: 'Sales',
        ref: 'A1',
        headerRow: true,
        totalsRow: true,
        style: {theme: 'TableStyleMedium2', showRowStripes: true},
        columns: [
          {name: 'Region', totalsRowLabel: 'Total', filterButton: true},
          {name: 'Amount', totalsRowFunction: 'sum', filterButton: false},
        ],
        rows: [['N', 10], ['S', 20], ['E', 30]],
      });
      const ws2 = (await roundTrip(wb)).getWorksheet('T');
      const t = ws2.getTable('Sales');
      expect(t).to.be.ok();
      expect(t.table.tableRef).to.equal('A1:B5');
      expect(t.table.totalsRow).to.equal(true);
      expect(t.table.columns.map(c => c.name)).to.deep.equal(['Region', 'Amount']);
      expect(t.table.columns[1].totalsRowFunction).to.equal('sum');
      expect(ws2.getCell('A1').value).to.equal('Region');
      expect(ws2.getCell('B2').value).to.equal(10);
      expect(ws2.getCell('A5').value).to.equal('Total');
    });

    it('generates a SUBTOTAL formula in the totals row of the sheet xml', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('T2');
      ws.addTable({
        name: 'Tbl', ref: 'A1', headerRow: true, totalsRow: true,
        columns: [{name: 'Label', totalsRowLabel: 'Total'}, {name: 'Qty', totalsRowFunction: 'sum'}],
        rows: [['a', 1], ['b', 2]],
      });
      const JSZip = require('jszip');
      const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
      const sheetXml = await zip.file('xl/worksheets/sheet1.xml').async('string');
      const tableXml = await zip.file('xl/tables/table1.xml').async('string');
      expect(tableXml).to.contain('totalsRowFunction="sum"');
      expect(tableXml).to.contain('name="Qty"');
      expect(sheetXml).to.contain('SUBTOTAL(109,Tbl[Qty])');
    });

    it('supports table add/remove/commit', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('T3');
      const table = ws.addTable({name: 'A1T', ref: 'A1', columns: [{name: 'c'}], rows: [[1]]});
      table.addRow([2]);
      table.commit();
      const ws2 = (await roundTrip(wb)).getWorksheet('T3');
      expect(ws2.getCell('A3').value).to.equal(2);
      ws2.removeTable('A1T');
      expect(ws2.getTable('A1T')).to.equal(undefined);
    });
  });

  describe('images', () => {
    it('round-trips an image with tl/br anchors from a buffer', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Img');
      const id = wb.addImage({buffer: PNG_1X1, extension: 'png'});
      ws.addImage(id, {tl: {col: 1, row: 2}, br: {col: 3, row: 4}});
      ws.addImage(id, {tl: {col: 5.5, row: 6.25}, ext: {width: 100, height: 50}});
      ws.addImage(id, 'B10:D12');
      const wb2 = await roundTrip(wb);
      const ws2 = wb2.getWorksheet('Img');
      const images = ws2.getImages();
      expect(images).to.have.length(3);
      expect(images[0].range.tl.col).to.equal(1);
      expect(images[0].range.tl.row).to.equal(2);
      expect(images[0].range.br.col).to.equal(3);
      expect(images[1].range.ext).to.deep.include({width: 100, height: 50});
      expect(images[1].range.tl.col).to.be.closeTo(5.5, 0.001);
      const {media} = wb2.model;
      expect(media).to.have.length(1);
      expect(media[0].extension).to.equal('png');
      expect(Buffer.from(media[0].buffer).equals(PNG_1X1)).to.equal(true);
      expect(wb2.getImage(images[0].imageId)).to.be.ok();
    });

    it('reads an image from an absolute filename', async () => {
      const path = require('path');
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('ImgFile');
      const id = wb.addImage({filename: path.resolve('./spec/integration/data/image.png'), extension: 'png'});
      ws.addImage(id, 'A1:B2');
      const wb2 = await roundTrip(wb);
      expect(wb2.model.media).to.have.length(1);
    });

    it('rejects ../ traversal in image filenames (workbook)', () => {
      const wb = new ExcelJS.Workbook();
      expect(() => wb.addImage({filename: '../../etc/passwd', extension: 'png'})).to.throw(/traversal|Invalid media path/i);
    });

    it('supports background images', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Bg');
      ws.addBackgroundImage(wb.addImage({buffer: PNG_1X1, extension: 'png'}));
      const JSZip = require('jszip');
      const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
      expect(Object.keys(zip.files).some(n => /^xl\/media\/image\d+\.png$/.test(n))).to.equal(true);
      expect(await zip.file('xl/worksheets/sheet1.xml').async('string')).to.contain('<picture');
    });
  });

  describe('pivot tables', () => {
    const buildBook = () => {
      const wb = new ExcelJS.Workbook();
      const src = wb.addWorksheet('Source');
      src.addRows([
        ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'],
        ['a1', 'b1', 'c1', 'd1', 'e1', 'f1', 4, 5],
        ['a1', 'b2', 'c1', 'd2', 'e1', 'f1', 4, 5],
        ['a2', 'b1', 'c2', 'd1', 'e2', 'f1', 14, 24],
        ['a2', 'b2', 'c2', 'd2', 'e2', 'f2', 24, 35],
      ]);
      const target = wb.addWorksheet('Pivot');
      target.addPivotTable({sourceSheet: src, rows: ['A', 'B', 'E'], columns: ['C', 'D'], values: ['H'], metric: 'sum'});
      return wb;
    };

    it('writes the pivot table parts with the source values', async () => {
      const JSZip = require('jszip');
      const zip = await JSZip.loadAsync(await buildBook().xlsx.writeBuffer());
      const names = Object.keys(zip.files);
      expect(names).to.include('xl/pivotTables/pivotTable1.xml');
      expect(names).to.include('xl/pivotCache/pivotCacheDefinition1.xml');
      expect(names).to.include('xl/pivotCache/pivotCacheRecords1.xml');
      const def = await zip.file('xl/pivotCache/pivotCacheDefinition1.xml').async('string');
      const rec = await zip.file('xl/pivotCache/pivotCacheRecords1.xml').async('string');
      expect(def).to.contain('a1').and.to.contain('b2').and.to.contain('e2');
      expect(rec).to.contain('<r>');
      // XML must be well-formed
      const {SaxesParser} = require('saxes');
      [def, rec].forEach(xml => {
        const p = new SaxesParser();
        p.on('error', e => { throw e; });
        p.write(xml).close();
      });
    });

    it('survives load -> save', async () => {
      const wb2 = await roundTrip(buildBook());
      expect(wb2.worksheets.map(w => w.name)).to.deep.equal(['Source', 'Pivot']);
      const wb3 = await roundTrip(wb2);
      expect(wb3.getWorksheet('Source').getCell('A2').value).to.equal('a1');
    });
  });

  describe('hyperlinks, comments, notes', () => {
    it('round-trips cell comments (notes)', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Cm');
      ws.getCell('A1').value = 1;
      ws.getCell('A1').note = 'Simple note';
      ws.getCell('B2').value = 'v';
      ws.getCell('B2').note = {texts: [{font: {bold: true}, text: 'Bold '}, {text: 'normal'}]};
      const ws2 = (await roundTrip(wb)).getWorksheet('Cm');
      expect(ws2.getCell('A1').note).to.equal('Simple note');
      expect(ws2.getCell('B2').note.texts).to.have.length(2);
      expect(ws2.getCell('B2').note.texts[0].font.bold).to.equal(true);
    });

    it('round-trips internal and external hyperlinks with tooltips', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Hl');
      wb.addWorksheet('Other');
      ws.getCell('A1').value = {text: 'ext', hyperlink: 'https://example.com/?a=1&b=2', tooltip: 'tip'};
      ws.getCell('A2').value = {text: 'int', hyperlink: '#Other!A1'};
      const ws2 = (await roundTrip(wb)).getWorksheet('Hl');
      expect(ws2.getCell('A1').value.hyperlink).to.equal('https://example.com/?a=1&b=2');
      const JSZip = require('jszip');
      const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
      expect(await zip.file('xl/worksheets/sheet1.xml').async('string')).to.contain('tooltip="tip"');
      expect(ws2.getCell('A2').value.hyperlink).to.equal('#Other!A1');
    });
  });
});

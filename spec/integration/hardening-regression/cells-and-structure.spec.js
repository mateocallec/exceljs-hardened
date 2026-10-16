const {roundTrip, ExcelJS} = require('../../utils/roundtrip');

describe('hardening regression: cells, rows, columns, sheets', () => {
  describe('cell values', () => {
    it('round-trips every common value type', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Values');
      const date = new Date(Date.UTC(2020, 5, 15, 12, 30, 0));
      ws.getCell('A1').value = 'text';
      ws.getCell('A2').value = 42;
      ws.getCell('A3').value = 3.14159;
      ws.getCell('A4').value = true;
      ws.getCell('A5').value = false;
      ws.getCell('A6').value = date;
      ws.getCell('A7').value = {formula: 'A2*2', result: 84};
      ws.getCell('A8').value = {richText: [{text: 'bold', font: {bold: true}}, {text: ' plain'}]};
      ws.getCell('A9').value = {text: 'link', hyperlink: 'https://example.com/'};
      ws.getCell('A10').value = {error: '#N/A'};
      ws.getCell('A11').value = 'with <xml> & "quotes" \'apos\'';
      ws.getCell('A12').value = 'line1\nline2';
      ws.getCell('A13').value = '  padded  ';
      ws.getCell('A14').value = 'ünïcödé 日本語 😀';
      ws.getCell('A15').value = null;

      const ws2 = (await roundTrip(wb)).getWorksheet('Values');
      expect(ws2.getCell('A1').value).to.equal('text');
      expect(ws2.getCell('A2').value).to.equal(42);
      expect(ws2.getCell('A3').value).to.equal(3.14159);
      expect(ws2.getCell('A4').value).to.equal(true);
      expect(ws2.getCell('A5').value).to.equal(false);
      expect(ws2.getCell('A6').value.getTime()).to.equal(date.getTime());
      expect(ws2.getCell('A7').value).to.deep.equal({formula: 'A2*2', result: 84});
      expect(ws2.getCell('A8').value.richText).to.have.length(2);
      expect(ws2.getCell('A8').value.richText[0].font.bold).to.equal(true);
      expect(ws2.getCell('A9').value).to.deep.equal({text: 'link', hyperlink: 'https://example.com/'});
      expect(ws2.getCell('A10').value).to.deep.equal({error: '#N/A'});
      expect(ws2.getCell('A11').value).to.equal('with <xml> & "quotes" \'apos\'');
      expect(ws2.getCell('A12').value).to.equal('line1\nline2');
      expect(ws2.getCell('A13').value).to.equal('  padded  ');
      expect(ws2.getCell('A14').value).to.equal('ünïcödé 日本語 😀');
      expect(ws2.getCell('A15').value).to.equal(null);
    });

    it('writes objects with a `formula` key as real formulas (documented behaviour)', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('F');
      ws.getCell('A1').value = 1;
      ws.getCell('A2').value = 2;
      ws.getCell('A3').value = {formula: 'SUM(A1:A2)'};
      const ws2 = (await roundTrip(wb)).getWorksheet('F');
      expect(ws2.getCell('A3').value.formula).to.equal('SUM(A1:A2)');
      expect(ws2.getCell('A3').type).to.equal(ExcelJS.ValueType.Formula);
    });

    it('does not turn plain strings that look like formulas into formulas in xlsx', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = '=1+1';
      ws.getCell('A2').value = '@SUM(1)';
      const ws2 = (await roundTrip(wb)).getWorksheet('S');
      expect(ws2.getCell('A1').value).to.equal('=1+1');
      expect(ws2.getCell('A2').value).to.equal('@SUM(1)');
    });

    it('keeps shared formulas working after round-trip', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('SF');
      ws.getCell('A1').value = 1;
      ws.getCell('A2').value = 2;
      ws.getCell('A3').value = 3;
      ws.getCell('B1').value = {formula: 'A1*2', result: 2};
      ws.getCell('B2').value = {sharedFormula: 'B1', result: 4};
      ws.getCell('B3').value = {sharedFormula: 'B1', result: 6};
      const ws2 = (await roundTrip(wb)).getWorksheet('SF');
      expect(ws2.getCell('B1').formula).to.equal('A1*2');
      expect(ws2.getCell('B2').formula).to.equal('A2*2');
      expect(ws2.getCell('B3').formula).to.equal('A3*2');
      expect(ws2.getCell('B3').result).to.equal(6);
    });

    it('slides sheet-qualified and absolute references in shared formulas', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('SF2');
      wb.addWorksheet('other-sheet_1');
      ws.getCell('B1').value = {formula: '\'other-sheet_1\'!A1+$C$1+D1', result: 0};
      ws.getCell('B2').value = {sharedFormula: 'B1', result: 0};
      const ws2 = (await roundTrip(wb)).getWorksheet('SF2');
      expect(ws2.getCell('B2').formula).to.equal('\'other-sheet_1\'!A2+$C$1+D2');
    });

    it('preserves cell styles', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Style');
      const c = ws.getCell('B2');
      c.value = 12.5;
      c.numFmt = '0.00%';
      c.font = {name: 'Arial', size: 14, bold: true, italic: true, color: {argb: 'FFFF0000'}};
      c.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF00FF00'}};
      c.border = {top: {style: 'thin'}, bottom: {style: 'double', color: {argb: 'FF0000FF'}}};
      c.alignment = {horizontal: 'center', vertical: 'middle', wrapText: true};
      c.protection = {locked: false};
      const c2 = (await roundTrip(wb)).getWorksheet('Style').getCell('B2');
      expect(c2.numFmt).to.equal('0.00%');
      expect(c2.font.name).to.equal('Arial');
      expect(c2.font.size).to.equal(14);
      expect(c2.font.bold).to.equal(true);
      expect(c2.font.color.argb).to.equal('FFFF0000');
      expect(c2.fill.fgColor.argb).to.equal('FF00FF00');
      expect(c2.border.top.style).to.equal('thin');
      expect(c2.border.bottom.style).to.equal('double');
      expect(c2.alignment.horizontal).to.equal('center');
      expect(c2.alignment.wrapText).to.equal(true);
      expect(c2.protection.locked).to.equal(false);
    });

    it('supports writing without shared strings / styles', async () => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('N').getCell('A1').value = 'inline-ish';
      const wb2 = await roundTrip(wb, {useSharedStrings: false, useStyles: false});
      expect(wb2.getWorksheet('N').getCell('A1').value).to.equal('inline-ish');
    });
  });

  describe('rows and columns', () => {
    it('round-trips row properties', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Rows');
      ws.getRow(1).values = ['a', 'b', 'c'];
      ws.getRow(1).height = 33;
      ws.getRow(2).values = [1, 2, 3];
      ws.getRow(2).hidden = true;
      ws.getRow(3).values = [4, 5, 6];
      ws.getRow(3).outlineLevel = 2;
      ws.getRow(5).values = ['gap'];
      const ws2 = (await roundTrip(wb)).getWorksheet('Rows');
      expect(ws2.getRow(1).height).to.equal(33);
      expect(ws2.getRow(2).hidden).to.equal(true);
      expect(ws2.getRow(3).outlineLevel).to.equal(2);
      expect(ws2.getRow(5).getCell(1).value).to.equal('gap');
      expect(ws2.rowCount).to.equal(5);
      expect(ws2.actualRowCount).to.equal(4);
    });

    it('round-trips column definitions (width, hidden, style, headers)', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Cols');
      ws.columns = [
        {header: 'Id', key: 'id', width: 10},
        {header: 'Name', key: 'name', width: 32, style: {font: {bold: true}}},
        {header: 'Hidden', key: 'hid', hidden: true, width: 5},
        {header: 'Out', key: 'out', outlineLevel: 1},
      ];
      ws.addRow({id: 1, name: 'x', hid: 'h', out: 'o'});
      const ws2 = (await roundTrip(wb)).getWorksheet('Cols');
      expect(ws2.getColumn(1).width).to.equal(10);
      expect(ws2.getColumn(2).width).to.equal(32);
      expect(ws2.getColumn(2).style.font.bold).to.equal(true);
      expect(ws2.getColumn(3).hidden).to.equal(true);
      expect(ws2.getColumn(4).outlineLevel).to.equal(1);
      expect(ws2.getRow(1).getCell(2).value).to.equal('Name');
      expect(ws2.getRow(2).getCell(2).value).to.equal('x');
    });

    it('handles grouped column ranges (<col min..max>)', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('ColRange');
      for (let i = 1; i <= 12; i++) ws.getColumn(i).width = 20;
      ws.getCell('A1').value = 'x';
      const ws2 = (await roundTrip(wb)).getWorksheet('ColRange');
      for (let i = 1; i <= 12; i++) expect(ws2.getColumn(i).width).to.equal(20);
    });

    it('supports the last row / column Excel allows', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Edge');
      ws.getCell('XFD1048576').value = 'corner';
      ws.getCell('A1').value = 'origin';
      const ws2 = (await roundTrip(wb)).getWorksheet('Edge');
      expect(ws2.getCell('XFD1048576').value).to.equal('corner');
      expect(ws2.getCell('A1').value).to.equal('origin');
      expect(ws2.rowCount).to.equal(1048576);
      expect(ws2.columnCount).to.equal(16384);
    });

    it('addRows/spliceRows/insertRow keep data consistent', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Splice');
      ws.addRows([[1], [2], [3], [4]]);
      ws.spliceRows(2, 1);
      ws.insertRow(1, ['new']);
      const ws2 = (await roundTrip(wb)).getWorksheet('Splice');
      expect(ws2.getSheetValues().slice(1).map(r => r[1])).to.deep.equal(['new', 1, 3, 4]);
    });
  });

  describe('merged cells', () => {
    it('round-trips several merges and keeps master value', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Merge');
      ws.getCell('A1').value = 'big';
      ws.mergeCells('A1:C3');
      ws.getCell('E5').value = 'wide';
      ws.mergeCells('E5:H5');
      ws.mergeCells(10, 1, 12, 1);
      const ws2 = (await roundTrip(wb)).getWorksheet('Merge');
      expect(Object.keys(ws2._merges).sort()).to.deep.equal(['A1', 'A10', 'E5']);
      expect(ws2.getCell('B2').master.address).to.equal('A1');
      expect(ws2.getCell('B2').value).to.equal('big');
      expect(ws2.getCell('G5').master.address).to.equal('E5');
      expect(ws2.getCell('A11').isMerged).to.equal(true);
      expect(ws2.getCell('D4').isMerged).to.equal(false);
    });

    it('rejects overlapping merges', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('Overlap');
      ws.mergeCells('A1:B2');
      expect(() => ws.mergeCells('B2:C3')).to.throw();
    });

    it('unmerge works', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Unmerge');
      ws.mergeCells('A1:B2');
      ws.unMergeCells('A1');
      const ws2 = (await roundTrip(wb)).getWorksheet('Unmerge');
      expect(Object.keys(ws2._merges)).to.have.length(0);
    });
  });

  describe('workbook / worksheet structure', () => {
    it('keeps sheet order, ids, names, state and tab colour', async () => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('First');
      wb.addWorksheet('Hidden', {state: 'hidden'});
      wb.addWorksheet('Very', {state: 'veryHidden'});
      wb.addWorksheet('Colour', {properties: {tabColor: {argb: 'FF123456'}}});
      const wb2 = await roundTrip(wb);
      expect(wb2.worksheets.map(w => w.name)).to.deep.equal(['First', 'Hidden', 'Very', 'Colour']);
      expect(wb2.worksheets.map(w => w.id)).to.deep.equal([1, 2, 3, 4]);
      expect(wb2.getWorksheet('Hidden').state).to.equal('hidden');
      expect(wb2.getWorksheet('Very').state).to.equal('veryHidden');
      expect(wb2.getWorksheet('Colour').properties.tabColor.argb).to.equal('FF123456');
      expect(wb2.getWorksheet(2).name).to.equal('Hidden');
    });

    it('eachSheet iterates all sheets with ids', async () => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('a');
      wb.addWorksheet('b');
      const seen = [];
      (await roundTrip(wb)).eachSheet((ws, id) => seen.push([ws.name, id]));
      expect(seen).to.deep.equal([['a', 1], ['b', 2]]);
    });

    it('removeWorksheet and re-save works', async () => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('a');
      wb.addWorksheet('b');
      wb.removeWorksheet('a');
      const wb2 = await roundTrip(wb);
      expect(wb2.worksheets.map(w => w.name)).to.deep.equal(['b']);
    });

    it('round-trips core properties and dates', async () => {
      const wb = new ExcelJS.Workbook();
      wb.creator = 'Alice';
      wb.lastModifiedBy = 'Bob';
      wb.title = 'Title <&>';
      wb.subject = 'Subj';
      wb.keywords = 'k1,k2';
      wb.category = 'cat';
      wb.description = 'desc';
      wb.company = 'ACME';
      wb.manager = 'Boss';
      wb.created = new Date(Date.UTC(2021, 0, 2, 3, 4, 5));
      wb.modified = new Date(Date.UTC(2022, 1, 3, 4, 5, 6));
      wb.addWorksheet('x');
      const wb2 = await roundTrip(wb);
      expect(wb2.creator).to.equal('Alice');
      expect(wb2.lastModifiedBy).to.equal('Bob');
      expect(wb2.title).to.equal('Title <&>');
      expect(wb2.keywords).to.equal('k1,k2');
      expect(wb2.company).to.equal('ACME');
      expect(wb2.manager).to.equal('Boss');
      expect(wb2.created.getTime()).to.equal(Date.UTC(2021, 0, 2, 3, 4, 5));
      expect(wb2.modified.getTime()).to.equal(Date.UTC(2022, 1, 3, 4, 5, 6));
    });

    it('can load -> modify -> save repeatedly (idempotent pipeline)', async () => {
      const wb = new ExcelJS.Workbook();
      wb.creator = 'c';
      wb.addWorksheet('s').addRows([['a', 1], ['b', 2]]);
      let current = wb;
      /* eslint-disable-next-line no-await-in-loop */
      for (let i = 0; i < 3; i++) {
        // eslint-disable-next-line no-await-in-loop
        current = await roundTrip(current);
        current.getWorksheet('s').getCell('B1').value = i + 10;
      }
      expect(current.getWorksheet('s').getCell('B1').value).to.equal(12);
      expect(current.getWorksheet('s').getCell('A2').value).to.equal('b');
    });

    it('round-trips views, freeze panes, page setup, header/footer, properties', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('View', {
        views: [{state: 'frozen', xSplit: 1, ySplit: 2, showGridLines: false, zoomScale: 80}],
        pageSetup: {orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9},
        headerFooter: {oddHeader: '&CHeader', oddFooter: '&LFooter'},
        properties: {defaultRowHeight: 20},
      });
      ws.getCell('A1').value = 1;
      ws.pageSetup.printArea = 'A1:D10';
      ws.pageSetup.printTitlesRow = '1:2';
      const ws2 = (await roundTrip(wb)).getWorksheet('View');
      expect(ws2.views[0].state).to.equal('frozen');
      expect(ws2.views[0].xSplit).to.equal(1);
      expect(ws2.views[0].ySplit).to.equal(2);
      expect(ws2.views[0].showGridLines).to.equal(false);
      expect(ws2.views[0].zoomScale).to.equal(80);
      expect(ws2.pageSetup.orientation).to.equal('landscape');
      expect(ws2.pageSetup.fitToPage).to.equal(true);
      expect(ws2.pageSetup.paperSize).to.equal(9);
      expect(ws2.pageSetup.printArea).to.equal('A1:D10');
      expect(ws2.pageSetup.printTitlesRow).to.equal('1:2');
      expect(ws2.headerFooter.oddHeader).to.equal('&CHeader');
      expect(ws2.properties.defaultRowHeight).to.equal(20);
    });

    it('round-trips sheet protection', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Prot');
      ws.getCell('A1').value = 1;
      await ws.protect('secret', {formatCells: true, insertRows: false});
      const ws2 = (await roundTrip(wb)).getWorksheet('Prot');
      expect(ws2.sheetProtection).to.be.an('object');
      expect(ws2.sheetProtection.sheet).to.equal(true);
      expect(ws2.sheetProtection.hashValue).to.be.a('string');
    });

    it('page breaks, auto filter and sheet names with special chars', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('It\'s <a> & b');
      ws.addRows([['h1', 'h2'], [1, 2], [3, 4]]);
      ws.autoFilter = 'A1:B3';
      ws.getRow(2).addPageBreak();
      const ws2 = (await roundTrip(wb)).worksheets[0];
      expect(ws2.name).to.equal('It\'s <a> & b');
      expect(ws2.autoFilter).to.equal('A1:B3');
      const JSZip = require('jszip');
      const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
      expect(await zip.file('xl/worksheets/sheet1.xml').async('string')).to.contain('<brk id="2"');
    });
  });
});

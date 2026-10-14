// Golden-master test: every real-world .xlsx fixture shipped in the repo is
// loaded and summarised (sheets, dimensions, values, merges, validations,
// conditional formats, tables, defined names, images ...). The summary is
// compared with spec/integration/hardening-regression/fixtures-snapshot.json,
// generated from the code BEFORE the security patches. A patch that changes
// how any of these files is parsed (even subtly) makes this test fail.
//
//   UPDATE_SNAPSHOTS=1 npx mocha --require spec/config/setup <this file>
// regenerates the snapshot; only do that if the change is intended.
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const {roundTrip, ExcelJS} = require('../../utils/roundtrip');

const SNAPSHOT = path.join(__dirname, 'fixtures-snapshot.json');
const DATA_DIRS = ['spec/integration/data', 'test/data'];
// 1842: pathological data-validation range (slow by design); huge: 14MB perf fixture;
// the other two are deliberately malformed inputs.
const SKIP = new Set(['test-issue-1842.xlsx', 'huge.xlsx', 'invalid-xml.xlsx', 'missing-bits.xlsx']);

const norm = v => {
  if (v instanceof Date) return {d: v.toISOString()};
  if (v && typeof v === 'object') {
    const out = {};
    Object.keys(v).sort().forEach(k => { out[k] = norm(v[k]); });
    return out;
  }
  return v;
};
const hash = o => crypto.createHash('sha1').update(JSON.stringify(o)).digest('hex');

function summarise(wb) {
  const sheets = wb.worksheets.map(ws => {
    const cells = [];
    ws.eachRow({includeEmpty: false}, row => {
      row.eachCell({includeEmpty: false}, cell => {
        cells.push([cell.address, cell.type, norm(cell.value), cell.numFmt || null, !!cell.font && !!cell.font.bold]);
      });
    });
    return {
      name: ws.name,
      id: ws.id,
      state: ws.state,
      rows: ws.rowCount,
      cols: ws.columnCount,
      actualRows: ws.actualRowCount,
      cells: cells.length,
      cellsHash: hash(cells),
      merges: Object.keys(ws._merges).sort(),
      dv: Object.keys(ws.dataValidations.model).length,
      cf: ws.conditionalFormattings.length,
      tables: Object.keys(ws.tables).sort(),
      images: ws.getImages().length,
      colWidths: hash(ws.columns ? ws.columns.map(c => c.width || null) : []),
      views: norm(ws.views),
      autoFilter: norm(ws.autoFilter) || null,
    };
  });
  return {
    sheets,
    definedNames: norm(wb.definedNames.model),
    media: wb.model.media.length,
    creator: wb.creator || null,
  };
}

describe('hardening regression: fixture golden master', () => {
  const files = [];
  DATA_DIRS.forEach(d => {
    if (!fs.existsSync(d)) return;
    fs.readdirSync(d).filter(f => f.endsWith('.xlsx') && !SKIP.has(f)).sort()
      .forEach(f => files.push(path.join(d, f)));
  });
  const snapshot = fs.existsSync(SNAPSHOT) ? JSON.parse(fs.readFileSync(SNAPSHOT, 'utf8')) : {};
  const update = !!process.env.UPDATE_SNAPSHOTS;
  const fresh = {};

  files.forEach(file => {
    it(`loads ${file} exactly as before and survives a re-save`, async function() {
      this.timeout(30000);
      const wb = new ExcelJS.Workbook();
      let summary;
      try {
        await wb.xlsx.readFile(file);
        summary = summarise(wb);
      } catch (e) {
        summary = {error: String(e.message).slice(0, 120)};
      }
      fresh[file] = summary;
      if (!update) {
        expect(summary, file).to.deep.equal(snapshot[file]);
      }
      if (!summary.error) {
        const wb2 = await roundTrip(wb);
        expect(wb2.worksheets.map(w => w.name)).to.deep.equal(wb.worksheets.map(w => w.name));
        expect(summarise(wb2).sheets.map(s => s.cells)).to.deep.equal(summary.sheets.map(s => s.cells));
      }
      expect(Object.keys(Object.prototype)).to.deep.equal([]);
    });
  });

  after(() => {
    if (update) {
      fs.writeFileSync(SNAPSHOT, `${JSON.stringify(fresh, null, 1)}\n`);
    }
  });
});

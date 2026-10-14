// Helpers shared by the hardening-regression specs.
const ExcelJS = verquire('exceljs');

async function roundTrip(workbook, writeOptions, readOptions) {
  const buffer = await workbook.xlsx.writeBuffer(writeOptions);
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.load(buffer, readOptions);
  return wb2;
}

// Bytes that are a valid 1x1 PNG.
const PNG_1X1 = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
  'base64'
);

module.exports = {roundTrip, PNG_1X1, ExcelJS};

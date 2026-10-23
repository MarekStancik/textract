import fs from 'node:fs';
import path from 'node:path';
import XLSX from 'xlsx';

function update_sheet_range(ws) {
  var range = { s: { r: 20000000, c: 20000000 }, e: { r: 0, c: 0 } };
  Object.keys(ws).filter(function (x) { return x.charAt(0) != "!"; }).map(XLSX.utils.decode_cell).forEach(function (x) {
    range.s.c = Math.min(range.s.c, x.c); range.s.r = Math.min(range.s.r, x.r);
    range.e.c = Math.max(range.e.c, x.c); range.e.r = Math.max(range.e.r, x.r);
  });
  ws['!ref'] = XLSX.utils.encode_range(range);
}

function extractText(filePath, options, cb) {
  var wb, error, csvSheets = [];

  try {
    wb = XLSX.read(fs.readFileSync(filePath), { type: 'buffer' });
    wb.SheetNames.forEach(name => {
      const sheet = wb.Sheets[name];
      update_sheet_range(sheet);
      csvSheets.push(XLSX.utils.sheet_to_csv(sheet, { strip: true, blankrows: false }))
    });
  } catch (err) {
    error = new Error('Could not extract ' + path.basename(filePath) + ', ' + err);
    cb(error, null);
    return;
  }

  cb(null, csvSheets.join());
}

export default {
  types: ['application/vnd.ms-excel',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel.sheet.binary.macroEnabled.12',
    'application/vnd.ms-excel.sheet.macroEnabled.12',
    'application/vnd.oasis.opendocument.spreadsheet',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.template',
    'application/vnd.oasis.opendocument.spreadsheet-template'
  ],
  extract: extractText
};

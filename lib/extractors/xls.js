var path = require( 'path' )
  , XLSX = require( 'xlsx' )
  ;

function extractText( filePath, options, cb ) {
  var wb, error, csvSheets = [];

  try {    
    wb = XLSX.readFile( filePath );
    wb.SheetNames.forEach(name => csvSheets.push(XLSX.utils.sheet_to_csv(wb.Sheets[name], { strip: true, blankrows: false } )));
    cb( null, csvSheets.join() );
  } catch ( err ) {
    error = new Error( 'Could not extract ' + path.basename( filePath ) + ', ' + err );
    cb( error, null );
    return;
  }
}

module.exports = {
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

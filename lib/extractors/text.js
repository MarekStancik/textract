import fs from 'node:fs';
import iconv from 'iconv-lite';
import jschardet from 'jschardet';
import path from 'node:path';

function extractText( filePath, options, cb ) {
  fs.readFile( filePath, function( error, data ) {
    var encoding, decoded, detectedEncoding;
    if ( error ) {
      cb( error, null );
      return;
    }
    try {
      detectedEncoding = jschardet.detect( data ).encoding;
      if ( !detectedEncoding ) {
        error = new Error( 'Could not detect encoding for file named [[ ' +
          path.basename( filePath ) + ' ]]' );
        cb( error, null );
        return;
      }
      encoding = detectedEncoding.toLowerCase();

      decoded = iconv.decode( data, encoding );
    } catch ( e ) {
      cb( e );
      return;
    }
    cb( null, decoded );
  });
}

export default {
  types: [/text\//, 'application/csv', 'application/javascript'],
  extract: extractText
};

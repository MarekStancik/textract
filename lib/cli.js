import path from 'node:path';
import textract from './index.js';

export default function( filePath, flags ) {
  filePath = path.resolve( process.cwd(), filePath );

  if ( flags.preserveLineBreaks === 'false' ) {
    flags.preserveLineBreaks = false;
  } else {
    flags.preserveLineBreaks = true;
  }

  textract.fromFileWithPath( filePath, flags, function( error, text ) {
    if ( error ) {
      // eslint-disable-next-line no-console
      console.error( error );
    } else {
      // eslint-disable-next-line no-console
      console.log( text );
    }
  });
};

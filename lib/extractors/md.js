import fs from 'node:fs';
import { marked } from 'marked';
import htmlExtract from './html.js';

function extractText( filePath, options, cb ) {
  fs.readFile( filePath, function( error, data ) {
    if ( error ) {
      cb( error, null );
      return;
    }

    try {
      htmlExtract.extractFromText( marked( data.toString() ), options, cb );
    } catch ( err ) {
      cb( err, null );
    }
  });
}

export const types = ['text/x-markdown', 'text/markdown'];
export const extract = extractText;

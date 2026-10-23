import epub2 from 'epub2';
import htmlExtract from './html.js';

var EPub = epub2.EPub;

function extractText( filePath, options, cb ) {
  var epub = new EPub( filePath )
    , chapterTexts = []
    , hasError = false
    , chapterCount = 0
    ;

  epub.on( 'end', function() {
    // Iterate over each chapter...
    epub.flow.forEach( function( chapter, index ) {
      // if already error, don't do anything
      if ( !hasError ) {
        // Get the chapter text
        epub.getChapterRaw( chapter.id, function( rawChaperError, text ) {
          if ( rawChaperError ) {
            hasError = true;
            cb( rawChaperError, null );
          } else {
            // Extract the raw text from the chapter text (it's html)
            htmlExtract.extractFromText( text, options, function( htmlExtractError, outText ) {
              if ( htmlExtractError ) {
                hasError = true;
                cb( htmlExtractError, null );
              } else {
                chapterTexts[index] = outText;
                chapterCount++;
                if ( chapterCount === epub.flow.length ) {
                  cb( null, chapterTexts.join( '' ) );
                }
              }
            });
          }
        });
      }
    });
  });

  epub.parse();
}

export default {
  types: ['application/epub+zip'],
  extract: extractText
};

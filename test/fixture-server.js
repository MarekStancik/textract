import { createServer } from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import mime from 'mime';

const testDirectory = path.dirname( fileURLToPath( import.meta.url ) );
const fixturesDirectory = path.join( testDirectory, 'files' );

export async function startFixtureServer() {
  const server = createServer( function( request, response ) {
    var requestPath = new URL( request.url, 'http://localhost' ).pathname
      , fixtureName = path.basename( decodeURIComponent( requestPath ) )
      , fixturePath
      ;

    if ( requestPath === '/misleading.aspx' || requestPath === '/summary.aspx' ) {
      fixtureName = 'test.html';
    }

    fixturePath = path.join( fixturesDirectory, fixtureName );
    fs.readFile( fixturePath, function( error, contents ) {
      if ( error ) {
        response.writeHead( 404 );
        response.end( 'Fixture not found' );
        return;
      }

      response.writeHead( 200, {
        'content-type': mime.getType( fixtureName ) || 'application/octet-stream',
        'content-length': contents.length
      });
      response.end( contents );
    });
  });

  await new Promise( function( resolve, reject ) {
    server.once( 'error', reject );
    server.listen( 0, '127.0.0.1', resolve );
  });

  return {
    url: 'http://127.0.0.1:' + server.address().port,
    close: function() {
      return new Promise( function( resolve, reject ) {
        server.close( function( error ) {
          if ( error ) {
            reject( error );
          } else {
            resolve();
          }
        });
      });
    }
  };
}
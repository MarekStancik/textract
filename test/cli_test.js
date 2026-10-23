import './common.js';
import { exec } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname( fileURLToPath( import.meta.url ) );
const cliPath = path.join( __dirname, '..', 'bin', 'textract' );
const testFilePath = path.join( __dirname, 'files', 'css.css' );

describe("cli", function(){
  it("will extract text", function(done) {
    exec( cliPath + " " + testFilePath,
      function( error, stdout, stderr ) {
        expect(stdout).to.eql(".foo {color:red}\n");
        done();
      }
    );
  })
});


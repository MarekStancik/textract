import chai from 'chai';
import textract from '../lib/index.js';

global.expect = chai.expect;

global.textract = textract;
global.fromBufferWithName = textract.fromBufferWithName;
global.fromBufferWithMime = textract.fromBufferWithMime;
global.fromFileWithPath = textract.fromFileWithPath;
global.fromFileWithMimeAndPath = textract.fromFileWithMimeAndPath;
global.fromUrl = textract.fromUrl;
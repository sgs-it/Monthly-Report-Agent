import test from 'node:test';
import assert from 'node:assert/strict';
import {zipSync,strToU8} from 'fflate';
import {readWordIdentity} from '../src/docx.js';
import {projectLabel} from '../src/project-name.js';

test('the project label uses adjacent Word cells and joins text split into formatting runs',()=>{
  const cell=parts=>'<w:tc><w:p>'+parts.map(part=>'<w:r><w:t>'+part+'</w:t></w:r>').join('')+'</w:p></w:tc>';
  const row=values=>'<w:tr>'+values.map(cell).join('')+'</w:tr>';
  const xml='<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>'+
    row([['Client'],['Mashreq Bank'],['Facilities Management'],['Location']])+
    row([['Location'],['Burj Al Shams (','Mashreq Bank Headquarters)'],['JOB ','CODE'],['22B','11401']])+
    '</w:tbl></w:body></w:document>';
  const wordIdentity=readWordIdentity(zipSync({'word/document.xml':strToU8(xml)}));
  assert.equal(projectLabel({name:'Old name',wordIdentity}),'Burj Al Shams (Mashreq Bank Headquarters) — 22B11401');
});

test('malformed job and location cells can use the master filename without changing Excel labels',()=>{
  assert.equal(projectLabel({name:'Location',templateName:'Landscaping Monthly Report - Site- October 2026.docx',
    wordIdentity:{location:'',jobCode:'82A11202'}}),'Site — 82A11202');
  assert.equal(projectLabel({name:'Site',templateName:'Indoor Plants Monthly Report - Site-32B00086 - October 2026.docx',
    wordIdentity:{location:'Al Barsha',jobCode:''}}),'Al Barsha — 32B00086');
  assert.equal(projectLabel({name:'MBZUAI 1A',format:'xlsx'}),'MBZUAI 1A');
});

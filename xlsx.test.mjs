import {existsSync} from 'node:fs';
import test from 'node:test';import assert from 'node:assert/strict';import {readFile,readdir} from 'node:fs/promises';import {unzipSync,strFromU8} from 'fflate';
import {generateExcelReport,excelToWordTemplate} from '../src/xlsx.js';import {inspectTemplate} from '../src/docx.js';
test('Excel output preserves all non-photo sheets, styles, validations, and formulas',{skip:!existsSync('Master report templates')},async()=>{
 const files=(await readdir('Master report templates')).filter(n=>/\.xlsx$/i.test(n));for(const name of files){const source=new Uint8Array(await readFile('Master report templates/'+name)),a=unzipSync(source),image=new Uint8Array(await readFile(new URL('./fixtures/test-photo.jpg',import.meta.url)));
 const out=generateExcelReport(source,'November 2026',[{bytes:image,width:320,height:200,caption:'Pruning shrubs',activity:'Pruning'}],'MBZUAI');const b=unzipSync(out);
 for(const key of Object.keys(a))if((/^xl\/worksheets\/sheet[1-578]\.xml$/.test(key))||key==='xl/styles.xml'||key==='xl/workbook.xml')assert.deepEqual(b[key],a[key],`${name}: ${key}`);
 assert.ok(Object.entries(b).some(([k,v])=>/worksheets\/sheet/.test(k)&&strFromU8(v).includes('November 2026')));
 assert.ok(inspectTemplate(excelToWordTemplate(source,'MBZUAI')).tokens.includes('{{PHOTOS}}'));
 }});

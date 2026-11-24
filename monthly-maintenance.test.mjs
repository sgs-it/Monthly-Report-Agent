import test from 'node:test';
import assert from 'node:assert/strict';
import {zipSync,unzipSync,strToU8,strFromU8} from 'fflate';
import {DOMParser} from '@xmldom/xmldom';
import {generateReport,demoTemplate} from '../src/docx.js';

const monthNames='Jan. Feb. Mar. April May June July Aug. Sep. Oct. Nov. Dec.'.split(' ');
const row=(values,plan=false)=>'<w:tr>'+values.map((v,i)=>'<w:tc><w:tcPr>'+
 (plan?'<w:shd w:fill="'+(i===11?'001D3D':'FFFFFF')+'"/>':'')+
 '</w:tcPr><w:p><w:r><w:t>'+v+'</w:t></w:r></w:p></w:tc>').join('')+'</w:tr>';
const sample=unzipSync(demoTemplate());
const activities=[['Mowing','Weekly'],['Watering','Daily'],['Pruning','Monthly']];
const schedule='<w:tbl>'+row(['Activities','Frequency',...monthNames],true)+activities.map(([name,freq],i)=>
 row([name,freq,...monthNames.map((_,m)=>m===9&&i===0?'●':m===10&&i===1?'✓':m===10&&i===2?'●':'')],true)).join('')+'</w:tbl>';
const operations='<w:tbl>'+row(['Activity','Frequency','Complete (✓)','Pending (✗)','Remarks'])+row(['Old activity','Weekly','✓','','October remarks'])+'</w:tbl>';
const other='<w:tbl>'+row(['S/N','Description'])+row(['1','Keep this remark'])+'</w:tbl>';
sample['word/document.xml']=strToU8(strFromU8(sample['word/document.xml']).replace('<w:body>','<w:body>'+schedule+operations+other));
const master=zipSync(sample);
for(const [month,index,expected] of [['November',10,activities.slice(1)],['October',9,activities.slice(0,1)],['December',11,[]]]){
 test(month+' highlights just the selected column and lists its scheduled work in the Word output',()=>{
  const report=generateReport(master,{}, {REPORT_MONTH:month,REPORT_PERIOD:month+' 2026',LOCATION:'Site',WORK_SUMMARY:'Review'},[]);
  const xml=strFromU8(unzipSync(report)['word/document.xml']);
  const doc=new DOMParser().parseFromString(xml,'application/xml');
  const [plan,performance]=Array.from(doc.getElementsByTagName('w:tbl'));
  for(const r of Array.from(plan.getElementsByTagName('w:tr'))){
   const cs=Array.from(r.getElementsByTagName('w:tc')).slice(2);
   const shades=cs.map(c=>c.getElementsByTagName('w:shd')[0].getAttribute('w:fill'));
   assert.deepEqual(shades.map((color,i)=>color==='001D3D'?i:-1).filter(i=>i>=0),[index]);
  }
  const actual=Array.from(performance.getElementsByTagName('w:tr')).slice(1)
   .map(r=>Array.from(r.getElementsByTagName('w:tc'),c=>c.textContent));
  assert.deepEqual(actual,expected.map(([name,freq])=>[name,freq,'','','']));
  assert.ok(xml.includes('Keep this remark'));
  assert.ok(!xml.includes('October remarks'));
 });
}

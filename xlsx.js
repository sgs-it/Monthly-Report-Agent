// Package-level edits retain original workbook styles, formulas, validations, and non-photo sheets.
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {demoTemplate} from './docx.js';
const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',REL='http://schemas.openxmlformats.org/package/2006/relationships',X='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',A='http://schemas.openxmlformats.org/drawingml/2006/main';
const parse=s=>new DOMParser().parseFromString(s,'application/xml'),ser=d=>new XMLSerializer().serializeToString(d),ns=(d,n,t)=>Array.from(d.getElementsByTagNameNS(n,t));
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
const safeUnzip=bytes=>{if(bytes.length>12*1024*1024)throw Error('Template exceeds 12 MB.');let size=0;return unzipSync(bytes,{filter:e=>{size+=e.originalSize;if(size>40*1024*1024)throw Error('Expanded workbook exceeds 40 MB.');return true;}});};
function open(bytes){const f=safeUnzip(bytes);if(!f['xl/workbook.xml'])throw Error('Invalid Excel .xlsx template.');const wb=parse(strFromU8(f['xl/workbook.xml'])),rels=parse(strFromU8(f['xl/_rels/workbook.xml.rels']));const sheets=ns(wb,S,'sheet').map(s=>{const r=ns(rels,REL,'Relationship').find(r=>r.getAttribute('Id')===s.getAttributeNS(R,'id'));let path=r.getAttribute('Target');path=path.startsWith('/')?path.slice(1):'xl/'+path;return {name:s.getAttribute('name'),path,doc:parse(strFromU8(f[path]))};});const strings=f['xl/sharedStrings.xml']?ns(parse(strFromU8(f['xl/sharedStrings.xml'])),S,'si').map(x=>ns(x,S,'t').map(t=>t.textContent).join('')):[];return {f,sheets,strings};}
function cellText(c,strings){if(c.getAttribute('t')==='s')return strings[Number(ns(c,S,'v')[0]?.textContent)]||'';if(c.getAttribute('t')==='inlineStr')return ns(c,S,'t').map(x=>x.textContent).join('');return ns(c,S,'v')[0]?.textContent||'';}
export function excelToWordTemplate(bytes,name){
 const {sheets,strings}=open(bytes),z=unzipSync(demoTemplate());let body='';
 for(const sheet of sheets){body+=`<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="30"/></w:rPr><w:t>${esc(sheet.name)}</w:t></w:r></w:p>`;
  if(sheet.name.toLowerCase()==='images'){body+='<w:p><w:r><w:t>{{PHOTOS}}</w:t></w:r></w:p>';continue;}
  const rows=ns(sheet.doc,S,'row').filter(r=>ns(r,S,'c').some(c=>cellText(c,strings))),max=Math.min(14,Math.max(...rows.flatMap(r=>ns(r,S,'c').filter(c=>cellText(c,strings)).map(c=>{let n=0;for(const ch of c.getAttribute('r').match(/[A-Z]+/)[0])n=n*26+ch.charCodeAt(0)-64;return n;}))));
  if(!rows.length)continue;
  body+=`<w:tbl><w:tblPr><w:tblW w:w="15000" w:type="dxa"/><w:tblBorders>${['top','left','bottom','right','insideH','insideV'].map(x=>`<w:${x} w:val="single" w:sz="4" w:color="999999"/>`).join('')}</w:tblBorders></w:tblPr><w:tblGrid>${Array.from({length:max},()=>`<w:gridCol w:w="${Math.floor(15000/max)}"/>`).join('')}</w:tblGrid>`;
  for(const [ri,row]of rows.entries()){const cells=new Map(ns(row,S,'c').map(c=>[c.getAttribute('r').match(/[A-Z]+/)[0],cellText(c,strings)]));body+='<w:tr>';for(let i=0;i<max;i++){let n=i+1,key='';while(n){n--;key=String.fromCharCode(65+n%26)+key;n=Math.floor(n/26);}body+=`<w:tc><w:tcPr><w:tcW w:w="${Math.floor(15000/max)}" w:type="dxa"/>${ri<2?'<w:shd w:fill="D9E2F3"/>':''}</w:tcPr><w:p><w:r><w:rPr><w:sz w:val="18"/>${ri<2?'<w:b/>':''}</w:rPr><w:t xml:space="preserve">${esc(cells.get(key)||'')}</w:t></w:r></w:p></w:tc>`;}body+='</w:tr>';}body+='</w:tbl>';
 }
 const doc=parse(strFromU8(z['word/document.xml'])),wbody=ns(doc,'http://schemas.openxmlformats.org/wordprocessingml/2006/main','body')[0];
 const start=`<w:p><w:r><w:rPr><w:b/><w:sz w:val="36"/></w:rPr><w:t>${esc(name)}</w:t></w:r></w:p><w:p><w:r><w:t>{{REPORT_PERIOD}}</w:t></w:r></w:p><w:p><w:r><w:t>{{WORK_SUMMARY}}</w:t></w:r></w:p>`;
 const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';const full=`<w:document xmlns:w="${W}"><w:body>${start+body}<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr></w:body></w:document>`;
 z['word/document.xml']=strToU8(full);return zipSync(z,{level:6});
}
export function generateExcelReport(bytes,period,photos,projectName){
 const {f,sheets}=open(bytes),sheet=sheets.find(s=>s.name.toLowerCase()==='images');if(!sheet)throw Error('The Excel template has no Images worksheet.');
 const relPath=sheet.path.replace('/worksheets/','/worksheets/_rels/')+'.rels';const sheetRels=parse(strFromU8(f[relPath]));const dr=ns(sheetRels,REL,'Relationship').find(r=>r.getAttribute('Type')===R+'/drawing');if(!dr)throw Error('Excel photo drawing is missing.');
 const drawingPath='xl/'+dr.getAttribute('Target').replace('../',''),drawing=parse(strFromU8(f[drawingPath]));const prototype=Array.from(drawing.documentElement.childNodes).find(n=>n.localName==='twoCellAnchor'||n.localName==='oneCellAnchor');if(!prototype)throw Error('Excel photo anchor is missing.');
 const drawingRelPath=drawingPath.replace('/drawings/','/drawings/_rels/')+'.rels';const oldRels=parse(strFromU8(f[drawingRelPath]));
 for(const rel of ns(oldRels,REL,'Relationship'))if(rel.getAttribute('Type')===R+'/image')delete f['xl/'+rel.getAttribute('Target').replace('../','')];
 const newDrawing=parse(`<xdr:wsDr xmlns:xdr="${X}" xmlns:a="${A}" xmlns:r="${R}"/>`),newRels=parse(`<Relationships xmlns="${REL}"/>`),sheetData=ns(sheet.doc,S,'sheetData')[0];
 const oldRows=ns(sheetData,S,'row'),rowPrototype=oldRows.find(r=>r.getAttribute('r')==='5')||oldRows.at(-1);
 oldRows.filter(r=>Number(r.getAttribute('r'))>=5).forEach(r=>sheetData.removeChild(r));
 for(const m of ns(sheet.doc,S,'mergeCell'))if(Number(m.getAttribute('ref').match(/\d+/)[0])>=5)m.parentNode.removeChild(m);
 const setCell=(row,col,value)=>{let c=ns(row,S,'c').find(c=>c.getAttribute('r').startsWith(col)&&c.getAttribute('r').match(/^[A-Z]+/)[0]===col);if(!c){c=sheet.doc.createElementNS(S,'c');c.setAttribute('r',col+row.getAttribute('r'));row.appendChild(c);}while(c.firstChild)c.removeChild(c.firstChild);c.setAttribute('t','inlineStr');const is=sheet.doc.createElementNS(S,'is'),t=sheet.doc.createElementNS(S,'t');t.textContent=String(value);is.appendChild(t);c.appendChild(is);};
 photos.forEach((photo,i)=>{const row=rowPrototype.cloneNode(true),r=i+5;row.setAttribute('r',String(r));row.setAttribute('ht','108');row.setAttribute('customHeight','1');ns(row,S,'c').forEach(c=>{c.setAttribute('r',c.getAttribute('r').replace(/\d+/,String(r)));while(c.firstChild)c.removeChild(c.firstChild);c.removeAttribute('t');});setCell(row,'A',i+1);setCell(row,'B',photo.caption);setCell(row,'C',photo.activity||'Work');setCell(row,'D',projectName);sheetData.appendChild(row);
  const anchor=newDrawing.importNode(prototype,true),from=ns(anchor,X,'from')[0];ns(from,X,'col')[0].textContent='4';ns(from,X,'row')[0].textContent=String(r-1);ns(from,X,'colOff')[0].textContent='10000';ns(from,X,'rowOff')[0].textContent='10000';
  const to=ns(anchor,X,'to')[0];if(to){ns(to,X,'col')[0].textContent='5';ns(to,X,'row')[0].textContent=String(r);ns(to,X,'colOff')[0].textContent='0';ns(to,X,'rowOff')[0].textContent='0';}
  for(const e of ns(anchor,X,'ext')){e.setAttribute('cx','1750000');e.setAttribute('cy','1300000');}
  const id=`rIdReport${i+1}`;ns(anchor,A,'blip').forEach(e=>e.setAttributeNS(R,'r:embed',id));ns(anchor,X,'cNvPr').forEach(e=>{e.setAttribute('id',String(i+1));e.setAttribute('name',photo.caption);});newDrawing.documentElement.appendChild(anchor);
  const filename=`report-${i+1}.jpeg`;f['xl/media/'+filename]=photo.bytes;const rel=newRels.createElementNS(REL,'Relationship');rel.setAttribute('Id',id);rel.setAttribute('Type',R+'/image');rel.setAttribute('Target','../media/'+filename);newRels.documentElement.appendChild(rel);
 });
 const title=ns(sheetData,S,'row').find(r=>r.getAttribute('r')==='1');if(title)setCell(title,'A',`${projectName} — ${period}`);
 const dim=ns(sheet.doc,S,'dimension')[0];if(dim)dim.setAttribute('ref',`A1:I${photos.length+4}`);
 const ct=parse(strFromU8(f['[Content_Types].xml']));if(!Array.from(ct.documentElement.childNodes).some(e=>e.getAttribute?.('Extension')==='jpeg')){const e=ct.createElementNS(ct.documentElement.namespaceURI,'Default');e.setAttribute('Extension','jpeg');e.setAttribute('ContentType','image/jpeg');ct.documentElement.appendChild(e);}
 f['[Content_Types].xml']=strToU8(ser(ct));f[sheet.path]=strToU8(ser(sheet.doc));f[drawingPath]=strToU8(ser(newDrawing));f[drawingRelPath]=strToU8(ser(newRels));return zipSync(f,{level:6});
}

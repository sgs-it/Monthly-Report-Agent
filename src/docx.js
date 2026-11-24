import {zipSync,unzipSync,strToU8,strFromU8} from 'fflate';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {applyMonthlyMaintenance} from './monthly-maintenance.js';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const A='http://schemas.openxmlformats.org/drawingml/2006/main';
const WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const REL='http://schemas.openxmlformats.org/package/2006/relationships';
const CT='http://schemas.openxmlformats.org/package/2006/content-types';
const parse=s=>new DOMParser({onError:(level)=>{if(level!=='warning')throw new Error('Invalid Word XML.');}}).parseFromString(s,'application/xml');
const serialize=d=>new XMLSerializer().serializeToString(d);
const nodes=(el,ns,tag)=>Array.from(el.getElementsByTagNameNS(ns,tag));
const text=p=>nodes(p,W,'t').map(n=>n.textContent).join('');
const xmlEscape=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
function readPackage(bytes){
  if(bytes.length>12*1024*1024)throw new Error('The master template must be smaller than 12 MB.');
  let total=0;
  const files=unzipSync(bytes,{filter:e=>{total+=e.originalSize;if(total>40*1024*1024)throw new Error('Expanded template exceeds 40 MB.');return true;}});
  if(!files['word/document.xml']||!files['[Content_Types].xml'])throw new Error('Please upload a Word .docx file.');
  if(Object.keys(files).some(k=>/vbaProject|activeX|embeddings\//i.test(k)))throw new Error('Use a .docx without macros, embedded files, or ActiveX.');
  return files;
}
function wordIdentity(doc) {
  const locationLabel = /^(?:location|site name)\s*:?$/i;
  const jobLabel = /^job\s*(?:code|number|no\.?|#)\s*:?$/i;
  const result = {location:'',jobCode:''};
  for (const row of nodes(doc,W,'tr')) {
    const values = Array.from(row.childNodes).filter(cell=>cell.localName==='tc')
      .map(cell=>nodes(cell,W,'p').map(text).join(' ').replace(/\s+/g,' ').trim());
    for (let i=0;i+1<values.length;i++) {
      const value=values[i+1];
      if (!value || locationLabel.test(value) || jobLabel.test(value) || value.includes('{{')) continue;
      if (!result.location && locationLabel.test(values[i])) result.location=value;
      if (!result.jobCode && jobLabel.test(values[i]) && /\d/.test(value)) result.jobCode=value;
    }
    if (result.location && result.jobCode) break;
  }
  return result;
}
export function readWordIdentity(bytes){
  const files=unzipSync(bytes,{filter:entry=>entry.name==='word/document.xml'&&entry.originalSize<=40*1024*1024});
  if(!files['word/document.xml'])throw Error('Cannot read the project fields from this master.');
  return wordIdentity(parse(strFromU8(files['word/document.xml'])));
}
export function inspectTemplate(bytes){
  const files=readPackage(bytes), doc=parse(strFromU8(files['word/document.xml']));
  const paragraphs=nodes(doc,W,'p').map((p,index)=>({index,text:text(p),images:nodes(p,W,'drawing').length+nodes(p,W,'pict').length}));
  const fullText=paragraphs.map(p=>p.text).join('\n');
  const body=nodes(doc,W,'body')[0];
  const tables=Array.from(body.childNodes).filter(n=>n.localName==='tbl').map((t,index)=>({index,images:nodes(t,W,'drawing').length,rows:nodes(t,W,'tr').length,columns:nodes(nodes(t,W,'tr')[0],W,'tc').length,text:text(t).slice(0,150)}));
  return {wordIdentity:wordIdentity(doc),paragraphs:paragraphs.filter(p=>p.text||p.images),tokens:[...new Set(fullText.match(/\{\{[A-Z_]+\}\}/g)||[])],bodyImageCount:paragraphs.reduce((a,p)=>a+p.images,0),tables};
}
export function appendPhotoSection(bytes,referenceBytes){
 const files=readPackage(bytes),doc=parse(strFromU8(files['word/document.xml'])),reference=readPackage(referenceBytes),ref=parse(strFromU8(reference['word/document.xml']));
 const refBody=nodes(ref,W,'body')[0],table=Array.from(refBody.childNodes).find(n=>n.localName==='tbl'&&nodes(n,W,'drawing').length>=2);
 if(!table)throw Error('A reference photo table is required.');
 const copied=doc.importNode(table,true),body=nodes(doc,W,'body')[0],section=Array.from(body.childNodes).find(n=>n.localName==='sectPr');
 const heading=parse(`<w:p xmlns:w="${W}"><w:pPr><w:keepNext/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>PHOTOGRAPHS</w:t></w:r></w:p>`).documentElement;
 body.insertBefore(doc.importNode(heading,true),section||null);body.insertBefore(copied,section||null);
 const refRels=parse(strFromU8(reference['word/_rels/document.xml.rels'])),rels=parse(strFromU8(files['word/_rels/document.xml.rels']));
 for(const blip of nodes(copied,A,'blip')){const old=blip.getAttributeNS(R,'embed'),rel=nodes(refRels,REL,'Relationship').find(r=>r.getAttribute('Id')===old);if(!rel)continue;const target=rel.getAttribute('Target'),newName=`media/donor-${crypto.randomUUID()}.${target.split('.').at(-1)}`,id=`rIdDonor${crypto.randomUUID().replaceAll('-','')}`;files['word/'+newName]=reference['word/'+target];const r=doc.importNode(rel,true);r.setAttribute('Id',id);r.setAttribute('Target',newName);rels.documentElement.appendChild(r);blip.setAttributeNS(R,'r:embed',id);}
 const types=parse(strFromU8(files['[Content_Types].xml'])),refTypes=parse(strFromU8(reference['[Content_Types].xml']));for(const d of nodes(refTypes,CT,'Default'))if(!nodes(types,CT,'Default').some(e=>e.getAttribute('Extension')===d.getAttribute('Extension')))types.documentElement.appendChild(types.importNode(d,true));
 files['word/document.xml']=strToU8(serialize(doc));files['word/_rels/document.xml.rels']=strToU8(serialize(rels));files['[Content_Types].xml']=strToU8(serialize(types));return zipSync(files,{level:6});
}
function replaceText(p, values){
  // Locate replacements across split Word runs without flattening unrelated formatting.
  for(const [needle,replacement] of Object.entries(values)){
    if(!needle)continue;
    let ts=nodes(p,W,'t'),combined=ts.map(t=>t.textContent).join(''),start=combined.indexOf(needle);
    while(start!==-1){
      let offset=0;const end=start+needle.length;
      for(const t of ts){const original=t.textContent, a=offset,b=a+original.length;offset=b;if(b<=start||a>=end)continue;
        const prefix=original.slice(0,Math.max(0,start-a)),suffix=original.slice(Math.min(original.length,end-a));
        t.textContent=prefix+(a<=start?String(replacement):'')+suffix;t.setAttribute('xml:space','preserve');
      }
      ts=nodes(p,W,'t');combined=ts.map(t=>t.textContent).join('');start=combined.indexOf(needle,start+String(replacement).length);
    }
  }
}
function setParagraph(p,value){const ts=nodes(p,W,'t');if(ts.length){ts[0].textContent=value;ts[0].setAttribute('xml:space','preserve');ts.slice(1).forEach(t=>t.textContent='');}else{const r=p.ownerDocument.createElementNS(W,'w:r'),t=p.ownerDocument.createElementNS(W,'w:t');t.textContent=value;r.appendChild(t);p.appendChild(r);}}
function resetPictureIds(root, serial){
  for(const e of [...nodes(root,WP,'docPr'),...nodes(root,'http://schemas.openxmlformats.org/drawingml/2006/picture','cNvPr')]){e.setAttribute('id',String(serial.next++));e.setAttribute('name',`Report photo ${serial.next}`);}
}
function fillGridPhoto(cell,photo,width,height){
 const box=width/height,source=photo.width/photo.height;
 const horizontal=source>box?Math.round(50000*(1-box/source)):0;
 const vertical=source<box?Math.round(50000*(1-source/box)):0;
 for(const fill of nodes(cell,'http://schemas.openxmlformats.org/drawingml/2006/picture','blipFill')){
  for(const old of [...nodes(fill,A,'srcRect'),...nodes(fill,A,'tile'),...nodes(fill,A,'stretch')])old.parentNode.removeChild(old);
  const crop=fill.ownerDocument.createElementNS(A,'a:srcRect');
  for(const [key,value] of Object.entries({l:horizontal,r:horizontal,t:vertical,b:vertical}))crop.setAttribute(key,String(value));
  fill.appendChild(crop);
  const stretch=fill.ownerDocument.createElementNS(A,'a:stretch');
  stretch.appendChild(fill.ownerDocument.createElementNS(A,'a:fillRect'));fill.appendChild(stretch);
 }
 for(const e of [...nodes(cell,WP,'extent'),...nodes(cell,A,'xfrm').flatMap(x=>nodes(x,A,'ext'))]){
  e.setAttribute('cx',String(width));e.setAttribute('cy',String(height));
 }
}
function generateTableReport(files,doc,mapping,photos,insertedTable){
  const body=nodes(doc,W,'body')[0],tables=Array.from(body.childNodes).filter(n=>n.localName==='tbl');
  const selected=insertedTable?[insertedTable]:mapping.photoTableIndices.map(i=>tables[i]);
  if(selected.some(t=>!t))throw new Error('Photo table mapping is no longer valid.');
  const base=selected[0],rows=Array.from(base.childNodes).filter(n=>n.localName==='tr');
  const imageRow=rows.find(r=>nodes(r,W,'drawing').length),imageRowIndex=rows.indexOf(imageRow),captionRow=rows[imageRowIndex+1];
  if(!imageRow||!captionRow)throw new Error('The photo table must have alternating picture and caption rows.');
  if(nodes(imageRow,WP,'anchor').length)throw new Error('Use inline pictures in the master photo table.');
  const cells=r=>Array.from(r.childNodes).filter(n=>n.localName==='tc');
  const columnCount=cells(imageRow).length;
  const templatePhoto=nodes(imageRow,W,'drawing')[0];
  const relPath='word/_rels/document.xml.rels';
  const rels=parse(files[relPath]?strFromU8(files[relPath]):`<Relationships xmlns="${REL}"/>`);
  const types=parse(strFromU8(files['[Content_Types].xml']));
  if(!nodes(types,CT,'Default').some(e=>e.getAttribute('Extension')==='jpg')){const e=types.createElementNS(CT,'Default');e.setAttribute('Extension','jpg');e.setAttribute('ContentType','image/jpeg');types.documentElement.appendChild(e);}
  const serial={next:Math.max(100,...nodes(doc,WP,'docPr').map(e=>Number(e.getAttribute('id'))||0))+1};
  const perPage=columnCount*Math.max(1,rows.filter(r=>nodes(r,W,'drawing').length).length);
  const grid=nodes(base,W,'tblGrid')[0],columns=grid?nodes(grid,W,'gridCol'):[];
  const firstBox=nodes(templatePhoto,WP,'extent')[0],height=Number(firstBox?.getAttribute('cy'))||2200000;
  const parent=base.parentNode;
  for(let offset=0;offset<photos.length;offset+=perPage){
    const table=base.cloneNode(true);Array.from(table.childNodes).filter(n=>n.localName==='tr').forEach(r=>table.removeChild(r));
    if(!offset)for(const heading of rows.slice(0,imageRowIndex))if(/^photographs?$/i.test(text(heading).trim()))table.appendChild(heading.cloneNode(true));
    for(let rowStart=offset;rowStart<Math.min(photos.length,offset+perPage);rowStart+=columnCount){
      const pics=imageRow.cloneNode(true),caps=captionRow.cloneNode(true);
      for(let c=0;c<columnCount;c++){
        const pc=cells(pics)[c],cc=cells(caps)[c];if(!cc)throw new Error('Photo and caption rows must have matching columns.');
        const photo=photos[rowStart+c];
        if(!photo||rowStart+c>=offset+perPage){for(const cell of [pc,cc]){Array.from(cell.childNodes).filter(n=>n.localName!=='tcPr').forEach(n=>cell.removeChild(n));cell.appendChild(doc.createElementNS(W,'w:p'));}continue;}
        let drawings=nodes(pc,W,'drawing');
        if(!drawings.length){const p=nodes(pc,W,'p')[0]||pc.appendChild(doc.createElementNS(W,'w:p'));const r=doc.createElementNS(W,'w:r');r.appendChild(templatePhoto.cloneNode(true));p.appendChild(r);}
        drawings=nodes(pc,W,'drawing');drawings.slice(1).forEach(n=>n.parentNode.removeChild(n));
        nodes(pc,W,'t').forEach(n=>n.textContent='');
        const id=`rIdPhoto${crypto.randomUUID().replaceAll('-','')}`,filename=`report-${crypto.randomUUID()}.jpg`;
        files[`word/media/${filename}`]=photo.bytes;
        const rel=rels.createElementNS(REL,'Relationship');rel.setAttribute('Id',id);rel.setAttribute('Type',`${R}/image`);rel.setAttribute('Target',`media/${filename}`);rels.documentElement.appendChild(rel);
        nodes(pc,A,'blip').forEach(n=>{n.setAttributeNS(R,'r:embed',id);n.removeAttributeNS(R,'link');});
        const tcMar=nodes(pc,W,'tcMar')[0]||nodes(base,W,'tblCellMar')[0];
        const margin=side=>tcMar?Number(nodes(tcMar,W,side)[0]?.getAttributeNS(W,'w')??72):72;
        const cellWidth=Number(columns[c]?.getAttributeNS(W,'w'));
        const width=cellWidth>margin('left')+margin('right')?Math.round((cellWidth-margin('left')-margin('right'))*635):Number(firstBox?.getAttribute('cx'))||1650000;
        fillGridPhoto(pc,photo,width,height);
        nodes(pc,WP,'docPr').forEach(e=>e.setAttribute('descr',photo.caption));resetPictureIds(pc,serial);
        const cp=nodes(cc,W,'p');if(cp.length){setParagraph(cp[0],photo.caption);cp.slice(1).forEach(p=>p.parentNode.removeChild(p));}
      }
      table.appendChild(pics);table.appendChild(caps);
    }
    parent.insertBefore(table,base);
  }
  // Remove only the mapped old photo region, including intervening old page breaks.
  let n=base;const last=selected.at(-1),selectedSet=new Set(selected);while(n){const next=n.nextSibling;if(selectedSet.has(n)||(n.localName==='p'&&!text(n).trim()&&!nodes(n,W,'drawing').length))parent.removeChild(n);if(n===last)break;n=next;}
  files['word/document.xml']=strToU8(serialize(doc));files[relPath]=strToU8(serialize(rels));files['[Content_Types].xml']=strToU8(serialize(types));
  return finishPackage(files);
}
function finishPackage(files){
  // Strip unused historic photo media and relationships; keep logos and signatures still referenced anywhere.
  const relPath='word/_rels/document.xml.rels',rels=parse(strFromU8(files[relPath]));
  const documentXML=strFromU8(files['word/document.xml']);
  for(const rel of nodes(rels,REL,'Relationship'))if(rel.getAttribute('Type')===`${R}/image`&&!documentXML.includes(`"${rel.getAttribute('Id')}"`))rel.parentNode.removeChild(rel);
  files[relPath]=strToU8(serialize(rels));
  const referenced=new Set();
  for(const [path,bytes] of Object.entries(files))if(path.endsWith('.rels')){
    const folder=path.slice(0,path.lastIndexOf('/_rels/'));
    for(const rel of nodes(parse(strFromU8(bytes)),REL,'Relationship')){const target=rel.getAttribute('Target');if(target.startsWith('media/'))referenced.add(`${folder}/${target}`);}
  }
  for(const path of Object.keys(files))if(path.startsWith('word/media/')&&!referenced.has(path))delete files[path];
  return zipSync(files,{level:6});
}
export function generateReport(templateBytes, mapping, values, photos){
  const files=readPackage(templateBytes),doc=parse(strFromU8(files['word/document.xml']));
  const paragraphs=nodes(doc,W,'p');
  for(const [index,value] of Object.entries(mapping.edits||{})){const p=paragraphs[Number(index)];if(p)setParagraph(p,value);}
  const monthMappings=mapping.periodText?.trim()?{[mapping.periodText.trim()]:values.REPORT_PERIOD}:{};
  for(const [part,bytes] of Object.entries(files))if(/^word\/(document|header\d*|footer\d*)\.xml$/.test(part)){
    const x=part==='word/document.xml'?doc:parse(strFromU8(bytes));
    nodes(x,W,'p').forEach((p,index)=>replaceText(p,{...Object.fromEntries(Object.entries(values).map(([k,v])=>[`{{${k}}}`,v])),...((part!=='word/document.xml'||!Number.isInteger(mapping.periodIndex)||index===mapping.periodIndex)?monthMappings:{})}));
    if(part!=='word/document.xml')files[part]=strToU8(serialize(x));
  }
  if(Number.isInteger(mapping.summaryIndex)){const p=paragraphs[mapping.summaryIndex];if(!p)throw new Error('Summary mapping is no longer valid.');setParagraph(p,values.WORK_SUMMARY);}
  applyMonthlyMaintenance(doc,values.REPORT_MONTH || values.REPORT_PERIOD);
  if(mapping.photoTableIndices?.length){if(/\{\{(?!PHOTOS)[A-Z_]+\}\}/.test(nodes(doc,W,'p').map(text).join('\n')))throw Error('An unmapped placeholder remains in the master.');return generateTableReport(files,doc,mapping,photos);}
  let anchor, prototype, captionPrototype;
  if(Number.isInteger(mapping.photoIndex)){
    prototype=paragraphs[mapping.photoIndex];
    if(!prototype||!nodes(prototype,A,'blip').length)throw new Error('Choose an inline photo paragraph in the master template.');
    if(nodes(prototype,WP,'anchor').length)throw new Error('Set the template photograph to In Line with Text in Word before mapping it.');
    anchor=prototype;captionPrototype=Number.isInteger(mapping.captionIndex)?paragraphs[mapping.captionIndex]:null;
    if(captionPrototype===prototype)throw new Error('The caption and photo mappings must be separate paragraphs.');
  }else{
    anchor=paragraphs.find(p=>text(p).includes('{{PHOTOS}}'));
    if(!anchor)throw new Error('Map a sample photo layout or add {{PHOTOS}} in its own paragraph.');
    if(text(anchor).trim()!=='{{PHOTOS}}')throw new Error('{{PHOTOS}} must be the only text in its paragraph.');
    prototype=parse(`<w:p xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:pPr><w:keepNext/><w:jc w:val="center"/></w:pPr><w:r><w:drawing><wp:inline><wp:extent cx="5486400" cy="3657600"/><wp:docPr id="1" name="Photo"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Photo"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdPhoto"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="3657600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`).documentElement;
  }
  // Masters with a standalone photo placeholder receive a four-column grid.
  const section=nodes(doc,W,'sectPr').at(-1),size=section&&nodes(section,W,'pgSz')[0],margins=section&&nodes(section,W,'pgMar')[0];
  const usable=(Number(size?.getAttributeNS(W,'w'))||11906)-(Number(margins?.getAttributeNS(W,'left'))||1134)-(Number(margins?.getAttributeNS(W,'right'))||1134);
  const width=Math.floor(usable/4),picWidth=(width-144)*635,picHeight=Math.round(picWidth*1.4);
  const cell=content=>`<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
  const pic=doc.importNode(prototype,true);
  for(const ext of nodes(pic,WP,'extent')){ext.setAttribute('cx',String(picWidth));ext.setAttribute('cy',String(picHeight));}
  const cap=captionPrototype?serialize(captionPrototype):'<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="18"/></w:rPr><w:t/></w:r></w:p>';
  const borders=['top','left','bottom','right','insideH','insideV'].map(edge=>`<w:${edge} w:val="single" w:sz="4" w:color="444444"/>`).join('');
  const row=cells=>'<w:tr><w:trPr><w:cantSplit/></w:trPr>'+cells+'</w:tr>';
  const grid=parse(`<w:tbl xmlns:w="${W}"><w:tblPr><w:tblW w:w="${width*4}" w:type="dxa"/><w:tblBorders>${borders}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="72" w:type="dxa"/><w:right w:w="72" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${('<w:gridCol w:w="'+width+'"/>').repeat(4)}</w:tblGrid>${row(cell(serialize(pic)).repeat(4))}${row(cell(cap).repeat(4))}</w:tbl>`).documentElement;
  const inserted=doc.importNode(grid,true);
  anchor.parentNode.insertBefore(inserted,anchor);
  if(captionPrototype?.parentNode)captionPrototype.parentNode.removeChild(captionPrototype);
  anchor.parentNode.removeChild(anchor);
  if(mapping.removeBodyPhotos)for(const p of paragraphs)if(p.parentNode&&nodes(p,W,'drawing').length)p.parentNode.removeChild(p);
  if(/\{\{[A-Z_]+\}\}/.test(nodes(doc,W,'p').map(text).join('\n')))throw Error('An unmapped placeholder remains in the master.');
  return generateTableReport(files,doc,mapping,photos,inserted);
}
export function demoTemplate(){
  const parts={
    '[Content_Types].xml':`<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    '_rels/.rels':`<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/document.xml':`<w:document xmlns:w="${W}"><w:body><w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="36"/></w:rPr><w:t>Demonstration monthly work report</w:t></w:r></w:p><w:p><w:r><w:t>{{LOCATION}}</w:t></w:r></w:p><w:p><w:r><w:t>{{REPORT_PERIOD}}</w:t></w:r></w:p><w:p><w:r><w:t>{{WORK_SUMMARY}}</w:t></w:r></w:p><w:p><w:r><w:t>{{PHOTOS}}</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/></w:sectPr></w:body></w:document>`};
  return zipSync(Object.fromEntries(Object.entries(parts).map(([k,v])=>[k,strToU8(v)])));
}

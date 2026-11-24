const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const months=['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
const children=(node,tag)=>Array.from(node.childNodes).filter(n=>n.namespaceURI===W&&n.localName===tag);
const descendants=(node,tag)=>Array.from(node.getElementsByTagNameNS(W,tag));
const value=node=>descendants(node,'p').map(p=>descendants(p,'t').map(t=>t.textContent).join('')).join(' ').trim();
const monthIndex=text=>months.indexOf(text.toLowerCase().trim().slice(0,3));
const cells=row=>children(row,'tc');
const create=(parent,tag)=>parent.ownerDocument.createElementNS(W,'w:'+tag);

function setFill(cell,source,defaultFill){
 const props=children(cell,'tcPr')[0]||cell.insertBefore(create(cell,'tcPr'),cell.firstChild);
 const current=children(props,'shd')[0];
 if(current)props.removeChild(current);
 if(!source&&!defaultFill)return;
 const shading=source?source.cloneNode(true):create(cell,'shd');
 if(!source){
  shading.setAttributeNS(W,'w:val','clear');
  shading.setAttributeNS(W,'w:fill',defaultFill);
 }
 const next=Array.from(props.childNodes).find(n=>['noWrap','tcMar','textDirection','tcFitText','vAlign','hideMark'].includes(n.localName));
 props.insertBefore(shading,next||null);
}
function fontColor(cell,color){
 for(const run of descendants(cell,'r')){
  const props=children(run,'rPr')[0]||run.insertBefore(create(run,'rPr'),run.firstChild);
  const ink=children(props,'color')[0]||props.appendChild(create(run,'color'));
  for(const attr of ['themeColor','themeTint','themeShade'])ink.removeAttributeNS(W,attr);
  ink.setAttributeNS(W,'w:val',color);
 }
}
const shading=cell=>descendants(cell,'shd')[0];
const fill=cell=>shading(cell)?.getAttributeNS(W,'fill')||'';
function isDark(color){
 if(!/^[a-f\d]{6}$/i.test(color))return false;
 const [r,g,b]=color.match(/../g).map(hex=>parseInt(hex,16));
 return 0.299*r+0.587*g+0.114*b<110;
}
function marked(cell){
 if(/^[●•✓✔☑√]+$/u.test(value(cell)))return true;
 return descendants(cell,'sym').some(s=>/wingdings/i.test(s.getAttributeNS(W,'font'))&&/^(?:F0)?(?:FC|FE)$/i.test(s.getAttributeNS(W,'char')));
}
function writeCell(cell,content){
 const paragraph=children(cell,'p')[0]?.cloneNode(false)||create(cell,'p');
 const paraStyle=descendants(cell,'pPr')[0],runStyle=descendants(cell,'rPr')[0];
 if(paraStyle)paragraph.appendChild(paraStyle.cloneNode(true));
 const run=create(cell,'r'),text=create(cell,'t');
 if(runStyle)run.appendChild(runStyle.cloneNode(true));
 text.textContent=content;run.appendChild(text);paragraph.appendChild(run);
 for(const node of Array.from(cell.childNodes))if(node.localName!=='tcPr')cell.removeChild(node);
 cell.appendChild(paragraph);
}
function populateOperations(table,activities){
 const [header,...oldRows]=children(table,'tr');
 if(!header)return;
 const labels=cells(header).map(cell=>value(cell).toLowerCase());
 const activity=labels.findIndex(label=>/^activit(y|ies)$/.test(label)),frequency=labels.indexOf('frequency');
 if(activity<0||frequency<0||!labels.some(label=>label.startsWith('complete')))return;
 const prototype=oldRows.find(row=>cells(row).length===labels.length)||header;
 for(const row of oldRows)table.removeChild(row);
 for(const entry of activities){
  const row=prototype.cloneNode(true);
  if(prototype===header)for(const prop of descendants(row,'tblHeader'))prop.parentNode.removeChild(prop);
  cells(row).forEach((cell,index)=>writeCell(cell,index===activity?entry.activity:index===frequency?entry.frequency:''));
  table.appendChild(row);
 }
}

export function applyMonthlyMaintenance(doc,reportMonth){
 const selected=monthIndex(reportMonth||'');
 if(selected<0)return;
 const tables=children(descendants(doc,'body')[0],'tbl');
 for(const [index,table] of tables.entries()){
  const rows=children(table,'tr');
  if(!rows.length)continue;
  const header=cells(rows[0]);
  if(header.length!==14||!/^activities$/i.test(value(header[0]))||!/^frequency$/i.test(value(header[1])))continue;
  const monthCells=header.slice(2),order=monthCells.map(cell=>monthIndex(value(cell)));
  if(order.includes(-1)||new Set(order).size!==12)continue;
  const target=order.indexOf(selected)+2;
  const counts=new Map();
  for(const cell of monthCells)counts.set(fill(cell),(counts.get(fill(cell))||0)+1);
  const [usual,usualCount]=[...counts].sort((a,b)=>b[1]-a[1])[0];
  const previous=monthCells.findIndex(cell=>fill(cell)!==usual&&(usualCount===11||isDark(fill(cell))));
  const accent=previous>=0&&isDark(fill(monthCells[previous]))?fill(monthCells[previous]):'001D3D';
  const entries=[];
  for(const [rowIndex,row] of rows.entries()){
   const cols=cells(row);
   if(cols.length!==header.length)continue;
   if(rowIndex&&marked(cols[target])&&value(cols[0]))entries.push({activity:value(cols[0]),frequency:value(cols[1])});
   if(previous>=0&&previous+2!==target){
    const old=cols[previous+2];
    const normal=cols.find((cell,col)=>col>=2&&col!==target&&col!==previous+2&&(!rowIndex?fill(cell)===usual:!isDark(fill(cell))));
    if(normal){
     setFill(old,shading(normal));
     fontColor(old,rowIndex?'0E6B0E':'000000');
    }
   }
   setFill(cols[target],null,accent);fontColor(cols[target],'FFFFFF');
  }
  if(tables[index+1])populateOperations(tables[index+1],entries);
 }
}

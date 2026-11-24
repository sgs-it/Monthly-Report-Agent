import {renderAsync} from 'docx-preview';
import {projectLabel} from '../src/project-name.js';

const escape=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
export const icon=name=>`<svg class="rg-icon" aria-hidden="true"><use href="#rg-icon-${name}"></use></svg>`;

export function dashboard(state,month,period,rows) {
  const hour=new Date().getHours();
  const greeting=hour<12?'Good morning':hour<18?'Good afternoon':'Good evening';
  const actions=[['integrations','link','Connect CompanyCam'],['upload','file-up','Upload master file'],['photos','file-plus-2','Generate monthly report']];
  const stats=[['building-2',state.locations.filter(l=>l.companyCamProjectId).length,'Linked projects'],['files',state.reports.filter(r=>r.month===month).length,'Reports for '+period],['layers',state.locations.length,'Master templates']];
  return `<div class="welcome"><div><h1>${greeting}</h1><p>Generate monthly site reports from your master files and CompanyCam photos.</p></div><div class="garden-detail" aria-hidden="true"></div></div>
    <div class="action-grid">${actions.map(([page,symbol,name])=>`<button class="action-tile" ${page==='upload'?'data-add':'data-page="'+page+'"'}><span class="action-symbol">${icon(symbol)}</span><strong>${name}</strong><span class="action-arrow">${icon('arrow-up-right')}</span></button>`).join('')}</div>
    <div class="metric-grid">${stats.map(([symbol,value,label])=>`<div class="metric"><span>${icon(symbol)}</span><div><strong>${value}</strong><small>${escape(label)}</small></div></div>`).join('')}</div>
    <section class="section-card"><div class="section-top"><h2>Projects</h2><button class="text-button" data-page="projects">View all projects →</button></div>${state.locations.length?rows(state.locations.slice(0,5)):'<div class="empty"><h3>Add your first site</h3><p>Upload its master file to start a report.</p><button class="secondary" data-add>Upload master file</button></div>'}</section>`;
}

export function integrations(sites,selected,connection){
  return `<div class="page-heading"><div><h1>Integrations</h1><p class="muted">Connect project photos to your site reports.</p></div></div>
  <section class="integration-panel"><div class="integration-head"><span class="camera-mark">${icon('camera')}</span><div><h2>CompanyCam</h2><p>Project and photo access</p></div><span class="tag">${escape(connection.connected?'Connected':connection.configured?'Check connection':'Not connected')}</span></div>
  <div class="integration-body"><div><p>${escape(connection.message)}</p><button class="secondary" id="rg-check">Check connection</button><hr><label>Report project<select id="rg-site">${sites.map(l=>`<option value="${escape(l.id)}" ${l.id===selected?.id?'selected':''}>${escape(projectLabel(l))}</option>`).join('')}</select></label>
  ${selected?`<p>Linked project · ${escape(selected.companyCamProjectName||selected.companyCamProjectId||'None selected')}</p><button class="primary" id="rg-link" ${connection.connected?'':'disabled'}>${icon('link')} Link CompanyCam project</button>`:'<button class="primary" data-add>Upload a site master first</button>'}</div><div class="garden-panel" aria-hidden="true"></div></div></section>`;
}

export function preparePreviewDialog(){
  const dialog=document.querySelector('#preview-dialog'),stage=dialog.querySelector('.preview-stage');
  const download=dialog.querySelector('[data-save-word]'),again=dialog.querySelector('[data-again]');
  let revision=0,blobURL='',report;
  function save(){
    const a=document.createElement('a');a.href=blobURL;a.download=report.filename;a.click();
    dialog.querySelector('.download-note').hidden=false;
  }
  download.onclick=again.onclick=save;
  dialog.querySelector('[data-dismiss-note]').onclick=()=>{dialog.querySelector('.download-note').hidden=true;};
  dialog.onclose=()=>{revision++;stage.replaceChildren();if(blobURL)URL.revokeObjectURL(blobURL);blobURL='';};
  return async r=>{
    const version=++revision;report=r;
    dialog.querySelector('.preview-name').textContent=r.projectName+' · '+r.period;
    dialog.querySelector('.download-note').hidden=true;
    const status=dialog.querySelector('.preview-status');status.textContent='Loading Word report…';
    download.disabled=true;dialog.showModal();
    try{
      const res=await fetch('/api/reports/'+encodeURIComponent(r.id)+'/download');
      if(!res.ok)throw Error('The report download failed. Close the preview and retry.');
      const blob=await res.blob();if(version!==revision)return;
      blobURL=URL.createObjectURL(blob);download.disabled=false;
      const frame=document.createElement('iframe');frame.title='Generated Word document';frame.setAttribute('sandbox','allow-same-origin');
      const ready=new Promise(resolve=>frame.onload=resolve);
      frame.srcdoc='<meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src &apos;none&apos;; img-src data: blob:; style-src &apos;unsafe-inline&apos;; font-src data: blob:"><style>body{margin:0;background:#edf1ef}a{pointer-events:none}</style><body></body>';
      stage.replaceChildren(frame);await ready;if(version!==revision)return;
      await renderAsync(blob,frame.contentDocument.body,frame.contentDocument.head,{renderAltChunks:false,useBase64URL:true});
      if(version!==revision)return;
      const pages=frame.contentDocument.querySelectorAll('section.docx');
      const fit=()=>{if(pages.length){frame.contentDocument.body.style.zoom='1';const w=pages[0].getBoundingClientRect().width+60;frame.contentDocument.body.style.zoom=String(Math.min(1,frame.clientWidth/w));}};
      fit();status.textContent=pages.length?`${pages.length} preview page${pages.length===1?'':'s'} · Word layout may vary`:'Word preview';
    }catch(e){if(version===revision)status.textContent=blobURL?'Preview unavailable. Your Word file is ready to download.':e.message;}
  };
}

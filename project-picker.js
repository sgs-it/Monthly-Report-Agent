import {projectLabel} from '../src/project-name.js';

export function createProjectPicker({api,reload,toast,onError}) {
  const dialog=document.createElement('dialog');
  dialog.className='cc-picker';
  dialog.setAttribute('aria-labelledby','cc-picker-title');
  dialog.innerHTML=`<div class="dialog-heading"><h2 id="cc-picker-title">Link CompanyCam project</h2><button type="button" class="secondary" data-close-picker>Close</button></div>
    <p class="cc-site"></p>
    <label>Search CompanyCam projects<input type="search" maxlength="200" placeholder="Search projects…" autocomplete="off"></label>
    <p role="status" class="cc-status"></p>
    <fieldset class="cc-results" aria-label="Matching projects"></fieldset>
    <div class="dialog-actions"><button class="secondary" data-more>Load more</button><button class="primary" data-link disabled>Link project</button></div>`;
  document.body.append(dialog);
  const find=selector=>dialog.querySelector(selector);
  const input=find('input'),results=find('fieldset'),status=find('[role=status]'),more=find('[data-more]'),link=find('[data-link]');
  let site,timer,generation=0,next=null,chosen='',saving=false;
  const ids=new Set();
  find('[data-close-picker]').onclick=()=>dialog.close();
  dialog.oncancel=e=>{if(saving)e.preventDefault();};
  dialog.onclose=()=>{generation++;clearTimeout(timer);};
  function addRow(project) {
    if(ids.has(project.id))return;
    ids.add(project.id);
    const row=document.createElement('label'),radio=document.createElement('input');
    radio.type='radio';radio.name='cc-match';radio.value=project.id;
    const info=document.createElement('span'),title=document.createElement('strong'),detail=document.createElement('small');
    title.textContent=project.name;
    detail.textContent=[project.address,'ID '+project.id,project.id===site.companyCamProjectId?'Currently linked':''].filter(Boolean).join(' · ');
    info.append(title,detail);row.append(radio,info);results.append(row);
  }
  results.onchange=e=>{chosen=e.target.value;link.disabled=!chosen;};
  async function load(after,version) {
    const params=new URLSearchParams();
    if(input.value.trim())params.set('query',input.value.trim());
    if(after)params.set('after',after);
    more.disabled=true;
    status.textContent='Searching CompanyCam…';
    try {
      const data=await api('/companycam/projects?'+params);
      if(version!==generation)return;
      if(after&&data.next===after)throw Error('CompanyCam returned the same page. Try searching again.');
      data.projects.forEach(addRow);
      next=data.next;
      more.hidden=!next;
      more.textContent='Load more';
      status.textContent=ids.size?`${ids.size} project${ids.size===1?'':'s'} shown`:'No matching projects. Try a different search.';
    }catch(error){
      if(version!==generation)return;
      status.textContent=error.message;more.hidden=false;more.textContent='Retry';
    }finally{if(version===generation)more.disabled=false;}
  }
  function search(delay) {
    clearTimeout(timer);
    const version=++generation;
    ids.clear();results.replaceChildren();chosen='';next=null;link.disabled=true;more.hidden=true;
    status.textContent='Searching CompanyCam…';
    timer=setTimeout(()=>load(null,version),delay);
  }
  input.oninput=()=>search(250);
  more.onclick=()=>load(next,generation);
  link.onclick=async()=>{
    if(!chosen||saving)return;
    saving=true;link.disabled=true;input.disabled=true;results.disabled=true;more.disabled=true;find('[data-close-picker]').disabled=true;
    try{
      await api('/companycam/link',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({locationId:site.id,companyCamProjectId:chosen})});
      dialog.close();await reload();toast('CompanyCam project linked.');
    }catch(error){if(dialog.open)status.textContent=error.message;else onError(error);}
    finally{saving=false;input.disabled=false;results.disabled=false;more.disabled=false;link.disabled=!chosen;find('[data-close-picker]').disabled=false;}
  };
  return location=>{
    if(!location)return;
    site=location;input.value='';find('.cc-site').textContent=projectLabel(site);
    dialog.showModal();input.focus();search(0);
  };
}

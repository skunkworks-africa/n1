import { PublicClientApplication, InteractionRequiredAuthError } from '@azure/msal-browser';
import {Planner,configured,SCOPES,referenceKey} from './graph.js';

const $=id=>document.getElementById(id);
const el=(tag,text,cls)=>{const n=document.createElement(tag);if(text!==undefined)n.textContent=text;if(cls)n.className=cls;return n;};
let auth,planner,account,config,view='board',editing=null,refreshing=false,saving=false,generation=0;
const form=$('task-form');
const field=name=>form.elements.namedItem(name);
const say=message=>$('notice').textContent=message;
function clearPrivate(){generation++;if($('editor').open)$('editor').close();editing=null;form.reset();$('checklist').replaceChildren();$('references').replaceChildren();$('labels').replaceChildren();field('assignees').replaceChildren();field('bucketId').replaceChildren();$('tasks').replaceChildren();$('metrics').replaceChildren();$('identity').textContent='';$('updated').textContent='';$('editor-notice').textContent='';$('native').removeAttribute('href');$('workspace').hidden=true;$('entry').hidden=false;$('principles').hidden=false;if(planner){planner.tasks.clear();planner.members=[];planner.buckets=[];planner.details=null;planner.plan=null;planner.me=null;}}
function handle(error,target='notice'){
  if([401,403,404].includes(error.status)||error instanceof InteractionRequiredAuthError){clearPrivate();$('connection').textContent='Access unavailable';target='notice';}
  $(target).textContent=error.message || 'The request could not be completed.';
}
function overdue(t){return t.percentComplete!==100&&t.dueDateTime&&new Date(t.dueDateTime)<new Date();}
function names(t){return Object.keys(t.assignments||{}).map(id=>planner.members.find(m=>m.id===id)?.displayName||'Plan member').join(', ')||'Unassigned';}
function progress(t){return t.percentComplete===100?100:t.percentComplete===0?0:50;}
function render(){
  const all=[...planner.tasks.values()];const q=$('search').value.trim().toLowerCase(),filter=$('filter').value;
  const tasks=all.filter(t=>t.title.toLowerCase().includes(q)).filter(t=>filter==='all'||filter==='mine'&&t.assignments?.[planner.me.id]||filter==='open'&&t.percentComplete!==100||filter==='overdue'&&overdue(t));
  $('metrics').replaceChildren(...[['Total tasks',all.length],['In progress',all.filter(t=>progress(t)===50).length],['Completed',all.filter(t=>progress(t)===100).length],['Overdue',all.filter(overdue).length]].map(([label,count])=>{const d=el('div',undefined,'metric');d.append(el('strong',String(count)),el('span',label));return d;}));
  const host=$('tasks');host.className=view;host.replaceChildren();
  const card=t=>{const b=el('button',undefined,'task-card');b.type='button';const bucket=planner.buckets.find(x=>x.id===t.bucketId);b.append(el('span',bucket?.name||'Task','bucket'),el('strong',t.title));const meta=el('div',undefined,'meta');meta.append(el('div',names(t)));if(t.dueDateTime)meta.append(el('span',`Due ${new Date(t.dueDateTime).toLocaleDateString()}`,overdue(t)?'overdue':''));b.append(meta);b.addEventListener('click',()=>openTask(t.id));return b;};
  if(view==='board')for(const [status,label]of [[0,'Not started'],[50,'In progress'],[100,'Completed']]){const col=el('section',undefined,'column'),items=tasks.filter(t=>progress(t)===status),h=el('h3',label);h.append(el('span',String(items.length)));col.append(h);col.append(...(items.length?items.map(card):[el('p','No tasks here.','empty')]));host.append(col);}
  else host.append(...(tasks.length?tasks.map(card):[el('p','No tasks match this view.','empty')]));
  $('board-view').setAttribute('aria-pressed',String(view==='board'));$('list-view').setAttribute('aria-pressed',String(view==='list'));
}
async function refresh(){if(refreshing||saving||$('editor').open)return;refreshing=true;const current=generation;$('refresh').disabled=true;try{await planner.load();if(current!==generation)return;$('identity').textContent=planner.me.displayName;$('connection').textContent='Connected to Microsoft';$('entry').hidden=true;$('principles').hidden=true;$('workspace').hidden=false;render();$('updated').textContent=`Last refreshed ${new Date().toLocaleTimeString()}. Refreshes every 60 seconds while this tab is active.`;say('');}catch(e){handle(e);}finally{refreshing=false;$('refresh').disabled=false;}}
function localDate(value){if(!value)return '';const d=new Date(value);return new Date(d.getTime()-d.getTimezoneOffset()*60000).toISOString().slice(0,16);}
function checklistRow(id=crypto.randomUUID(),item={title:'',isChecked:false}){const row=el('div',undefined,'check-row');row.dataset.id=id;const c=el('input');c.type='checkbox';c.checked=item.isChecked;c.setAttribute('aria-label','Checklist item complete');const title=el('input');title.value=item.title;title.maxLength=100;title.required=true;title.setAttribute('aria-label','Checklist item title');const remove=el('button','×');remove.type='button';remove.setAttribute('aria-label','Remove checklist item');remove.onclick=()=>row.remove();row.append(c,title,remove);$('checklist').append(row);}
function linkRow(key='',value={}){const row=el('div',undefined,'link-row');row.dataset.original=key;const url=el('input');url.type='url';url.required=true;url.value=key?decodeURIComponent(key):'';url.setAttribute('aria-label','Link URL');url.placeholder='https://';const alias=el('input');alias.value=value.alias||'';alias.maxLength=100;alias.setAttribute('aria-label','Link label');alias.placeholder='Link label';const remove=el('button','×');remove.type='button';remove.setAttribute('aria-label','Remove link');remove.onclick=()=>row.remove();row.append(url,alias,remove);$('references').append(row);}
async function openTask(id){
  if(saving)return;const current=generation;
  try{const data=id?await planner.task(id):{task:null,details:{}};if(current!==generation)return;editing=data;form.reset();$('editor-notice').textContent='';const t=data.task||{priority:5,percentComplete:0};$('editor-title').textContent=id?'Edit task':'Add task';
    for(const key of ['title','priority','percentComplete'])field(key).value=t[key]??'';
    if(![0,50,100].includes(t.percentComplete)){const opt=new Option(`${t.percentComplete}% complete`,String(t.percentComplete));field('percentComplete').add(opt);field('percentComplete').value=t.percentComplete;}
    if(![1,3,5,9].includes(t.priority)){field('priority').add(new Option(`Priority ${t.priority}`,String(t.priority)));field('priority').value=t.priority;}
    for(const key of ['startDateTime','dueDateTime'])field(key).value=localDate(t[key]);
    field('bucketId').replaceChildren(...planner.buckets.map(b=>new Option(b.name,b.id,false,b.id===t.bucketId)));
    const members=[...planner.members];for(const key of Object.keys(t.assignments||{}))if(!members.some(m=>m.id===key))members.push({id:key,displayName:'Former plan member (remove to unassign)'});
    field('assignees').replaceChildren(...members.map(m=>new Option(m.displayName||'Plan member',m.id,false,!!t.assignments?.[m.id])));
    field('description').value=data.details.description||'';$('checklist').replaceChildren();for(const [key,item] of Object.entries(data.details.checklist||{}))checklistRow(key,item);
    $('references').replaceChildren();for(const [key,item] of Object.entries(data.details.references||{}))linkRow(key,item);
    $('labels').replaceChildren(el('legend','Labels'));for(const [key,name]of Object.entries(planner.details.categoryDescriptions||{})){if(!name&&!t.appliedCategories?.[key])continue;const label=el('label');const c=el('input');c.type='checkbox';c.dataset.category=key;c.checked=!!t.appliedCategories?.[key];label.append(c,document.createTextNode(name||key));$('labels').append(label);}
    $('delete').hidden=!id;$('native').hidden=!id;$('native').href=`https://planner.cloud.microsoft/webui/plan/${encodeURIComponent(config.planId)}?tid=${encodeURIComponent(config.tenantId)}`;
    $('save').disabled=!planner.buckets.length;if(!planner.buckets.length)$('editor-notice').textContent='This plan needs a bucket. Add one in Microsoft Planner first.';
    $('editor').showModal();
  }catch(e){handle(e);}
}
function collect(){
  const t=editing.task,body={title:field('title').value.trim(),bucketId:field('bucketId').value,percentComplete:Number(field('percentComplete').value),priority:Number(field('priority').value),assignments:{},appliedCategories:{}};
  for(const k of ['startDateTime','dueDateTime'])body[k]=field(k).value?new Date(field(k).value).toISOString():null;
  const selected=new Set([...field('assignees').selectedOptions].map(o=>o.value));
  for(const id of Object.keys(t?.assignments||{}))if(!selected.has(id))body.assignments[id]=null;
  for(const id of selected)if(!t?.assignments?.[id])body.assignments[id]={'@odata.type':'#microsoft.graph.plannerAssignment',orderHint:' !'};
  for(const c of $('labels').querySelectorAll('input'))body.appliedCategories[c.dataset.category]=c.checked;
  const details={description:field('description').value,checklist:{},references:{}};
  for(const id of Object.keys(editing.details.checklist||{}))details.checklist[id]=null;
  for(const row of $('checklist').children)details.checklist[row.dataset.id]={'@odata.type':'#microsoft.graph.plannerChecklistItem',title:row.querySelector('input:not([type=checkbox])').value.trim(),isChecked:row.querySelector('[type=checkbox]').checked};
  if(Object.values(details.checklist).filter(Boolean).length>20)throw Error('A task can have at most 20 checklist items.');
  for(const key of Object.keys(editing.details.references||{}))details.references[key]=null;
  for(const row of $('references').children){const inputs=row.querySelectorAll('input');const key=referenceKey(inputs[0].value);details.references[key]={'@odata.type':'#microsoft.graph.plannerExternalReference',alias:inputs[1].value.trim()||inputs[0].value,type:'Other'};}
  return {body,details};
}
form.addEventListener('submit',async event=>{
  event.preventDefault();if(saving)return;
  let phase='task';try{const {body,details}=collect();planner.validate(body);saving=true;$('save').disabled=true;$('delete').disabled=true;$('close-editor').disabled=true;
    const saved=await planner.saveTask(editing.task,body);phase='details';
    // Record creation immediately: retries must not create a duplicate task.
    editing.task=saved;
    const current=await planner.task(saved.id);editing.task=current.task;
    const detailVersion=editing.details['@odata.etag']?editing.details:current.details;
    await planner.saveDetails(current.task,detailVersion,details);
    $('editor').close();editing=null;phase='done';
  }catch(e){handle(e,'editor-notice');if(phase==='task'&&saving&&!editing?.task&&(!e.status||e.status>=500)){$('editor-notice').textContent='Task creation could not be confirmed. Close this form and refresh the board to check whether the task was created before trying again.';$('save').dataset.reload='true';}if(phase==='details')$('editor-notice').textContent='Task fields were saved, but details were not confirmed. Close and reopen the task before making further changes. '+e.message;
    if(phase==='details'){$('save').dataset.reload='true';$('delete').disabled=true;}
  }finally{saving=false;$('save').disabled=$('save').dataset.reload==='true';$('delete').disabled=$('save').disabled;$('close-editor').disabled=false;if(phase==='done'){await refresh();say('Task saved in Microsoft Planner.');}}
});
$('editor').addEventListener('close',()=>{delete $('save').dataset.reload;form.reset();editing=null;$('checklist').replaceChildren();$('references').replaceChildren();});
$('editor').addEventListener('cancel',e=>{if(saving)e.preventDefault();});
$('delete').onclick=async()=>{if(!editing?.task||saving||!confirm('Delete this task from Microsoft Planner? This cannot be undone.'))return;saving=true;$('delete').disabled=true;try{await planner.remove(editing.task);$('editor').close();saving=false;await refresh();say('Task deleted.');}catch(e){handle(e,'editor-notice');}finally{saving=false;$('delete').disabled=false;}};
$('close-editor').onclick=()=>$('editor').close();$('add-check').onclick=()=>checklistRow();$('add-link').onclick=()=>linkRow();$('create').onclick=()=>openTask();$('refresh').onclick=refresh;
$('search').oninput=render;$('filter').onchange=render;
for(const name of ['board','list'])$(name+'-view').onclick=()=>{view=name;render();};
$('theme').onclick=()=>{document.documentElement.dataset.theme=(document.documentElement.dataset.theme|| (matchMedia('(prefers-color-scheme: dark)').matches?'dark':'light'))==='dark'?'light':'dark';};
$('privacy-button').onclick=()=>$('privacy').showModal();$('close-privacy').onclick=()=>$('privacy').close();
$('signin').onclick=async()=>{try{$('signin').disabled=true;const result=await auth.loginPopup({scopes:SCOPES,prompt:'select_account'});account=result.account;if(account.tenantId!==config.tenantId)throw Error('Use an approved account in the organization that owns this plan.');auth.setActiveAccount(account);$('signout').hidden=false;await refresh();}catch(e){say(e instanceof InteractionRequiredAuthError?'Sign in again to continue.':e.message||'Sign-in could not be completed.');}finally{$('signin').disabled=false;}};
$('signout').onclick=async()=>{const old=account;clearPrivate();account=null;$('signout').hidden=true;$('connection').textContent='Signed out';say('Signed out of the workspace.');try{await auth.logoutPopup({account:old,postLogoutRedirectUri:config.redirectUri});}catch{await auth.clearCache();}};
async function init(){try{config=await fetch('./config.json',{cache:'no-store'}).then(r=>{if(!r.ok)throw Error();return r.json();});if(!configured(config,location))return;
  auth=new PublicClientApplication({auth:{clientId:config.clientId,authority:`https://login.microsoftonline.com/${config.tenantId}`,redirectUri:config.redirectUri,postLogoutRedirectUri:config.redirectUri},cache:{cacheLocation:'memoryStorage'},system:{allowRedirectInIframe:false,loggerOptions:{piiLoggingEnabled:false,loggerCallback:()=>{}}}});
  await auth.initialize();
  planner=new Planner(config,async()=>{if(!account)throw Error('Sign in to access the plan.');const result=await auth.acquireTokenSilent({account,scopes:SCOPES});return result.accessToken;});
  $('signin').disabled=false;$('setup').textContent='Use your work account or an invited guest account in the N1 organization.';$('connection').textContent='Sign-in required';
}catch{say('Workspace configuration could not be loaded. Please try again later.');}}
setInterval(()=>{if(account&&!document.hidden&&!$('workspace').hidden)void refresh();},60000);
window.addEventListener('pagehide',()=>{clearPrivate();account=null;});
window.addEventListener('pageshow',e=>{if(e.persisted)location.reload();});
void init();


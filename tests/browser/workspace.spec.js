import {test,expect} from '@playwright/test';
import {readFileSync} from 'node:fs';
const guid='11111111-1111-4111-8111-111111111111';
async function mock(page,{denied=false,conflict=false,detailsFail=false}={}){
  const writes=[];let count=0;
  const task={id:'task-one',planId:'synthetic-plan',title:'Synthetic project task',bucketId:'bucket-one',percentComplete:0,priority:5,assignments:{},'@odata.etag':'task-etag'};
  await page.route('**/config.json',r=>r.fulfill({json:{enabled:true,tenantId:guid,clientId:guid,groupId:guid,planId:'synthetic-plan',redirectUri:'http://127.0.0.1:4173/redirect.html'}}));
  await page.route('**/node_modules/.vite/deps/@azure_msal-browser.js*',r=>r.fulfill({contentType:'application/javascript',body:`export class InteractionRequiredAuthError extends Error {} export class PublicClientApplication { async initialize(){} async loginPopup(){return {account:{tenantId:'${guid}'}}} setActiveAccount(){} async acquireTokenSilent(){return {accessToken:'synthetic-token'}} async logoutPopup(){} async clearCache(){} }`}));
  await page.route('https://graph.microsoft.com/**',async r=>{
    const request=r.request(),path=new URL(request.url()).pathname.replace('/v1.0','');
    if(request.method()!=='GET'){writes.push({path,method:request.method(),body:request.postDataJSON(),headers:request.headers()});if(conflict&&request.method()==='PATCH')return r.fulfill({status:412});if(detailsFail&&path.endsWith('/details'))return r.fulfill({status:503});if(request.method()==='POST'){count++;return r.fulfill({json:{...task,id:'created-task'}});}return r.fulfill({status:204});}
    if(denied&&path==='/planner/plans/synthetic-plan')return r.fulfill({status:403});
    const payload={
      '/planner/plans/synthetic-plan':{id:'synthetic-plan',owner:guid},
      [`/groups/${guid}`]:{visibility:'Private',groupTypes:['Unified']},
      '/me':{id:'member-one',displayName:'Synthetic member'},
      '/me/memberOf':{value:[{id:guid}]},
      '/planner/plans/synthetic-plan/tasks':{value:[task]},
      '/planner/plans/synthetic-plan/buckets':{value:[{id:'bucket-one',planId:'synthetic-plan',name:'Work'}]},
      [`/groups/${guid}/members`]:{value:[{id:'member-one',displayName:'Synthetic member','@odata.type':'#microsoft.graph.user'}]},
      '/planner/plans/synthetic-plan/details':{categoryDescriptions:{category1:'Review'}},
      '/planner/tasks/task-one':task,
      '/planner/tasks/created-task':{...task,id:'created-task'},
      '/planner/tasks/task-one/details':{'@odata.etag':'details-etag',description:'Synthetic description',checklist:{},references:{}},
      '/planner/tasks/created-task/details':{'@odata.etag':'details-etag',description:'',checklist:{},references:{}}
    }[path];return r.fulfill({json:payload??{}});
  });
  await page.goto('/');await page.getByRole('button',{name:'Sign in with Microsoft'}).click();
  return {writes,created:()=>count};
}
test('unconfigured public page is private, accessible and responsive',async({page})=>{
  const external=[];page.on('request',r=>{if(r.url().startsWith('https:'))external.push(r.url());});
  await page.goto('/');await expect(page.getByRole('button',{name:'Sign in with Microsoft'})).toBeDisabled();await expect(page.locator('#workspace')).toBeHidden();
  await page.route('**/test-axe.js',r=>r.fulfill({contentType:'application/javascript',body:readFileSync('node_modules/axe-core/axe.min.js','utf8')}));await page.addScriptTag({url:'/test-axe.js'});
  const violations=await page.evaluate(async()=> (await axe.run(document,{runOnly:{type:'tag',values:['wcag2a','wcag2aa','wcag21aa']}})).violations);expect(violations).toEqual([]);
  await page.screenshot({path:'.private/n1-desktop.png',fullPage:true});
  await page.setViewportSize({width:390,height:844});expect(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth)).toBe(true);
  await page.screenshot({path:'.private/n1-mobile.png',fullPage:true});
  await page.getByRole('button',{name:'Privacy & access'}).click();await expect(page.getByRole('dialog')).toBeVisible();expect(external).toEqual([]);
});
test('denied member never sees task content',async({page})=>{const {writes}=await mock(page,{denied:true});await expect(page.locator('#notice')).toContainText('denied access');await expect(page.locator('#workspace')).toBeHidden();expect(writes).toEqual([]);});
test('board, filters, edit fields and details use separate ETags',async({page})=>{
  const {writes}=await mock(page);await expect(page.locator('#workspace')).toBeVisible();await page.getByRole('button',{name:'List',exact:true}).click();await page.locator('#search').fill('no-match');await expect(page.locator('#tasks')).toContainText('No tasks match');await page.locator('#search').fill('');
  await page.getByRole('button',{name:/Synthetic project task/}).click();await page.getByLabel('Task title').fill('Updated synthetic title');await page.getByRole('combobox',{name:'Progress',exact:true}).selectOption('50');await page.getByRole('listbox',{name:'Assigned to',exact:true}).selectOption('member-one');await page.getByLabel('Description',{exact:true}).fill('Changed synthetic description');await page.getByRole('button',{name:'Add checklist item'}).click();await page.getByLabel('Checklist item title').fill('Synthetic checklist');
  await page.getByRole('button',{name:'Save task'}).click();await expect(page.locator('#notice')).toHaveText('Task saved in Microsoft Planner.');expect(writes).toHaveLength(2);expect(writes[0].headers['if-match']).toBe('task-etag');expect(writes[1].headers['if-match']).toBe('details-etag');expect(writes[0].body.title).toBe('Updated synthetic title');expect(writes[0].body.assignments['member-one']).toBeTruthy();
  await page.getByRole('button',{name:'Sign out',exact:true}).click();await expect(page.locator('#workspace')).toBeHidden();expect(await page.locator('body').innerText()).not.toContain('Synthetic project task');expect(await page.evaluate(()=>({local:{...localStorage},session:{...sessionStorage}}))).toEqual({local:{},session:{}});
});
test('edit conflict keeps changes and does not retry writes',async({page})=>{const {writes}=await mock(page,{conflict:true});await page.getByRole('button',{name:/Synthetic project task/}).click();await page.getByLabel('Task title').fill('Unsaved edit');await page.getByRole('button',{name:'Save task'}).click();await expect(page.locator('#editor-notice')).toContainText('changed elsewhere');await expect(page.getByLabel('Task title')).toHaveValue('Unsaved edit');expect(writes).toHaveLength(1);});
test('partial create failure cannot duplicate the new task',async({page})=>{const state=await mock(page,{detailsFail:true});await page.getByRole('button',{name:'Add task',exact:false}).click();await page.getByLabel('Task title').fill('New synthetic task');await page.getByRole('button',{name:'Save task'}).click();await expect(page.locator('#editor-notice')).toContainText('Task fields were saved');await expect(page.getByRole('button',{name:'Save task'})).toBeDisabled();expect(state.created()).toBe(1);});
test('board passes automated accessibility checks',async({page})=>{await mock(page);await expect(page.locator('#workspace')).toBeVisible();await page.screenshot({path:'.private/n1-board-synthetic.png',fullPage:true});await page.route('**/test-axe.js',r=>r.fulfill({contentType:'application/javascript',body:readFileSync('node_modules/axe-core/axe.min.js','utf8')}));await page.addScriptTag({url:'/test-axe.js'});expect(await page.evaluate(async()=> (await axe.run(document,{runOnly:{type:'tag',values:['wcag2a','wcag2aa','wcag21aa']}})).violations)).toEqual([]);});



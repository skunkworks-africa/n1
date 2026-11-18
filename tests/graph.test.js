import test from 'node:test';
import assert from 'node:assert/strict';
import {Planner,configured,referenceKey,GraphError} from '../src/graph.js';
const group='11111111-1111-4111-8111-111111111111';
const config={enabled:true,tenantId:group,clientId:group,groupId:group,planId:'synthetic-plan',redirectUri:'https://example.test/redirect.html'};
const task={id:'task-1',planId:config.planId,title:'Synthetic task',bucketId:'bucket-1',percentComplete:0,priority:5,'@odata.etag':'task-original'};
function setup(overrides={}){const calls=[];const responses={
  '/planner/plans/synthetic-plan':{id:config.planId,owner:group},
  [`/groups/${group}`]:{id:group,visibility:'Private',groupTypes:['Unified']},
  '/me':{id:'user-1',displayName:'Synthetic member'},
  '/me/memberOf':{value:[{id:group}]},
  '/planner/plans/synthetic-plan/tasks':{value:[task]},
  '/planner/plans/synthetic-plan/buckets':{value:[{id:'bucket-1',planId:config.planId,name:'Work'}]},
  [`/groups/${group}/members`]:{value:[{id:'user-1',displayName:'Synthetic member','@odata.type':'#microsoft.graph.user'}]},
  '/planner/plans/synthetic-plan/details':{categoryDescriptions:{}},
  '/planner/tasks/task-1':task,
  '/planner/tasks/task-1/details':{'@odata.etag':'details-original',description:''},...overrides};
  const p=new Planner(config,async()=>'synthetic-token',async(url,options)=>{calls.push({url,options});const key=new URL(url).pathname.replace('/v1.0','');const response=responses[key];if(typeof response==='number')return new Response('',{status:response});if(options.method!=='GET')return new Response(null,{status:204});return Response.json(response);});return {p,calls};}
test('configuration is disabled by default and bound to an exact HTTPS origin/path',()=>{assert.equal(configured({...config,enabled:false},new URL(config.redirectUri)),false);assert.equal(configured(config,new URL(config.redirectUri)),true);assert.equal(configured(config,new URL('https://attacker.test/')),false);assert.equal(configured({...config,redirectUri:'http://example.test/'},new URL('http://example.test/')),false);});
test('a successful sign-in is insufficient without private group membership',async()=>{const {p,calls}=setup({'/me/memberOf':{value:[]}});await assert.rejects(p.load(),e=>e.status===403);assert.equal(calls.some(c=>c.url.endsWith('/tasks')),false);});
test('public group and wrong owner fail before task loading',async()=>{for(const responses of [{[`/groups/${group}`]:{visibility:'Public',groupTypes:['Unified']}},{'/planner/plans/synthetic-plan':{owner:'other'}}]){const {p,calls}=setup(responses);await assert.rejects(p.load());assert.equal(calls.some(c=>c.url.endsWith('/tasks')),false);}});
test('Graph authorization denial propagates without fetching private tasks',async()=>{const {p,calls}=setup({'/planner/plans/synthetic-plan':403});await assert.rejects(p.load(),e=>e.status===403);assert.equal(calls.length,1);});
test('tasks outside the configured plan cannot be read or changed',async()=>{const {p}=setup();await p.load();await assert.rejects(p.task('other-task'),e=>e.status===403);await assert.rejects(p.saveTask({...task,planId:'other'},task),e=>e.status===403);});
test('external pagination URL is rejected before token lookup',async()=>{let tokens=0;const p=new Planner(config,async()=>{tokens++;return 'secret';});await assert.rejects(p.request('https://attacker.test/v1.0/steal'));assert.equal(tokens,0);});
test('patch retains original ETag and strips unexpected fields',async()=>{const {p,calls}=setup();await p.load();await p.saveTask(task,{...task,title:'Changed',planId:'different'});const patch=calls.find(c=>c.options.method==='PATCH');assert.equal(patch.options.headers['If-Match'],'task-original');assert.equal(JSON.parse(patch.options.body).planId,undefined);assert.equal(patch.options.cache,'no-store');});
test('details use their own ETag',async()=>{const {p,calls}=setup();await p.load();await p.saveDetails(task,{'@odata.etag':'detail-version'},{description:'Synthetic details'});assert.equal(calls.find(c=>c.options.method==='PATCH').options.headers['If-Match'],'detail-version');});
test('conflict is reported without overwrite retry',async()=>{const {p,calls}=setup();await p.load();p.fetcher=async(url,options)=>{calls.push({url,options});if(options.method==='PATCH')return new Response('',{status:412});const q=setup();return q.p.fetcher(url,options);};await assert.rejects(p.saveTask(task,task),e=>e.status===412);assert.equal(calls.filter(c=>c.options.method==='PATCH').length,1);});
test('assignment, bucket and date validation prevent invalid writes',async()=>{const {p}=setup();await p.load();assert.throws(()=>p.validate({...task,assignments:{outsider:{}}}));assert.throws(()=>p.validate({...task,bucketId:'another'}));assert.throws(()=>p.validate({...task,startDateTime:'2026-02-02',dueDateTime:'2026-01-01'}));});
test('unsafe reference URLs rejected and Planner URL encoding applied',()=>{assert.throws(()=>referenceKey('javascript:alert(1)'));assert.throws(()=>referenceKey('https://user:pass@example.test'));assert.equal(referenceKey('https://example.test/path#anchor'),'https%3A//example%2Etest/path%23anchor');});
test('all pagination pages are loaded',async()=>{const {p}=setup();let n=0;p.fetcher=async()=>Response.json(++n===1?{value:[1],'@odata.nextLink':'https://graph.microsoft.com/v1.0/next'}:{value:[2]});assert.deepEqual(await p.list('/start'),[1,2]);});
test('authorization loss at write time prevents mutation',async()=>{const {p,calls}=setup();await p.load();p.fetcher=async()=>new Response('',{status:403});await assert.rejects(p.saveTask(task,task),e=>e.status===403);assert.equal(calls.filter(c=>c.options.method!=='GET').length,0);});

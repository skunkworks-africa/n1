export const SCOPES = ['User.Read', 'Tasks.ReadWrite', 'GroupMember.Read.All'];
const ROOT = 'https://graph.microsoft.com/v1.0';
const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
export function configured(c, location) {
  return c?.enabled === true && [c.tenantId,c.clientId,c.groupId].every(x=>guid.test(x)) && /^[\w-]{10,100}$/.test(c.planId) && c.redirectUri === new URL('redirect.html',location.origin+location.pathname).href && (location.protocol === 'https:' || location.hostname === 'localhost' || location.hostname === '127.0.0.1');
}
export class GraphError extends Error {
  constructor(status) {
    super(({401:'Your session has expired. Sign out and sign in again.',403:'Microsoft has denied access. Ask the plan owner to check your membership and permissions.',404:'The plan or task is unavailable. Refresh or contact the plan owner.',409:'This task changed elsewhere. Close and reopen it to review the latest version before saving.',412:'This task changed elsewhere. Close and reopen it to review the latest version before saving.',429:'Microsoft is limiting requests. Please wait a minute before trying again.'})[status] || 'Microsoft could not complete the request. Please try again.');
    this.status=status;
  }
}
export class Planner {
  constructor(config, getToken, fetcher=(...args)=>globalThis.fetch(...args)) { this.config=config; this.getToken=getToken; this.fetcher=fetcher; this.tasks=new Map(); this.members=[]; this.buckets=[]; }
  async request(path, {method='GET',body,etag}={}) {
    const url = new URL(path.startsWith('https:') ? path : ROOT+path);
    if(url.origin!=='https://graph.microsoft.com'||!url.pathname.startsWith('/v1.0/')||url.username||url.password) throw Error('Unsafe Graph destination.');
    const token = await this.getToken();
    const headers={Authorization:`Bearer ${token}`,Accept:'application/json'};
    if(body!==undefined) headers['Content-Type']='application/json';
    if(etag) headers['If-Match']=etag;
    const response=await this.fetcher(url.href,{method,headers,body:body===undefined?undefined:JSON.stringify(body),cache:'no-store',credentials:'omit',redirect:'error'});
    if(!response.ok) throw new GraphError(response.status);
    return response.status===204?null:response.json();
  }
  async list(path) {
    const all=[];const seen=new Set(); let next=path;
    while(next) { if(seen.has(next)||seen.size>100) throw Error('Unable to finish loading the list.'); seen.add(next);const page=await this.request(next);all.push(...page.value);next=page['@odata.nextLink']; }
    return all;
  }
  async authorize() {
    const {planId,groupId}=this.config;
    const plan=await this.request(`/planner/plans/${encodeURIComponent(planId)}`);
    const owner=plan.container?.containerId || plan.owner;
    if(owner!==groupId || (plan.container?.type && plan.container.type!=='group')) throw Error('The plan is not in the approved Microsoft 365 group.');
    const group=await this.request(`/groups/${groupId}?$select=id,visibility,groupTypes`);
    if(group.visibility!=='Private' || !group.groupTypes?.includes('Unified')) throw Error('N1 requires a private Microsoft 365 group. Ask the administrator to review the plan.');
    const me=await this.request('/me?$select=id,displayName');
    const memberships=await this.list('/me/memberOf?$select=id');
    if(!memberships.some(g=>g.id===groupId)) throw new GraphError(403);
    this.me=me;this.plan=plan;
    return plan;
  }
  async load() {
    await this.authorize();
    const id=encodeURIComponent(this.config.planId);
    const [tasks,buckets,members,details]=await Promise.all([this.list(`/planner/plans/${id}/tasks`),this.list(`/planner/plans/${id}/buckets`),this.list(`/groups/${this.config.groupId}/members?$select=id,displayName`),this.request(`/planner/plans/${id}/details`)]);
    if(tasks.some(t=>t.planId!==this.config.planId)||buckets.some(b=>b.planId!==this.config.planId)) throw Error('Microsoft returned an unexpected plan.');
    this.tasks=new Map(tasks.map(t=>[t.id,t]));this.buckets=buckets;this.members=members.filter(m=>m['@odata.type']==='#microsoft.graph.user');this.details=details;
    return tasks;
  }
  async task(id) {
    if(!this.tasks.has(id)) throw new GraphError(403);
    const t=await this.request(`/planner/tasks/${encodeURIComponent(id)}`);
    if(t.planId!==this.config.planId) throw new GraphError(403);
    const details=await this.request(`/planner/tasks/${encodeURIComponent(id)}/details`);
    return {task:t,details};
  }
  validate(body) {
    if(!body.title?.trim() || body.title.length>255) throw Error('Enter a task title of 1–255 characters.');
    if(!this.buckets.some(b=>b.id===body.bucketId)) throw Error('Select a bucket in this plan.');
    if(!Number.isInteger(body.percentComplete)||body.percentComplete<0||body.percentComplete>100||!Number.isInteger(body.priority)||body.priority<0||body.priority>10) throw Error('Select a valid progress and priority.');
    if(body.startDateTime&&body.dueDateTime&&new Date(body.startDateTime)>new Date(body.dueDateTime)) throw Error('The due date must be after the start date.');
    for(const [id,value] of Object.entries(body.assignments||{})) if(value!==null && !this.members.some(m=>m.id===id)) throw Error('Assign only members of this plan’s group.');
    return Object.fromEntries(['title','bucketId','percentComplete','priority','startDateTime','dueDateTime','assignments','appliedCategories'].filter(k=>k in body).map(k=>[k,body[k]]));
  }
  async saveTask(original, body) {
    await this.authorize();
    const safe=this.validate(body);
    if(!original) {const created=await this.request('/planner/tasks',{method:'POST',body:{...safe,planId:this.config.planId}}); this.tasks.set(created.id,created);return created;}
    if(!this.tasks.has(original.id)||original.planId!==this.config.planId||!original['@odata.etag']) throw new GraphError(403);
    await this.request(`/planner/tasks/${encodeURIComponent(original.id)}`,{method:'PATCH',body:safe,etag:original['@odata.etag']});
    return original;
  }
  async saveDetails(task, original, body) {
    await this.authorize();
    if(!this.tasks.has(task.id)||task.planId!==this.config.planId||!original['@odata.etag']) throw new GraphError(403);
    const safe=Object.fromEntries(['description','checklist','references'].filter(k=>k in body).map(k=>[k,body[k]]));
    await this.request(`/planner/tasks/${encodeURIComponent(task.id)}/details`,{method:'PATCH',etag:original['@odata.etag'],body:safe});
  }
  async remove(task) {
    await this.authorize();
    if(!this.tasks.has(task.id)||task.planId!==this.config.planId||!task['@odata.etag']) throw new GraphError(403);
    await this.request(`/planner/tasks/${encodeURIComponent(task.id)}`,{method:'DELETE',etag:task['@odata.etag']});
    this.tasks.delete(task.id);
  }
}
export function referenceKey(url) {
  const parsed=new URL(url);
  if(!['https:','http:'].includes(parsed.protocol)||parsed.username||parsed.password) throw Error('Links must use http or https without embedded credentials.');
  return parsed.href.replace(/[.%:#@]/g,c=>'%'+c.charCodeAt(0).toString(16).toUpperCase());
}

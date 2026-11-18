/* 9D IT Asset Pro: bilingual UI and tenant-backed user management.
   Frontend roles support UI workflows only. Secure list/backend permissions separately.
 */
(() => {
'use strict';
const SITE='ninedistribution.sharepoint.com,13d8bc75-67a0-4dbf-b579-8b26c2dd2e9f,dedddf0b-d21b-4660-84ad-259b6ffb0ab4';
const BASE='https://graph.microsoft.com/v1.0/sites/'+SITE+'/lists';
const LIST='9D Portal Users',ADMIN=['nutsarun@ninedistribution.com','satit@ninedistribution.com'];
const roles={admin:'Super Admin',staff:'IT Staff',auditor:'Auditor',viewer:'Viewer'};
const $=id=>document.getElementById(id);
const safe=x=>String(x??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const email=x=>String(x||'').trim().toLowerCase();
let users=[],listId='',onlyAllowlisted=false,loading=false;
let lang=localStorage.getItem('nine-ui-lang')==='en'?'en':'th';
const isRoot=()=>ADMIN.includes(email(currentUser?.email));
const roleFor=()=>isRoot()?'admin':(users.find(u=>u.email===email(currentUser?.email)&&u.enabled)?.role||'viewer');
const permission=(name)=>{
 if(!currentUser)return false;
 const r=roleFor();
 if(name==='superadmin'||name==='users')return r==='admin';
 if(name==='write'||name==='checkout'||name==='employee')return r==='admin'||r==='staff';
 if(name==='audit')return r==='admin'||r==='staff'||r==='auditor';
 return true;
};
window.nineCan=permission;
window.nineLanguage=()=>lang;
const labels={
 'ภาพรวมระบบ (Dashboard)':'Dashboard',
 'คลังทรัพย์สิน (Inventory)':'Asset Inventory',
 'เบิก-ยืมอุปกรณ์ (Check-Out)':'Check Out Assets',
 'รับคืนอุปกรณ์ (Check-In)':'Check In Assets',
 'ประวัติ & พิมพ์เอกสาร':'History & Documents',
 'ตั้งค่า Microsoft SSO':'Microsoft SSO Settings',
 'ข้อมูลพนักงาน (Employees)':'Employees',
 'รวมข้อมูล SharePoint':'Merge SharePoint Lists',
 'ภาพรวมทรัพย์สินไอทีในองค์กร':'IT Asset Dashboard',
 'สรุปสถานะอุปกรณ์ทั้งหมด ผู้ถือครองปัจจุบัน และรายการเคลื่อนไหวล่าสุด':'Overview of all assets, current custodians and recent activity',
 'คลังฐานข้อมูลทรัพย์สินไอที':'IT Asset Inventory',
 'จัดการทะเบียนคอมพิวเตอร์ โน๊ตบุ๊ค จอภาพ และอุปกรณ์ต่อพ่วงทั้งหมด':'Manage notebooks, computers, monitors and other IT equipment',
 'เพิ่มอุปกรณ์ใหม่':'Add New Asset',
 'ทำรายการเบิก-ยืม':'Check Out',
 'รับคืนอุปกรณ์':'Check In',
 'สรุปทรัพย์สินพนักงาน / ตรวจสภาพ':'Employee Asset Summary / Inspection',
 'ทรัพย์สินทั้งหมด':'Total Assets',
 'พร้อมใช้งาน (คลัง)':'In Stock',
 'ถูกยืมใช้งาน':'Assigned',
 'ส่งซ่อม':'Under Repair',
 'รายการในทะเบียน':'Registered assets',
 'พร้อมจ่ายให้พนักงาน':'Available for issue',
 'ค้นหา Asset ID, ยี่ห้อ, รุ่น, Serial Number, หรือชื่อผู้ใช้งาน...':'Search Asset ID, brand, model, serial or custodian...',
 'ทุกหมวดหมู่ (All Categories)':'All Categories',
 'ทุกสถานะ (All Statuses)':'All Statuses',
 'สรุปการถือครองทรัพย์สินรายบุคคล':'Employee Asset Possession Summary',
 'เลือกทั้งหมด':'Select All',
 'ยกเลิกทั้งหมด':'Clear Selection',
 'ออกเอกสาร PDF รายที่เลือก':'Export Selected as PDF',
 'การรับรองรายการและสภาพทรัพย์สิน':'Asset Condition Confirmation',
 'ข้อมูลผู้ถือครอง':'Custodian Information',
 'ดูมูลค่าแยกประเภท':'View Valuation by Category',
 'มูลค่าทรัพย์สินรวม (เฉพาะรายการที่มีข้อมูล)':'Total Recorded Asset Value',
 'เพิ่มพนักงาน':'Add Employee',
 'แก้ไข':'Edit',
 'ลบ':'Delete',
 'บันทึก':'Save',
 'ยกเลิก':'Cancel',
 'ออกจากระบบ':'Sign Out',
 'ชื่อ-นามสกุล':'Full Name',
 'แผนก':'Department',
 'ตำแหน่ง':'Position',
 'สังกัดบริษัท':'Company',
 'สถานะ':'Status',
 'หมายเหตุ':'Notes',
 'มูลค่าทรัพย์สิน (บาท)':'Asset Value (THB)',
 'ผู้ครอบครอง / วันที่รับ':'Custodian / Assigned Date',
 'จัดการ':'Actions',
 'ประเภท':'Category',
 'สภาพ':'Condition',
 'ปกติ':'Normal',
 'ผิดปกติ':'Abnormal',
 'ปีที่ตรวจ':'Inspection Year',
 'ทุกบริษัท':'All Companies',
 'สรุปผล':'Summary',
 'จัดการผู้ใช้งาน':'User Management',
 'ตั้งค่าระบบ':'Settings',
 'ภาษาเว็บไซต์':'Website Language',
 'ภาษาไทย':'Thai',
 'ภาษาอังกฤษ':'English'
};
const backwards=Object.fromEntries(Object.entries(labels).map(([a,b])=>[b,a]));
const originals=new WeakMap();
function translate(){
 const app=$('app-shell');
 if(!app)return;
 const walker=document.createTreeWalker(app,NodeFilter.SHOW_TEXT);
 let n;
 while((n=walker.nextNode())){
  const parent=n.parentElement;
  if(!parent||parent.closest('script,style,textarea,input,option,[data-no-translate]')||!n.textContent.trim())continue;
  const raw=n.textContent;
  const trimmed=raw.trim();
  if(!originals.has(n))originals.set(n,backwards[trimmed]||trimmed);
  const original=originals.get(n);
  const value=lang==='en'?(labels[original]||original):original;
  if(trimmed!==value)n.textContent=raw.replace(trimmed,value);
 }
 app.querySelectorAll('input[placeholder]').forEach(input=>{
  const orig=input.dataset.originalPlaceholder||input.placeholder;
  input.dataset.originalPlaceholder=orig;
  input.placeholder=lang==='en'?(labels[orig]||orig):orig;
 });
 for(const b of document.querySelectorAll('[data-language-select]'))b.value=lang;
}
let scheduled=false;
function schedule(){if(scheduled)return;scheduled=true;requestAnimationFrame(()=>{scheduled=false;translate()})}
function setLang(value){
 lang=value==='en'?'en':'th';localStorage.setItem('nine-ui-lang',lang);
 document.documentElement.lang=lang;translate();
}
function gql(url,method='GET',body=null,scope='Sites.Read.All'){
 const auth=window.nineGraphAuth;
 if(!auth)throw Error('Please sign in with Microsoft');
 return (async()=>{
  const req={scopes:[scope],account:auth.account};
  let token;try{token=(await auth.client.acquireTokenSilent(req)).accessToken}
  catch(e){token=(await auth.client.acquireTokenPopup(req)).accessToken}
  const headers={Authorization:'Bearer '+token};
  if(body!==null)headers['Content-Type']='application/json';
  const res=await fetch(url,{method,headers,body:body===null?undefined:JSON.stringify(body)});
  if(!res.ok)throw Error('SharePoint '+res.status+': '+(await res.text()).slice(0,240));
  return res.status===204?null:res.json();
 })();
}
async function all(url,scope='Sites.Read.All'){
 let list=[];while(url){const page=await gql(url,'GET',null,scope);list.push(...page.value||[]);url=page['@odata.nextLink']}return list;
}
async function loadUsers(){
 loading=true;
 try{
  const lists=await all(BASE+'?$top=100');
  const found=lists.find(x=>x.displayName===LIST);listId=found?.id||'';
  users=[];onlyAllowlisted=false;
  if(listId){
   const items=await all(BASE+'/'+listId+'/items?$top=100&$expand=fields');
   for(const row of items){
    const f=row.fields||{};
    if(f.Title==='__PORTAL_CONFIG__'){onlyAllowlisted=String(f.Enabled).toLowerCase()==='true';continue}
    const mail=email(f.Title);
    if(!mail)continue;
    users.push({email:mail,role:roles[f.Role]?f.Role:'viewer',enabled:String(f.Enabled).toLowerCase()!=='false',id:row.id});
   }
  }
  const self=email(currentUser?.email);
  const configured=users.find(u=>u.email===self);
  if(!isRoot()&&((configured&&!configured.enabled)||(onlyAllowlisted&&!configured))){
   currentUser=null;showLoginScreen();showToast('บัญชี Microsoft นี้ยังไม่ได้รับสิทธิ์ใช้งานเว็บ กรุณาติดต่อ Super Admin','error');
   return;
  }
  if(currentUser){
   const r=roleFor();
   currentUser.role=r==='staff'?'support':r==='auditor'?'viewer':r==='admin'?'admin':'viewer';
   currentUser.roleLabel=roles[r];
   applyUserProfileAndRBAC();
   renderAll();
   window.renderEmployees?.();
   updateMenus();
   renderUsers();
  }
 }catch(e){
  console.warn('User directory read not available',e);
  if(isRoot())showToast('ยังไม่สามารถอ่านรายชื่อผู้ใช้จาก SharePoint: '+e.message,'error');
 }finally{loading=false}
}
async function createList(){
 if(!isRoot())return;
 if(!confirm('สร้าง SharePoint List "'+LIST+'" เพื่อเก็บบัญชีและบทบาทใช่หรือไม่?'))return;
 try{
  let found=(await all(BASE+'?$top=100')).find(x=>x.displayName===LIST);
  if(!found)found=await gql(BASE,'POST',{displayName:LIST,list:{template:'genericList'}},'Sites.Manage.All');
  listId=found.id;
  const cols=await all(BASE+'/'+listId+'/columns?$top=100','Sites.Manage.All');
  for(const name of ['Role','Enabled']){
   if(!cols.some(x=>x.name===name))await gql(BASE+'/'+listId+'/columns','POST',{name,text:{}},'Sites.Manage.All');
  }
  const rows=await all(BASE+'/'+listId+'/items?$top=100&$expand=fields');
  if(!rows.some(x=>x.fields?.Title==='__PORTAL_CONFIG__'))
   await gql(BASE+'/'+listId+'/items','POST',{fields:{Title:'__PORTAL_CONFIG__',Role:'config',Enabled:'false'}},'Sites.ReadWrite.All');
  await loadUsers();showToast('สร้างฐานข้อมูลผู้ใช้งานเรียบร้อยแล้ว','success');
 }catch(e){showToast('สร้างรายการผู้ใช้ไม่สำเร็จ: '+e.message,'error')}
}
async function saveUser(ev){
 ev.preventDefault();
 if(!isRoot()||!listId)return;
 const mail=email($('portal-user-email').value);
 if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(mail))return showToast('รูปแบบอีเมลไม่ถูกต้อง','error');
 if(ADMIN.includes(mail))return showToast('บัญชี Super Admin หลักไม่สามารถเปลี่ยนสิทธิ์ผ่านหน้านี้','error');
 const role=$('portal-user-role').value,enabled=$('portal-user-enabled').checked;
 if(!roles[role]||role==='admin')return showToast('เฉพาะบัญชี Super Admin เดิมเท่านั้นที่เป็น Super Admin ได้','error');
 const existing=users.find(u=>u.email===mail);
 try{
  const fields={Title:mail,Role:role,Enabled:String(enabled)};
  if(existing)await gql(BASE+'/'+listId+'/items/'+existing.id+'/fields','PATCH',fields,'Sites.ReadWrite.All');
  else await gql(BASE+'/'+listId+'/items','POST',{fields},'Sites.ReadWrite.All');
  $('portal-user-email').value='';await loadUsers();showToast('บันทึกสิทธิ์ผู้ใช้งานแล้ว','success');
 }catch(e){showToast('บันทึกสิทธิ์ไม่สำเร็จ: '+e.message,'error')}
}
async function removeUser(mail){
 if(!isRoot()||!listId)return;
 const u=users.find(x=>x.email===mail);if(!u||!confirm('ลบสิทธิ์ '+mail+' ออกจากรายการใช่หรือไม่?'))return;
 try{await gql(BASE+'/'+listId+'/items/'+u.id,'DELETE',null,'Sites.ReadWrite.All');await loadUsers();showToast('ลบสิทธิ์แล้ว','success')}
 catch(e){showToast('ลบไม่สำเร็จ: '+e.message,'error')}
}
async function toggleRestriction(flag){
 if(!isRoot()||!listId)return;
 if(flag&&!confirm('เมื่อเปิดโหมดจำกัดเฉพาะผู้ที่อนุญาต ผู้ใช้ที่ไม่มีรายชื่อหรือถูกปิดใช้งานจะเข้าสู่เว็บไม่ได้ ยืนยันหรือไม่?')){$('portal-allowlist').checked=false;return}
 try{
  const rows=await all(BASE+'/'+listId+'/items?$top=100&$expand=fields');
  const config=rows.find(x=>x.fields?.Title==='__PORTAL_CONFIG__');
  const fields={Title:'__PORTAL_CONFIG__',Role:'config',Enabled:String(flag)};
  if(config)await gql(BASE+'/'+listId+'/items/'+config.id+'/fields','PATCH',fields,'Sites.ReadWrite.All');
  else await gql(BASE+'/'+listId+'/items','POST',{fields},'Sites.ReadWrite.All');
  onlyAllowlisted=flag;renderUsers();showToast('อัปเดตข้อกำหนดการเข้าใช้งานแล้ว','success');
 }catch(e){$('portal-allowlist').checked=!flag;showToast(e.message,'error')}
}
function renderUsers(){
 const root=$('portal-users-table');if(!root)return;
 const rows=ADMIN.map(x=>({email:x,role:'admin',enabled:true,immutable:true})).concat(users.filter(x=>!ADMIN.includes(x.email)));
 root.innerHTML='<div class="overflow-x-auto"><table class="min-w-full text-sm"><thead><tr><th class="p-2 text-left">Microsoft 365 Email</th><th class="p-2 text-left">Role</th><th class="p-2 text-left">Access</th><th class="p-2 text-left">Actions</th></tr></thead><tbody>'+rows.map(u=>
  '<tr class="border-t"><td class="p-2">'+safe(u.email)+'</td><td class="p-2">'+safe(roles[u.role])+'</td><td class="p-2">'+(u.enabled?'Active':'Disabled')+'</td><td class="p-2">'+(u.immutable?'Protected':'<button type="button" class="portal-edit text-blue-600 mr-2" data-email="'+safe(u.email)+'">Edit</button><button type="button" class="portal-remove text-red-600" data-email="'+safe(u.email)+'">Remove</button>')+'</td></tr>').join('')+'</tbody></table></div>';
 root.querySelectorAll('.portal-edit').forEach(b=>b.onclick=()=>{const u=users.find(x=>x.email===b.dataset.email);if(!u)return;$('portal-user-email').value=u.email;$('portal-user-role').value=u.role;$('portal-user-enabled').checked=u.enabled});
 root.querySelectorAll('.portal-remove').forEach(b=>b.onclick=()=>removeUser(b.dataset.email));
 $('portal-allowlist').checked=onlyAllowlisted;
 $('portal-config-state').textContent=listId?'SharePoint: '+LIST+' (connected)':'Not configured: create the SharePoint user list first';
 $('portal-setup-users').hidden=!!listId;
 $('portal-user-form').querySelectorAll('input,select,button').forEach(e=>e.disabled=!listId);
 $('portal-allowlist').disabled=!listId;
}
function updateMenus(){
 const r=roleFor();
 for(const id of ['nav-checkout','nav-checkin','btn-add-asset']){
  const el=$(id);if(el)el.classList.toggle('hidden',!permission('write'));
 }
 // Keep existing Dashboard controls visible (summary button lives here).
 if($('dashboard-quick-actions'))$('dashboard-quick-actions').classList.remove('hidden');
 $('dashboard-quick-actions')?.querySelectorAll('button[onclick*="checkout"],button[onclick*="checkin"]').forEach(b=>b.classList.toggle('hidden',!permission('write')));
 if($('emp-add'))$('emp-add').hidden=!permission('employee');
 for(const id of ['nav-settings','nav-portal-users'])$(id)?.classList.toggle('hidden',!isRoot()&&id==='nav-portal-users');
 const badge=$('topbar-role-badge');if(badge)badge.textContent=roles[r]||'Viewer';
 const sum=$('employee-summary-button');if(sum)sum.classList.toggle('hidden',!permission('audit'));
}
function setup(){
 if($('tab-portal-users'))return;
 const nav=document.querySelector('#sidebar nav'),main=document.querySelector('#app-shell main');
 if(!nav||!main)return;
 const link=(id,title,tab)=>{
  const b=document.createElement('button');b.id=id;b.dataset.tab=tab;
  b.className='nav-btn w-full flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-medium hover:bg-slate-800 hover:text-white text-slate-300';
  b.textContent=title;b.onclick=()=>switchTab(tab);nav.append(b);
 };
 link('nav-settings','ตั้งค่าระบบ','settings');link('nav-portal-users','จัดการผู้ใช้งาน','portal-users');
 const sec=document.createElement('section');sec.id='tab-settings';sec.className='tab-content hidden bg-white rounded-2xl border p-5 space-y-4';
 sec.innerHTML='<h2 class="font-bold text-2xl">ตั้งค่าระบบ (Settings)</h2><div class="max-w-md space-y-2"><label class="block font-semibold">ภาษาเว็บไซต์ (Website Language)</label><select id="portal-language" data-language-select class="w-full border rounded-xl p-3"><option value="th">🇹🇭 ภาษาไทย</option><option value="en">🇬🇧 English</option></select><p class="text-xs text-slate-500">การเลือกภาษาจะจดจำในเบราว์เซอร์นี้และไม่เปลี่ยนข้อมูลใน SharePoint</p></div>';
 const usersTab=document.createElement('section');usersTab.id='tab-portal-users';usersTab.className='tab-content hidden bg-white rounded-2xl border p-5 space-y-4';
 usersTab.innerHTML='<h2 class="text-2xl font-bold">จัดการผู้ใช้งาน (User Management)</h2><p class="text-sm text-amber-800 bg-amber-50 p-3 border rounded-lg">ต้องกำหนดสิทธิ์ของ SharePoint Lists ให้สัมพันธ์กับบทบาทด้วย การซ่อนปุ่มในเว็บไซต์ไม่ใช่การป้องกัน API โดยตัวมันเอง</p><p id="portal-config-state" class="text-sm text-slate-500">Loading...</p><button id="portal-setup-users" class="border rounded-lg px-4 py-2">สร้าง SharePoint User List</button><form id="portal-user-form" class="grid grid-cols-1 sm:grid-cols-4 gap-3 items-end"><label class="sm:col-span-2 text-sm">Microsoft 365 Email<input id="portal-user-email" type="email" required class="mt-1 block w-full p-2 border rounded-lg" placeholder="user@ninedistribution.com"></label><label class="text-sm">Role<select id="portal-user-role" class="mt-1 block w-full p-2 border rounded-lg"><option value="staff">IT Staff</option><option value="auditor">Auditor</option><option value="viewer">Viewer</option></select></label><div class="space-y-2"><label class="flex gap-2"><input type="checkbox" id="portal-user-enabled" checked> Active</label><button type="submit" class="px-3 py-2 bg-blue-600 text-white rounded-lg w-full">Save User</button></div></form><label class="flex items-start gap-2 text-sm"><input type="checkbox" id="portal-allowlist"> <span>จำกัดการเข้าถึงเว็บเฉพาะบัญชีที่อนุญาต (Allowlist Only)</span></label><div id="portal-users-table"></div>';
 main.append(sec,usersTab);
 $('portal-language').value=lang;
 $('portal-language').onchange=e=>setLang(e.target.value);
 $('portal-setup-users').onclick=createList;
 $('portal-user-form').onsubmit=saveUser;
 $('portal-allowlist').onchange=e=>toggleRestriction(e.target.checked);
 renderUsers();updateMenus();schedule();
 const obs=new MutationObserver(schedule);obs.observe(main,{subtree:true,childList:true,characterData:true});
}
const oldHook=window.nineAfterLogin;
window.nineAfterLogin=async function(){
 await oldHook?.();
 setup();
 await loadUsers();
 updateMenus();schedule();
};
document.addEventListener('DOMContentLoaded',()=>{setup();schedule()});
window.ninePortalSettings={setLang,loadUsers,permission,updateMenus};
})();

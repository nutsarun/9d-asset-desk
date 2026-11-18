/* 9D Asset Pro SharePoint integration — no data or tokens committed to GitHub. */
(() => {
'use strict';
const ADMINS=new Set(['nutsarun@ninedistribution.com','satit@ninedistribution.com']);
const SITE='ninedistribution.sharepoint.com,13d8bc75-67a0-4dbf-b579-8b26c2dd2e9f,dedddf0b-d21b-4660-84ad-259b6ffb0ab4';
const EMP_LIST_NAME='9D Asset Employees';
let directory=[], employeeListId=null, employeeSource='local';
const esc=(v)=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const norm=s=>String(s||'').replace(/\s+/g,' ').trim().toLocaleLowerCase();
const admin=()=>!!(window.nineGraphAuth && currentUser && ADMINS.has(currentUser.email.toLowerCase()));
const path=(listId,itemId)=>'https://graph.microsoft.com/v1.0/sites/'+SITE+'/lists/'+listId+'/items'+(itemId?'/'+itemId:'');
async function graph(url,method='GET',body=null,write=false,etag=null){
 const auth=window.nineGraphAuth;
 if(!auth)throw Error('กรุณาเข้าสู่ระบบ Microsoft ใหม่');
 if(write && !admin())throw Error('บัญชีนี้ไม่ใช่ Super Admin');
 let response;
 try {
  const request={scopes:write?['Sites.ReadWrite.All']:['Sites.Read.All'],account:auth.account};
  try{response=await auth.client.acquireTokenSilent(request);}
  catch(e){response=await auth.client.acquireTokenPopup(request);}
 } catch(e){throw Error((write?'โปรดเพิ่ม Delegated permission Sites.ReadWrite.All และ Grant admin consent ก่อนบันทึกข้อมูล: ':'ไม่ได้รับสิทธิ์ SharePoint: ')+(e.message||e));}
 const headers={'Authorization':'Bearer '+response.accessToken};
 if(body!==null)headers['Content-Type']='application/json';
 if(etag)headers['If-Match']=etag;
 const res=await fetch(url,{method,headers,body:body===null?undefined:JSON.stringify(body)});
 if(!res.ok){const payload=await res.text();throw Error('SharePoint HTTP '+res.status+': '+(res.status===412?'มีผู้อื่นแก้ไขข้อมูลรายการนี้แล้ว กรุณารีเฟรช':payload.slice(0,210)));}
 return res.status===204?null:await res.json();
}
async function currentAsset(a){
 const item=await graph(path(a._sp.listId,a._sp.itemId)+'?$expand=fields');
 const f=item.fields||{};
 const status=norm(f.Status);
 const holder=norm(f[a._sp.kind==='computer'?'_x0e4a_UserName':'UserName0']);
 const inuse=/in\s*-\s*use|in\s*use|borrow|assign|ยืม|ใช้งาน/.test(status);
 return {item,fields:f,holder,busy:inuse || !!holder,status};
}
function employeeForName(name){return directory.find(e=>norm(e.name)===norm(name));}
function enrich(){
 for(const a of assets){
  const owner=a.assignedTo?.name;
  if(owner){
   const e=employeeForName(owner);
   if(e){a.assignedTo.empId=e.empId;a.assignedTo.department=e.department||'';a.assignedTo.position=e.position||'';a.assignedTo.contact=e.email||e.contact||'';}
  }
 }
}
function inferred(){
 const map=new Map();
 for(const a of assets)if(a.assignedTo?.name){
  const name=a.assignedTo.name.trim(),key=norm(name);
  if(!key)continue;
  if(!map.has(key))map.set(key,{empId:'',name,department:'',position:'',email:'',contact:'',_derived:true});
 }
 return [...map.values()];
}
async function loadDirectory(){
 const list=await graph('https://graph.microsoft.com/v1.0/sites/'+SITE+'/lists?$select=id,displayName&$top=200');
 const found=(list.value||[]).find(x=>x.displayName===EMP_LIST_NAME);
 employeeListId=found?.id||null; employeeSource=employeeListId?'sharepoint':'local';
 let remote=[];
 if(employeeListId){
  let url=path(employeeListId)+'?$top=100&$expand=fields';
  while(url){const page=await graph(url);remote.push(...(page.value||[]));url=page['@odata.nextLink'];}
  remote=remote.map(x=>({empId:x.fields.EmployeeCode||'',name:x.fields.Title||'',department:x.fields.Department||'',position:x.fields.Position||'',email:x.fields.Email||'',contact:x.fields.Contact||'',company:x.fields.Company||'Nine distribution co.,ltd',_itemId:x.id}));
 }
 const local=JSON.parse(localStorage.getItem('nine-employee-local-v2')||'[]');
 const merged=new Map();
 for(const e of inferred())merged.set(norm(e.name),e);
 for(const e of (employeeListId?remote:local))merged.set(norm(e.name),e);
 directory=[...merged.values()].sort((a,b)=>a.name.localeCompare(b.name,'th'));
 enrich();renderEmployees();populateEmployeePresets();
}
function ensurePanel(){
 const nav=document.querySelector('#sidebar nav');
 if(!nav||document.getElementById('nav-employees'))return;
 const b=document.createElement('button');
 b.id='nav-employees';b.className='nav-btn w-full flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-medium transition-all hover:bg-slate-800 hover:text-white';
 b.dataset.tab='employees';b.innerHTML='<i data-lucide="users" class="w-5 h-5"></i><span>ข้อมูลพนักงาน (Employees)</span>';
 b.onclick=()=>switchTab('employees');nav.insertBefore(b,document.querySelector('#nav-sso-settings'));
 const sec=document.createElement('section');sec.id='tab-employees';sec.className='tab-content hidden space-y-5';
 sec.innerHTML='<div class="flex flex-wrap justify-between gap-3 items-center"><div><h2 class="text-2xl font-bold">ข้อมูลพนักงาน</h2><p class="text-sm text-slate-500">เชื่อมโยงผู้ถือครองจาก SharePoint และจัดการข้อมูลพนักงาน</p></div><button id="emp-add" class="bg-blue-600 text-white px-4 py-2 rounded-xl" onclick="nineEmployeeEdit(-1)">+ เพิ่มพนักงาน</button></div><div id="emp-source" class="rounded-xl p-3 border text-sm"></div><input id="emp-search" oninput="renderEmployees()" class="w-full bg-white border rounded-xl p-3" placeholder="ค้นหาชื่อพนักงาน รหัส แผนก และอีเมล"><div id="employee-list" class="bg-white rounded-xl border overflow-x-auto"></div><div id="emp-modal" class="fixed inset-0 z-50 bg-black/50 hidden items-center justify-center p-4"><form id="emp-form" onsubmit="nineEmployeeSave(event)" class="bg-white rounded-2xl p-6 w-full max-w-lg space-y-3"><h3 class="font-bold text-xl">เพิ่ม / แก้ไขพนักงาน</h3><input type="hidden" id="emp-index"><div class="grid grid-cols-2 gap-3"><label class="text-sm">รหัสพนักงาน<input id="emp-code" class="w-full border p-2 rounded-lg"></label><label class="text-sm">ชื่อ-นามสกุล *<input id="emp-name" required class="w-full border p-2 rounded-lg"></label><label class="text-sm">แผนก<input id="emp-department" class="w-full border p-2 rounded-lg"></label><label class="text-sm">ตำแหน่ง<input id="emp-position" class="w-full border p-2 rounded-lg"></label><label class="text-sm">อีเมล<input type="email" id="emp-email" class="w-full border p-2 rounded-lg"></label><label class="text-sm">เบอร์โทร<input id="emp-contact" class="w-full border p-2 rounded-lg"></label></div><div class="flex justify-end gap-2"><button type="button" onclick="document.getElementById(\'emp-modal\').classList.add(\'hidden\');document.getElementById(\'emp-modal\').classList.remove(\'flex\')" class="border rounded-lg px-4 py-2">ยกเลิก</button><button class="bg-blue-600 text-white rounded-lg px-4 py-2" type="submit">บันทึก</button></div></form></div>';
 document.querySelector('main').appendChild(sec);window.lucide?.createIcons();
}
function renderEmployees(){
 const root=document.getElementById('employee-list');if(!root)return;
 const isAdmin=admin();
 document.getElementById('emp-add').hidden=!isAdmin;
 const src=document.getElementById('emp-source');
 src.className='rounded-xl p-3 border text-sm '+(employeeListId?'bg-emerald-50 border-emerald-200 text-emerald-800':'bg-amber-50 border-amber-200 text-amber-900');
 src.textContent=employeeListId?'ข้อมูลพนักงานเชื่อมกับ SharePoint List: '+EMP_LIST_NAME:'ยังไม่มี SharePoint List "'+EMP_LIST_NAME+'" — ข้อมูลพนักงานที่เพิ่ม/แก้ไขจะเก็บเฉพาะเบราว์เซอร์นี้ ไม่แชร์กับผู้อื่น กรุณาสร้าง List ตามคู่มือใน GitHub เพื่อเปิดการซิงก์';
 const query=norm(document.getElementById('emp-search')?.value);
 const filtered=directory.map((e,i)=>({...e,_index:i})).filter(e=>norm(Object.values(e).join(' ')).includes(query));
 const counts=new Map();
 for(const a of assets)if(a.assignedTo?.name){const k=norm(a.assignedTo.name);counts.set(k,(counts.get(k)||0)+1);}
 root.innerHTML='<table class="min-w-full text-sm"><thead class="bg-slate-50"><tr><th class="p-3 text-left">ชื่อพนักงาน</th><th class="p-3 text-left">แผนก / อีเมล</th><th class="p-3">อุปกรณ์ที่ถือครอง</th><th class="p-3">จัดการ</th></tr></thead><tbody>'+filtered.map(e=>'<tr class="border-t"><td class="p-3 font-semibold">'+esc(e.name)+'<p class="text-xs text-slate-500">'+esc(e.empId)+'</p></td><td class="p-3">'+esc(e.department)+'<p class="text-xs text-slate-500">'+esc(e.email)+'</p></td><td class="p-3 text-center">'+(counts.get(norm(e.name))||0)+'</td><td class="p-3 text-center">'+(isAdmin?'<button class="text-blue-700 mr-3" onclick="nineEmployeeEdit('+e._index+')">แก้ไข</button><button class="text-rose-600" onclick="nineEmployeeDelete('+e._index+')">ลบ</button>':'ดูอย่างเดียว')+'</td></tr>').join('')+'</tbody></table>';
}
window.renderEmployees=renderEmployees;
window.nineEmployeeEdit=function(index){
 if(!admin())return showToast('สำหรับ Super Admin เท่านั้น','error');
 const e=directory[index]||{};
 const v=(id,k)=>document.getElementById(id).value=e[k]||'';
 document.getElementById('emp-index').value=index;
 v('emp-code','empId');v('emp-name','name');v('emp-department','department');v('emp-position','position');v('emp-email','email');v('emp-contact','contact');v('emp-company','company');
 document.getElementById('emp-modal').classList.remove('hidden');document.getElementById('emp-modal').classList.add('flex');
};
window.nineEmployeeSave=async function(ev){
 ev.preventDefault();if(!admin())return;
 const index=Number(document.getElementById('emp-index').value);
 const old=directory[index];
 const e={empId:document.getElementById('emp-code').value.trim(),name:document.getElementById('emp-name').value.trim(),department:document.getElementById('emp-department').value.trim(),position:document.getElementById('emp-position').value.trim(),email:document.getElementById('emp-email').value.trim(),contact:document.getElementById('emp-contact').value.trim(),company:document.getElementById('emp-company')?.value||'Nine distribution co.,ltd'};
 if(!e.name)return;
 if(directory.some((x,i)=>i!==index && norm(x.name)===norm(e.name)))return showToast('ชื่อพนักงานซ้ำกับรายการเดิม','error');
 try{
  if(employeeListId){
   const fields={Title:e.name,EmployeeCode:e.empId,Department:e.department,Position:e.position,Email:e.email,Contact:e.contact,Company:e.company};
   if(old?._itemId){await graph(path(employeeListId,old._itemId)+'/fields','PATCH',fields,true); }
   else await graph(path(employeeListId),'POST',{fields},true);
   await loadDirectory();
  }else{
   if(index<0)directory.push(e);else directory[index]=e;
   localStorage.setItem('nine-employee-local-v2',JSON.stringify(directory.filter(x=>!x._derived)));
  }
  document.getElementById('emp-modal').classList.add('hidden');document.getElementById('emp-modal').classList.remove('flex');renderEmployees();populateEmployeePresets();
  showToast(employeeListId?'บันทึกข้อมูลพนักงานลง SharePoint แล้ว':'บันทึกเฉพาะเบราว์เซอร์นี้แล้ว','success');
 }catch(err){showToast(err.message,'error');}
};
window.nineEmployeeDelete=async function(index){
 if(!admin())return;
 const e=directory[index];if(!e)return;
 if(assets.some(a=>norm(a.assignedTo?.name)===norm(e.name)))return showToast('ไม่สามารถลบพนักงานที่กำลังถือครองทรัพย์สินอยู่','error');
 if(!confirm('ยืนยันลบข้อมูล '+e.name+' ?'))return;
 try{
  if(employeeListId && e._itemId){await graph(path(employeeListId,e._itemId),'DELETE',null,true);await loadDirectory();}
  else{directory.splice(index,1);localStorage.setItem('nine-employee-local-v2',JSON.stringify(directory.filter(x=>!x._derived)));}
  renderEmployees();showToast('ลบข้อมูลพนักงานแล้ว','success');
 }catch(err){showToast(err.message,'error');}
};
async function refreshAssetData(){
 const all=[];
 for(const list of NINE_GRAPH.lists)all.push(...await readGraphList(await readToken(),list));
 assets=all;enrich();renderAll();
}
async function readToken(){
 const a=window.nineGraphAuth;
 try{return (await a.client.acquireTokenSilent({account:a.account,scopes:['Sites.Read.All']})).accessToken;}
 catch(e){return (await a.client.acquireTokenPopup({account:a.account,scopes:['Sites.Read.All']})).accessToken;}
}
async function patchAsset(a,fields){
 const item=await graph(path(a._sp.listId,a._sp.itemId)+'?$expand=fields');
 const etag=item['@odata.etag']||item.fields?.['@odata.etag'];
 await graph(path(a._sp.listId,a._sp.itemId)+'/fields','PATCH',fields,true,etag);
}
async function checkoutItem(a,employee,date){
 const live=await currentAsset(a);
 if(live.busy||live.status==='lost'||live.status==='not available'||live.status==='repair')throw Error('อุปกรณ์ '+a.id+' ไม่พร้อมให้เบิก: '+(live.fields.Status||'')+(live.holder?' (ผู้ถือครอง: '+live.holder+')':''));
 const col=a._sp.kind==='computer'?'_x0e4a_UserName':'UserName0';
 await graph(path(a._sp.listId,a._sp.itemId)+'/fields','PATCH',{Status:a._sp.kind==='computer'?'In -Use':'Borrowed',[col]:employee.name,ReceivedDate:a._sp.kind==='master'?date:date+'T07:00:00Z'},true,live.item['@odata.etag']||live.fields?.['@odata.etag']);
}
async function checkinItem(a){
 const live=await currentAsset(a);
 if(!live.busy)throw Error('อุปกรณ์ '+a.id+' ไม่ได้อยู่ในสถานะถูกยืม');
 const col=a._sp.kind==='computer'?'_x0e4a_UserName':'UserName0';
 await graph(path(a._sp.listId,a._sp.itemId)+'/fields','PATCH',{Status:'Stock',[col]:''},true,live.item['@odata.etag']||live.fields?.['@odata.etag']);
}
window.nineCheckout=async function(ev){
 ev.preventDefault();if(!admin())return showToast('เฉพาะ Super Admin','error');
 const selected=checkoutCart.map(id=>assets.find(a=>a.id===id)).filter(Boolean);
 const employee={empId:document.getElementById('co-emp-id').value.trim(),name:document.getElementById('co-emp-name').value.trim(),department:document.getElementById('co-emp-dept').value.trim(),position:document.getElementById('co-emp-pos').value.trim(),contact:document.getElementById('co-emp-contact').value.trim(),company:directory.find(x=>norm(x.name)===norm(document.getElementById('co-emp-name').value))?.company||'Nine distribution co.,ltd'};
 const date=document.getElementById('co-date').value;
 if(!employee.name||!employee.empId||!selected.length)return showToast('กรุณาระบุพนักงานและเลือกอุปกรณ์','error');
 const button=ev.submitter;if(button)button.disabled=true;
 const done=[],errors=[];
 try{
  for(const a of selected) {
   try{await checkoutItem(a,employee,date);done.push(a);}
   catch(e){errors.push(e.message);}
  }
  if(done.length){
   const tx={docNo:generateDocNumber('HO'),type:'CHECKOUT',date,employee,items:done.map(a=>({id:a.id,category:a.category,brand:a.brand,model:a.model,serial:a.serial,condition:document.getElementById('co-note').value})),note:document.getElementById('co-note').value,officerName:currentUser.name+' ('+currentUser.email+')'};
   transactions.unshift(tx);
   sessionStorage.setItem('nine-transactions-session',JSON.stringify(transactions));
  }
  checkoutCart=[];await refreshAssetData();
  showToast(done.length+' รายการบันทึกลง SharePoint'+(errors.length?' / '+errors.length+' รายการไม่สำเร็จ: '+errors.join(' | '):''),errors.length?'error':'success');
 }catch(e){showToast(e.message,'error');}finally{if(button)button.disabled=false;}
};
window.nineCheckin=async function(ev,groupKey){
 ev.preventDefault();if(!admin())return showToast('เฉพาะ Super Admin','error');
 const inputs=[...ev.target.querySelectorAll('input[name="return_asset"]:checked')];
 if(!inputs.length)return showToast('กรุณาเลือกอุปกรณ์ที่จะคืน','error');
 const grp=groupAssignedAssetsByEmployee()[groupKey];if(!grp)return;
 const done=[],errors=[];
 for(const cb of inputs){
  const a=assets.find(x=>x.id===cb.value);if(!a)continue;
  try{await checkinItem(a);done.push(a);}catch(e){errors.push(e.message);}
 }
 if(done.length){
  transactions.unshift({docNo:generateDocNumber('RT'),type:'CHECKIN',date:ev.target.querySelector('input[name="return_date"]').value,employee:grp.employee,items:done.map(a=>({id:a.id,category:a.category,brand:a.brand,model:a.model,serial:a.serial,condition:'รับคืน'})),note:ev.target.querySelector('input[name="return_note"]').value,officerName:currentUser.name+' ('+currentUser.email+')'});
  sessionStorage.setItem('nine-transactions-session',JSON.stringify(transactions));
 }
 await refreshAssetData();showToast('คืนสำเร็จ '+done.length+' รายการ'+(errors.length?' / ผิดพลาด: '+errors.join(' | '):''),errors.length?'error':'success');
};

const ASSET_CODES={Notebook:'N',Monitor:'M',Smartphone:'S',Accessory:'A',Network:'W',Other:'O'};
function nextAssetId(company,cat,ids=assets.map(a=>a.id)){
 const prefix=(String(company).startsWith('After')?'A9IT-':'9DIT-')+(ASSET_CODES[cat]||'O');
 let max=0;for(const id of ids){const match=String(id).match(new RegExp('^'+prefix+'(\\d{4,})$','i'));if(match)max=Math.max(max,Number(match[1]));}
 return prefix+String(max+1).padStart(4,'0');
}
window.nineNextAssetId=nextAssetId;
async function liveAssetIds(){
 const ids=[];
 for(const list of NINE_GRAPH.lists){let url=path(list.id)+'?$top=100&$expand=fields';while(url){const result=await graph(url);for(const x of result.value||[]){if(x.fields?.Title)ids.push(x.fields.Title);if(x.fields?.HostName)ids.push(x.fields.HostName);}url=result['@odata.nextLink'];}}
 return ids;
}
const columnsCache=new Map();
async function hasAssetColumn(listId,column){
 if(!columnsCache.has(listId)){let cols=[],url='https://graph.microsoft.com/v1.0/sites/'+SITE+'/lists/'+listId+'/columns?$top=200';while(url){const page=await graph(url);cols.push(...page.value||[]);url=page['@odata.nextLink'];}columnsCache.set(listId,new Set(cols.map(x=>x.name)));}
 return columnsCache.get(listId).has(column);
}

window.nineSaveAsset=async function(ev){
 ev.preventDefault();if(!admin())return showToast('เฉพาะ Super Admin','error');
 const id=document.getElementById('asset-edit-mode').value;
 const a=assets.find(x=>x.id===id);
 const cat=document.getElementById('asset-category').value;
 const list=NINE_GRAPH.lists.find(x=>x.kind==='master')||NINE_GRAPH.lists.find(x=>x.kind===(cat==='Accessory'?'accessory':'computer'));
 const company=document.getElementById('asset-company')?.value||'Nine distribution co.,ltd';
 const hostname=document.getElementById('asset-id')?.value.trim()||'';
 const valueInput=document.getElementById('asset-value')?.value.trim()||'';
 if(valueInput!==''&&(!Number.isFinite(Number(valueInput))||Number(valueInput)<0))return showToast('กรุณาระบุมูลค่าทรัพย์สินเป็นตัวเลขไม่ติดลบ','error');
 const fields={Title:document.getElementById('asset-id').value.trim(),Vendor:document.getElementById('asset-brand').value.trim(),Model:document.getElementById('asset-model').value.trim(),ItemType:cat,Status:document.getElementById('asset-status').value==='available'?'Stock':document.getElementById('asset-status').value==='repair'?'Repair':'Not available'};
 fields[list.kind==='accessory'?'AssetorserialNumber':'Assetorserialnumber']=document.getElementById('asset-serial').value.trim();
 fields[list.kind==='accessory'?'Owner':'Company']=list.kind==='master'?company:list.kind==='computer'?(company.startsWith('After')?'After9':'9Distribution'):(company.startsWith('After')?'After Nine':'Ninedistribution');
 try{
  const hostSupported=await hasAssetColumn(list.id,'HostName');
  if(list.kind==='master'){
   if(!(await hasAssetColumn(list.id,'AssetValue')))throw Error('ยังไม่มีคอลัมน์ AssetValue ใน Master List กรุณากดปุ่ม เตรียมมูลค่าและซิงก์ Asset ID ที่ Dashboard');
   fields.AssetValue=valueInput===''?null:Number(valueInput);
   if(a){fields.Title=a.id;fields.AssetID=a.id;fields.HostName=a.id;}
  }
  if(hostname && !hostSupported)throw Error('SharePoint List นี้ยังไม่มีคอลัมน์ HostName กรุณาเพิ่มคอลัมน์ Single line of text ชื่อ HostName');
  if(hostSupported)fields.HostName=list.kind==='master'?(a?.id||fields.Title):(hostname||fields.Title);
  if(a){
   if(a._sp.kind!==list.kind)throw Error('ไม่สามารถย้ายรายการข้าม List กรุณาสร้างรายการใหม่');
   if(a.status==='assigned'){
    if(list.kind!=='master')throw Error('ไม่สามารถแก้รายการที่มีผู้ถือครองบน List เดิม');
    if(fields.AssetValue===undefined)throw Error('ไม่มีคอลัมน์มูลค่า');
    await patchAsset(a,{AssetValue:fields.AssetValue});
   }else await patchAsset(a,fields);
  }else{
   const allIds=await liveAssetIds();
   fields.Title=nextAssetId(company,cat,allIds);
   if(list.kind==='master'){fields.AssetID=fields.Title;fields.HostName=fields.Title;}
   if(hostSupported)fields.HostName=fields.Title;
   if(allIds.some(x=>String(x).toLowerCase()===fields.Title.toLowerCase()))throw Error('Asset ID ซ้ำ');
   await graph(path(list.id),'POST',{fields},true);
  }
  closeAssetModal();await refreshAssetData();showToast('บันทึกทะเบียนทรัพย์สินลง SharePoint แล้ว','success');
 }catch(e){showToast(e.message,'error');}
};
window.nineDeleteAsset=async function(id){
 if(!admin())return showToast('เฉพาะ Super Admin','error');
 const a=assets.find(x=>x.id===id);if(!a)return;
 if(a.status==='assigned')return showToast('ห้ามลบอุปกรณ์ที่มีผู้ถือครอง','error');
 if(!confirm('ยืนยันลบ Asset '+a.id+' ออกจาก SharePoint?'))return;
 try{
  const live=await currentAsset(a);
  if(live.busy)throw Error('ห้ามลบเพราะมีผู้ใช้งาน');
  await graph(path(a._sp.listId,a._sp.itemId),'DELETE',null,true,live.item['@odata.etag']);
  await refreshAssetData();showToast('ลบอุปกรณ์จาก SharePoint แล้ว','success');
 }catch(e){showToast(e.message,'error');}
};
const previousRenderAll=renderAll;
renderAll=function(){previousRenderAll();renderEmployees();};
const previousPopulate=populateEmployeePresets;
populateEmployeePresets=function(){
 previousPopulate();
 const select=document.getElementById('co-preset-emp');
 const keys=new Set([...select.options].map(o=>o.value));
 for(const e of directory){const k=e.empId||e.name;if(!keys.has(k)){select.add(new Option(e.name+' ('+(e.department||'')+')',k));keys.add(k);}}
};
const prevFill=fillEmployeePreset;
fillEmployeePreset=function(v){
 const e=directory.find(x=>x.empId===v||x.name===v);
 if(!e)return prevFill(v);
 [['co-emp-id',e.empId||e.name],['co-emp-name',e.name],['co-emp-dept',e.department||'-'],['co-emp-pos',e.position],['co-emp-contact',e.email||e.contact]].forEach(([id,val])=>document.getElementById(id).value=val||'');
};
const prevToggle=toggleCartItem;
toggleCartItem=function(id){
 const a=assets.find(x=>x.id===id);
 if(a && (a.status!=='available'||a.assignedTo)){return showToast('ไม่สามารถเบิกซ้ำ: อุปกรณ์ '+a.id+' มีผู้ใช้งานอยู่แล้ว','error');}
 prevToggle(id);
};
const prevInventory=renderInventory;
renderInventory=function(){
 prevInventory();
 if(!admin())return;
 // Existing buttons dispatch to the safe Graph-backed functions.
};
submitCheckout=function(e){return window.nineCheckout(e);};
submitCheckin=function(e,key){return window.nineCheckin(e,key);};
saveAsset=function(e){return window.nineSaveAsset(e);};
deleteAsset=function(id){return window.nineDeleteAsset(id);};
const prevOpen=openAssetModal;
openAssetModal=function(id=null){if(!admin())return showToast('เฉพาะ Super Admin','error');return prevOpen(id);};
document.addEventListener('DOMContentLoaded',()=>{ensurePanel();document.getElementById('emp-add').hidden=true;});

window.nineDirectory=()=>directory.map(x=>({...x}));
window.nineEmployeeHasSharedList=()=>Boolean(employeeListId);
window.nineEmployeeBulkUpsert=async function(rows){
 if(!admin())throw Error('เฉพาะ Super Admin');
 const done=[],errors=[];
 for(const e of rows){
  try{
   const match=directory.find(x=>(e.empId&&x.empId===e.empId)||norm(x.name)===norm(e.name));
   const fields={Title:e.name,EmployeeCode:e.empId||'',Department:e.department||'',Position:e.position||'',Email:e.email||'',Contact:e.contact||'',Company:e.company||'Nine distribution co.,ltd'};
   if(employeeListId){if(match?._itemId)await graph(path(employeeListId,match._itemId)+'/fields','PATCH',fields,true);else await graph(path(employeeListId),'POST',{fields},true);}
   else{if(match){Object.assign(match,e);match._derived=false;}else directory.push({...e,_derived:false});}
   done.push(e.name);
  }catch(err){errors.push(e.name+': '+err.message);}
 }
 if(employeeListId)await loadDirectory();else{localStorage.setItem('nine-employee-local-v2',JSON.stringify(directory.filter(x=>!x._derived)));renderEmployees();}
 return {done,errors,shared:Boolean(employeeListId)};
};

window.nineAfterLogin=async function(){
 ensurePanel();
 const isAdmin=admin();
 ['nav-checkout','nav-checkin','dashboard-quick-actions','btn-add-asset'].forEach(id=>document.getElementById(id)?.classList.toggle('hidden',!isAdmin));
 const ele=document.getElementById('topbar-role-badge');if(ele&&isAdmin)ele.textContent='Super Admin — SharePoint (Nutsarun / Satit)';
 transactions=JSON.parse(sessionStorage.getItem('nine-transactions-session')||'[]');
 try{await loadDirectory();renderAll();}catch(e){console.warn(e);showToast('อ่านข้อมูลพนักงาน: '+e.message,'error');}
};
})();

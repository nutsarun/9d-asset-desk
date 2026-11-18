/* 9D confirmation and nested settings UI. Does not change SharePoint data itself. */
(() => {
 'use strict';
 const $=id=>document.getElementById(id);
 const escapeHTML=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
 const root=()=>['nutsarun@ninedistribution.com','satit@ninedistribution.com'].includes(String(currentUser?.email||'').toLowerCase());
 const en=()=>window.nineLanguage?.()==='en';
 let pending=false,dialogPromise=null;
 const text=(th,english)=>en()?english:th;
 function createDialog(){
  if($('asset-action-confirm'))return;
  const modal=document.createElement('div');
  modal.id='asset-action-confirm';
  modal.className='fixed inset-0 z-[100] hidden items-center justify-center bg-slate-900/70 p-4';
  modal.setAttribute('role','presentation');
  modal.innerHTML='<div role="alertdialog" aria-modal="true" aria-labelledby="asset-confirm-title" aria-describedby="asset-confirm-description" class="bg-white rounded-2xl shadow-2xl max-w-lg w-full p-6 space-y-4">'+
   '<div class="flex gap-3 items-start"><div class="bg-amber-100 rounded-full p-2 text-amber-700"><i data-lucide="alert-triangle" class="w-5 h-5"></i></div>'+
   '<div class="flex-1"><h3 id="asset-confirm-title" class="text-lg font-bold text-slate-900"></h3><p id="asset-confirm-description" class="text-sm text-slate-600 mt-1"></p></div></div>'+
   '<div id="asset-confirm-details" class="bg-slate-50 border rounded-xl p-3 text-sm max-h-56 overflow-y-auto"></div>'+
   '<div class="bg-amber-50 border border-amber-200 rounded-lg p-3 text-xs text-amber-900" id="asset-confirm-warning"></div>'+
   '<div class="flex justify-end gap-3"><button type="button" id="asset-confirm-cancel" class="rounded-xl px-4 py-2.5 border bg-white text-slate-700 font-medium"></button>'+
   '<button type="button" id="asset-confirm-accept" class="rounded-xl px-5 py-2.5 bg-blue-600 text-white font-semibold"></button></div></div>';
  document.body.append(modal);
  $('asset-confirm-cancel').onclick=()=>settle(false);
  $('asset-confirm-accept').onclick=()=>settle(true);
  modal.addEventListener('mousedown',event=>{if(event.target===modal)settle(false)});
  document.addEventListener('keydown',event=>{if(event.key==='Escape'&&dialogPromise)settle(false)});
  window.lucide?.createIcons();
 }
 function settle(ok){
  if(!dialogPromise)return;
  const resolve=dialogPromise;dialogPromise=null;
  const m=$('asset-action-confirm');
  m.classList.add('hidden');m.classList.remove('flex');
  resolve(ok);
 }
 async function confirmAction(kind,employee,items){
  createDialog();
  if(dialogPromise)return false;
  const isReturn=kind==='return';
  $('asset-confirm-title').textContent=isReturn?text('ยืนยันการรับคืนทรัพย์สิน','Confirm Asset Return'):text('ยืนยันการส่งมอบทรัพย์สิน','Confirm Asset Handover');
  $('asset-confirm-description').textContent=text('ตรวจสอบพนักงานและรายการอุปกรณ์ก่อนบันทึกลง SharePoint','Review the employee and asset list before saving to SharePoint');
  $('asset-confirm-details').innerHTML='<div class="mb-2"><b>'+text('พนักงาน','Employee')+':</b> '+escapeHTML(employee||'-')+'</div>'+
   '<div class="mb-1"><b>'+text('จำนวนอุปกรณ์','Asset Count')+':</b> '+items.length+' '+text('รายการ','items')+'</div>'+
   '<ol class="list-decimal pl-5 space-y-1">'+items.map(a=>'<li><span class="font-mono font-semibold">'+escapeHTML(a.id||a)+'</span><span class="text-slate-500"> '+escapeHTML([a.brand,a.model].filter(Boolean).join(' '))+'</span></li>').join('')+'</ol>';
  $('asset-confirm-warning').textContent=isReturn?
   text('หากยืนยัน ระบบจะบันทึกรับคืนและเปลี่ยนสถานะอุปกรณ์ใน SharePoint โปรดตรวจสอบรายการให้ถูกต้อง','Confirming will record the return and update asset status in SharePoint.'):
   text('หากยืนยัน ระบบจะบันทึกการส่งมอบและเปลี่ยนผู้ถือครองใน SharePoint โปรดตรวจสอบรายการให้ถูกต้อง','Confirming will record the handover and update asset custody in SharePoint.');
  $('asset-confirm-cancel').textContent=text('ยกเลิก','Cancel');
  $('asset-confirm-accept').textContent=isReturn?text('ยืนยันรับคืน','Confirm Return'):text('ยืนยันส่งมอบ','Confirm Handover');
  const m=$('asset-action-confirm');m.classList.remove('hidden');m.classList.add('flex');
  $('asset-confirm-cancel').focus();
  return new Promise(resolve=>{dialogPromise=resolve});
 }
 function installGuards(){
  if(window.nineConfirmGuardsInstalled)return;
  // The original handlers already perform actual SharePoint updates. Only invoke them after confirmation.
  const saveHandover=submitCheckout,saveReturn=submitCheckin;
  submitCheckout=async function(event){
   event?.preventDefault?.();
   if(pending)return;
   const items=checkoutCart.map(id=>assets.find(a=>a.id===id)).filter(Boolean);
   const employee=$('co-emp-name')?.value.trim()||'';
   if(!items.length||!employee||!$('co-emp-id')?.value.trim()){
    showToast(text('กรุณาระบุพนักงาน รหัสพนักงาน และเลือกอุปกรณ์ก่อน','Please select assets, employee and employee ID first'),'error');return;
   }
   pending=true;
   try{if(await confirmAction('handover',employee,items))await saveHandover(event)}
   catch(err){showToast(err?.message||String(err),'error')}
   finally{pending=false}
  };
  submitCheckin=async function(event,groupKey){
   event?.preventDefault?.();
   if(pending)return;
   const inputs=[...event.target.querySelectorAll('input[name="return_asset"]:checked')];
   if(!inputs.length){showToast(text('กรุณาเลือกอุปกรณ์ที่จะรับคืน','Select assets to return'),'error');return}
   const group=groupAssignedAssetsByEmployee()[groupKey];
   if(!group){showToast(text('ไม่พบพนักงาน','Employee not found'),'error');return}
   const items=inputs.map(input=>assets.find(a=>a.id===input.value)).filter(Boolean);
   pending=true;
   try{if(await confirmAction('return',group.employee?.name,items))await saveReturn(event,groupKey)}
   catch(err){showToast(err?.message||String(err),'error')}
   finally{pending=false}
  };
  window.nineConfirmGuardsInstalled=true;
 }
 function nestedSettings(){
  const settings=$('tab-settings'),users=$('tab-portal-users'),master=$('tab-master');
  if(!settings||!users||!master||$('settings-submenu'))return;
  // Keep the existing settings language form, then create a second-level menu.
  const original=[...settings.children];
  const general=document.createElement('div');general.id='settings-general-content';general.className='space-y-4';
  original.forEach(node=>general.append(node));
  const nav=document.createElement('div');nav.id='settings-submenu';nav.className='grid grid-cols-1 sm:grid-cols-2 gap-3 mb-5';
  const actions=[
   {key:'general',title:text('ภาษาและการแสดงผล','Language & Display'),info:text('ภาษาไทย / English','Thai / English'),icon:'languages'},
   {key:'sso',title:'Microsoft SSO',info:text('ตั้งค่าการเข้าสู่ระบบ Microsoft 365','Microsoft 365 sign-in settings'),icon:'shield-check',privileged:true},
   {key:'users',title:text('จัดการผู้ใช้งาน','User Management'),info:text('บัญชีผู้ใช้และสิทธิ์การเข้าถึง','Accounts and access roles'),icon:'users',privileged:true},
   {key:'master',title:text('รวมข้อมูล SharePoint','SharePoint Data Migration'),info:text('Master Asset Inventory','Master Asset Inventory'),icon:'database',privileged:true}
  ];
  for(const a of actions){
   const b=document.createElement('button');b.type='button';b.dataset.settingsTab=a.key;
   b.className='settings-inner-btn text-left flex items-center gap-3 border border-slate-200 rounded-xl p-4 bg-white hover:border-blue-400 hover:bg-blue-50';
   b.innerHTML='<i data-lucide="'+a.icon+'" class="w-5 h-5 text-blue-600"></i><span class="flex-1"><b class="block text-sm">'+a.title+'</b><small class="block text-slate-500 mt-1">'+a.info+'</small></span>';
   b.onclick=()=>switchInner(a.key);
   if(a.privileged)b.dataset.adminOnly='true';
   nav.append(b);
  }
  settings.append(nav,general);
  for(const [node,id] of [[users,'users'],[master,'master']]){
   node.classList.remove('tab-content');node.classList.add('hidden');
   node.dataset.settingsPanel=id;
   node.classList.add('mt-4');settings.append(node);
  }
  const existing=[['nav-sso-settings'],['nav-portal-users'],['nav-master']];
  for(const [id] of existing){const b=$(id);if(b){b.classList.add('hidden');b.style.display='none'}}
  window.nineSettingsNav={open:switchInner,refresh:refreshSettings};
  refreshSettings();
  switchInner('general');
  window.lucide?.createIcons();
 }
 function refreshSettings(){
  for(const node of document.querySelectorAll('#settings-submenu [data-admin-only]')){
   node.classList.toggle('hidden',!root());
  }
  for(const id of ['nav-sso-settings','nav-portal-users','nav-master']){
   const node=$(id);if(node)node.style.display='none';
  }
 }
 function switchInner(key){
  refreshSettings();
  if(key==='sso'){
   if(!root())return showToast(text('เฉพาะ Super Admin','Super Admin only'),'error');
   // Reuse the existing Microsoft SSO modal; do not expose its config on the main menu.
   openSsoSetupModal(false);return;
  }
  if((key==='users'||key==='master')&&!root())return showToast(text('เฉพาะ Super Admin','Super Admin only'),'error');
  for(const id of ['settings-general-content','tab-portal-users','tab-master'])$(id)?.classList.add('hidden');
  if(key==='general')$('settings-general-content')?.classList.remove('hidden');
  if(key==='users')$('tab-portal-users')?.classList.remove('hidden');
  if(key==='master')$('tab-master')?.classList.remove('hidden');
  document.querySelectorAll('#settings-submenu [data-settings-tab]').forEach(button=>{
   const current=button.dataset.settingsTab===key;
   button.classList.toggle('border-blue-600',current);button.classList.toggle('bg-blue-50',current);
  });
 }
 document.addEventListener('DOMContentLoaded',()=>{
  createDialog();installGuards();nestedSettings();
 });
 // Current code instantiates the settings panels in DOMContentLoaded too; guard against order changes.
 const originalLogin=window.nineAfterLogin;
 window.nineAfterLogin=async function(){
  await originalLogin?.();
  nestedSettings();refreshSettings();
 };
})();

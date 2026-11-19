/* 9D Master Inventory — single authoritative asset list; no legacy migration. */
(function(){
 'use strict';
 const SITE='ninedistribution.sharepoint.com,13d8bc75-67a0-4dbf-b579-8b26c2dd2e9f,dedddf0b-d21b-4660-84ad-259b6ffb0ab4';
 const MASTER_ID='509b2348-ae5e-410e-b8af-eff3595e752b';
 const ROOT='https://graph.microsoft.com/v1.0/sites/'+SITE+'/lists/'+MASTER_ID;
 const isAdmin=()=>['nutsarun@ninedistribution.com','satit@ninedistribution.com'].includes(String(currentUser?.email||'').toLowerCase());
 async function chooseListForLogin(){
   // Must never silently fall back to the old Computer and Accessories lists.
   NINE_GRAPH.lists=[{id:MASTER_ID,kind:'master',category:'Other'}];
   return {id:MASTER_ID,displayName:'9D Master Asset Inventory'};
 }
 function buildMenu(){
   if(document.getElementById('nav-master'))return;
   const nav=document.querySelector('#sidebar nav');
   if(!nav)return;
   const btn=document.createElement('button');btn.id='nav-master';btn.dataset.tab='master';
   btn.className='nav-btn w-full flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-medium hover:bg-slate-800 hover:text-white';
   btn.innerHTML='<i data-lucide="database-backup" class="w-5 h-5"></i><span>สำรองและกู้คืนข้อมูล</span>';
   btn.onclick=()=>switchTab('master');
   nav.append(btn);
   const sec=document.createElement('section');sec.id='tab-master';sec.className='tab-content hidden space-y-4';
   sec.innerHTML='<h2 class="text-2xl font-bold">Database Backup & Restore</h2><p class="text-sm text-slate-600">ฐานข้อมูลหลัก: 9D Master Asset Inventory</p><div id="master-backup-mount"></div>';
   document.querySelector('#app-shell main')?.append(sec);
   btn.classList.toggle('hidden',!isAdmin());
   window.nineDatabaseBackup?.mount?.();
   window.lucide?.createIcons();
 }
 window.nineMaster={chooseListForLogin,buildMenu,getActive:()=>({id:MASTER_ID,kind:'master',category:'Other'})};
 document.addEventListener('DOMContentLoaded',buildMenu);
})();
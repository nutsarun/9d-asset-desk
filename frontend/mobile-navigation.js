/* 9D mobile navigation and blank-dashboard recovery */
(() => {
'use strict';
const names=[
 ['dashboard','ภาพรวม','layout-dashboard'],
 ['inventory','คลังทรัพย์สิน','hard-drive'],
 ['employees','พนักงาน','users'],
 ['annual','ตรวจนับประจำปี','clipboard-check']
];
function isLoggedIn(){return !!document.querySelector('#app-shell:not(.hidden)') && !!window.nineGraphAuth;}
function make(){
 const main=document.querySelector('#app-shell main');
 if(!main||document.getElementById('mobile-quick-nav'))return;
 const bar=document.createElement('nav');bar.id='mobile-quick-nav';
 bar.setAttribute('aria-label','เมนูหลักสำหรับมือถือ');
 bar.className='lg:hidden bg-white border border-slate-200 rounded-xl p-2 mb-4 shadow-sm';
 bar.innerHTML='<div class="grid grid-cols-2 gap-2">'+names.map(([id,title])=>
  '<button type="button" data-mobile-tab="'+id+'" class="mobile-quick-btn border border-slate-200 rounded-lg px-3 py-3 text-sm text-slate-700 font-semibold bg-slate-50">'+title+'</button>'
 ).join('')+'</div>';
 const banner=main.querySelector('#topbar-role-badge')?.closest('.mb-6');
 if(banner)banner.insertAdjacentElement('afterend',bar);else main.prepend(bar);
 bar.querySelectorAll('[data-mobile-tab]').forEach(b=>b.onclick=()=>show(b.dataset.mobileTab));
}
function show(name){
 const target=document.getElementById('tab-'+name);
 if(!target)return;
 try {switchTab(name);}catch(e){console.warn('Navigation refresh failed',e);}
 // Do not hide the app when a dashboard render fails; navigation stays available.
 document.querySelectorAll('.tab-content').forEach(node=>node.classList.toggle('hidden',node!==target));
 target.classList.remove('hidden');
 const nav=document.getElementById('mobile-quick-nav');
 nav?.querySelectorAll('[data-mobile-tab]').forEach(b=>{
  const active=b.dataset.mobileTab===name;
  b.classList.toggle('bg-blue-600',active);
  b.classList.toggle('text-white',active);
  b.classList.toggle('bg-slate-50',!active);
  b.setAttribute('aria-current',active?'page':'false');
 });
 if(name==='annual')window.nineAnnualV2?.render?.();
}
window.nineMobileTabs={show};
document.addEventListener('DOMContentLoaded',make);
})();
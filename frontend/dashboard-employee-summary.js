/* Dashboard employee custody summary + annual verification printouts */
(() => {
'use strict';
const $=id=>document.getElementById(id);
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const sel=new Set();let people=[];
const hasValue=a=>a.assetValue!==null&&a.assetValue!==undefined&&a.assetValue!==''&&Number.isFinite(Number(a.assetValue))&&Number(a.assetValue)>=0;
const money=n=>Number(n).toLocaleString('th-TH',{minimumFractionDigits:2,maximumFractionDigits:2});
const valueSummary=items=>({total:items.reduce((sum,a)=>sum+(hasValue(a)?Number(a.assetValue):0),0),missing:items.filter(a=>!hasValue(a)).length});
const formattedValue=a=>hasValue(a)?'฿'+money(a.assetValue):'ยังไม่ระบุ';
const yr=()=>Number($('emp-summary-year')?.value)||new Date().getFullYear();
function groups(){
 const m=new Map(),filter=$('emp-summary-company')?.value||'ALL';
 for(const a of assets){
  if(a.status!=='assigned'||!a.assignedTo?.name)continue;
  if(filter!=='ALL'&&a.company!==filter)continue;
  const name=String(a.assignedTo.name).trim(),id=name.toLowerCase();
  if(!m.has(id))m.set(id,{id,name,employee:a.assignedTo,companies:new Set(),items:[]});
  const p=m.get(id);p.items.push(a);p.companies.add(a.company||'Nine distribution co.,ltd');
 }
 return [...m.values()].sort((a,b)=>a.name.localeCompare(b.name,'th'));
}
const storage=a=>'9d-annual-'+yr()+'-'+a._sp?.listId+'-'+a._sp?.itemId;
function review(a){try{return JSON.parse(localStorage.getItem(storage(a))||'{}')}catch{return{}}}
function save(a,change){localStorage.setItem(storage(a),JSON.stringify({...review(a),...change,checkedAt:new Date().toISOString(),checkedBy:currentUser?.email||''}))}
function build(){
 document.getElementById('nav-annual')?.remove();
 const old=$('tab-annual');if(old)old.style.display='none';
 if($('employee-summary-panel'))return;
 const dashboard=$('tab-dashboard');if(!dashboard)return;
 const btn=document.createElement('button');btn.id='employee-summary-button';btn.className='bg-indigo-600 text-white px-4 py-2.5 rounded-xl font-semibold text-sm';btn.textContent='สรุปทรัพย์สินพนักงาน / ตรวจสภาพ';
 btn.onclick=()=>{const p=$('employee-summary-panel');p.classList.toggle('hidden');if(!p.classList.contains('hidden')){draw();p.scrollIntoView({behavior:'smooth',block:'start'})}};
 ($('dashboard-quick-actions')||dashboard).append(btn);
 const panel=document.createElement('section');panel.id='employee-summary-panel';panel.className='hidden bg-white border rounded-2xl p-4 space-y-4';
 panel.innerHTML='<h3 class="font-bold text-xl">สรุปการถือครองทรัพย์สินรายบุคคล</h3><div class="grid grid-cols-2 gap-3"><label>ปีตรวจนับ<input type="number" id="emp-summary-year" class="block w-full p-2 border rounded-lg"></label><label>บริษัท<select id="emp-summary-company" class="block w-full p-2 border rounded-lg"><option value="ALL">ทุกสังกัด</option><option value="Nine distribution co.,ltd">Nine Distribution</option><option value="After Nine Co.,ltd">After Nine</option></select></label></div><div class="flex flex-wrap gap-2"><button id="summary-choose-all" class="border rounded-lg px-3 py-2">เลือกทั้งหมด</button><button id="summary-choose-none" class="border rounded-lg px-3 py-2">ยกเลิกทั้งหมด</button><button id="summary-document" class="bg-blue-600 text-white rounded-lg px-4 py-2">ออกเอกสาร PDF รายที่เลือก</button></div><p id="summary-count" class="text-sm text-slate-500"></p><div id="summary-people" class="space-y-3"></div><p class="text-xs text-amber-700 bg-amber-50 p-3 rounded-lg">สรุปข้อมูลผู้ถือครองจาก SharePoint ปัจจุบัน ส่วนผลตรวจสภาพบันทึกเฉพาะเบราว์เซอร์นี้จนกว่าจะมีทะเบียนตรวจสอบกลาง</p>';
 dashboard.querySelector('#dashboard-kpis')?.insertAdjacentElement('beforebegin',panel);
 $('emp-summary-year').value=new Date().getFullYear();
 $('emp-summary-year').onchange=draw;$('emp-summary-company').onchange=draw;
 $('summary-choose-all').onclick=()=>{people.forEach(g=>sel.add(g.id));draw()};
 $('summary-choose-none').onclick=()=>{sel.clear();draw()};
 $('summary-document').onclick=()=>print(null);
 const css=document.createElement('style');css.textContent='.custody-paper{font:11px/1.45 Sarabun,sans-serif;color:#0f172a}.custody-paper .page{break-after:page;page-break-after:always}.custody-paper .page:last-child{break-after:auto;page-break-after:auto}.custody-paper .head{display:flex;align-items:center;gap:12px;border-bottom:2px solid #334155;padding-bottom:8px;margin-bottom:10px}.custody-paper .head img{width:110px;height:75px;object-fit:contain}.custody-paper h1{font-size:18px;font-weight:bold}.custody-paper h3{background:#f1f5f9;border-left:4px solid #475569;padding:6px;margin:12px 0 7px}.custody-paper table{table-layout:fixed;width:100%;border-collapse:collapse;font-size:10px}.custody-paper th,.custody-paper td{border:1px solid #94a3b8;padding:6px;vertical-align:top;overflow-wrap:anywhere}.custody-paper tr{break-inside:avoid}.custody-paper .sign{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-top:14px}.custody-paper .sign div{border:1px solid #cbd5e1;padding:10px 3px;min-height:96px;text-align:center}@media print{@page{size:A4 portrait;margin:12mm}#print-modal{display:block!important;position:static!important;inset:auto!important;overflow:visible!important;padding:0!important}#print-modal-wrapper{max-width:none!important;width:100%!important;margin:0!important}#print-modal-wrapper>.bg-slate-200\\/80{padding:0!important;background:white!important;overflow:visible!important}#a4-document-content{padding:0!important;min-height:0!important;box-shadow:none!important}.no-print{display:none!important}}';
 document.head.append(css);
}
function draw(){
 if(!$('summary-people'))return;
 people=groups();
 const valid=new Set(people.map(x=>x.id));for(const k of [...sel])if(!valid.has(k))sel.delete(k);
 $('summary-count').textContent='พนักงาน '+people.length+' คน / อุปกรณ์ '+people.reduce((n,g)=>n+g.items.length,0)+' รายการ / เลือก '+sel.size+' คน';
 $('summary-document').disabled=!sel.size;
 $('summary-people').innerHTML=people.map(g=>{
  const summary={};for(const a of g.items){const cat=a.assetType||a.category;summary[cat]=(summary[cat]||0)+1}
  const valuation=valueSummary(g.items);
  const rows=g.items.map(a=>{const i=assets.indexOf(a),r=review(a);return '<tr class="border-t"><td class="p-2 text-xs">'+esc(a.id)+'<br>'+esc(a.serial||'')+'</td><td class="p-2">'+esc(a.assetType||a.category)+' / '+esc(a.brand)+' '+esc(a.model)+'</td><td class="p-2 whitespace-nowrap '+(hasValue(a)?'text-blue-700':'text-amber-700')+'">'+formattedValue(a)+'</td><td class="p-2"><select class="employee-condition border rounded p-1" data-i="'+i+'"><option value="">รอตรวจ</option><option value="good" '+(r.condition==='good'?'selected':'')+'>สมบูรณ์ดี</option><option value="bad" '+(r.condition==='bad'?'selected':'')+'>ไม่สมบูรณ์</option></select></td><td class="p-2"><input class="employee-reason border rounded p-1 w-40" data-i="'+i+'" '+(r.condition==='bad'?'':'disabled')+' placeholder="ไม่สมบูรณ์เพราะ..." value="'+esc(r.reason||'')+'"></td></tr>'}).join('');
  return '<div class="border rounded-xl overflow-hidden"><div class="bg-slate-50 p-3 flex justify-between flex-wrap gap-2"><label class="flex items-center gap-2"><input type="checkbox" class="employee-pick" data-id="'+esc(g.id)+'" '+(sel.has(g.id)?'checked':'')+'><span><b>'+esc(g.name)+'</b> · '+g.items.length+' รายการ<br><small>'+esc(Object.entries(summary).map(x=>x[0]+': '+x[1]).join(' / '))+'</small><br><span class="text-sm font-semibold text-blue-800">มูลค่าทรัพย์สินรวม ฿'+money(valuation.total)+'</span>'+(valuation.missing?'<br><small class="text-amber-700">ยังไม่ระบุมูลค่า '+valuation.missing+' รายการ</small>':'')+'</span></label><button data-id="'+esc(g.id)+'" class="employee-single border bg-white rounded-lg px-3 py-2">เอกสารรายคน</button></div><div class="overflow-x-auto"><table class="w-full text-sm"><thead><tr><th>Asset ID / Serial</th><th>ประเภท / รุ่น</th><th>มูลค่าทรัพย์สิน</th><th>สภาพ</th><th>กรณีไม่สมบูรณ์เพราะ</th></tr></thead><tbody>'+rows+'</tbody></table></div></div>';
 }).join('')||'<p class="p-4">ไม่พบพนักงานที่ครอบครองทรัพย์สิน</p>';
 $('summary-people').querySelectorAll('.employee-pick').forEach(x=>x.onchange=()=>{if(x.checked)sel.add(x.dataset.id);else sel.delete(x.dataset.id);draw()});
 $('summary-people').querySelectorAll('.employee-single').forEach(x=>x.onclick=()=>print(x.dataset.id));
 $('summary-people').querySelectorAll('.employee-condition').forEach(x=>x.onchange=()=>{const a=assets[Number(x.dataset.i)];save(a,{condition:x.value,reason:x.value==='bad'?review(a).reason||'':''});draw()});
 $('summary-people').querySelectorAll('.employee-reason').forEach(x=>x.onchange=()=>save(assets[Number(x.dataset.i)],{reason:x.value.trim()}));
}
function personPaper(g){
 const parts=Math.ceil(g.items.length/7),output=[];
 for(let p=0;p<parts;p++){
  const items=g.items.slice(p*7,(p+1)*7);
  const valuation=valueSummary(g.items);
  const header='<div class="head"><img src="./nine-distribution-brand.webp"><div style="flex:1"><b>'+esc([...g.companies].join(' / '))+'</b><h1>แบบสรุปและตรวจสภาพทรัพย์สินรายบุคคล</h1>EMPLOYEE ASSET CONDITION CONFIRMATION</div><div>ปี '+yr()+'<br>หน้า '+(p+1)+'/'+parts+'</div></div>';
  const rows=items.map((a,i)=>{const r=review(a);return '<tr><td>'+(p*7+i+1)+'</td><td>'+esc(a.id)+'<br>'+esc(a.serial||'-')+'</td><td>'+esc(a.assetType||a.category)+'<br>'+esc(a.brand)+' '+esc(a.model)+'</td><td style="text-align:right;white-space:nowrap">'+formattedValue(a)+'</td><td>'+(r.condition==='good'?'☑':'☐')+' สมบูรณ์ดี<br>'+(r.condition==='bad'?'☑':'☐')+' ไม่สมบูรณ์</td><td>'+esc(r.condition==='bad'?r.reason||'ยังไม่ระบุสาเหตุ':'..........................')+'</td></tr>'}).join('');
  const table='<h3>รายการอุปกรณ์และผลการตรวจ</h3><table><colgroup><col style="width:5%"><col style="width:17%"><col style="width:23%"><col style="width:17%"><col style="width:17%"><col style="width:21%"></colgroup><thead><tr><th>#</th><th>Asset ID / Serial</th><th>ประเภท / รุ่น</th><th>มูลค่า (บาท)</th><th>สภาพ</th><th>ไม่สมบูรณ์เพราะ</th></tr></thead><tbody>'+rows+'</tbody></table>';
  const info='<h3>ข้อมูลผู้ถือครอง</h3>ชื่อ-นามสกุล: <b>'+esc(g.name)+'</b>　รหัสพนักงาน: '+esc(g.employee.empId||'-')+'<br>แผนก: '+esc(g.employee.department||'-')+'　ตำแหน่ง: '+esc(g.employee.position||'-')+'<br>จำนวนอุปกรณ์ทั้งหมด: <b>'+g.items.length+' รายการ</b>'+
   '<br>มูลค่าทรัพย์สินที่ครอบครองรวม: <b>฿'+money(valuation.total)+'</b>'+
   (valuation.missing?' <span style="color:#92400e">(ยังไม่ระบุมูลค่า '+valuation.missing+' รายการ)</span>':'');
  const sign=p===parts-1?'<h3>การรับรองรายการและสภาพทรัพย์สิน</h3>☐ อุปกรณ์อยู่ครบตามรายการ　☐ มีรายการต้องติดตามแก้ไข<p>ข้าพเจ้าได้ตรวจสอบรายการทรัพย์สินที่ครอบครองและแจ้งสภาพตามความเป็นจริงแล้ว</p><div class="sign"><div><b>ผู้ถือครอง</b><p style="margin-top:23px">.................................</p>'+esc(g.name)+'<br>วันที่ ...../...../.....</div><div><b>ผู้ตรวจนับ IT</b><p style="margin-top:23px">.................................</p>ชื่อ ..........................<br>วันที่ ...../...../.....</div><div><b>ผู้อนุมัติ</b><p style="margin-top:23px">.................................</p>ชื่อ ..........................<br>วันที่ ...../...../.....</div></div>':'';
  output.push('<div class="page">'+header+info+table+sign+'</div>');
 }
 return output.join('');
}
function print(person){
 people=groups();
 const chosen=person?people.filter(g=>g.id===person):people.filter(g=>sel.has(g.id));
 if(!chosen.length)return showToast('กรุณาเลือกพนักงาน','error');
 const bad=chosen.flatMap(g=>g.items.filter(a=>review(a).condition==='bad'&&!review(a).reason).map(a=>a.id));
 if(bad.length)return showToast('กรุณาระบุสาเหตุที่ไม่สมบูรณ์: '+bad.slice(0,3).join(', '),'error');
 $('a4-document-content').innerHTML='<div class="custody-paper">'+chosen.map(personPaper).join('')+'</div>';
 const m=$('print-modal');m.classList.remove('hidden');m.classList.add('block');
 const h=m.querySelector('.no-print span.font-semibold');if(h)h.textContent='เอกสารสรุปทรัพย์สิน '+chosen.length+' พนักงาน · เลือก Save as PDF';
}
const prior=window.nineUpgradeAfterLogin;
window.nineUpgradeAfterLogin=function(){prior?.();build();draw()};
document.addEventListener('DOMContentLoaded',()=>{build();draw()});
window.nineEmployeeSummary={build,draw,print};
})();
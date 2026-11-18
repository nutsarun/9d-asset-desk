/* Annual employee audit and printable PDF forms */
(() => {
 const $=id=>document.getElementById(id),safe=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
 const year=()=>Number($('annual-year')?.value)||new Date().getFullYear(),selected=new Set();
 let people=[];
 function groups(){
  const result=new Map(),filter=$('annual-company')?.value||'ALL';
  assets.forEach(a=>{
   if(a.status!=='assigned'||!a.assignedTo?.name||(filter!=='ALL'&&a.company!==filter))return;
   const name=a.assignedTo.name.replace(/\s+/g,' ').trim(),id=name.toLowerCase();
   if(!result.has(id))result.set(id,{id,name,employee:a.assignedTo,items:[],companies:new Set()});
   result.get(id).items.push(a);result.get(id).companies.add(a.company||'Nine distribution co.,ltd');
  });
  return [...result.values()].sort((a,b)=>a.name.localeCompare(b.name,'th'));
 }
 const key=a=>'9d-annual-'+year()+'-'+a._sp?.listId+'-'+a._sp?.itemId;
 const check=a=>{try{return JSON.parse(localStorage.getItem(key(a))||'{}')}catch{return{}}};
 function setup(){
  if($('audit-control-v2')||!$('tab-annual'))return;
  $('annual-items').style.display='none';$('annual-print').hidden=true;
  const section=document.createElement('div');section.id='audit-control-v2';section.className='space-y-3';
  section.innerHTML='<div class="bg-white border rounded-xl p-4 flex gap-2 items-center flex-wrap justify-between"><b id="audit-summary-v2">รายงานพนักงาน</b><div class="flex gap-2 flex-wrap"><button id="audit-all-v2" class="border px-3 py-2 rounded-lg">เลือกทั้งหมด</button><button id="audit-none-v2" class="border px-3 py-2 rounded-lg">ล้างที่เลือก</button><button id="audit-pdf-v2" class="bg-blue-600 text-white px-4 py-2 rounded-lg">Export PDF รายที่เลือก</button></div></div><p class="text-xs bg-amber-50 border rounded-lg p-3">เลือก Save as PDF ในหน้าต่างพิมพ์ของเบราว์เซอร์ ข้อมูลผู้ถือครองเป็นข้อมูลปัจจุบัน ผลตรวจสภาพยังเก็บเฉพาะในเบราว์เซอร์</p><div id="audit-employees-v2" class="space-y-3"></div>';
  $('tab-annual').append(section);
  $('audit-all-v2').onclick=()=>{people.forEach(g=>selected.add(g.id));render()};
  $('audit-none-v2').onclick=()=>{selected.clear();render()};
  $('audit-pdf-v2').onclick=()=>print(null);
  ['annual-year','annual-company'].forEach(id=>$(id)?.addEventListener('change',render));
  const css=document.createElement('style');
  css.textContent='.annual-form{font:11px Sarabun,sans-serif;line-height:1.45;color:#0f172a}.annual-form .sheet{break-after:page;page-break-after:always}.annual-form .sheet:last-child{break-after:auto;page-break-after:auto}.annual-form .head{display:flex;gap:12px;align-items:center;border-bottom:2px solid #334155;margin-bottom:10px;padding-bottom:8px}.annual-form .head img{width:90px;height:60px;object-fit:contain}.annual-form h1{font-size:18px;font-weight:bold}.annual-form h2{font-size:12px;background:#f1f5f9;padding:6px;border-left:4px solid #334155;margin:12px 0 7px}.annual-form table{border-collapse:collapse;table-layout:fixed;width:100%;font-size:10px}.annual-form th,.annual-form td{border:1px solid #94a3b8;padding:5px;vertical-align:top;overflow-wrap:anywhere}.annual-form tr{break-inside:avoid}.annual-form .sign{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-top:15px}.annual-form .sign>div{border:1px solid #cbd5e1;min-height:95px;text-align:center;padding:8px 3px}@media print{@page{size:A4 portrait;margin:12mm}#print-modal{display:block!important;position:static!important;inset:auto!important;overflow:visible!important;padding:0!important}#print-modal-wrapper{max-width:none!important;width:100%!important;margin:0!important}#print-modal-wrapper>.bg-slate-200\\/80{background:white!important;padding:0!important;overflow:visible!important}#a4-document-content{min-height:0!important;width:100%!important;padding:0!important;box-shadow:none!important}.no-print{display:none!important}}';
  document.head.append(css);
 }
 function render(){
  if(!$('audit-employees-v2'))return;
  people=groups();const keys=new Set(people.map(x=>x.id));[...selected].forEach(x=>{if(!keys.has(x))selected.delete(x)});
  $('audit-summary-v2').textContent='พนักงาน '+people.length+' คน · อุปกรณ์ '+people.reduce((n,g)=>n+g.items.length,0)+' รายการ · เลือก '+selected.size+' คน';
  $('audit-pdf-v2').disabled=!selected.size;
  $('audit-pdf-v2').classList.toggle('opacity-50',!selected.size);
  $('audit-employees-v2').innerHTML=people.map(g=>{
   const totals=new Map();g.items.forEach(a=>totals.set(a.assetType||a.category,(totals.get(a.assetType||a.category)||0)+1));
   const rows=g.items.map(a=>{const idx=assets.indexOf(a),c=check(a);
    return '<tr class="border-t"><td class="p-2 text-xs">'+safe(a.id)+'<br>'+safe(a.serial||'-')+'</td><td class="p-2">'+safe(a.assetType||a.category)+' · '+safe(a.brand)+' '+safe(a.model)+'</td><td class="p-2"><select class="audit-entry border p-1 rounded" data-i="'+idx+'" data-field="condition"><option value="">รอตรวจ</option><option value="ปกติ" '+(c.condition==='ปกติ'?'selected':'')+'>ปกติ</option><option value="ผิดปกติ" '+(c.condition==='ผิดปกติ'?'selected':'')+'>ชำรุด/ผิดปกติ</option></select></td><td class="p-2"><select class="audit-entry border p-1 rounded" data-i="'+idx+'" data-field="complete"><option value="">รอตรวจ</option><option value="ครบ" '+(c.complete==='ครบ'?'selected':'')+'>ครบ</option><option value="ไม่ครบ" '+(c.complete==='ไม่ครบ'?'selected':'')+'>ไม่ครบ</option></select></td><td class="p-2"><input class="audit-entry border p-1 rounded w-36" data-i="'+idx+'" data-field="note" value="'+safe(c.note||'')+'" placeholder="รายละเอียด"></td></tr>';
   }).join('');
   return '<section class="bg-white border rounded-xl overflow-hidden"><div class="p-3 bg-slate-50 flex justify-between flex-wrap gap-2"><label class="flex gap-3 items-center"><input type="checkbox" class="audit-person-v2" data-id="'+safe(g.id)+'" '+(selected.has(g.id)?'checked':'')+'><span><b>'+safe(g.name)+'</b> — '+g.items.length+' รายการ<div class="text-xs text-slate-500">'+safe([...g.companies].join(' / '))+' | '+[...totals].map(([k,n])=>safe(k)+' '+n).join(' · ')+'</div></span></label><button class="audit-single-v2 border bg-white px-3 py-2 rounded-lg" data-id="'+safe(g.id)+'">Export PDF</button></div><div class="overflow-x-auto"><table class="w-full text-sm"><thead><tr><th>Asset ID / Serial</th><th>อุปกรณ์</th><th>สภาพ</th><th>ครบถ้วน</th><th>หมายเหตุ</th></tr></thead><tbody>'+rows+'</tbody></table></div></section>';
  }).join('')||'<div class="p-6 bg-white rounded-xl">ไม่พบพนักงานที่ครอบครองทรัพย์สิน</div>';
  $('audit-employees-v2').querySelectorAll('.audit-person-v2').forEach(c=>c.onchange=()=>{c.checked?selected.add(c.dataset.id):selected.delete(c.dataset.id);render()});
  $('audit-employees-v2').querySelectorAll('.audit-single-v2').forEach(b=>b.onclick=()=>print(b.dataset.id));
  $('audit-employees-v2').querySelectorAll('.audit-entry').forEach(input=>input.onchange=()=>{const a=assets[Number(input.dataset.i)],c=check(a);c[input.dataset.field]=input.value;c.checkedAt=new Date().toISOString();c.checkedBy=currentUser?.email;localStorage.setItem(key(a),JSON.stringify(c))});
 }
 function form(g){
  const pages=[],size=8,count=Math.ceil(g.items.length/size);
  for(let p=0;p<count;p++){
   const header='<div class="head"><img src="./nine-distribution-logo.webp"><div style="flex:1"><b>'+safe([...g.companies].join(' / '))+'</b><h1>แบบตรวจนับและตรวจสภาพทรัพย์สินประจำปี</h1>ANNUAL IT ASSET VERIFICATION</div><div>ปี '+year()+'<br>หน้า '+(p+1)+'/'+count+'</div></div>';
   const info='<h2>ข้อมูลผู้ถือครอง</h2>ชื่อ <b>'+safe(g.name)+'</b>　รหัส '+safe(g.employee.empId||'-')+'<br>แผนก '+safe(g.employee.department||'-')+'　ตำแหน่ง '+safe(g.employee.position||'-')+'<br>จำนวนทรัพย์สินทั้งหมด <b>'+g.items.length+' รายการ</b>';
   const rows=g.items.slice(p*size,(p+1)*size).map((a,i)=>{const c=check(a);
    return '<tr><td>'+(p*size+i+1)+'</td><td><b>'+safe(a.id)+'</b><br>'+safe(a.serial||'-')+'</td><td>'+safe(a.assetType||a.category)+'<br>'+safe(a.brand)+' '+safe(a.model)+'</td><td>'+(c.condition==='ปกติ'?'☑':'☐')+' ปกติ<br>'+(c.condition==='ผิดปกติ'?'☑':'☐')+' ชำรุด</td><td>'+(c.complete==='ครบ'?'☑':'☐')+' ครบ<br>'+(c.complete==='ไม่ครบ'?'☑':'☐')+' ไม่ครบ</td><td>'+safe(c.note||'..................................')+'</td></tr>';
   }).join('');
   const table='<h2>รายการตรวจสอบทรัพย์สิน</h2><table><colgroup><col style="width:5%"><col style="width:20%"><col style="width:25%"><col style="width:12%"><col style="width:12%"><col style="width:26%"></colgroup><thead><tr><th>#</th><th>Asset ID / Serial</th><th>ประเภท / รุ่น</th><th>สภาพ</th><th>ครบถ้วน</th><th>หมายเหตุ / เสียหาย</th></tr></thead><tbody>'+rows+'</tbody></table>';
   const sign=p===count-1?'<h2>สรุปผลการตรวจและรับรอง</h2>☐ อยู่ครบ　☐ ไม่ครบ　☐ พบอุปกรณ์ชำรุด<p>หมายเหตุเพิ่มเติม: ...........................................................................................................</p><p>ข้าพเจ้าได้ตรวจสอบรายการและสภาพทรัพย์สินข้างต้นแล้ว และยืนยันผลตามที่ระบุ</p><div class="sign"><div><b>ผู้ถือครอง</b><p style="margin-top:25px">................................</p>'+safe(g.name)+'<br>วันที่ ....../....../......</div><div><b>ผู้ตรวจนับ (IT)</b><p style="margin-top:25px">................................</p>ชื่อ ................................<br>วันที่ ....../....../......</div><div><b>ผู้อนุมัติ</b><p style="margin-top:25px">................................</p>ชื่อ ................................<br>วันที่ ....../....../......</div></div>':'';
   pages.push('<div class="sheet">'+header+info+table+sign+'</div>');
  }
  return pages.join('');
 }
 function print(person){
  people=groups();const chosen=person?people.filter(x=>x.id===person):people.filter(x=>selected.has(x.id));
  if(!chosen.length)return showToast('กรุณาเลือกพนักงาน','error');
  $('a4-document-content').innerHTML='<div class="annual-form">'+chosen.map(form).join('')+'</div>';
  const modal=$('print-modal');modal.classList.remove('hidden');modal.classList.add('block');
  const btn=modal.querySelector('button[onclick*="print"] span');if(btn)btn.textContent='Export PDF (Save as PDF)';
 }
 const prior=window.nineUpgradeAfterLogin;
 window.nineUpgradeAfterLogin=()=>{prior?.();setup();render()};
 document.addEventListener('DOMContentLoaded',()=>{setup();render()});
 window.nineAnnualV2={groups,render,print};
})();
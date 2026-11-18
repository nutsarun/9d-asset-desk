/* Nine Distribution enterprise UI extensions */
(function(){
'use strict';
const COMPANY=['Nine distribution co.,ltd','After Nine Co.,ltd'],LOGO='./nine-distribution-logo.webp';
const $=id=>document.getElementById(id);
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const clean=v=>String(v||'').replace(/\s+/g,' ').trim();
const company=v=>/after|a9/i.test(v||'')?COMPANY[1]:COMPANY[0];
const superAdmin=()=>currentUser&&['nutsarun@ninedistribution.com','satit@ninedistribution.com'].includes(currentUser.email.toLowerCase());
function select(id,label){
 const field=document.createElement('label');field.className='text-sm';field.textContent=label;
 const s=document.createElement('select');s.id=id;s.className='w-full border p-2 rounded-lg bg-white';
 COMPANY.forEach(x=>s.add(new Option(x,x)));field.append(s);return field;
}
function setup(){
 const side=document.querySelector('#sidebar .p-6.border-b .flex.items-center.gap-3');
 if(side&&!$('brand-9d')){const d=document.createElement('span');d.id='brand-9d';d.className='bg-white rounded p-1';d.innerHTML='<img src="'+LOGO+'" alt="Nine Distribution" style="height:45px;width:95px;object-fit:contain">';side.prepend(d);}
 const egrid=$('emp-form')?.querySelector('.grid');if(egrid&&!$('emp-company'))egrid.append(select('emp-company','สังกัดบริษัท'));
 const employeeHead=$('tab-employees')?.querySelector('.flex.flex-wrap.justify-between');
 if(employeeHead&&!$('nine-csv-upload')){
  const buttons=document.createElement('div');buttons.className='flex gap-2 flex-wrap';
  buttons.innerHTML='<button id="nine-export-employee" class="bg-white border rounded-lg px-3 py-2 text-sm">Export พนักงาน CSV</button><label class="bg-white border rounded-lg px-3 py-2 text-sm cursor-pointer">Import CSV<input type="file" id="nine-csv-upload" accept=".csv,text/csv" hidden></label><button id="nine-csv-sample" class="bg-white border rounded-lg px-3 py-2 text-sm">ไฟล์ตัวอย่าง CSV</button>';
  employeeHead.append(buttons);
  $('nine-export-employee').onclick=()=>csv('9D-Employees',(window.nineDirectory?.()||[]).map(e=>({EmployeeCode:e.empId,Name:e.name,Company:e.company||COMPANY[0],Department:e.department,Position:e.position,Email:e.email,Contact:e.contact})));
  $('nine-csv-sample').onclick=()=>csv('9D-Employee-Template',[{EmployeeCode:'EMP001',Name:'ชื่อ นามสกุล',Company:COMPANY[0],Department:'IT',Position:'Engineer',Email:'example@ninedistribution.com',Contact:'-'}]);
  $('nine-csv-upload').onchange=importCSV;
 }
 const assetForm=$('asset-form');
 if(assetForm&&!$('asset-company')){
  const d=document.createElement('div');d.className='grid grid-cols-2 gap-3';
  d.append(select('asset-company','สังกัดบริษัท'));
  const label=document.createElement('label');label.className='text-sm';label.textContent='Host name (SharePoint)';
  label.innerHTML+='<input id="asset-hostname" readonly class="w-full border p-2 rounded-lg bg-slate-100" placeholder="เท่ากับ Asset ID เสมอ"><span class="text-xs text-slate-500">Host name = Asset ID โดยไม่เปลี่ยนรหัสเดิม</span>';d.append(label);
  const v=document.createElement('label');v.className='text-sm';v.textContent='มูลค่าทรัพย์สิน (บาท)';v.innerHTML+='<input type="number" min="0" step="0.01" id="asset-value" class="w-full border p-2 rounded-lg" placeholder="ระบุมูลค่า เช่น 25000.00">';d.append(v);
  $('asset-edit-mode')?.insertAdjacentElement('afterend',d);
  $('asset-company').onchange=autoId;$('asset-category').onchange=autoId;
 }
 for(const id of ['asset-category','inv-filter-category','co-cat-filter']){const sel=$(id);if(sel&&!sel.querySelector('option[value="Other"]'))sel.add(new Option('อื่นๆ (Other)','Other'));}
 const inventoryHeader=$('tab-inventory')?.querySelector('h2')?.parentElement?.parentElement;
 if(inventoryHeader&&!$('nine-backup-csv')){
  const b=document.createElement('button');b.id='nine-backup-csv';b.className='bg-white border rounded-xl px-4 py-2 text-sm';b.textContent='Export CSV Backup';
  b.onclick=()=>csv('9D-Assets-Backup',assets.map(a=>({AssetID:a.id,Company:a.company||COMPANY[0],HostName:a.id,AssetValueTHB:a.assetValue??'',Category:a.category,Brand:a.brand,Model:a.model,SerialNumber:a.serial,Status:a.status,UserName:a.assignedTo?.name,ReceivedDate:a.assignedTo?.assignedDate,SharePointItemID:a._sp?.itemId,SharePointListID:a._sp?.listId})));
  inventoryHeader.append(b);
 }
 const nav=$('sidebar')?.querySelector('nav');
 if(nav&&!$('nav-annual')){
  const b=document.createElement('button');b.id='nav-annual';b.dataset.tab='annual';b.className='nav-btn w-full flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-medium hover:bg-slate-800 hover:text-white';b.innerHTML='<i data-lucide="clipboard-check" class="w-5 h-5"></i>ตรวจนับทรัพย์สินประจำปี';b.onclick=()=>switchTab('annual');nav.insertBefore(b,$('nav-sso-settings'));
 }
 if(!$('tab-annual')){
  const sec=document.createElement('section');sec.id='tab-annual';sec.className='tab-content hidden space-y-5';
  sec.innerHTML='<div class="flex flex-wrap justify-between gap-3"><div><h2 class="text-2xl font-bold">ตรวจนับทรัพย์สินประจำปี</h2><p class="text-sm text-slate-500">รายชื่อผู้ถือครองและรายการอุปกรณ์ที่ต้องตรวจ</p></div><div class="flex gap-2"><button id="annual-export" class="border rounded-xl bg-white px-3 py-2">Export CSV</button><button id="annual-print" class="bg-blue-600 text-white rounded-xl px-3 py-2">พิมพ์รายงาน A4</button></div></div><div class="grid sm:grid-cols-2 gap-3"><label>ปีที่ตรวจ<input type="number" id="annual-year" class="block w-full border rounded-lg p-2" min="2020" max="2100"></label><label>สังกัด<select id="annual-company" class="block w-full border rounded-lg p-2"><option value="ALL">ทุกบริษัท</option></select></label></div><div class="bg-amber-50 border border-amber-200 rounded-lg p-3 text-xs">รายงานเป็นข้อมูลผู้ถือครองปัจจุบัน ไม่ใช่ภาพย้อนหลังของปีที่เลือก ผลตรวจสภาพเก็บเฉพาะเบราว์เซอร์นี้จนกว่าจะมี SharePoint Audit List</div><div id="annual-items" class="space-y-4"></div>';
  $('main').append(sec);$('annual-year').value=new Date().getFullYear();COMPANY.forEach(x=>$('annual-company').add(new Option(x,x)));
  $('annual-year').onchange=renderAnnual;$('annual-company').onchange=renderAnnual;
  $('annual-export').onclick=exportAudit;$('annual-print').onclick=()=>printAudit(null);
 }
 const st=document.createElement('style');
 st.textContent='.nine-doc{color:#0f172a;font:12px Sarabun,sans-serif;line-height:1.45}.nine-doc h1{font-size:20px;font-weight:700}.nine-doc .doc-head{display:flex;align-items:center;gap:14px;justify-content:space-between;border-bottom:2px solid #1e293b;padding-bottom:8px}.nine-doc .doc-logo{height:65px;width:100px;object-fit:contain}.nine-doc .doc-section{background:#f1f5f9;padding:5px 8px;margin:14px 0 8px;border-left:4px solid #334155;font-weight:bold}.nine-doc table{border-collapse:collapse;width:100%;table-layout:fixed;font-size:10px}.nine-doc th,.nine-doc td{padding:5px;border:1px solid #94a3b8;overflow-wrap:anywhere;vertical-align:top}.nine-doc .sign-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:9px;margin-top:25px}.nine-doc .sign{border:1px solid #cbd5e1;border-radius:6px;text-align:center;min-height:92px;padding:10px 3px;font-size:10px}.nine-doc .audit-page{break-after:page;page-break-after:always}.nine-doc .audit-page:last-child{break-after:auto;page-break-after:auto}@media print{@page{size:A4 portrait;margin:12mm}#print-modal{display:block!important;position:static!important;inset:auto!important;overflow:visible!important;padding:0!important}#print-modal-wrapper{max-width:none!important;width:100%!important;margin:0!important}#print-modal-wrapper>.bg-slate-200\\/80{padding:0!important;background:white!important;overflow:visible!important}#a4-document-content{padding:0!important;min-height:0!important;width:100%!important;box-shadow:none!important}.nine-doc tr,.nine-doc .sign-grid{break-inside:avoid;page-break-inside:avoid}.no-print{display:none!important}}';
 document.head.append(st);window.lucide?.createIcons();
}
function autoId(){
 if($('asset-edit-mode')?.value!=='add')return;
 $('asset-id').value=window.nineNextAssetId?.($('asset-company').value,$('asset-category').value)||'';
 if($('asset-hostname'))$('asset-hostname').value=$('asset-id').value;
}
const previousOpen=openAssetModal;
openAssetModal=function(id=null){
 const ret=previousOpen(id);const a=id?assets.find(x=>x.id===id):null;
 $('asset-company').value=a?.company||COMPANY[0];$('asset-hostname').value=a?.id||'';
 if($('asset-value'))$('asset-value').value=a?.assetValue??'';
 $('asset-id').readOnly=true;if(!a)autoId();return ret;
};
function csvValue(value){
 let s=String(value??'');if(/^[=+\-@\t\r]/.test(s))s="'"+s;
 return '"'+s.replace(/"/g,'""')+'"';
}
function csv(file,rows){
 if(!rows.length)return showToast('ไม่มีข้อมูลให้ Export','error');
 const keys=Object.keys(rows[0]);
 const data='\ufeff'+[keys.map(csvValue).join(','),...rows.map(row=>keys.map(k=>csvValue(row[k])).join(','))].join('\r\n');
 const url=URL.createObjectURL(new Blob([data],{type:'text/csv;charset=utf-8'}));
 const a=document.createElement('a');a.href=url;a.download=file+'-'+new Date().toISOString().slice(0,10)+'.csv';a.click();setTimeout(()=>URL.revokeObjectURL(url),1500);
}
function parseCSV(text){
 text=text.replace(/^\ufeff/,'');const rows=[];let row=[],value='',inside=false;
 for(let i=0;i<text.length;i++){
  const c=text[i];if(inside){if(c==='"'&&text[i+1]==='"'){value+='"';i++;}else if(c==='"')inside=false;else value+=c;}
  else if(c==='"')inside=true;
  else if(c===','){row.push(value);value='';}
  else if(c==='\r'||c==='\n'){if(c==='\r'&&text[i+1]==='\n')i++;row.push(value);if(row.some(x=>x.trim()))rows.push(row);row=[];value='';}
  else value+=c;
 }
 if(inside)throw Error('รูปแบบ CSV ไม่ถูกต้อง');row.push(value);if(row.some(x=>x.trim()))rows.push(row);return rows;
}
async function importCSV(event){
 const file=event.target.files?.[0];if(!file)return;
 try{
  if(!superAdmin())throw Error('เฉพาะ Super Admin');
  if(file.size>2000000)throw Error('ไฟล์ใหญ่เกิน 2 MB');
  const rows=parseCSV(await file.text());
  if(rows.length<2||rows.length>1001)throw Error('CSV ต้องมีข้อมูล 1–1000 คน');
  const headers=rows.shift().map(x=>x.trim().toLowerCase().replace(/\s+/g,''));
  const get=(r,...names)=>{for(const name of names){const i=headers.indexOf(name.toLowerCase().replace(/\s+/g,''));if(i>=0)return (r[i]||'').trim();}return'';};
  const parsed=rows.map(r=>({empId:get(r,'EmployeeCode','EmployeeID','รหัสพนักงาน'),name:get(r,'Name','Title','ชื่อ-นามสกุล'),company:get(r,'Company','สังกัด')||COMPANY[0],department:get(r,'Department','แผนก'),position:get(r,'Position','ตำแหน่ง'),email:get(r,'Email','อีเมล'),contact:get(r,'Contact','Phone','เบอร์โทร')}));
  if(parsed.some(e=>!e.name||!COMPANY.includes(e.company)))throw Error('ตรวจสอบ Name และ Company ให้ตรงตามไฟล์ตัวอย่าง');
  if(new Set(parsed.map(e=>e.name.toLocaleLowerCase())).size!==parsed.length)throw Error('ชื่อพนักงานซ้ำใน CSV');
  const shared=window.nineEmployeeHasSharedList?.();
  if(!confirm('นำเข้าพนักงาน '+parsed.length+' คน ('+(shared?'SharePoint':'เฉพาะเบราว์เซอร์')+') ใช่หรือไม่?'))return;
  const result=await window.nineEmployeeBulkUpsert(parsed);
  showToast('นำเข้า '+result.done.length+' คน / ผิดพลาด '+result.errors.length+(result.errors.length?' : '+result.errors.slice(0,2).join('; '):''),result.errors.length?'error':'success');
 }catch(err){showToast(err.message,'error');}finally{event.target.value='';}
}
generateDocNumber=function(prefix){
 const date=new Date(),day=date.getFullYear()+String(date.getMonth()+1).padStart(2,'0')+String(date.getDate()).padStart(2,'0');
 const key='9d-doc-counter-'+prefix+'-'+day;
 const exists=transactions.filter(x=>String(x.docNo).startsWith(prefix+'-'+day+'-')).map(x=>Number(x.docNo.split('-').at(-1))||0);
 const n=Math.max(+(localStorage.getItem(key)||0),0,...exists)+1;localStorage.setItem(key,String(n));
 return prefix+'-'+day+'-'+String(n).padStart(3,'0');
};
function heading(title,subtitle,number,date,org){
 return '<div class="doc-head"><img class="doc-logo" src="'+LOGO+'"><div style="flex:1"><b>'+esc(org||COMPANY[0])+'</b><h1>'+esc(title)+'</h1><small>'+esc(subtitle)+'</small></div><div style="font-size:10px;text-align:right">เลขที่: <b>'+esc(number)+'</b><br>วันที่: '+esc(date)+'</div></div>';
}
function displayPrint(body){
 $('a4-document-content').innerHTML='<div class="nine-doc">'+body+'</div>';
 const modal=$('print-modal');modal.classList.remove('hidden');modal.classList.add('block');
}
function signatures(name,officer){
 return '<div class="sign-grid"><div class="sign"><b>ผู้ถือครองทรัพย์สิน</b><p style="margin-top:22px">...................................</p>'+esc(name||'-')+'<p>วันที่ ....../....../......</p></div><div class="sign"><b>ผู้ตรวจรับ / ผู้ส่งมอบ</b><p style="margin-top:22px">...................................</p>'+esc(officer||'-')+'<p>วันที่ ....../....../......</p></div><div class="sign"><b>ผู้อนุมัติ</b><p style="margin-top:22px">...................................</p>...................................<p>วันที่ ....../....../......</p></div></div>';
}
openPrintDocument=function(no){
 const tx=transactions.find(x=>x.docNo===no);if(!tx)return;
 const returned=tx.type==='CHECKIN',e=tx.employee||{},org=e.company||COMPANY[0];
 const items=tx.items.map((a,i)=>'<tr><td>'+Number(i+1)+'</td><td>'+esc(a.id)+'</td><td>'+esc(a.category)+'</td><td>'+esc(a.brand)+' '+esc(a.model)+'</td><td>'+esc(a.serial)+'</td><td>'+esc(a.condition||'-')+'</td></tr>').join('');
 displayPrint(heading(returned?'ใบตรวจรับคืนทรัพย์สินและอุปกรณ์ไอที':'ใบส่งมอบทรัพย์สินและอุปกรณ์ไอที',returned?'IT ASSET RETURN FORM':'IT ASSET HANDOVER FORM',tx.docNo,tx.date,org)+
 '<div class="doc-section">ข้อมูลพนักงาน</div><div>ชื่อ-นามสกุล: <b>'+esc(e.name)+'</b>　รหัส: '+esc(e.empId)+'</div><div>สังกัด: '+esc(org)+'　แผนก: '+esc(e.department)+'</div>'+
 '<div class="doc-section">รายการทรัพย์สิน '+tx.items.length+' รายการ</div><table><colgroup><col style="width:6%"><col style="width:17%"><col style="width:12%"><col style="width:27%"><col style="width:21%"><col style="width:17%"></colgroup><thead><tr><th>#</th><th>Asset ID</th><th>ประเภท</th><th>ยี่ห้อ / รุ่น</th><th>Serial</th><th>หมายเหตุ</th></tr></thead><tbody>'+items+'</tbody></table><div class="doc-section">หมายเหตุ</div><p>'+esc(tx.note||'-')+'</p>'+signatures(e.name,tx.officerName));
};
function auditYear(){return Number($('annual-year')?.value)||new Date().getFullYear();}
function auditKey(a){return '9d-annual-'+auditYear()+'-'+a._sp?.listId+'-'+a._sp?.itemId;}
function audit(a){try{return JSON.parse(localStorage.getItem(auditKey(a))||'{}')}catch{return{}}}
window.nineAuditChange=function(index,field,value){
 const a=assets[index];if(!a)return;const item=audit(a);item[field]=value;item.checkedAt=new Date().toISOString();item.checkedBy=currentUser?.email;localStorage.setItem(auditKey(a),JSON.stringify(item));
};
function groups(){
 const filter=$('annual-company')?.value||'ALL',map=new Map();
 for(const a of assets){
  if(!a.assignedTo?.name||(filter!=='ALL'&&company(a.company)!==filter))continue;
  const name=clean(a.assignedTo.name),key=name.toLocaleLowerCase();
  if(!map.has(key))map.set(key,{name,company:company(a.company),employee:a.assignedTo,items:[]});
  map.get(key).items.push(a);
 }
 return [...map.values()].sort((a,b)=>a.name.localeCompare(b.name,'th'));
}
function renderAnnual(){
 const root=$('annual-items');if(!root)return;
 const list=groups();if(!list.length){root.innerHTML='<p class="p-6 bg-white border rounded-xl">ไม่มีผู้ถือครองทรัพย์สิน</p>';return;}
 root.innerHTML=list.map((g,gi)=>'<div class="bg-white border rounded-xl p-4 space-y-2"><div class="flex justify-between"><div><b>'+esc(g.name)+'</b> ('+g.items.length+' รายการ)<p class="text-xs text-slate-500">'+esc(g.company)+'</p></div><button class="border rounded-lg px-3 py-1" onclick="nineAuditPrint('+gi+')">พิมพ์ A4</button></div><div class="overflow-x-auto"><table class="w-full text-sm"><thead><tr class="bg-slate-50"><th class="text-left p-2">Asset / Serial</th><th class="text-left p-2">อุปกรณ์</th><th class="p-2">สภาพ</th><th class="p-2">ครบถ้วน</th><th class="p-2">หมายเหตุ</th></tr></thead><tbody>'+g.items.map(a=>{const i=assets.indexOf(a),v=audit(a);return '<tr class="border-t"><td class="p-2 text-xs">'+esc(a.id)+'<br>'+esc(a.serial)+'</td><td class="p-2">'+esc(a.brand)+' '+esc(a.model)+'</td><td><select class="border rounded p-1" onchange="nineAuditChange('+i+',\'condition\',this.value)"><option>รอตรวจ</option><option value="ปกติ" '+(v.condition==='ปกติ'?'selected':'')+'>ปกติ</option><option value="ผิดปกติ" '+(v.condition==='ผิดปกติ'?'selected':'')+'>ผิดปกติ</option></select></td><td><select class="border rounded p-1" onchange="nineAuditChange('+i+',\'complete\',this.value)"><option>รอตรวจ</option><option value="ครบ" '+(v.complete==='ครบ'?'selected':'')+'>ครบ</option><option value="ไม่ครบ" '+(v.complete==='ไม่ครบ'?'selected':'')+'>ไม่ครบ</option></select></td><td><input class="border rounded p-1 w-40" value="'+esc(v.note||'')+'" onchange="nineAuditChange('+i+',\'note\',this.value)"></td></tr>'}).join('')+'</tbody></table></div></div>').join('');
}
function auditPrintBody(g){
 const y=auditYear();
 const rows=g.items.map((a,i)=>{const x=audit(a);return '<tr><td>'+(i+1)+'</td><td>'+esc(a.id)+'</td><td>'+esc(a.brand)+' '+esc(a.model)+'</td><td>'+esc(a.serial)+'</td><td>□ ปกติ　□ ผิดปกติ<br>'+esc(x.condition||'')+'</td><td>□ ครบ　□ ไม่ครบ<br>'+esc(x.complete||'')+'</td></tr>'}).join('');
 return '<div class="audit-page">'+heading('ใบตรวจนับทรัพย์สินประจำปี '+y,'ANNUAL ASSET VERIFICATION','AUD-'+y+'-'+esc(g.name).slice(0,6),new Date().toLocaleDateString('th-TH'),g.company)+
 '<div class="doc-section">ผู้ถือครอง: '+esc(g.name)+'</div><div>แผนก: '+esc(g.employee.department||'-')+'</div><div class="doc-section">รายการทรัพย์สิน</div><table><colgroup><col style="width:6%"><col style="width:18%"><col style="width:25%"><col style="width:21%"><col style="width:15%"><col style="width:15%"></colgroup><thead><tr><th>#</th><th>Asset ID</th><th>อุปกรณ์</th><th>Serial</th><th>สภาพ</th><th>ครบถ้วน</th></tr></thead><tbody>'+rows+'</tbody></table><div class="doc-section">ผลการตรวจ</div>□ อยู่ครบ　□ ไม่ครบ　□ ปกติ　□ ผิดปกติ <p>หมายเหตุ: .......................................................................................................</p>'+signatures(g.name,'ผู้ตรวจนับ')+'</div>';
}
function printAudit(index){const g=groups(),selected=index===null?g:(g[index]?[g[index]]:[]);if(!selected.length)return showToast('ไม่มีข้อมูล','error');displayPrint(selected.map(auditPrintBody).join(''));}
window.nineAuditPrint=index=>printAudit(index);
function exportAudit(){
 const rows=groups().flatMap(g=>g.items.map(a=>{const v=audit(a);return {Year:auditYear(),Company:g.company,Employee:g.name,AssetID:a.id,HostName:a.id,AssetValueTHB:a.assetValue??'',Category:a.category,Brand:a.brand,Model:a.model,Serial:a.serial,Condition:v.condition||'',Complete:v.complete||'',Note:v.note||'',CheckedAt:v.checkedAt||'',CheckedBy:v.checkedBy||''}}));
 csv('Annual-Asset-Audit-'+auditYear(),rows);
}
window.nineUpgradeAfterLogin=function(){renderAnnual()};
document.addEventListener('DOMContentLoaded',setup);
})();
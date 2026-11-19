/* Master Asset Inventory backup, preview and guarded recovery.
 * Manual SharePoint document-library snapshots. No writes to legacy Lists.
 * Browser is not a scheduler; automatic nightly backup needs server-side automation.
 */
(function(){
'use strict';
const SITE='ninedistribution.sharepoint.com,13d8bc75-67a0-4dbf-b579-8b26c2dd2e9f,dedddf0b-d21b-4660-84ad-259b6ffb0ab4';
const MASTER='509b2348-ae5e-410e-b8af-eff3595e752b';
const PREFIX='https://graph.microsoft.com/v1.0/sites/'+SITE;
const LIST=PREFIX+'/lists/'+MASTER;
const FOLDER='9D Asset Backups';
const FOLDER_ESC=encodeURIComponent(FOLDER);
const FIELDS=['Title','AssetID','HostName','Vendor','Model','ItemType','Assetorserialnumber','Status','UserName0','Company','ReceivedDate','ReturnedDate','Datepurchased','Warranty','SourceListId','SourceItemId','SourceKind','OriginalTitle','OldUser','Remark','AttachmentsFlag','SourceLink','AssetValue'];
const SYS_FIELDS=['SourceListId','SourceItemId','SourceKind','OriginalTitle','OldUser','AttachmentsFlag','SourceLink'];
const WRITE_FIELDS=FIELDS.filter(x=>!SYS_FIELDS.includes(x));
const $=id=>document.getElementById(id);
const safe=x=>String(x??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const admin=()=>['nutsarun@ninedistribution.com','satit@ninedistribution.com'].includes(String(currentUser?.email||'').toLowerCase());
const status=(msg,error=false)=>{const el=$('db-backup-status');if(el){el.textContent=msg;el.className='rounded-lg p-3 text-sm border '+(error?'bg-rose-50 border-rose-200 text-rose-800':'bg-slate-50 border-slate-200 text-slate-700')}};
let loading=false,prepared=null,selectedFile=null,working=false;
async function graph(url,method='GET',body=undefined,scope='Sites.Read.All',headersOverride={}){
 if(!window.nineGraphAuth)throw Error('กรุณาเข้าสู่ระบบ Microsoft ก่อน');
 const auth=window.nineGraphAuth;
 const token=await auth.client.acquireTokenSilent({account:auth.account,scopes:[scope]})
  .catch(()=>auth.client.acquireTokenPopup({account:auth.account,scopes:[scope]}));
 const headers={Authorization:'Bearer '+token.accessToken,...headersOverride};
 if(body!==undefined && typeof body!=='string')headers['Content-Type']='application/json';
 const res=await fetch(url,{method,headers,body:body===undefined?undefined:typeof body==='string'?body:JSON.stringify(body)});
 if(!res.ok)throw Error('Microsoft Graph HTTP '+res.status+': '+(await res.text()).slice(0,290));
 return res.status===204?null:res.json();
}
async function collect(url,scope='Sites.Read.All'){
 let out=[];
 while(url){const result=await graph(url,'GET',undefined,scope);out.push(...(result.value||[]));url=result['@odata.nextLink']}
 return out;
}
async function masterRows(){
 const rows=await collect(LIST+'/items?$top=100&$expand=fields');
 if(!rows.length)throw Error('Master Inventory ว่างหรืออ่านข้อมูลไม่ได้');
 return rows;
}
function normalize(row){
 const fields={};for(const key of FIELDS)if(Object.prototype.hasOwnProperty.call(row.fields||{},key))fields[key]=row.fields[key]??null;
 return {sharePointItemId:String(row.id),fields};
}
async function digest(text){
 const buf=await crypto.subtle.digest('SHA-256',new TextEncoder().encode(text));
 return Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,'0')).join('');
}
async function snapshot(){
 const rows=await masterRows();
 const result={schema:'9d.asset.master.backup.v1',siteId:SITE,listId:MASTER,listName:'9D Master Asset Inventory',
  createdAt:new Date().toISOString(),createdBy:String(currentUser?.email||''),rowCount:rows.length,fields:FIELDS,records:rows.map(normalize)};
 const content=JSON.stringify(result);
 result.sha256=await digest(content);
 return result;
}
function fileName(created){
 const time=created.replace(/[-:]/g,'').replace('T','_').replace('.','_');
 return '9D_Master_Assets_'+time+'.json';
}
function downloadJson(data,name){
 const blob=new Blob([JSON.stringify(data,null,2)],{type:'application/json;charset=utf-8'});
 const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();
 setTimeout(()=>URL.revokeObjectURL(url),10000);
}
async function folderExists(create=false){
 const url=PREFIX+'/drive/root:/'+FOLDER_ESC;
 try{return await graph(url,'GET')}catch(err){if(!/HTTP 404/.test(err.message)||!create)throw err}
 return graph(PREFIX+'/drive/root/children','POST',{name:FOLDER,folder:{},'@microsoft.graph.conflictBehavior':'fail'},'Sites.ReadWrite.All');
}
async function uploadJson(data){
 await folderExists(true);
 const name=fileName(data.createdAt);
 const path=PREFIX+'/drive/root:/'+FOLDER_ESC+'/'+encodeURIComponent(name)+':/content';
 // Never silently overwrite an existing backup file.
 try{await graph(PREFIX+'/drive/root:/'+FOLDER_ESC+'/'+encodeURIComponent(name));
   throw Error('มี Snapshot ชื่อนี้อยู่แล้ว ไม่อนุญาตให้เขียนทับ');
 }catch(err){if(!/HTTP 404/.test(err.message))throw err}
 const payload=JSON.stringify(data,null,2);
 // Only small snapshots supported with this simple PUT; large snapshots need upload sessions.
 if(new Blob([payload]).size>3.8*1024*1024)throw Error('ไฟล์เกิน 3.8MB กรุณาดาวน์โหลดสำรองหรือใช้ Upload Session');
 await graph(path,'PUT',payload,'Sites.ReadWrite.All',{'Content-Type':'application/json'});
 return name;
}
function lock(flag){working=flag;for(const id of ['db-save','db-download','db-history','db-restore-apply','db-preview']){const e=$(id);if(e)e.disabled=flag}}
async function performBackup(remote){
 if(working||!admin())return;
 lock(true);prepared=null;
 try{
  status('กำลังอ่าน Master List และตรวจสอบรายการ...');
  const data=await snapshot(),name=fileName(data.createdAt);
  if(remote){status('กำลังสร้างไฟล์ Snapshot บน SharePoint...');await uploadJson(data);status('สำรองข้อมูลบน SharePoint สำเร็จ '+data.rowCount+' รายการ · '+name);await history()}
  else{downloadJson(data,name);status('ดาวน์โหลด Snapshot '+data.rowCount+' รายการแล้ว · '+name)}
  showToast('สำรองฐานข้อมูล Master Asset Inventory สำเร็จ','success');
 }catch(err){status('สำรองข้อมูลไม่สำเร็จ: '+err.message,true);showToast('Backup: '+err.message,'error')}
 finally{lock(false)}
}
async function history(){
 if(!admin())return;
 const el=$('db-backup-history');if(!el)return;
 try{
  const folder=await folderExists(false);
  const rows=await collect(PREFIX+'/drive/items/'+folder.id+'/children?$top=100');
  const files=rows.filter(x=>x.name?.startsWith('9D_Master_Assets_')&&x.name.endsWith('.json'))
   .sort((a,b)=>String(b.createdDateTime).localeCompare(String(a.createdDateTime)));
  el.innerHTML=files.length?'<div class="overflow-x-auto"><table class="w-full text-sm"><thead><tr><th class="text-left p-2">ไฟล์สำรอง</th><th class="text-left p-2">สร้างเมื่อ</th><th class="text-right p-2">ขนาด</th><th class="p-2">ไฟล์</th></tr></thead><tbody>'+
   files.map(f=>'<tr class="border-t"><td class="p-2 font-mono text-xs">'+safe(f.name)+'</td><td class="p-2">'+safe(new Date(f.createdDateTime).toLocaleString('th-TH'))+'</td><td class="p-2 text-right">'+Number(f.size||0).toLocaleString()+' bytes</td><td class="p-2"><a class="text-blue-700 underline" target="_blank" rel="noopener noreferrer" href="'+safe(f.webUrl)+'">เปิด</a></td></tr>').join('')+'</tbody></table></div>':
   '<p class="text-sm text-slate-500">ยังไม่มีไฟล์สำรองใน SharePoint</p>';
 }catch(err){
  if(/HTTP 404/.test(err.message))el.textContent='ยังไม่พบโฟลเดอร์สำรองใน SharePoint — กด Backup Now เพื่อสร้าง';
  else el.textContent='ไม่สามารถอ่านประวัติสำรอง: '+err.message;
 }
}
async function validateFile(file){
 if(!file||file.size>5*1024*1024)throw Error('ไฟล์ต้องเป็น JSON ขนาดไม่เกิน 5MB');
 const data=JSON.parse(await file.text());
 if(data.schema!=='9d.asset.master.backup.v1'||data.siteId!==SITE||data.listId!==MASTER)
  throw Error('ไม่ใช่ไฟล์ Snapshot สำหรับ 9D Master Asset Inventory');
 if(!Array.isArray(data.records)||!data.records.length||data.records.length>20000||data.rowCount!==data.records.length)
  throw Error('จำนวนรายการในไฟล์ไม่ถูกต้อง');
 if(!data.sha256)throw Error('ไม่พบ checksum สำหรับตรวจสอบไฟล์');
 const clone={...data};delete clone.sha256;
 // SHA256 created from the compact original JSON before the checksum was appended.
 if(await digest(JSON.stringify(clone))!==data.sha256)throw Error('Checksum ไม่ตรง — ไฟล์อาจถูกแก้ไขหรือเสียหาย');
 const ids=new Set(),assets=new Set();
 for(const r of data.records){
  if(!r.fields||!r.sharePointItemId||!r.fields.AssetID)throw Error('ไฟล์มีแถวที่ไม่มี Asset ID');
  const id=String(r.sharePointItemId),asset=String(r.fields.AssetID);
  if(ids.has(id)||assets.has(asset))throw Error('พบ SharePoint ID หรือ Asset ID ซ้ำในไฟล์');
  ids.add(id);assets.add(asset);
 }
 return data;
}
async function preview(){
 if(working||!admin())return;
 prepared=null;const file=$('db-restore-file')?.files?.[0];
 if(!file)return status('กรุณาเลือกไฟล์ JSON ที่สำรองไว้',true);
 lock(true);
 try{
  const backup=await validateFile(file),rows=await masterRows();
  const live=new Map(rows.map(r=>[String(r.id),r]));
  let updates=0,same=0,missing=0,conflicts=0;const changes=[],samples=[];
  for(const original of backup.records){
   const current=live.get(String(original.sharePointItemId));
   if(!current){missing++;continue}
   if(String(current.fields?.AssetID||'')!==String(original.fields.AssetID)){conflicts++;continue}
   const fields={};
   for(const k of WRITE_FIELDS){
    if(!Object.prototype.hasOwnProperty.call(original.fields,k))continue;
    const a=original.fields[k],b=current.fields?.[k]??null;
    if(JSON.stringify(a)!==JSON.stringify(b))fields[k]=a;
   }
   if(Object.keys(fields).length){updates++;changes.push({id:String(current.id),asset:String(current.fields.AssetID),fields,currentEtag:current['@odata.etag']||null});if(samples.length<8)samples.push(String(current.fields.AssetID)+': '+Object.keys(fields).join(', '))}
   else same++;
  }
  if(conflicts)throw Error('พบ '+conflicts+' รายการที่ Asset ID เปลี่ยนจากในไฟล์ หยุดเพื่อความปลอดภัย');
  prepared={backup,changes,updateCount:updates,liveCount:rows.length,missing,conflicts,createdAt:Date.now()};
  $('db-preview-results').innerHTML='<div class="rounded-lg border bg-blue-50 p-3 text-sm">'+
   '<b>ตัวอย่างการกู้คืน (ยังไม่เขียนข้อมูล)</b><br>Snapshot: '+safe(backup.createdAt)+' · '+backup.rowCount+' รายการ'+
   '<br>ปัจจุบัน: '+rows.length+' รายการ · จะอัปเดต '+updates+' · ไม่เปลี่ยน '+same+' · ไม่มีใน Master '+missing+
   '<br><b>ความปลอดภัย:</b> ไม่ลบรายการใหม่ ไม่สร้างรายการที่หายไป ไม่แก้คอลัมน์ประวัติ Source และไม่กู้คืนไฟล์แนบ</div>'+
   (samples.length?'<div class="text-xs p-2 bg-slate-50 rounded-lg">'+samples.map(safe).join('<br>')+'</div>':'');
  $('db-restore-apply').disabled=!updates||missing>0;
  $('db-restore-apply').classList.toggle('hidden',!updates||missing>0);
  if(missing)status('มี '+missing+' รายการไม่พบใน Master — ไม่อนุญาตกู้คืน โปรดตรวจสอบข้อมูลก่อน',true);
  else status('ตรวจสอบ Snapshot ผ่านแล้ว สามารถเลือกยืนยันกู้คืนรายการที่ต่างกันได้');
 }catch(err){status('ตรวจสอบไฟล์ไม่ได้: '+err.message,true);$('db-preview-results').textContent=err.message;$('db-restore-apply').classList.add('hidden')}
 finally{lock(false)}
}
async function restore(){
 if(working||!admin()||!prepared?.changes.length)return;
 if(prepared.missing||prepared.conflicts)return status('ไม่สามารถกู้คืนเมื่อมีรายการที่ไม่ตรงกัน',true);
 if(Date.now()-prepared.createdAt>5*60*1000)return status('ผลตรวจสอบเกิน 5 นาที กรุณากดตรวจสอบใหม่',true);
 const phrase=prompt('การกู้คืนจะแทนที่ข้อมูลปัจจุบันของ '+prepared.changes.length+' รายการ\nโปรดพิมพ์ RESTORE เพื่อยืนยัน');
 if(phrase!=='RESTORE')return;
 if(!confirm('ยืนยันอีกครั้ง: เปลี่ยนค่าทะเบียนทรัพย์สิน '+prepared.changes.length+' รายการจาก Backup? การดำเนินการนี้ไม่สามารถย้อนกลับโดยอัตโนมัติ'))return;
 lock(true);let done=0,errors=[];
 try{
  for(const record of prepared.changes){
   try{
    const current=await graph(LIST+'/items/'+record.id+'?$expand=fields');
    if(String(current.fields?.AssetID||'')!==record.asset||current['@odata.etag']!==record.currentEtag)throw Error('มีการแก้ไขรายการนี้ระหว่างตรวจสอบ');
    const headers=current['@odata.etag']?{'If-Match':current['@odata.etag']}:{};
    await graph(LIST+'/items/'+record.id+'/fields','PATCH',record.fields,'Sites.ReadWrite.All',headers);
    done++;
   }catch(err){errors.push(record.asset+': '+err.message)}
   status('กู้คืน '+done+'/'+prepared.changes.length+' · ผิดพลาด '+errors.length);
  }
  if(typeof window.nineRefreshDashboard==='function')await window.nineRefreshDashboard();
  status('กู้คืนสำเร็จ '+done+' รายการ'+(errors.length?' · ไม่สำเร็จ '+errors.length+' รายการ ('+errors.slice(0,3).join('; ')+')':''));
  showToast(errors.length?'กู้คืนบางรายการไม่สำเร็จ':'กู้คืนข้อมูลเรียบร้อยแล้ว',errors.length?'error':'success');
 }catch(err){status('เกิดข้อผิดพลาด: '+err.message,true)}
 finally{prepared=null;lock(false);$('db-restore-apply')?.classList.add('hidden')}
}
function mount(){
 if($('db-backup-root')||!admin())return;
 const root=$('master-backup-mount');if(!root)return;
 root.innerHTML='<div id="db-backup-root" class="space-y-4">'+
  '<div class="rounded-xl border bg-white p-4 space-y-3"><h3 class="text-lg font-semibold">Backup Now — สำรองฐานข้อมูลหลัก</h3>'+
  '<p class="text-sm text-slate-600">สำรองคอลัมน์ทะเบียนทรัพย์สิน รวมสถานะ ผู้ถือครอง และ AssetValue โดยไม่ดึงจาก Lists เก่า ข้อมูลจะไม่ถูกบันทึกลง GitHub</p>'+
  '<div class="flex flex-wrap gap-2"><button type="button" id="db-save" class="bg-blue-600 text-white rounded-lg px-4 py-2">สำรองลง SharePoint</button>'+
  '<button type="button" id="db-download" class="border rounded-lg px-4 py-2">ดาวน์โหลด JSON</button>'+
  '<button type="button" id="db-history" class="border rounded-lg px-4 py-2">รีเฟรชประวัติ Backup</button></div>'+
  '<p class="text-xs text-amber-800 bg-amber-50 rounded-lg p-3">การสำรองนี้ทำงานเมื่อกดปุ่มเท่านั้น ไม่ได้สำรองอัตโนมัติทุกวัน หากต้องการ Daily Backup โดยไม่มีใครเปิดเว็บ ต้องตั้ง Power Automate หรือระบบฝั่งเซิร์ฟเวอร์เพิ่ม</p></div>'+
  '<div class="rounded-xl border bg-white p-4 space-y-3"><h3 class="text-lg font-semibold">Backup History — ประวัติไฟล์สำรองใน SharePoint</h3><div id="db-backup-history" class="text-sm text-slate-500">กดรีเฟรชประวัติ Backup</div></div>'+
  '<div class="rounded-xl border bg-white p-4 space-y-3"><h3 class="text-lg font-semibold">Restore Preview — ตรวจสอบก่อนกู้คืน</h3>'+
  '<p class="text-sm text-slate-600">เลือก Snapshot JSON ที่สร้างจากเมนูนี้ ตรวจสอบรายการที่จะเปลี่ยนก่อนกู้คืน ระบบไม่ลบข้อมูลใหม่และไม่กู้คืนไฟล์แนบ</p>'+
  '<input id="db-restore-file" type="file" accept=".json,application/json" class="w-full border rounded-lg p-2 text-sm">'+
  '<button type="button" id="db-preview" class="border rounded-lg px-4 py-2">ตรวจสอบไฟล์ก่อนกู้คืน</button>'+
  '<div id="db-preview-results" class="space-y-2"></div>'+
  '<button type="button" id="db-restore-apply" class="hidden bg-rose-700 text-white rounded-lg px-4 py-2">ยืนยันกู้คืนข้อมูลที่แตกต่าง</button></div>'+
  '<div id="db-backup-status" role="status" class="rounded-lg p-3 text-sm border bg-slate-50">พร้อมสำรองข้อมูลจาก 9D Master Asset Inventory</div>'+
  '</div>';
 $('db-save').onclick=()=>performBackup(true);
 $('db-download').onclick=()=>performBackup(false);
 $('db-history').onclick=history;
 $('db-preview').onclick=preview;
 $('db-restore-apply').onclick=restore;
 history();
}
window.nineDatabaseBackup={mount,history,preview};
document.addEventListener('DOMContentLoaded',mount);
})();
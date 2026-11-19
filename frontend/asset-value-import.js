/* Private, admin-initiated value import. CSV stays in the user's browser; no financial values bundled in this public repo. */
(() => {
 'use strict';
 const SITE='ninedistribution.sharepoint.com,13d8bc75-67a0-4dbf-b579-8b26c2dd2e9f,dedddf0b-d21b-4660-84ad-259b6ffb0ab4';
 const LIST='509b2348-ae5e-410e-b8af-eff3595e752b';
 const ROOT='https://graph.microsoft.com/v1.0/sites/'+SITE+'/lists/'+LIST;
 const $=id=>document.getElementById(id);
 const admin=()=>['nutsarun@ninedistribution.com','satit@ninedistribution.com'].includes(String(currentUser?.email||'').toLowerCase());
 const esc=x=>String(x??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
 const cash=x=>Number(x).toLocaleString('th-TH',{minimumFractionDigits:2,maximumFractionDigits:2});
 let pending=[],working=false,lastChecked=null;
 async function token(scopes){
  const auth=window.nineGraphAuth;if(!auth)throw Error('กรุณาเข้าสู่ระบบ Microsoft ก่อน');
  const request={account:auth.account,scopes};
  try{return (await auth.client.acquireTokenSilent(request)).accessToken}
  catch(e){return (await auth.client.acquireTokenPopup(request)).accessToken}
 }
 async function graph(url,method='GET',body=null,scope='Sites.Read.All',etag=null){
  const headers={Authorization:'Bearer '+await token([scope])};
  if(body!==null)headers['Content-Type']='application/json';
  if(etag)headers['If-Match']=etag;
  const result=await fetch(url,{method,headers,body:body===null?undefined:JSON.stringify(body)});
  if(!result.ok)throw Error('Microsoft Graph '+result.status+': '+(await result.text()).slice(0,210));
  return result.status===204?null:result.json();
 }
 async function all(url,scope='Sites.Read.All'){
  const out=[];while(url){const data=await graph(url,'GET',null,scope);out.push(...(data.value||[]));url=data['@odata.nextLink']}return out;
 }
 // RFC-style CSV parser: preserves quoted commas, quotes and newlines.
 function parseCSV(text){
  text=text.replace(/^\uFEFF/,'');
  const records=[];let line=[],cell='',quote=false;
  for(let i=0;i<text.length;i++){
   const c=text[i];
   if(quote){
    if(c==='"'&&text[i+1]==='"'){cell+='"';i++}
    else if(c==='"')quote=false;
    else cell+=c;
   }else if(c==='"'&&cell==='')quote=true;
   else if(c===','){line.push(cell);cell=''}
   else if(c==='\r'||c==='\n'){
    if(c==='\r'&&text[i+1]==='\n')i++;
    line.push(cell);if(line.some(s=>s.trim()))records.push(line);
    line=[];cell='';
   }else cell+=c;
  }
  if(quote)throw Error('รูปแบบ CSV ไม่ถูกต้อง: เครื่องหมายคำพูดไม่ครบ');
  line.push(cell);if(line.some(s=>s.trim()))records.push(line);
  if(records.length<2)throw Error('CSV ไม่พบข้อมูล');
  const headers=records.shift().map(x=>x.trim());
  if(new Set(headers).size!==headers.length)throw Error('พบชื่อคอลัมน์ซ้ำ');
  return records.map((row,index)=>Object.fromEntries(headers.map((h,i)=>[h,row[i]??'']))).map((r,i)=>({...r,__line:i+2}));
 }
 function parseAmount(value){
  const v=String(value).trim().replace(/[฿,\s]/g,'');
  if(!/^\d+(?:\.\d{1,2})?$/.test(v))throw Error('มูลค่าไม่ถูกต้อง: '+value);
  const n=Number(v);if(!Number.isFinite(n)||n<0||n>1000000000)throw Error('มูลค่าผิดช่วงที่ยอมรับ');
  return Math.round(n*100)/100;
 }
 function mount(){
  if($('nine-value-import')||!admin())return;
  const card=$('nine-valuation-card');if(!card)return;
  const panel=document.createElement('section');panel.id='nine-value-import';
  panel.className='border rounded-xl bg-white p-4 my-4 space-y-3';
  panel.innerHTML='<div class="flex flex-wrap items-center justify-between gap-2"><div><b>นำเข้ามูลค่าทรัพย์สินจาก CSV</b><p class="text-xs text-slate-500">เฉพาะ Super Admin · บันทึกลง 9D Master Asset Inventory เท่านั้น</p></div><button id="nine-values-toggle" type="button" class="bg-white border px-3 py-2 rounded-lg text-sm">เลือกไฟล์ CSV</button></div>'+
    '<div id="nine-import-body" class="hidden space-y-3">'+
    '<p class="text-xs bg-amber-50 border border-amber-200 rounded-lg p-3">ใช้ไฟล์นำเข้ามูลค่าที่มีคอลัมน์ SharePointItemID, AssetID และ AssetValueTHB เท่านั้น ระบบจะตรวจสอบรหัส Asset ID ก่อนบันทึก ช่อง Value ที่ว่างจะไม่ถูกนำเข้า ไม่แก้ข้อมูลอื่น</p>'+
    '<input id="nine-value-csv" type="file" accept=".csv,text/csv" class="w-full text-sm border rounded-lg p-2">'+
    '<div id="nine-import-preview" class="text-sm text-slate-600"></div>'+
    '<button id="nine-import-apply" type="button" disabled class="hidden bg-blue-600 text-white rounded-lg px-4 py-2 disabled:opacity-40">ยืนยันบันทึกมูลค่าลง SharePoint</button></div>';
  card.insertAdjacentElement('afterend',panel);
  $('nine-values-toggle').onclick=()=>$('nine-import-body').classList.toggle('hidden');
  $('nine-value-csv').onchange=preview;
  $('nine-import-apply').onclick=apply;
 }
 async function preview(){
  pending=[];lastChecked=null;
  const target=$('nine-import-preview'),applyBtn=$('nine-import-apply');
  applyBtn.disabled=true;applyBtn.classList.add('hidden');
  const file=$('nine-value-csv')?.files?.[0];if(!file)return;
  if(file.size>1000000){target.textContent='ไฟล์ใหญ่เกิน 1MB';return}
  if(NINE_GRAPH.lists.length!==1||NINE_GRAPH.lists[0].kind!=='master'){
   target.textContent='เว็บไซต์ยังไม่ได้เลือกใช้ 9D Master Asset Inventory เป็น List หลัก กรุณาโหลดหน้าใหม่';return;
  }
  try{
   target.textContent='กำลังตรวจสอบรายการกับ SharePoint...';
   const records=parseCSV(await file.text());
   for(const h of ['SharePointItemID','AssetID','AssetValueTHB']){
    if(!Object.prototype.hasOwnProperty.call(records[0],h))throw Error('ไม่มีคอลัมน์ '+h);
   }
   const allItems=await all(ROOT+'/items?$top=100&$expand=fields');
   const byId=new Map(allItems.map(x=>[String(x.id),x]));
   const unique=new Set();const toApply=[];
   for(const record of records){
    const id=String(record.SharePointItemID||'').trim();
    const asset=String(record.AssetID||'').trim();
    const original=String(record.AssetValueTHB||'').trim();
    if(!original)continue;
    if(!/^\d+$/.test(id)||!asset)throw Error('แถว '+record.__line+': ไม่มี SharePointItemID หรือ AssetID');
    if(unique.has(id))throw Error('แถว '+record.__line+': SharePointItemID ซ้ำ');
    unique.add(id);
    const live=byId.get(id);
    if(!live)throw Error('ไม่พบ SharePoint Item ID '+id);
    if(String(live.fields?.AssetID||'').trim()!==asset||String(live.fields?.Title||'').trim()!==asset)
     throw Error('ข้อมูลไม่ตรง: '+asset+' กับ SharePoint row '+id+' — ยกเลิกทั้งชุด');
    const amount=parseAmount(original);
    const old=live.fields.AssetValue;
    const prior=(old===undefined||old===null||old==='')?null:Number(old);
    toApply.push({id,asset,value:amount,prior,etag:live['@odata.etag']||null});
   }
   if(!toApply.length)throw Error('ไม่พบรายการที่ระบุมูลค่า');
   if(toApply.length>500)throw Error('มากกว่า 500 รายการ กรุณาแบ่งไฟล์');
   const changed=toApply.filter(x=>x.prior===null||Math.abs(x.prior-x.value)>0.001);
   const overwrite=toApply.filter(x=>x.prior!==null&&Math.abs(x.prior-x.value)>0.001);
   const identical=toApply.length-changed.length;
   pending=changed;lastChecked=Date.now();
   const total=toApply.reduce((a,b)=>a+b.value,0);
   target.innerHTML='<div class="p-3 bg-blue-50 border rounded-lg">ตรวจสอบแล้ว <b>'+toApply.length+' รายการ</b> | มูลค่าในไฟล์ <b>฿'+cash(total)+'</b><br>ต้องบันทึก '+changed.length+' รายการ · ค่าเดิมเท่ากัน '+identical+' · มีค่าที่จะถูกแทนที่ '+overwrite.length+'</div>'+
    '<div class="mt-2 max-h-40 overflow-auto text-xs">'+toApply.map(x=>'<div>'+esc(x.asset)+' — ฿'+cash(x.value)+(x.prior!==null?' (เดิม ฿'+cash(x.prior)+')':'')+'</div>').join('')+'</div>';
   applyBtn.classList.remove('hidden');applyBtn.disabled=changed.length===0;
  }catch(err){target.textContent='ไม่สามารถนำเข้า: '+err.message}
 }
 async function ensureColumn(){
  const cols=await all(ROOT+'/columns?$top=100');
  if(cols.some(x=>x.name==='AssetValue'))return;
  await graph(ROOT+'/columns','POST',{name:'AssetValue',number:{decimalPlaces:'two',displayAs:'number'}},'Sites.Manage.All');
 }
 async function apply(){
  if(working||!admin()||!pending.length||!lastChecked)return;
  if(Date.now()-lastChecked>5*60*1000){$('nine-import-preview').textContent='ผลตรวจสอบเกิน 5 นาที กรุณาเลือกไฟล์ใหม่';return}
  const overwrites=pending.filter(x=>x.prior!==null).length;
  if(!confirm('ต้องการบันทึกมูลค่า '+pending.length+' รายการใน SharePoint หรือไม่?'+(overwrites?'\\nมี '+overwrites+' รายการที่จะเปลี่ยนมูลค่าเดิม':'')+'\\nการบันทึกนี้ไม่แก้หมายเลขทรัพย์สินหรือคอลัมน์อื่น'))return;
  working=true;
  const button=$('nine-import-apply');button.disabled=true;
  let done=0,skipped=0;const errors=[];
  try{
   await ensureColumn();
   for(const row of pending){
    try{
     const live=await graph(ROOT+'/items/'+row.id+'?$expand=fields');
     if(String(live.fields?.AssetID||'')!==row.asset||String(live.fields?.Title||'')!==row.asset)throw Error('Asset ID เปลี่ยนระหว่างการทำงาน');
     const current=live.fields.AssetValue===undefined||live.fields.AssetValue===null||live.fields.AssetValue===''?null:Number(live.fields.AssetValue);
     if(current!==row.prior){skipped++;errors.push(row.asset+': มีการเปลี่ยนมูลค่าจากผู้ใช้อื่น');continue}
     await graph(ROOT+'/items/'+row.id+'/fields','PATCH',{AssetValue:row.value},'Sites.ReadWrite.All',live['@odata.etag']);
     const verify=await graph(ROOT+'/items/'+row.id+'?$expand=fields');
     if(Math.abs(Number(verify.fields?.AssetValue)-row.value)>0.001)throw Error('อ่านกลับแล้วไม่ตรง');
     const a=assets.find(x=>String(x._sp?.itemId)===row.id&&x.id===row.asset);if(a)a.assetValue=row.value;
     done++;
    }catch(err){errors.push(row.asset+': '+err.message)}
    $('nine-import-preview').textContent='บันทึก '+done+' / '+pending.length+' รายการ'+(errors.length?' · ผิดพลาด '+errors.length:'');
   }
   renderDashboard();
   mount();
   const msg='บันทึกมูลค่า SharePoint สำเร็จ '+done+' รายการ'+(errors.length?' · ไม่สำเร็จ '+errors.length+' รายการ: '+errors.slice(0,2).join('; '):'');
   showToast(msg,errors.length?'error':'success');
   if($('nine-import-preview'))$('nine-import-preview').textContent=msg+' — รีเฟรชหน้าเพื่ออ่านข้อมูลใหม่';
   if($('nine-import-apply'))$('nine-import-apply').disabled=true;
   pending=[];lastChecked=null;
  }catch(err){showToast('ไม่สามารถตั้งค่าคอลัมน์ AssetValue: '+err.message,'error')}
  finally{working=false}
 }
 const old=renderDashboard;
 renderDashboard=function(){old();mount()};
 document.addEventListener('DOMContentLoaded',mount);
})();

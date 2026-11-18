/* Nine Distribution Master Asset Migration (admin-initiated via delegated Microsoft Graph) */
(function(){
"use strict";
const SITE="ninedistribution.sharepoint.com,13d8bc75-67a0-4dbf-b579-8b26c2dd2e9f,dedddf0b-d21b-4660-84ad-259b6ffb0ab4";
const BASE="https://graph.microsoft.com/v1.0/sites/"+SITE;
const MASTER_NAME="9D Master Asset Inventory";
const MARKER="9D_MASTER_V1_VERIFIED";
const SOURCES=[
 {id:"ac67c228-aaa2-44f2-8f8e-72d61dd020f0",kind:"computer",name:"Computer"},
 {id:"7a610ea3-7308-4870-a273-9770713f3594",kind:"accessory",name:"Accessories"}
];
const COLS=["AssetID","HostName","Vendor","Model","ItemType","Assetorserialnumber","Status","UserName0","Company","ReceivedDate","ReturnedDate","Datepurchased","Warranty","SourceListId","SourceItemId","SourceKind","OriginalTitle","OldUser","Remark","AttachmentsFlag","SourceLink","AssetValue"];
let running=false,masterList=null,activeList=null;
const role=()=>["nutsarun@ninedistribution.com","satit@ninedistribution.com"].includes(String(currentUser?.email||"").toLowerCase());
const norm=v=>String(v??"").replace(/\s+/g," ").trim();
const key=(list,item)=>list+"|"+item;
function status(label){
 const e=document.getElementById("master-status");
 if(e)e.textContent=label;
}
async function acquire(scopes){
 const login=window.nineGraphAuth;
 if(!login)throw Error("กรุณาเข้าสู่ระบบ Microsoft อีกครั้ง");
 try{return (await login.client.acquireTokenSilent({account:login.account,scopes})).accessToken;}
 catch(e){return (await login.client.acquireTokenPopup({account:login.account,scopes})).accessToken;}
}
async function request(url,method="GET",body=null,scopes=["Sites.Read.All"]){
 const token=await acquire(scopes);
 const headers={Authorization:"Bearer "+token};
 if(body!==null)headers["Content-Type"]="application/json";
 const result=await fetch(url,{method,headers,body:body!==null?JSON.stringify(body):undefined});
 if(!result.ok){
  const response=await result.text();
  throw Error(method+" "+url.split("?")[0]+" HTTP "+result.status+": "+response.slice(0,320));
 }
 if(result.status===204)return null;
 return result.json();
}
async function pages(url,scopes=["Sites.Read.All"]){
 const items=[];
 while(url){const p=await request(url,"GET",null,scopes);items.push(...p.value||[]);url=p["@odata.nextLink"];}
 return items;
}
async function findMaster(){
 const lists=await pages(BASE+"/lists?$top=100");
 return lists.find(x=>x.displayName===MASTER_NAME && x.list?.template!=="documentLibrary")||null;
}
async function sourceRecords(){
 const items=[];
 for(const src of SOURCES){
  const rows=await pages(BASE+"/lists/"+src.id+"/items?$top=100&$expand=fields");
  items.push(...rows.map(row=>({source:src,row})));
 }
 return items;
}
function projection(s){
 const f=s.row.fields||{},isComputer=s.source.kind==="computer";
 const host=norm(f.HostName);
 const title=norm(f.Title);
 const code=isComputer?title:(host||"LEGACY-ACC-"+s.row.id);
 const company=(/after|a9/i.test(String(isComputer?f.Company:f.Owner)))?"After Nine Co.,ltd":"Nine distribution co.,ltd";
 const owner=norm(isComputer?f._x0e4a_UserName:f.UserName0);
 const serial=norm(isComputer?f.Assetorserialnumber:f.AssetorserialNumber);
 const original=String(f.Remark||"").slice(0,245);
 const out={
  Title:code,AssetID:code,HostName:code,Vendor:String(f.Vendor||"").slice(0,250),
  Model:String(f.Model||"").slice(0,250),ItemType:String(f.ItemType||"").slice(0,250),
  Assetorserialnumber:serial,Status:String(f.Status||""),
  UserName0:owner,Company:company,
  ReceivedDate:String(f.ReceivedDate||""),ReturnedDate:String(f.ReturnedDate||""),
  Datepurchased:String(f.Datepurchased||""),Warranty:String(f.Warranty||"").slice(0,245),
  SourceListId:s.source.id,SourceItemId:String(s.row.id),SourceKind:s.source.kind,
  OriginalTitle:title,OldUser:String(f.OldUser||"").slice(0,245),
  Remark:original,AttachmentsFlag:f.Attachments?"Yes":"No",
  SourceLink:String(s.row.web_url||"").slice(0,250),AssetValue:""
 };
 for(const k of Object.keys(out)){out[k]=String(out[k]).slice(0,255);}
 return out;
}
function matched(a,b){
 const keys=["Title","AssetID","Vendor","Model","ItemType","Assetorserialnumber","Status","UserName0","Company","SourceListId","SourceItemId","SourceKind","OriginalTitle","AttachmentsFlag"];
 return keys.every(k=>String(a[k]||"")===String(b[k]||""));
}
async function inspectMaster(source,master){
 const masterRows=await pages(BASE+"/lists/"+master.id+"/items?$top=100&$expand=fields");
 const index=new Map(),duplicate=[];
 for(const row of masterRows){
  const f=row.fields||{},ref=key(f.SourceListId,f.SourceItemId);
  if(index.has(ref))duplicate.push(ref);
  index.set(ref,row);
 }
 const missing=[],different=[];
 for(const record of source){
  const shape=projection(record),ref=key(shape.SourceListId,shape.SourceItemId);
  const target=index.get(ref);
  if(!target)missing.push(ref);
  else if(!matched(shape,target.fields||{}))different.push(ref);
 }
 const extras=masterRows.filter(row=>!source.some(s=>key(s.source.id,s.row.id)===key(row.fields?.SourceListId,row.fields?.SourceItemId)));
 return {masterRows,index,missing,different,duplicate,extras,ready:missing.length===0&&different.length===0&&duplicate.length===0&&extras.length===0&&masterRows.length===source.length};
}
async function makeMaster(){
 const exists=await findMaster();
 if(exists)return exists;
 const made=await request(BASE+"/lists","POST",{displayName:MASTER_NAME,description:"Migration in progress. Read-only until verified.",list:{template:"genericList"}},["Sites.Manage.All"]);
 return made;
}
async function ensureColumns(master){
 const uri=BASE+"/lists/"+master.id+"/columns";
 const list=await pages(uri+"?$top=100",["Sites.Manage.All"]);
 const seen=new Set(list.map(x=>x.name));
 for(const column of COLS){
  if(seen.has(column))continue;
  status("กำลังสร้างคอลัมน์ "+column);
  await request(uri,"POST",column==="AssetValue"?{name:column,number:{decimalPlaces:"two"}}:{name:column,text:{}},["Sites.Manage.All"]);
 }
}
async function copyAndVerify(){
 if(running||!role())return;
 if(!window.nineGraphAuth)return status("กรุณาเข้าสู่ระบบ Microsoft ใหม่");
 if(!confirm("เริ่มสร้าง SharePoint List ใหม่และคัดลอกข้อมูลจาก Computer/Accessories โดยไม่ลบต้นฉบับใช่หรือไม่?\n\nจำเป็นต้องอนุมัติ Sites.Manage.All และ Sites.ReadWrite.All ใน Microsoft Entra ก่อน"))return;
 running=true;
 const btn=document.getElementById("master-run");if(btn)btn.disabled=true;
 try {
  status("อ่านข้อมูลต้นฉบับ...");
  const source=await sourceRecords();
  if(source.length!==70){
   throw Error("ปัจจุบันต้นฉบับมี "+source.length+" รายการ (เคยตรวจพบ 70) โปรดตรวจสอบการเปลี่ยนแปลงก่อนย้าย");
  }
  const sourceRefs=new Set(source.map(s=>key(s.source.id,s.row.id)));
  if(sourceRefs.size!==source.length)throw Error("รายการต้นฉบับซ้ำ");
  status("สร้าง Master List และคอลัมน์...");
  let master=await makeMaster();
  await ensureColumns(master);
  masterList=master;
  const existing=await pages(BASE+"/lists/"+master.id+"/items?$top=100&$expand=fields");
  const found=new Map(existing.map(x=>[key(x.fields?.SourceListId,x.fields?.SourceItemId),x]));
  let created=0;
  for(const sourceItem of source){
   const fields=projection(sourceItem),ref=key(fields.SourceListId,fields.SourceItemId);
   if(found.has(ref))continue;
   status("คัดลอกข้อมูล "+(created+1)+" รายการ (จาก 70)...");
   await request(BASE+"/lists/"+master.id+"/items","POST",{fields},["Sites.ReadWrite.All"]);
   created++;
  }
  status("กำลังตรวจสอบการย้ายข้อมูล...");
  const result=await inspectMaster(source,master);
  if(!result.ready){
   const msg="ยังไม่ผ่านการตรวจ: Missing "+result.missing.length+", Mismatch "+result.different.length+", Duplicates "+result.duplicate.length+", Extra "+result.extras.length;
   status(msg+" — เก็บสอง Lists เดิมและยังไม่สลับเว็บไซต์");
   return;
  }
  const attach=source.filter(x=>x.row.fields?.Attachments).length;
  status("คัดลอกครบ "+result.masterRows.length+" รายการ ตรวจคีย์และฟิลด์สำเร็จ พบไฟล์แนบเดิม "+attach+" รายการ (ยังอยู่ต้นฉบับ)");
  await request(BASE+"/lists/"+master.id,"PATCH",{description:MARKER+"; source entries="+source.length+"; original attachments remain in legacy lists."},["Sites.Manage.All"]);
  // Switch active list in the current browser. Other users auto-detect the verified marker on their next login.
  NINE_GRAPH.lists=[{id:master.id,kind:"master",category:"Other"}];
  activeList=master;
  const token=await acquire(["Sites.Read.All"]);
  assets=await readGraphList(token,NINE_GRAPH.lists[0]);
  checkoutCart=[];renderAll();
  window.nineUpgradeAfterLogin?.();
  status("สำเร็จ: เปิดใช้ 9D Master Asset Inventory "+assets.length+" รายการแล้ว (ไฟล์แนบเดิมยังเก็บใน Lists ต้นฉบับ)");
  showToast("ตรวจสอบครบ "+assets.length+" รายการ และเปลี่ยนไปใช้ Master List แล้ว","success");
 }catch(e){status("ยังไม่เปลี่ยนฐานข้อมูล: "+e.message);showToast(e.message,"error");}
 finally{running=false;if(btn)btn.disabled=false;}
}
async function chooseListForLogin(){
 try{
  const found=await findMaster();
  if(!found||!String(found.description||"").startsWith(MARKER))return null;
  // Revalidate unique source IDs and row count before using the newly built list.
  const rows=await pages(BASE+"/lists/"+found.id+"/items?$top=100&$expand=fields");
  const migrated=rows.filter(x=>x.fields?.SourceListId&&x.fields?.SourceItemId);
  const refs=new Set(migrated.map(x=>key(x.fields.SourceListId,x.fields.SourceItemId)));
  if(migrated.length<70||refs.size!==migrated.length)return null; // Allow subsequent newly created assets without legacy source keys.
  NINE_GRAPH.lists=[{id:found.id,kind:"master",category:"Other"}];
  masterList=found;activeList=found;
  return found;
 }catch(err){console.warn("Master discovery failed, keep legacy",err);return null;}
}
function buildMenu(){
 if(document.getElementById("master-page"))return;
 const nav=document.querySelector("#sidebar nav");
 const button=document.createElement("button");
 button.id="nav-master";button.dataset.tab="master";
 button.className="nav-btn w-full flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-medium hover:bg-slate-800 hover:text-white";
 button.innerHTML='<i data-lucide="database" class="w-5 h-5"></i><span>รวมข้อมูล SharePoint</span>';
 button.onclick=()=>switchTab("master");
 nav?.append(button);
 const sec=document.createElement("section");sec.id="tab-master";sec.className="tab-content hidden space-y-4";
 sec.innerHTML='<h2 class="text-2xl font-bold">9D Master Asset Inventory</h2><p class="text-sm text-slate-600">คัดลอก Computer และ Accessories ไปยัง Master List ใหม่ โดยเก็บต้นฉบับไว้ 100% และตรวจสอบข้อมูลก่อนเปิดใช้</p><div class="bg-white p-5 border rounded-xl space-y-2"><p class="font-semibold">รายการต้นฉบับ: Computer 32 + Accessories 38 = 70</p><p class="text-sm">ตรวจสอบการย้ายด้วย SourceListId + SourceItemId, Serial, Owner, Company, Model และ Status</p><p class="text-sm text-amber-700">ไฟล์แนบเดิม 9 รายการยังคงอยู่ใน List ต้นฉบับ สคริปต์นี้ไม่คัดลอกไฟล์แนบและไม่ลบต้นฉบับ</p><p id="master-status" class="text-sm bg-slate-50 rounded-lg p-3">พร้อมตรวจสอบสิทธิ์และเริ่มการย้ายข้อมูล</p><button id="master-run" class="bg-blue-600 text-white rounded-xl px-5 py-3 font-semibold">สร้าง Master List + ย้ายข้อมูล 70 รายการ</button></div>';
 document.querySelector("main")?.append(sec);
 document.getElementById("master-run").onclick=copyAndVerify;
 document.getElementById("nav-master").classList.toggle("hidden",!role());
 window.lucide?.createIcons();
}
window.nineMaster={chooseListForLogin,copyAndVerify,buildMenu,getActive:()=>activeList};
document.addEventListener("DOMContentLoaded",buildMenu);
})();
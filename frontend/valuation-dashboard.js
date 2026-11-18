/* Financial dashboard: SharePoint-backed values, HostName/AssetID alignment */
(function(){
"use strict";
const SITE="ninedistribution.sharepoint.com,13d8bc75-67a0-4dbf-b579-8b26c2dd2e9f,dedddf0b-d21b-4660-84ad-259b6ffb0ab4";
const LIST="509b2348-ae5e-410e-b8af-eff3595e752b";
const ROOT="https://graph.microsoft.com/v1.0/sites/"+SITE+"/lists/"+LIST;
const admin=()=>["nutsarun@ninedistribution.com","satit@ninedistribution.com"].includes(String(currentUser?.email||"").toLowerCase());
const escaped=v=>String(v??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const valid=a=>a.assetValue!==null && a.assetValue!==undefined && Number.isFinite(Number(a.assetValue)) && Number(a.assetValue)>=0;
const money=x=>Number(x).toLocaleString("th-TH",{minimumFractionDigits:2,maximumFractionDigits:2});
let busy=false,expanded=false;
async function token(scope){
 const a=window.nineGraphAuth;if(!a)throw Error("กรุณาเข้าสู่ระบบ Microsoft ใหม่");
 const request={account:a.account,scopes:[scope]};
 try{return (await a.client.acquireTokenSilent(request)).accessToken;}
 catch(e){return (await a.client.acquireTokenPopup(request)).accessToken;}
}
async function call(url,method="GET",body=null,scope="Sites.Read.All",etag=null){
 const headers={Authorization:"Bearer "+await token(scope)};
 if(body!==null)headers["Content-Type"]="application/json";
 if(etag)headers["If-Match"]=etag;
 const res=await fetch(url,{method,headers,body:body===null?undefined:JSON.stringify(body)});
 if(!res.ok)throw Error("Microsoft Graph HTTP "+res.status+": "+(await res.text()).slice(0,230));
 return res.status===204?null:await res.json();
}
async function all(url,scope="Sites.Read.All"){
 const found=[];while(url){const result=await call(url,"GET",null,scope);found.push(...result.value||[]);url=result["@odata.nextLink"];}return found;
}
function valueRows(){
 const map=new Map();
 for(const a of assets){
  const type=String(a.assetType||a.category||"Other").trim()||"Other";
  if(!map.has(type))map.set(type,{name:type,total:0,known:0,unknown:0,assets:[]});
  const g=map.get(type);g.assets.push(a);
  if(valid(a)){g.total+=Number(a.assetValue);g.known++}else g.unknown++;
 }
 return [...map.values()].sort((a,b)=>b.total-a.total||a.name.localeCompare(b.name,"th"));
}
function addUI(){
 if(document.getElementById("nine-valuation-card"))return;
 const container=document.createElement("div");container.id="nine-valuation-card";
 container.className="bg-white border border-slate-200 rounded-2xl p-5 mb-5";
 const dash=document.getElementById("tab-dashboard");
 const kpis=document.getElementById("dashboard-kpis");
 if(kpis?.parentElement)kpis.insertAdjacentElement("afterend",container);
 else dash?.prepend(container);
}
function draw(){
 const card=document.getElementById("nine-valuation-card");if(!card)return;
 const data=valueRows(),total=data.reduce((n,g)=>n+g.total,0);
 const unknown=data.reduce((n,g)=>n+g.unknown,0),known=data.reduce((n,g)=>n+g.known,0);
 const master=NINE_GRAPH.lists.length===1 && NINE_GRAPH.lists[0].kind==="master";
 card.innerHTML='<div class="flex flex-wrap items-center justify-between gap-3"><div><div class="text-sm font-semibold text-slate-500">มูลค่าทรัพย์สินรวม (เฉพาะรายการที่มีข้อมูล)</div><div class="text-3xl font-bold text-slate-900">฿'+money(total)+'</div><div class="text-sm text-slate-500 mt-1">ระบุมูลค่าแล้ว '+known+' / '+assets.length+' รายการ'+(unknown?' · ยังไม่ระบุ '+unknown+' รายการ':'')+'</div></div><div class="flex gap-2 flex-wrap"><button id="nine-breakdown-btn" class="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-xl text-sm">ดูมูลค่าแยกประเภท</button>'+(master&&admin()?'<button id="nine-prepare-value" class="border bg-white px-4 py-2 rounded-xl text-sm">เตรียมคอลัมน์มูลค่า / ซิงก์ Host name</button>':'')+'</div></div><div id="nine-breakdown" class="mt-4 '+(expanded?'':'hidden')+'">'+(unknown?'<div class="text-xs bg-amber-50 border border-amber-200 rounded-lg p-3 mb-3">ผลรวมไม่ใช่มูลค่าทรัพย์สินทั้งองค์กรจนกว่าจะกรอกมูลค่าครบทุกชิ้น (ข้อมูลว่างไม่นับเป็น 0 บาท)</div>':'')+'<div class="overflow-x-auto"><table class="w-full text-sm"><thead><tr class="bg-slate-50"><th class="p-2 text-left">หมวดหมู่</th><th class="p-2 text-right">จำนวน</th><th class="p-2 text-right">มีมูลค่า</th><th class="p-2 text-right">มูลค่ารวม (บาท)</th><th></th></tr></thead><tbody>'+data.map((g,i)=>'<tr class="border-t"><td class="p-2">'+escaped(g.name)+'</td><td class="p-2 text-right">'+g.assets.length+'</td><td class="p-2 text-right">'+g.known+'</td><td class="p-2 text-right font-semibold">'+money(g.total)+'</td><td class="p-2"><button data-group="'+i+'" class="nine-finance-detail text-blue-700 underline">ดูอุปกรณ์</button></td></tr>').join('')+'</tbody></table></div><div id="nine-finance-items" class="mt-3"></div></div>';
 document.getElementById("nine-breakdown-btn").onclick=()=>{expanded=!expanded;draw()};
 document.getElementById("nine-prepare-value")?.addEventListener("click",prepare);
 card.querySelectorAll(".nine-finance-detail").forEach(b=>b.onclick=()=>{
  const g=data[Number(b.dataset.group)],out=document.getElementById("nine-finance-items");
  out.innerHTML='<div class="font-semibold mb-2">'+escaped(g.name)+' ('+g.assets.length+' ชิ้น)</div><div class="overflow-x-auto"><table class="text-sm w-full"><thead><tr class="bg-slate-50"><th class="p-2 text-left">Asset ID / Host name</th><th class="p-2 text-left">ยี่ห้อ/รุ่น</th><th class="p-2 text-left">สถานะ</th><th class="p-2 text-right">มูลค่า</th></tr></thead><tbody>'+g.assets.map(a=>'<tr class="border-t"><td class="p-2 font-mono">'+escaped(a.id)+'</td><td class="p-2">'+escaped(a.brand)+' '+escaped(a.model)+'</td><td class="p-2">'+escaped(a.status)+'</td><td class="p-2 text-right">'+(valid(a)?'฿'+money(a.assetValue):'<span class="text-amber-700">ยังไม่ระบุ</span>')+'</td></tr>').join('')+'</tbody></table></div>';
 });
}
async function prepare(){
 if(busy||!admin())return;
 if(!confirm("จะสร้างคอลัมน์ตัวเลข AssetValue (ถ้ายังไม่มี) และปรับ HostName ให้เท่ากับ AssetID สำหรับทรัพย์สินเดิมใน Master List โดยไม่เปลี่ยน Asset ID หรือข้อมูลต้นฉบับทั้งสอง Lists ยืนยันหรือไม่?"))return;
 busy=true;
 const btn=document.getElementById("nine-prepare-value");if(btn)btn.disabled=true;
 let changed=0;const errors=[];
 try{
  const columns=await all(ROOT+"/columns?$top=100","Sites.Manage.All");
  if(!columns.some(x=>x.name==="AssetValue")){
   await call(ROOT+"/columns","POST",{name:"AssetValue",number:{decimalPlaces:"two",displayAs:"number"}},"Sites.Manage.All");
  }
  const rows=await all(ROOT+"/items?$top=100&$expand=fields");
  const expected=assets.length;
  if(rows.length!==expected)throw Error("ข้อมูล Master List ไม่ตรงกับหน้าเว็บ ("+rows.length+" กับ "+expected+") กรุณาโหลดหน้าเว็บใหม่");
  for(const row of rows){
   const id=String(row.fields?.AssetID||"").trim();
   if(!id){errors.push("SharePoint row "+row.id+" ไม่มี AssetID");continue;}
   if(row.fields.HostName===id&&row.fields.Title===id)continue;
   try{
    const current=await call(ROOT+"/items/"+row.id+"?$expand=fields");
    if(!current.fields?.AssetID||String(current.fields.AssetID)!==id)throw Error("Asset ID เปลี่ยนระหว่างการทำงาน");
    await call(ROOT+"/items/"+row.id+"/fields","PATCH",{HostName:id,Title:id},"Sites.ReadWrite.All",current["@odata.etag"]);
    changed++;
   }catch(err){errors.push(id+": "+err.message);}
  }
  const reread=await all(ROOT+"/items?$top=100&$expand=fields");
  const mismatched=reread.filter(x=>!x.fields.AssetID||x.fields.Title!==x.fields.AssetID||x.fields.HostName!==x.fields.AssetID);
  const msg="ซิงก์ HostName เสร็จ "+changed+" รายการ; คงเหลือไม่ตรง "+mismatched.length+" รายการ";
  if(errors.length)throw Error(msg+". ปัญหา: "+errors.slice(0,2).join("; "));
  if(mismatched.length)throw Error(msg);
  showToast("ตั้งค่า AssetValue และ Host name = Asset ID สำเร็จแล้ว","success");
  if(typeof refreshAssetData==="function")await refreshAssetData();
 }catch(err){showToast("การตั้งค่า Master List: "+err.message,"error");}
 finally{busy=false;draw();}
}
const original=renderDashboard;
renderDashboard=function(){original();addUI();draw();};
document.addEventListener("DOMContentLoaded",addUI);
})();
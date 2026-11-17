const statuses = {
  available: ["พร้อมใช้งาน", "#1c9d83", "#e4f6f0", "check"],
  assigned: ["กำลังใช้งาน", "#7956e7", "#eee8ff", "user"],
  loaned: ["ให้ยืม", "#286bb1", "#e4f1ff", "transfer"],
  unavailable: ["ไม่พร้อมใช้งาน", "#9c6c2b", "#fff1db", "tool"],
  lost: ["สูญหาย", "#b94b35", "#fff0ec", "info"],
  unknown: ["ยังไม่ได้ระบุ", "#657085", "#edf0f6", "info"],
  repair: ["ส่งซ่อม", "#e49c44", "#fff1db", "tool"],
  retired: ["ปลดระวาง", "#8490aa", "#edf0f6", "archive"]
};
const sources={computer:"Computer",accessories:"Accessories"};
const inventoryCompanies={
  computer:[["9Distribution","Nine Distribution"],["After9","After Nine"]],
  accessories:[["Ninedistribution","Nine Distribution"],["After Nine","After Nine"]]
};
const inventoryCategories={
  computer:["PC","Notebook","All in one","Phone"],
  accessories:["Mouse","Keyboard","Docking Station","Cable","Monitor","USB-C Adapter","USB-A Adapter","Printer","wifi dongle"]
};
const inCustody = asset => ["assigned","loaned"].includes(asset.status);
const categories = ["โน้ตบุ๊ก","คอมพิวเตอร์","จอภาพ","อุปกรณ์เครือข่าย","เครื่องพิมพ์","เครื่องมือทดสอบ","อุปกรณ์เสริม","อื่น ๆ"];
const paths = {
  grid:"M3 3h7v7H3zM14 3h7v7h-7zM3 14h7v7H3zM14 14h7v7h-7z",
  box:"M3 7l9-4 9 4v10l-9 4-9-4V7zm0 0l9 4 9-4m-9 4v10M7.5 5l9 4",
  user:"M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2m12-13a4 4 0 1 1-8 0 4 4 0 0 1 8 0z",
  doc:"M5 3h10l4 4v14H5zM14 3v5h5M8 12h8m-8 4h8",
  check:"M20 6 9 17l-5-5",tool:"M14 7l3 3 4-4a6 6 0 0 1-8 8l-7 7-3-3 7-7a6 6 0 0 1 8-8z",
  transfer:"M3 8h17m-5-5 5 5-5 5M21 16H4m5-5-5 5 5 5",
  archive:"M4 8v12h16V8M3 4h18v4H3zM9 12h6",
  plus:"M12 5v14M5 12h14", refresh:"M20 8a8 8 0 1 0 0 8m0-12v5h-5",info:"M12 11v5m0-8v.1M22 12a10 10 0 1 1-20 0 10 10 0 0 1 20 0z",arrow:"M5 12h14m-6-6 6 6-6 6",close:"m6 6 12 12M6 18 18 6"
};
const icon = name => `<svg aria-hidden="true" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.65" stroke-linejoin="round" stroke-linecap="round"><path d="${paths[name]}"/></svg>`;
const esc = value => String(value ?? "").replace(/[&<>"']/g, char => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[char]));
const today = () => new Intl.DateTimeFormat("en-CA",{timeZone:"Asia/Bangkok",year:"numeric",month:"2-digit",day:"2-digit"}).format(new Date());
const dateLabel = value => value ? new Intl.DateTimeFormat("th-TH",{dateStyle:"medium",timeZone:"Asia/Bangkok"}).format(new Date(`${value}T12:00:00+07:00`)) : "—";
const timeLabel = value => new Intl.DateTimeFormat("th-TH",{dateStyle:"medium",timeStyle:"short",timeZone:"Asia/Bangkok"}).format(new Date(value));
const button = (label,action,options={}) => `<button type="button" class="ad-button ${options.className ?? ""}" data-do="${action}"${options.id ? ` data-id="${esc(options.id)}"`:""}>${options.icon ? icon(options.icon):""}${esc(label)}</button>`;
const late = asset => inCustody(asset) && asset.assignment?.due && asset.assignment.due<today();
const statusTag = asset => {const [label,color,bg]=statuses[asset.status];return `<span class="ad-tag" style="--tag-bg:${bg};--tag-ink:${color}"><i class="ad-dot" style="background:${color}"></i>${label}</span>${late(asset)?'<span class="ad-tag ad-late">เกินกำหนด</span>':""}`;};
const field = (name,label,value="",opts={}) => `<label class="ad-label ${opts.wide?"ad-wide":""}">${esc(label)}${opts.required?" *":""}${opts.choices?`<select class="ad-control" name="${name}" ${opts.required?"required":""}>${opts.choices.map(([key,title])=>`<option value="${esc(key)}" ${key===value?"selected":""}>${esc(title)}</option>`).join("")}</select>`:opts.area?`<textarea class="ad-control" maxlength="${opts.maxLength??500}" name="${name}">${esc(value)}</textarea>`:`<input class="ad-control" name="${name}" type="${opts.type??"text"}" value="${esc(value)}" ${opts.type==="date"?"":'maxlength="120"'} ${opts.required?"required":""}>`}</label>`;
function assetSnapshot(asset){const {code,brand,model,category,serial,site}=asset;return {code,brand,model,category,serial,site};}
function printMarkup(doc){
  const {asset,person}=doc;
  const rows=[["ชื่อ–นามสกุล",person.name],...(doc.previousHolder?[["ผู้ใช้งานเดิม",doc.previousHolder.name]]:[]),["รหัสพนักงาน",person.employeeCode],["แผนก",person.department],["ตำแหน่ง",person.role],["วันที่รับ",dateLabel(doc.date)],["กำหนดคืน",dateLabel(doc.due)],["รหัสทรัพย์สิน",asset.code],["ประเภทอุปกรณ์",asset.category],["ยี่ห้อ / รุ่น",[asset.brand,asset.model].filter(Boolean).join(" / ")],["หมายเลขเครื่อง (S/N)",asset.serial],["สถานที่",asset.site],["อุปกรณ์เสริม",doc.accessories],["สภาพและหมายเหตุ",doc.condition]];
  return `<!doctype html><html lang="th"><meta charset="utf-8"><title>${esc(doc.id)} | ใบรับมอบอุปกรณ์</title><style>@page{size:A4;margin:15mm}*{box-sizing:border-box}body{font:14px/1.5 "Noto Sans Thai",Tahoma,sans-serif;color:#1b1830;max-width:185mm;margin:20px auto;padding:0 9px}header{border-bottom:2px solid #29223d;margin-bottom:14px;padding:6px 0 13px}h1{font-size:24px;margin:4px 0}p{margin:6px 0}table{border-collapse:collapse;width:100%;table-layout:fixed}th,td{padding:6px 10px;border:1px solid #bdb9c6;text-align:left;overflow-wrap:anywhere}th{width:34%;background:#f2f0f8}footer{display:flex;gap:30px;margin-top:27px;break-inside:avoid}footer div{flex:1;text-align:center}@media print{body{margin:0;padding:0}}</style><body><header><b>บริษัท ไนน์ ดิสทริบิวชั่น จำกัด</b><h1>ใบรับมอบ / ยืมทรัพย์สินสำหรับพนักงาน</h1><p>เลขที่เอกสาร ${esc(doc.id)}</p></header><table><tbody>${rows.map(([label,value])=>`<tr><th scope="row">${esc(label)}</th><td>${esc(value||"—")}</td></tr>`).join("")}</tbody></table><p style="margin-top:15px">ข้าพเจ้าได้รับและตรวจสอบอุปกรณ์ตามรายการข้างต้นแล้ว จะเก็บรักษาให้อยู่ในสภาพเหมาะสม และส่งคืนตามกำหนดที่ระบุไว้ (ถ้ามี)</p><p><strong>*ข้าพเจ้าทำการตรวจสอบสภาพอุปกรณ์ทั้งหมดดังกล่าวแล้วว่าอยู่ในสภาพสมบูรณ์พร้อมใช้งานและจะเก็บรักษาให้เป็นอย่างดี และข้าพเจ้าจะนำทรัพย์สินทั้งหมดมาคืนภายใน 1 วันนับจากวันที่มีผลลาออกไม่ว่าจะเป็นทางการหรือไม่ หากไม่นำกลับมาคืนภายในระยะเวลาที่กำหนด ข้าพเจ้ายินยอมให้ทาง บริษัท ไนน์ ดิสทริบิวชั่น จำกัด ดำเนินขั้นตอนตามกฎหมาย</strong></p><footer><div>ลงชื่อ ................................................ ผู้รับมอบ<p>(${esc(person.name)})</p>วันที่ ........../........../..........</div><div>ลงชื่อ ................................................ ผู้ส่งมอบ<p>(............................................................)</p>วันที่ ........../........../..........</div></footer></body></html>`;
}
function mountAssetDesk(root,store){
  let page="dashboard", filter="all", category="all", sourceFilter="all", query="", data={assets:[],people:[]}, problem="", loading=false, ready=false, toastTimer, destroyed=false;
  const pages={dashboard:["แดชบอร์ด","ภาพรวมและสถานะอุปกรณ์","grid"],assets:["ทะเบียนทรัพย์สิน","ค้นหา เพิ่ม และติดตามอุปกรณ์บริษัท","box"],transfers:["เปลี่ยนผู้ใช้งาน","โอนอุปกรณ์ที่มีผู้ใช้งานอยู่ให้พนักงานคนใหม่","transfer"],people:["พนักงาน","จัดการรายชื่อผู้รับผิดชอบทรัพย์สิน","user"],documents:["เอกสารรับมอบ","เปิดใบรับมอบที่บันทึกไว้เพื่อพิมพ์หรือบันทึก PDF","doc"]};
  root.classList.add("assetdesk");
  root.innerHTML=`<div class="ad-layout"><aside class="ad-sidebar"><div class="ad-brand"><div class="ad-logo">9D</div><div><b>Asset Desk</b><small>Nine Distribution</small></div></div><nav class="ad-nav" aria-label="เมนูหลัก"></nav><div class="ad-sidefoot">ทะเบียนทรัพย์สินบริษัท</div></aside><main class="ad-main"></main></div><dialog class="ad-dialog"></dialog><div class="ad-toast" role="status"></div>`;
  const nav=root.querySelector("nav"), main=root.querySelector("main"), dialog=root.querySelector("dialog"), toast=root.querySelector(".ad-toast");
  const byId=(kind,id)=>data[kind].find(row=>row.id===id);
  const navigate=(next,status="all")=>{page=next;filter=status;category="all";sourceFilter="all";query="";render();};
  const alertUser=message=>{toast.textContent=message;clearTimeout(toastTimer);toastTimer=setTimeout(()=>{toast.textContent="";},4400);};
  function assetRows(items, transferOnly=false){
    const action=transferOnly?"transfer":"detail";
    const label=transferOnly?"เปลี่ยนผู้ใช้งาน":"รายละเอียด";
    return items.length?`<div class="ad-overflow"><table class="ad-table"><thead><tr><th scope="col">รหัสทรัพย์สิน</th><th scope="col">อุปกรณ์</th><th scope="col">ประเภท</th><th scope="col">ผู้ใช้งาน</th><th scope="col">สถานะ</th><th scope="col">จัดการ</th></tr></thead><tbody>${items.map(asset=>`<tr><td>${esc(asset.code)}</td><td>${esc([asset.brand,asset.model].filter(Boolean).join(" "))}</td><td>${esc(asset.category)}</td><td>${esc(byId("people",asset.assignment?.person.id)?.name??asset.assignment?.person.name??"—")}</td><td>${statusTag(asset)}</td><td>${button(label,action,{id:asset.id,className:"ad-tablebtn"})}</td></tr>`).join("")}</tbody></table></div>`:`<p class="ad-empty">${loading?"กำลังโหลด…":transferOnly?"ไม่พบทรัพย์สินที่มีผู้ใช้งาน":"ไม่พบรายการทรัพย์สิน"}</p>`;
  }
  function dashboard(){
    const counts=Object.fromEntries(Object.keys(statuses).map(key=>[key,data.assets.filter(asset=>asset.status===key).length]));
    let offset=0;
    const slices=Object.entries(counts).filter(([,n])=>n>0).map(([key,n])=>{const start=offset;offset+=n/data.assets.length*100;return `${statuses[key][1]} ${start}% ${offset}%`;});
    const categoryCounts=Object.entries(data.assets.reduce((acc,asset)=>{acc[asset.category]=(acc[asset.category]??0)+1;return acc;},Object.create(null))).sort((a,b)=>b[1]-a[1]);
    const metric=[["all",["ทรัพย์สินทั้งหมด","","","box"],data.assets.length],...Object.entries(statuses).map(([key,value])=>[key,value,counts[key]])];
    const recent=[...data.assets].sort((a,b)=>String(b.updatedAt).localeCompare(String(a.updatedAt))).slice(0,5);
    const overdue=data.assets.filter(late).length;
    return `<section class="ad-metrics" aria-label="สรุปจำนวนทรัพย์สิน">${metric.map(([key,[label,color,bg,symbol],count])=>`<button class="ad-card" data-do="filter-status" data-id="${key}"><span class="ad-metric-label">${label}<span class="ad-iconbox" ${color?`style="background:${bg};color:${color}"`:""}>${icon(symbol)}</span></span><strong>${count}</strong></button>`).join("")}</section>
    <div class="ad-charts"><section class="ad-panel"><h2>สัดส่วนตามสถานะ</h2><div class="ad-donut-wrap"><div class="ad-donut" aria-hidden="true" style="background:${slices.length?`conic-gradient(${slices.join(",")})`:"#eeeaf8"}"><div><strong>${data.assets.length}</strong>รายการ</div></div><div class="ad-legend">${Object.entries(statuses).map(([key,[label,color]])=>`<div><i class="ad-dot" style="background:${color}"></i><span>${label}</span><b>${counts[key]}</b></div>`).join("")}</div></div></section>
    <section class="ad-panel"><div class="ad-panelhead"><h2>อุปกรณ์ตามประเภท</h2>${overdue?button(`เกินกำหนด ${overdue} รายการ`,"filter-status",{id:"overdue",className:"ad-tablebtn"}):""}</div>${categoryCounts.length?`<div class="ad-bars">${categoryCounts.map(([label,count])=>`<div class="ad-bar"><span title="${esc(label)}">${esc(label)}</span><div class="ad-track" aria-hidden="true"><i style="width:${100*count/categoryCounts[0][1]}%"></i></div><b>${count}</b></div>`).join("")}</div>`:'<p class="ad-empty">ยังไม่มีข้อมูล</p>'}</section></div>
    <section class="ad-panel"><div class="ad-panelhead"><h2>ทรัพย์สินที่อัปเดตล่าสุด</h2>${button("ดูทะเบียน","go-assets",{className:"ad-quiet",icon:"arrow"})}</div>${assetRows(recent)}</section>`;
  }
  function filteredAssets(){
    const term=query.trim().toLocaleLowerCase();
    return data.assets.filter(asset=>(page!=="transfers"||inCustody(asset))&&(filter==="all"||asset.status===filter||filter==="overdue"&&late(asset))&&(category==="all"||asset.category===category)&&(sourceFilter==="all"||asset.inventorySource===sourceFilter)&&[asset.code,asset.assetNumber,asset.company,asset.brand,asset.model,asset.serial,asset.category,asset.site,asset.assignment?.person.name,byId("people",asset.assignment?.person.id)?.name].map(value=>String(value??"").toLocaleLowerCase()).some(value=>value.includes(term))).sort((a,b)=>a.code.localeCompare(b.code,undefined,{numeric:true}));
  }
  function registry(){
    const transferOnly=page==="transfers";
    const option=(keys,current)=>keys.map(([key,value])=>`<option value="${esc(key)}" ${key===current?"selected":""}>${esc(value)}</option>`).join("");
    return `<section class="ad-panel"><h2>${transferOnly?"ทรัพย์สินที่มีผู้ใช้งาน":"รายการอุปกรณ์"}</h2><div class="ad-toolbar"><input class="ad-control" type="search" aria-label="ค้นหาทรัพย์สิน" placeholder="ค้นหารหัส รุ่น S/N หรือผู้รับ" value="${esc(query)}" data-query><div class="ad-actions">${data.inventorySources?`<select class="ad-control" aria-label="เลือกทะเบียน" data-source>${option([["all","ทุกทะเบียน"],...data.inventorySources.map(key=>[key,sources[key]])],sourceFilter)}</select>`:""}${transferOnly?"":`<select class="ad-control" aria-label="กรองสถานะ" data-filter>${option([["all","ทุกสถานะ"],...Object.entries(statuses).map(([k,v])=>[k,v[0]]),["overdue","เกินกำหนดคืน"]],filter)}</select>`}<select class="ad-control" aria-label="กรองประเภท" data-category>${option([["all","ทุกประเภท"],...[...new Set(data.assets.map(asset=>asset.category))].sort().map(key=>[key,key])],category)}</select></div></div><div data-results>${assetRows(filteredAssets(),page==="transfers")}</div></section>`;
  }
  function people(){
    return `<section class="ad-panel"><h2>รายชื่อผู้รับมอบ</h2>${data.people.length?`<div class="ad-overflow"><table class="ad-table"><thead><tr><th scope="col">ชื่อ–นามสกุล</th><th scope="col">รหัสพนักงาน</th><th scope="col">แผนก</th><th scope="col">ตำแหน่ง</th><th scope="col">อุปกรณ์ที่ถืออยู่</th><th scope="col">จัดการ</th></tr></thead><tbody>${[...data.people].sort((a,b)=>a.name.localeCompare(b.name,"th")).map(person=>`<tr><td>${esc(person.name)}</td><td>${esc(person.employeeCode||"—")}</td><td>${esc(person.department||"—")}</td><td>${esc(person.role||"—")}</td><td>${data.assets.filter(asset=>asset.assignment?.person.id===person.id).length}</td><td>${button("แก้ไข","edit-person",{id:person.id,className:"ad-tablebtn"})}</td></tr>`).join("")}</tbody></table></div>`:'<p class="ad-empty">ยังไม่มีรายชื่อพนักงาน</p>'}</section>`;
  }
  function allDocs(){
    return data.assets.flatMap(asset=>asset.history.filter(item=>item.doc).map(item=>item.doc)).sort((a,b)=>b.date.localeCompare(a.date)||b.id.localeCompare(a.id));
  }
  function documents(){
    const docs=allDocs();
    return `<section class="ad-panel"><h2>ใบรับมอบที่บันทึกไว้</h2>${docs.length?`<div class="ad-overflow"><table class="ad-table"><thead><tr><th scope="col">เลขที่เอกสาร</th><th scope="col">ทรัพย์สิน</th><th scope="col">ผู้รับ</th><th scope="col">วันที่รับ</th><th scope="col">กำหนดคืน</th><th scope="col">เอกสาร</th></tr></thead><tbody>${docs.map(doc=>`<tr><td>${esc(doc.id)}</td><td>${esc(doc.asset.code)} · ${esc(doc.asset.model)}</td><td>${esc(doc.person.name)}</td><td>${dateLabel(doc.date)}</td><td>${dateLabel(doc.due)}</td><td>${button("เปิดใบรับมอบ","document",{id:doc.id,className:"ad-tablebtn",icon:"doc"})}</td></tr>`).join("")}</tbody></table></div>`:'<p class="ad-empty">เอกสารจะปรากฏเมื่อบันทึกการส่งมอบ</p>'}</section>`;
  }
  function render(){
    if(destroyed)return;
    nav.innerHTML=Object.entries(pages).map(([key,[label,,symbol]])=>`<button type="button" data-do="nav" data-id="${key}" ${page===key?'aria-current="page"':""}>${icon(symbol)}${label}</button>`).join("");
    const [title,desc]=pages[page], view={dashboard,assets:registry,transfers:registry,people,documents}[page];
    const add=page==="people"?button("เพิ่มพนักงาน","new-person",{className:"ad-primary",icon:"plus"}):["documents","transfers"].includes(page)?"":button("เพิ่มทรัพย์สิน","new-asset",{className:"ad-primary",icon:"plus"});
    main.innerHTML=`<div class="ad-top"><span>การจัดการอุปกรณ์บริษัท</span><span class="ad-pill">SharePoint</span></div><header class="ad-intro"><div><h1>${title}</h1><p class="ad-sub">${desc}</p></div><div class="ad-actions">${button("รีเฟรช","refresh",{icon:"refresh"})}${ready?add:""}</div></header>${problem?`<div class="ad-notice ad-notice-error" role="alert">${icon("info")}<span>${esc(problem)}${ready?" · แสดงข้อมูลจากการโหลดครั้งก่อน":""}</span>${store.setup?button("ตั้งค่าฐานข้อมูล","setup",{className:"ad-tablebtn"}):""}</div>`:""}${loading&&!ready?'<p class="ad-empty" role="status">กำลังโหลดข้อมูล…</p>':ready?view():""}`;
  }
  async function refresh(){
    if(loading)return;
    loading=true;render();
    try{data=await store.read();problem="";ready=true;}
    catch(err){problem=err.message||"โหลดข้อมูลไม่สำเร็จ";}
    finally{loading=false;render();}
  }
  function show(title,body,footer=""){
    dialog.innerHTML=`<div class="ad-panelhead"><h2 id="ad-dialog-title">${esc(title)}</h2><button class="ad-button" type="button" data-do="close" aria-label="ปิด">${icon("close")}</button></div>${body}${footer}`;
    dialog.setAttribute("aria-labelledby","ad-dialog-title");
    if(!dialog.open)dialog.showModal();
  }
  function form(title,type,fields,submit,id){
    show(title,`<form data-form="${type}" ${id?`data-id="${esc(id)}"`:""}><div class="ad-form-grid">${fields}</div><output class="ad-form-error" role="alert"></output><footer>${button("ยกเลิก","close")}<button type="submit" class="ad-button ad-primary">${esc(submit)}</button></footer></form>`);
  }
  function assetForm(asset={}){
    if(data.inventorySources&&!asset.inventorySource){
      show("เพิ่มทรัพย์สิน","<p>เลือกทะเบียนสำหรับอุปกรณ์</p>",`<footer>${data.inventorySources.map(source=>button(sources[source],"new-native",{id:source,className:"ad-primary"})).join("")}</footer>`);
      return;
    }
    const native=asset.inventorySource;
    const types=[...new Set([...(inventoryCategories[native]??categories),asset.category].filter(Boolean))];
    const companies=[["","ไม่ระบุ"],...(inventoryCompanies[native]??[])];
    if(asset.company&&!companies.some(([value])=>value===asset.company)) companies.push([asset.company,asset.company]);
    form(asset.id?"แก้ไขทรัพย์สิน":native?`เพิ่มทรัพย์สิน · ${sources[native]}`:"เพิ่มทรัพย์สิน","asset",[
      native?`<input type="hidden" name="inventorySource" value="${esc(native)}">`:"",
      field("code",native?"Host Name / รหัสทรัพย์สิน":"รหัสทรัพย์สิน",asset.code,{required:true}),field("category","ประเภท",asset.category,{choices:[["","เลือกประเภท"],...types.map(cat=>[cat,cat])],required:true}),
      field("brand","ยี่ห้อ",asset.brand),field("model","รุ่น / ชื่ออุปกรณ์",asset.model,{required:true}),
      native?field("assetNumber","Asset or serial number",asset.assetNumber):"",
      native?field("company","บริษัท",asset.company,{choices:companies}):"",
      native==="accessories"?"":field("serial","หมายเลขเครื่อง (S/N)",asset.serial),field("site","สถานที่",asset.site),
      field("note","หมายเหตุ",asset.note,{wide:true,area:true,maxLength:native?255:500})
    ].join(""),"บันทึกทรัพย์สิน",asset.id);
  }
  function personForm(person={}){
    form(person.id?"แก้ไขพนักงาน":"เพิ่มพนักงาน","person",[field("name","ชื่อ–นามสกุล",person.name,{required:true}),field("employeeCode","รหัสพนักงาน",person.employeeCode),field("department","แผนก",person.department),field("role","ตำแหน่ง",person.role)].join(""),"บันทึกพนักงาน",person.id);
  }
  function detail(asset){
    const name=byId("people",asset.assignment?.person.id)?.name??asset.assignment?.person.name;
    const details=[...(asset.inventorySource?[["ทะเบียน",sources[asset.inventorySource]],["บริษัท",inventoryCompanies[asset.inventorySource]?.find(([value])=>value===asset.company)?.[1]??asset.company],["Asset or serial number",asset.assetNumber],["ผู้ใช้งานก่อนหน้า",asset.previousUsers],["วันที่คืนล่าสุด",dateLabel(asset.returnedDate)]]:[]),["ประเภท",asset.category],["หมายเลขเครื่อง",asset.serial],["สถานที่",asset.site],["ผู้ใช้งาน",name??asset.currentHolder],["วันที่รับ",dateLabel(asset.assignment?.date||asset.receivedDate)],["กำหนดคืน",dateLabel(asset.assignment?.due)],["หมายเหตุ",asset.note]];
    show(`${asset.code} · ${[asset.brand,asset.model].filter(Boolean).join(" ")}`,`${statusTag(asset)}<div class="ad-fields">${details.map(([title,value])=>`<div><small>${title}</small><span>${esc(value||"—")}</span></div>`).join("")}</div><h3>ประวัติ</h3>${asset.history.length?`<ol class="ad-history">${[...asset.history].reverse().map(item=>`<li>${esc(timeLabel(item.at))} · ${esc(item.label)}${item.doc?.previousHolder?` · ${esc(item.doc.previousHolder.name)} → ${esc(item.doc.person.name)}`:""}${item.note?` · ${esc(item.note)}`:""} ${item.doc?button("ใบรับมอบ","document",{id:item.doc.id,className:"ad-tablebtn"}):""}</li>`).join("")}</ol>`:'<p class="ad-sub">ยังไม่มีประวัติการเปลี่ยนแปลง</p>'}`,`<footer>${button("แก้ไข","edit-asset",{id:asset.id})}${asset.status==="available"?button("ส่งมอบ","assign",{id:asset.id,className:"ad-primary"}):""}${inCustody(asset)?button("เปลี่ยนผู้ใช้งาน","transfer",{id:asset.id,className:"ad-primary"}):""}${button(inCustody(asset)?"รับคืน / เปลี่ยนสถานะ":"เปลี่ยนสถานะ","change-status",{id:asset.id})}</footer>`);
  }
  let transferSource;
  function handoverForm(asset, transfer=false){
    if(transfer && !inCustody(asset))return;
    const current=transfer?asset.assignment:null;
    transferSource=current?asset:null;
    const title=transfer?"เปลี่ยนผู้ใช้งาน":"ส่งมอบ";
    const recipients=data.people.filter(person=>person.id!==current?.person.id)
      .sort((a,b)=>a.name.localeCompare(b.name,"th"))
      .map(person=>[person.id,[person.name,person.department,person.employeeCode].filter(Boolean).join(" · ")]);
    const summary=current?`<div class="ad-wide"><strong>${esc([asset.brand,asset.model].filter(Boolean).join(" "))}</strong><div>ผู้ใช้งานเดิม ${esc(current.person.name)}</div></div>`:"";
    if(!recipients.length){
      show(`${title} · ${asset.code}`,`${summary}<p>เพิ่มรายชื่อผู้รับมอบก่อนบันทึก</p>`,`<footer>${button("เพิ่มพนักงาน",transfer?"person-for-transfer":"person-for-asset",{id:asset.id,className:"ad-primary"})}</footer>`);
      return;
    }
    const receiptDate=today();
    form(`${title} · ${asset.code}`,transfer?"transfer":"assign",[
      summary,
      field("person",transfer?"ผู้ใช้งานใหม่":"ผู้รับมอบ","",{wide:true,required:true,choices:[["","เลือกพนักงาน"],...recipients]}),
      data.inventorySources?field("custodyStatus","รูปแบบการรับมอบ",asset.status==="loaned"?"loaned":"assigned",{wide:true,choices:[["assigned",statuses.assigned[0]],["loaned",statuses.loaned[0]]]}):"",
      field("date","วันที่รับ",current?.date>receiptDate?current.date:receiptDate,{type:"date",required:true}),field("due","กำหนดคืน","",{type:"date"}),
      field("accessories","อุปกรณ์เสริมและหมายเลขเครื่อง",current?.accessories,{wide:true,area:true}),field("condition","สภาพและหมายเหตุ","",{wide:true,area:true})
    ].join(""),transfer?"บันทึกการเปลี่ยนและเปิดใบรับมอบ":"บันทึกและเปิดใบรับมอบ",asset.id);
  }
  function statusOptions(asset){
    const active=asset.inventorySource&&asset.assignment?["assigned","loaned"]:[];
    return ["available","repair","unavailable","lost","retired",...active].filter(key=>key!==asset.status);
  }
  function changeStatus(asset){
    const options=statusOptions(asset).map(key=>[key,statuses[key][0]]);
    form(inCustody(asset)?`รับคืน / เปลี่ยนสถานะ · ${asset.code}`:`เปลี่ยนสถานะ · ${asset.code}`,"status",field("status","สถานะใหม่",options[0][0],{choices:options,wide:true})+field("note","สภาพอุปกรณ์ / หมายเหตุ","",{wide:true,area:true}),"บันทึกสถานะ",asset.id);
  }
  function documentPreview(doc){
    const iframe=document.createElement("iframe");
    iframe.title="ตัวอย่างใบรับมอบ";
    iframe.setAttribute("sandbox","allow-same-origin allow-modals");
    iframe.style.cssText="width:100%;height:62vh;border:1px solid #e9e6f1;border-radius:8px";
    iframe.srcdoc=printMarkup(doc);
    show(`ใบรับมอบ ${doc.id}`,`<div class="ad-panelhead"><p class="ad-sub">เอกสารสำหรับลงชื่อผู้รับและผู้ส่งมอบ</p>${button("พิมพ์ / บันทึก PDF","print",{icon:"doc"})}</div><div data-preview></div>`);
    dialog.querySelector("[data-preview]").append(iframe);
  }
  let pendingAsset;
  async function submit(event){
    const formElement=event.target.closest("[data-form]");
    if(!formElement)return;
    event.preventDefault();
    if(formElement.dataset.busy)return;
    formElement.dataset.busy="true";
    const submitButton=formElement.querySelector('button[type="submit"]');
    const output=formElement.querySelector("output");
    submitButton.disabled=true;output.textContent="";
    const values=Object.fromEntries([...new FormData(formElement)].map(([key,value])=>[key,String(value).trim()]));
    const id=formElement.dataset.id;
    const type=formElement.dataset.form;
    const now=new Date().toISOString();
    try{
      if(loading)throw Error("กำลังโหลดข้อมูล กรุณารอสักครู่");
      let kind="assets", record, doc;
      if(type==="setup"){await store.setup();dialog.close();await refresh();return;}
      if(type==="person"){
        kind="people";
        if(!values.name)throw Error("กรุณากรอกชื่อ–นามสกุล");
        if(values.employeeCode&&data.people.some(person=>person.id!==id&&person.employeeCode?.toLowerCase()===values.employeeCode.toLowerCase()))throw Error("รหัสพนักงานนี้มีอยู่แล้ว");
        const old=byId(kind,id);
        record={...old,...values,id:id??crypto.randomUUID(),updatedAt:now};
      }else{
        const old=byId(kind,id);
        if(id&&!old)throw Error("ไม่พบทรัพย์สิน กรุณารีเฟรช");
        if(type==="asset"){
          if(!values.code||!values.model||!(inventoryCategories[values.inventorySource]??categories).includes(values.category)&&values.category!==old?.category)throw Error("กรุณากรอกรหัส รุ่น และประเภทอุปกรณ์");
          if(data.assets.some(asset=>asset.id!==id&&asset.inventorySource===values.inventorySource&&asset.code.toLowerCase()===values.code.toLowerCase()))throw Error("รหัสทรัพย์สินนี้มีอยู่แล้ว");
          record={status:"available",...old,...values,id:id??crypto.randomUUID(),updatedAt:now,history:[...(old?.history??[]),{at:now,label:old?"แก้ไขข้อมูลอุปกรณ์":"เพิ่มทรัพย์สิน"}]};
        }else if(type==="assign"||type==="transfer"){
          const transfer=type==="transfer";
          if(transfer && (old!==transferSource||!inCustody(old)))throw Error("ข้อมูลทรัพย์สินเปลี่ยนแล้ว กรุณาปิดและเปิดรายการใหม่");
          const current=transfer?old.assignment:null;
          const person=byId("people",values.person);
          if(!transfer&&old.status!=="available")throw Error("อุปกรณ์ต้องอยู่ในสถานะพร้อมใช้งาน");
          if(!person||person.id===current?.person.id||!values.date)throw Error("กรุณาเลือกผู้รับมอบและระบุวันที่รับ");
          if(current&&values.date<current.date)throw Error("วันที่เปลี่ยนผู้ใช้งานต้องตรงกับหรือหลังวันที่รับครั้งก่อน");
          if(values.due&&values.due<values.date)throw Error("กำหนดคืนต้องตรงกับหรือหลังวันที่รับ");
          const {id:personId,name,employeeCode,department,role}=person;
          doc={id:`HO-${values.date.replaceAll("-","")}-${crypto.randomUUID().slice(0,12).toUpperCase()}`,asset:assetSnapshot(old),person:{id:personId,name,employeeCode,department,role},date:values.date,due:values.due,accessories:values.accessories,condition:values.condition};
          if(current)doc.previousHolder={id:current.person.id,name:current.person.name};
          const label=transfer?"เปลี่ยนผู้ใช้งาน":`ส่งมอบให้ ${name}`;
          record={...old,status:values.custodyStatus==="loaned"?"loaned":"assigned",assignment:doc,updatedAt:now,history:[...old.history,{at:now,label,doc}]};
        }else{
          if(!statusOptions(old).includes(values.status))throw Error("เลือกสถานะใหม่ที่รองรับ");
          const assignment=inCustody(values)?old.assignment:null;
          const returned=inCustody(old)&&["available","repair","unavailable"].includes(values.status);
          const label=returned?`รับคืนจาก ${old.assignment?.person.name??"ผู้ใช้งาน"} · ${statuses[values.status][0]}`:statuses[values.status][0];
          record={...old,status:values.status,assignment,updatedAt:now,history:[...old.history,{at:now,label,note:values.note}]};
        }
      }
      const saved=await store.save(kind,record);
      data[kind]=[...data[kind].filter(row=>row.id!==saved.id),saved];
      dialog.close();render();alertUser("บันทึกสำเร็จ");
      const savedDoc=doc&&saved.history.find(item=>item.doc?.id===doc.id)?.doc;
      if(savedDoc)documentPreview(savedDoc);
      else if(pendingAsset&&kind==="people"){
        const {assetId,transfer}=pendingAsset;
        pendingAsset=undefined;
        const asset=byId("assets",assetId);
        if(asset)handoverForm(asset,transfer);
      }
      else if(kind==="assets")detail(saved);
    }catch(err){output.textContent=err.message||"บันทึกไม่สำเร็จ";}
    finally{delete formElement.dataset.busy;submitButton.disabled=false;}
  }
  function click(event){
    const trigger=event.target.closest("[data-do]");
    if(!trigger||!root.contains(trigger))return;
    if(dialog.querySelector('[data-busy="true"]'))return;
    const {do:action,id}=trigger.dataset;
    const asset=byId("assets",id);
    const actions={
      nav:()=>navigate(id),"go-assets":()=>navigate("assets"),"filter-status":()=>navigate("assets",id),
      refresh,close:()=>{pendingAsset=undefined;dialog.close();},"new-asset":()=>assetForm(),"new-native":()=>Object.hasOwn(sources,id)&&assetForm({inventorySource:id}),"new-person":()=>personForm(),
      detail:()=>asset&&detail(asset),"edit-asset":()=>asset&&assetForm(asset),"edit-person":()=>personForm(byId("people",id)),
      assign:()=>asset&&handoverForm(asset),transfer:()=>asset&&handoverForm(asset,true),"change-status":()=>asset&&changeStatus(asset),
      "person-for-asset":()=>{pendingAsset={assetId:id,transfer:false};personForm();},
      "person-for-transfer":()=>{pendingAsset={assetId:id,transfer:true};personForm();},
      print:()=>{const frame=dialog.querySelector("iframe");if(frame?.contentDocument?.readyState==="complete")frame.contentWindow.print();},
      document:()=>{const doc=allDocs().find(item=>item.id===id);if(doc)documentPreview(doc);},
      setup:()=>form("ตั้งค่าฐานข้อมูล SharePoint","setup","<p class='ad-wide'>สร้างรายการทรัพย์สินและพนักงานในไซต์นี้ ต้องมีสิทธิ์จัดการรายการ</p>","สร้างรายการ")
    };
    actions[action]?.();
  }
  function change(event){
    if(event.target.matches("[data-query]"))query=event.target.value;
    else if(event.target.matches("[data-filter]"))filter=event.target.value;
    else if(event.target.matches("[data-category]"))category=event.target.value;
    else if(event.target.matches("[data-source]"))sourceFilter=event.target.value;
    else return;
    main.querySelector("[data-results]").innerHTML=assetRows(filteredAssets(),page==="transfers");
  }
  function cancel(event){if(dialog.querySelector('[data-busy="true"]'))event.preventDefault();else pendingAsset=undefined;}
  root.addEventListener("click",click);root.addEventListener("submit",submit);root.addEventListener("input",change);dialog.addEventListener("cancel",cancel);
  refresh();
  return ()=>{destroyed=true;clearTimeout(toastTimer);root.removeEventListener("click",click);root.removeEventListener("submit",submit);root.removeEventListener("input",change);dialog.removeEventListener("cancel",cancel);root.replaceChildren();};
}
export { mountAssetDesk };

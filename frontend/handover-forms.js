/* Hand-over and return A4 forms only. Employee summary PDF remains unchanged. */
(() => {
'use strict';
const $=id=>document.getElementById(id);
const esc=x=>String(x??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const en=()=>window.nineLanguage?.()==='en';
const label=(th,english)=>en()?english:th;
function style(){
 if($('ho-forms-v3-style'))return;
 const s=document.createElement('style');s.id='ho-forms-v3-style';
 s.textContent='.ho-paper{font:11px/1.48 Sarabun,sans-serif;color:#1e293b}.ho-paper .ho-page{break-after:page;page-break-after:always}.ho-paper .ho-page:last-child{break-after:auto;page-break-after:auto}.ho-paper .ho-header{display:grid;grid-template-columns:110px minmax(0,1fr) 145px;align-items:center;gap:12px;border-bottom:2px solid #315b95;padding-bottom:10px;margin-bottom:14px}.ho-paper .ho-logo{width:110px;height:76px;object-fit:contain}.ho-paper h1{font-size:18px;font-weight:700;line-height:1.3}.ho-paper .ho-docid{text-align:right;font-size:10px;overflow-wrap:anywhere}.ho-paper .ho-section{font-weight:bold;font-size:12px;background:#f1f5f9;border-left:4px solid #315b95;padding:6px 9px;margin:13px 0 8px}.ho-paper .ho-fields{display:grid;grid-template-columns:1fr 1fr;gap:7px 17px}.ho-paper .ho-fields div{border-bottom:1px dotted #94a3b8;min-height:22px;overflow-wrap:anywhere}.ho-paper table{width:100%;table-layout:fixed;border-collapse:collapse;font-size:10px}.ho-paper th,.ho-paper td{border:1px solid #b9c5d4;padding:6px 5px;vertical-align:top;overflow-wrap:anywhere}.ho-paper th{background:#f1f5f9}.ho-paper tr{break-inside:avoid;page-break-inside:avoid}.ho-paper .ho-note{min-height:47px;padding:8px;border:1px solid #cbd5e1;border-radius:4px}.ho-paper .ho-signs{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px;margin-top:17px;break-inside:avoid}.ho-paper .ho-signs div{border:1px solid #cbd5e1;text-align:center;padding:9px 3px;min-height:106px}.ho-paper .ho-footer{margin-top:11px;font-size:9px;color:#64748b;text-align:right}.ho-paper .ho-acknowledgement{display:grid;grid-template-columns:1fr 1fr;gap:15px;margin-top:12px;break-inside:avoid;page-break-inside:avoid}.ho-paper .ho-recipient{border:1px solid #cbd5e1;border-radius:5px;text-align:center;padding:10px;min-height:108px}.ho-paper .ho-signature-line{margin-top:28px}.ho-paper .ho-system-issuer{border:1px solid #cbd5e1;border-radius:5px;padding:12px;display:flex;flex-direction:column;justify-content:space-between;gap:13px;overflow-wrap:anywhere}.ho-paper .ho-system-stamp{font-size:10px;color:#475569;text-align:right}.ho-paper .ho-note p{overflow-wrap:anywhere}@media print{@page{size:A4 portrait;margin:12mm}#print-modal{display:block!important;position:static!important;inset:auto!important;overflow:visible!important;padding:0!important}#print-modal-wrapper{max-width:none!important;width:100%!important;margin:0!important}#print-modal-wrapper>.bg-slate-200\/80{background:white!important;padding:0!important;overflow:visible!important}#a4-document-content{padding:0!important;width:100%!important;min-height:0!important;box-shadow:none!important}.no-print{display:none!important}}';
 document.head.append(s);
}
function paper(tx){
 const returned=tx.type==='CHECKIN',employee=tx.employee||{},org=employee.company||'Nine distribution co.,ltd';
 const title=returned?label('ใบตรวจรับคืนทรัพย์สินและอุปกรณ์ไอที','IT Asset Return & Inspection'):label('ใบส่งมอบและรับทรัพย์สินไอที','IT Asset Handover & Receipt');
 const subtitle=returned?'IT ASSET RETURN & INSPECTION FORM':'IT ASSET HANDOVER & RECEIPT FORM';
 const section=(n,th,eng)=>'<div class="ho-section">'+n+'. '+label(th,eng)+'</div>';
 const field=(name,value)=>'<div><span style="color:#64748b">'+name+':</span> <b>'+esc(value||'-')+'</b></div>';
 const date=tx.date?new Date(tx.date).toLocaleDateString(en()?'en-GB':'th-TH',{year:'numeric',month:'short',day:'numeric'}):'-';
 const max=6,count=Math.max(1,Math.ceil((tx.items||[]).length/max));let result='';
 for(let p=0;p<count;p++){
  const header='<div class="ho-header"><img class="ho-logo" alt="Nine Distribution" src="./nine-distribution-brand.webp"><div><b style="font-size:10px;color:#315b95">'+esc(org)+'</b><h1>'+title+'</h1><small>'+subtitle+'</small></div><div class="ho-docid">'+label('เลขที่เอกสาร','Doc No.')+': <b>'+esc(tx.docNo)+'</b><br>'+label('วันที่','Date')+': '+esc(date)+'<br>'+label('หน้า','Page')+' '+(p+1)+'/'+count+'</div></div>';
  const details=section(1,'ข้อมูลพนักงาน','Employee Information')+'<div class="ho-fields">'+
   field(label('ชื่อ-นามสกุล','Full Name'),employee.name)+field(label('รหัสพนักงาน','Employee ID'),employee.empId)+
   field(label('สังกัด','Company'),org)+field(label('แผนก','Department'),employee.department)+
   field(label('ตำแหน่ง','Position'),employee.position)+field(label('ข้อมูลติดต่อ','Contact'),employee.contact)+'</div>';
  const rows=(tx.items||[]).slice(p*max,(p+1)*max).map((a,i)=>'<tr><td style="text-align:center">'+(p*max+i+1)+'</td><td><b>'+esc(a.id)+'</b></td><td>'+esc(a.category||'-')+'</td><td>'+esc([a.brand,a.model].filter(Boolean).join(' '))+'</td><td>'+esc(a.serial||'-')+'</td><td>'+(returned?'☐ '+label('ปกติ','Normal')+'<br>☐ '+label('ชำรุด','Damaged'):esc(a.condition||'-'))+'</td></tr>').join('');
  const table=section(2,'รายการทรัพย์สิน','Asset Details')+'<div style="font-size:10px;margin-bottom:5px">'+label('รวม','Total')+' '+(tx.items||[]).length+' '+label('รายการ','items')+'</div><table><colgroup><col style="width:6%"><col style="width:17%"><col style="width:12%"><col style="width:27%"><col style="width:20%"><col style="width:18%"></colgroup><thead><tr><th>#</th><th>Asset ID</th><th>'+label('ประเภท','Type')+'</th><th>'+label('ยี่ห้อ / รุ่น','Brand / Model')+'</th><th>Serial</th><th>'+label('สภาพ','Condition')+'</th></tr></thead><tbody>'+rows+'</tbody></table>';
  let footer='';
  if(p===count-1){
   const terms=returned?label('เจ้าหน้าที่รับอุปกรณ์คืนเพื่อทำการตรวจสอบสภาพและความครบถ้วน หากพบความเสียหายหรืออุปกรณ์สูญหายให้ระบุไว้ในเอกสาร','Returned equipment is subject to condition and completeness inspection. Record any damage or missing items.'):label("ข้าพเจ้าได้ทำการตรวจสอบสภาพอุปกรณ์ทั้งหมดดังกล่าวแล้วว่าอยู่ในสภาพสมบูรณ์พร้อมใช้งาน และจะเก็บรักษาทรัพย์สินทั้งหมดให้เป็นอย่างดี ข้าพเจ้าจะนำทรัพย์สินทั้งหมดมาคืนภายใน 1 วันนับจากวันที่มีผลลาออก ไม่ว่าจะเป็นการลาออกอย่างเป็นทางการหรือไม่ก็ตาม หากไม่นำทรัพย์สินกลับมาคืนภายในระยะเวลาที่กำหนด ข้าพเจ้ายินยอมให้บริษัท ไนน์ ดิสทริบิวชั่น จำกัด ดำเนินการตามขั้นตอนของกฎหมายที่เกี่ยวข้อง","I have inspected all the listed equipment and confirm that it is in good condition and ready for use. I will take proper care of these assets and return all of them within one day from the effective date of my resignation, whether formally notified or otherwise. If I fail to return the assets within the prescribed period, I acknowledge that Nine Distribution Co., Ltd. may pursue remedies available under applicable law.");
   const officerName=(currentUser?.name||tx.officerName||'-').trim();
   const officerMail=currentUser?.email||'';
   const issuer='<div class="ho-system-issuer"><div><b>'+label('ผู้จัดทำเอกสาร / ผู้ส่งมอบ','Document Issuer / Handed Over By')+'</b><div>'+esc(officerName)+'</div>'+(officerMail?'<small>'+esc(officerMail)+'</small>':'')+'</div><div class="ho-system-stamp">'+label('(ออกโดยระบบ)','(System Generated)')+'</div></div>';
   const signature='<div class="ho-recipient"><b>'+ (returned?label('ผู้ส่งคืนทรัพย์สิน (พนักงาน)','Returned By (Employee)'):label('ผู้รับมอบทรัพย์สิน (พนักงาน)','Received By (Employee)')) +'</b><div class="ho-signature-line">........................................................</div><div>('+esc(employee.name||'................................')+')</div><div>'+label('วันที่','Date')+' ........../........../..........</div></div>';
   footer=section(3,'หมายเหตุและเงื่อนไข','Conditions & Notes')+
     '<div class="ho-note">'+(tx.note?'<p style="margin:0 0 9px"><b>'+label('หมายเหตุ','Notes')+':</b> '+esc(tx.note)+'</p>':'')+
     '<p style="margin:0;line-height:1.75;text-align:justify">'+terms+'</p></div>'+
     section(4,'การยืนยันรับมอบและออกเอกสาร','Acknowledgement & Issuance')+
     '<div class="ho-acknowledgement">'+signature+issuer+'</div>';
  }
  result+='<div class="ho-page">'+header+details+table+footer+'<div class="ho-footer">'+esc(tx.docNo)+' · 9D IT Asset Pro</div></div>';
 }
 return result;
}
window.nineRenderHandoverPaper=function(tx){style();return '<div class="ho-paper">'+paper(tx)+'</div>';};
openPrintDocument=function(docNo){
 const tx=transactions.find(x=>x.docNo===docNo);
 if(!tx||!['CHECKOUT','CHECKIN'].includes(tx.type))return showToast('ไม่พบเอกสาร','error');
 style();
 $('a4-document-content').innerHTML='<div class="ho-paper">'+paper(tx)+'</div>';
 const modal=$('print-modal');modal.classList.remove('hidden');modal.classList.add('block');
 const header=modal.querySelector('.no-print span.font-semibold');
 if(header)header.textContent=label('ตัวอย่างเอกสาร A4 พร้อมพิมพ์ / บันทึก PDF','A4 Print Preview / Save as PDF');
};
document.addEventListener('DOMContentLoaded',style);
})();
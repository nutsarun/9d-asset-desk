/* 9D IT Asset Pro — automatic PDF archive of completed handover/return receipts.
 * Destination: Inventory site's SharePoint Documents library.
 * Important: the underlying transaction is already saved before this runs; never re-run it on archival failure.
 */
(() => {
'use strict';
const SITE='ninedistribution.sharepoint.com,13d8bc75-67a0-4dbf-b579-8b26c2dd2e9f,dedddf0b-d21b-4660-84ad-259b6ffb0ab4';
const ROOT='https://graph.microsoft.com/v1.0/sites/'+SITE+'/drive/root:';
const PARENT='9D Asset Documents';
const DIRS={CHECKOUT:'รับมอบ (HO)',CHECKIN:'ส่งคืน (RT)'};
const $=id=>document.getElementById(id);
let pdfLibPromise=null;
const safeName=s=>String(s??'').normalize('NFC').replace(/[\\/:*?"<>|#%{}\u0000-\u001f]/g,' ').replace(/\s+/g,' ').trim().replace(/[. ]+$/g,'').slice(0,115);
const filename=tx=>safeName(tx.docNo)+'_'+safeName(tx.employee?.name||'ไม่ระบุชื่อ')+'.pdf';
const folderPath=tx=>[PARENT,DIRS[tx.type]].map(encodeURIComponent).join('/');
async function token(){
 const a=window.nineGraphAuth;
 if(!a?.client||!a.account)throw Error('ยังไม่ได้เข้าสู่ระบบ Microsoft 365');
 const req={account:a.account,scopes:['Sites.ReadWrite.All']};
 try{return (await a.client.acquireTokenSilent(req)).accessToken}
 catch(err){return (await a.client.acquireTokenPopup(req)).accessToken}
}
async function graph(url,method='GET',body=null,ctype='application/json',extra={}){
 const headers={Authorization:'Bearer '+await token(),...extra};
 if(body!==null)headers['Content-Type']=ctype;
 const res=await fetch(url,{method,headers,body:body===null?undefined:body});
 if(res.status===404)return null;
 if(!res.ok)throw Error('SharePoint '+res.status+': '+(await res.text()).slice(0,230));
 if(res.status===204)return {};
 return res.json();
}
function pdfLibrary(){
 if(window.html2pdf)return Promise.resolve(window.html2pdf);
 if(pdfLibPromise)return pdfLibPromise;
 pdfLibPromise=new Promise((resolve,reject)=>{
  const script=document.createElement('script');script.async=true;
  script.src='https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js';
  script.onload=()=>window.html2pdf?resolve(window.html2pdf):reject(Error('โหลดตัวสร้าง PDF ไม่สมบูรณ์'));
  script.onerror=()=>reject(Error('โหลดระบบ PDF ไม่ได้ โปรดตรวจสอบการอนุญาต CDN ของบริษัท'));
  document.head.appendChild(script);
 }).catch(e=>{pdfLibPromise=null;throw e});
 return pdfLibPromise;
}
async function pdfBlob(tx){
 if(typeof window.nineRenderHandoverPaper!=='function')throw Error('ไม่พบแม่แบบเอกสาร HO/RT');
 const lib=await pdfLibrary();
 const html=window.nineRenderHandoverPaper(tx);
 if(!html||!html.includes('ho-paper')||!html.includes(safeHTML(tx.docNo)))
  throw Error('ไม่สามารถสร้างเนื้อหาเอกสาร HO/RT ได้');
 // html2canvas does not reliably capture elements positioned far outside the viewport.
 // Render the document at an actual on-screen coordinate, behind an opaque progress panel.
 const overlay=document.createElement('div');
 overlay.id='nine-pdf-render-overlay';
 overlay.style.cssText='position:fixed;inset:0;z-index:99999;overflow:auto;background:#f1f5f9;padding:16px;display:flex;flex-direction:column;align-items:center;gap:12px;';
 overlay.innerHTML='<div style="width:100%;max-width:720px;font:600 14px Sarabun,sans-serif;color:#334155;text-align:center">กำลังสร้าง PDF จากใบรับมอบ/รับคืน โปรดอย่าปิดหน้านี้</div>';
 const container=document.createElement('div');
 container.className='ho-pdf-capture';
 container.style.cssText='position:relative;left:0;top:0;display:block;width:186mm;max-width:none;min-height:100mm;background:#fff;color:#1e293b;padding:0;margin:0;box-sizing:border-box;flex:none;';
 container.innerHTML=html;
 overlay.appendChild(container);
 document.body.appendChild(overlay);
 try{
  await document.fonts?.ready;
  await Promise.all([...container.querySelectorAll('img')].map(img=>new Promise((resolve,reject)=>{
   if(img.complete){if(img.naturalWidth)resolve();else reject(Error('โหลดโลโก้ใน PDF ไม่สำเร็จ'));return}
   img.onload=resolve;img.onerror=()=>reject(Error('โหลดโลโก้ใน PDF ไม่สำเร็จ'));
  })));
  // Force browser layout and a painted frame before rasterizing.
  if(container.getBoundingClientRect().width<200 || container.getBoundingClientRect().height<100)
   throw Error('แบบฟอร์มไม่แสดงบนหน้าเว็บ จึงยกเลิกการสร้าง PDF');
  await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
  const worker=lib().set({
   margin:[12,12,12,12],
   image:{type:'jpeg',quality:0.98},
   html2canvas:{scale:2,useCORS:true,backgroundColor:'#ffffff',scrollX:0,scrollY:0,logging:false},
   jsPDF:{unit:'mm',format:'a4',orientation:'portrait'},
   pagebreak:{mode:['css','legacy'],before:'.ho-page:not(:first-child)'}
  }).from(container);
  await worker.toCanvas();
  const canvas=await worker.get('canvas');
  // Detect blank/white pages before we upload anything to SharePoint.
  if(!canvas || canvas.width<400 || canvas.height<300)throw Error('ตัวสร้าง PDF ไม่พบภาพเอกสาร');
  const ctx=canvas.getContext('2d',{willReadFrequently:true});
  if(!ctx)throw Error('ไม่สามารถตรวจสอบภาพเอกสาร');
  const {data}=ctx.getImageData(0,0,canvas.width,canvas.height);
  let visible=0;const pixelCount=canvas.width*canvas.height, stride=Math.max(1,Math.floor(pixelCount/50000));
  for(let p=0;p<pixelCount;p+=stride){
   const n=p*4;
   if(data[n]<205 || data[n+1]<205 || data[n+2]<205)visible++;
  }
  if(visible<35)throw Error('ตรวจพบหน้ากระดาษเปล่า ระบบยกเลิกการบันทึก PDF เพื่อป้องกันไฟล์เสีย');
  const result=await worker.toPdf().outputPdf('blob');
  if(!(result instanceof Blob)||result.size<3000)throw Error('สร้างไฟล์ PDF ไม่สมบูรณ์');
  const beginning=await result.slice(0,5).text();
  if(beginning!=='%PDF-')throw Error('ไฟล์ที่สร้างไม่ใช่ PDF ที่ถูกต้อง');
  return result;
 }finally{overlay.remove()}
}
function safeHTML(value){return String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function persist(){
 try{sessionStorage.setItem('nine-transactions-session',JSON.stringify(transactions))}catch(e){console.warn('Cannot cache document link',e)}
 if(typeof renderHistory==='function')renderHistory();
}
async function archiveTransaction(tx){
 if(!tx||!DIRS[tx.type]||!tx.docNo||!tx.employee?.name)return {ok:false,error:'ข้อมูลเอกสารไม่ครบ'};
 if(tx.pdfStatus==='saved'&&tx.pdfUrl)return {ok:true,url:tx.pdfUrl,alreadySaved:true};
 if(tx.pdfStatus==='saving')return {ok:false,error:'กำลังบันทึกไฟล์ PDF'};
 tx.pdfStatus='saving';persist();
 const name=filename(tx),folder=folderPath(tx),url=ROOT+'/'+folder+'/'+encodeURIComponent(name);
 try{
  // Collision avoidance: never silently overwrite a previous HO/RT document.
  const old=await graph(url,'GET');
  if(old){
   tx.pdfStatus='conflict';tx.pdfError='มีไฟล์ชื่อ '+name+' อยู่แล้ว ต้องตรวจสอบเลขที่เอกสารใน SharePoint ก่อน';persist();
   throw Error(tx.pdfError);
  }
  const pdf=await pdfBlob(tx);
  // If-None-Match is an additional race-condition guard; folders must be writeable by the logged-in staff member.
  const uploaded=await graph(url+':/content','PUT',pdf,'application/pdf',{'If-None-Match':'*'});
  if(!uploaded?.id||!uploaded.webUrl)throw Error('ไม่ได้รับข้อมูลยืนยันไฟล์จาก SharePoint');
  tx.pdfStatus='saved';tx.pdfUrl=uploaded.webUrl;tx.pdfName=name;tx.pdfError='';
  persist();
  showToast('จัดเก็บ PDF ใน SharePoint สำเร็จ: '+name,'success');
  return {ok:true,url:uploaded.webUrl,name};
 }catch(err){
  tx.pdfStatus=tx.pdfStatus==='conflict'?'conflict':'error';
  tx.pdfError=err.message||String(err);persist();
  showToast('บันทึกทรัพย์สินสำเร็จ แต่ยังจัดเก็บ PDF ไม่ได้: '+tx.pdfError,'error');
  return {ok:false,error:tx.pdfError};
 }
}
window.nineDocumentArchive={archiveTransaction,filename,folderPath};
window.nineRetryArchive=async function(docNo){
 const tx=transactions.find(t=>t.docNo===docNo);
 if(!tx)return showToast('ไม่พบข้อมูลเอกสารในเบราว์เซอร์นี้','error');
 if(!confirm('ต้องการบันทึก PDF เลขที่ '+docNo+' ลง SharePoint อีกครั้งหรือไม่? (ไม่ทำรายการเบิก–คืนซ้ำ)'))return;
 return archiveTransaction(tx);
};
})();
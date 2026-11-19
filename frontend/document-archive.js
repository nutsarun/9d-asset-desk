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
 const container=document.createElement('div');
 container.setAttribute('aria-hidden','true');container.className='ho-pdf-capture';
 container.style.cssText='position:fixed;left:-12000px;top:0;width:186mm;background:#fff;color:#1e293b;z-index:-10;padding:0;margin:0;';
 container.innerHTML=window.nineRenderHandoverPaper(tx);
 document.body.appendChild(container);
 try{
  await document.fonts?.ready;
  await Promise.all([...container.querySelectorAll('img')].map(img=>new Promise((resolve,reject)=>{
   if(img.complete){if(img.naturalWidth)resolve();else reject(Error('โหลดโลโก้ใน PDF ไม่สำเร็จ'));return}
   img.onload=resolve;img.onerror=()=>reject(Error('โหลดโลโก้ใน PDF ไม่สำเร็จ'));
  })));
  const result=await lib().set({
   margin:[12,12,12,12],
   image:{type:'jpeg',quality:0.98},
   html2canvas:{scale:2,useCORS:true,backgroundColor:'#ffffff',scrollX:0,scrollY:0,logging:false},
   jsPDF:{unit:'mm',format:'a4',orientation:'portrait'},
   pagebreak:{mode:['css','legacy'],before:'.ho-page:not(:first-child)'}
  }).from(container).outputPdf('blob');
  if(!(result instanceof Blob)||result.size<1000)throw Error('สร้างไฟล์ PDF ไม่สมบูรณ์');
  return result;
 }finally{container.remove()}
}
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
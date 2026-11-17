import {test} from 'node:test';
import assert from 'node:assert/strict';
import {createStore} from '../src/store.js';
import {schemas} from '../src/schema.js';

function mockStore(){
  const calls=[];
  let row;
  const store=createStore({siteId:'tenant.sharepoint.com,collection,web',lists:{assets:'assets-list',people:'people-list'}},
    {getToken:async()=>({token:'test-token'})},async(url,opts)=>{
      calls.push([url.pathname,opts.method,opts.headers]);
      const body=opts.body?JSON.parse(opts.body):undefined;
      if (opts.method==='POST') {
        row={id:'1', eTag:'"1"', fields:body.fields};
        return Response.json({id:'1'});
      }
      if (opts.method==='PATCH'){
        if (opts.headers['If-Match']!==row.eTag) return Response.json({}, {status:412});
        row={...row,eTag:JSON.stringify(String(Number(JSON.parse(row.eTag))+1)),fields:body};
        return Response.json(body);
      }
      if (url.pathname.endsWith('/1')) return Response.json(row);
      return Response.json({value:url.pathname.includes('assets-list')?[row]:[]});
    });
  return {store,calls};
}

test('saves to the mapped lists and preserves history and eTag conflicts',async()=>{
  const now='2026-11-17T09:00:00.000Z';
  const record={id:crypto.randomUUID(),code:'9D-01',model:'Test',category:'อื่น ๆ',status:'available',updatedAt:now,history:[{at:now,label:'เพิ่มทรัพย์สิน'}]};
  const {store,calls}=mockStore();
  const created=await store.save('assets',record);
  const updated=await store.save('assets',{...created,status:'repair',history:[{at:now,label:'attempted overwrite'},{at:now,label:'ส่งซ่อม'}]});
  assert.equal(updated.status,'repair');
  assert.equal(updated.history[0].label,'เพิ่มทรัพย์สิน');
  assert.equal(updated.etag,'"2"');
  await assert.rejects(store.save('assets',{...created,history:updated.history}),error=>error.status===412);
  assert.deepEqual((await store.read()).assets,[updated]);
  assert.ok(calls.every(([, ,headers])=>headers.Authorization==='Bearer test-token'));
});

test('changes holders with a new document and preserves the previous assignment history',async()=>{
  const {store,calls}=mockStore();
  const now='2026-11-17T09:00:00.000Z';
  const asset={code:'9D-01',model:'Laptop',category:'โน้ตบุ๊ก'};
  const former={id:crypto.randomUUID(),name:'ผู้รับเดิม'};
  const recipient={id:crypto.randomUUID(),name:'ผู้รับใหม่'};
  const previousDoc={id:'HO-OLD',asset,person:former,date:'2026-11-16'};
  const original=await store.save('assets',{
    ...asset,id:crypto.randomUUID(),status:'assigned',assignment:previousDoc,updatedAt:now,
    history:[{at:now,label:'ส่งมอบ',doc:previousDoc}]
  });
  const doc={...previousDoc,id:'HO-NEW',person:recipient,date:'2026-11-17',previousHolder:former};
  const transfer=next=>schemas.assets.parse({
    ...original,assignment:next,
    history:[...original.history,{at:now,label:'เปลี่ยนผู้ใช้งาน',doc:next}]
  });

  await assert.rejects(store.save('assets',transfer({...doc,previousHolder:recipient})),{status:400});
  await assert.rejects(store.save('assets',transfer({...doc,date:'2026-11-15'})),{status:400});
  const incorrectHistory=transfer(doc);
  incorrectHistory.history[1].doc=previousDoc;
  await assert.rejects(store.save('assets',incorrectHistory),{status:400});
  assert.equal(calls.filter(([,method])=>method==='PATCH').length,0);

  const saved=await store.save('assets',transfer(doc));
  assert.equal(saved.status,'assigned');
  assert.deepEqual(saved.assignment,doc);
  assert.deepEqual(saved.history.map(event=>event.doc),[previousDoc,doc]);
  assert.deepEqual((await store.read()).assets,[saved]);
  await assert.rejects(store.save('assets',transfer(doc)),{status:412});
});

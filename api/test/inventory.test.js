import {test} from 'node:test';
import assert from 'node:assert/strict';
import {createStore} from '../src/store.js';
import {schemas} from '../src/schema.js';

test('reads both native registries and writes handovers back to the exact source item',async()=>{
  const modified='2026-11-17T09:00:00.000Z';
  const specimen=(fields)=>({id:'8',eTag:'"1"',fields:{Modified:modified,...fields}});
  const rows={
    pc:specimen({Title:'PC-08',Assetorserialnumber:'BOOK-8',serialnumber:'S-08',Model:'Notebook',ItemType:'Notebook',
      Status:'In -Use',_x0e4a_UserName:'เดิม',Company:'9Distribution',OldUser:'ชื่อก่อนหน้า',
      ReturnedDate:'2026-01-01T08:00:00Z',Warranty:'field outside the app'}),
    acc:specimen({Title:'หมายเหตุ',Owner:'After Nine',Model:'K120',ItemType:'Keyboard',
      Status:'Borrowed',UserName0:'ผู้ยืม',ReceivedDate:'2026-11-15T17:00:00Z'}),
  };
  const writes=[];
  const store=createStore({siteId:'tenant.sharepoint.com,collection,web',lists:{people:'staff'},inventory:{computer:'pc',accessories:'acc'}},
    {getToken:async()=>({token:'test-token'})},async(url,options)=>{
      const [,list,identity]=url.pathname.match(/\/lists\/([^/]+)\/items\/?([^/]*)/);
      if (options.method==='PATCH') {
        if (options.headers['If-Match']!==rows[list].eTag) return Response.json({}, {status:412});
        const fields=JSON.parse(options.body);
        writes.push({list,fields});
        rows[list].fields={...rows[list].fields,...fields};
        rows[list].eTag=JSON.stringify(String(Number(JSON.parse(rows[list].eTag))+1));
        return Response.json(fields);
      }
      if (identity) return Response.json(rows[list]);
      return Response.json({value:rows[list]?[rows[list]]:[]});
    });

  const {assets:[pc,acc],people}=await store.read();
  assert.notEqual(pc.id,acc.id);
  assert.deepEqual([pc.status,acc.status],['assigned','loaned']);
  assert.equal(pc.assignment.date,'');
  assert.equal(acc.assignment.date,'2026-11-16');
  assert.equal(acc.code,'ACC#8');
  assert.equal(acc.note,'หมายเหตุ');
  assert.equal(pc.serial,'S-08');
  assert.equal(people.length,2);

  const doc={id:'HO-SAMPLE',date:'2026-11-17',asset:{code:pc.code,model:pc.model,category:pc.category},
    person:acc.assignment.person,previousHolder:pc.assignment.person};
  const next=schemas.assets.parse({...pc,assignment:doc,status:'loaned',history:[{at:modified,label:'เปลี่ยนผู้ใช้งาน',doc}]});
  const saved=await store.save('assets',next);
  assert.deepEqual(writes[0],{list:'pc',fields:{
    AssetDeskPayload:JSON.stringify({id:next.id,site:next.site,history:next.history,updatedAt:next.updatedAt}),
    Status:'Borrowed',_x0e4a_UserName:'ผู้ยืม',ReceivedDate:'2026-11-16T17:00:00.000Z',
    ReturnedDate:null,OldUser:'ชื่อก่อนหน้า / เดิม'
  }});
  assert.deepEqual(saved.assignment,doc);
  assert.equal(rows.pc.fields.Warranty,'field outside the app');
  assert.equal(rows.acc.eTag,'"1"');
  await assert.rejects(store.save('assets',next),{status:412});

  const reclassified=await store.save('assets',schemas.assets.parse({...saved,status:'assigned',
    history:[...saved.history,{at:modified,label:'กำลังใช้งาน'}]}));
  assert.equal(writes[1].fields.Status,'In -Use');
  assert.equal(writes[1].fields._x0e4a_UserName,undefined);
  assert.deepEqual(reclassified.assignment,doc);

  const returned=await store.save('assets',schemas.assets.parse({...reclassified,status:'available',assignment:null,
    history:[...reclassified.history,{at:modified,label:'รับคืน'}]}));
  assert.equal(writes[2].fields._x0e4a_UserName,null);
  assert.equal(returned.returnedDate,'2026-11-17');
  assert.deepEqual(returned.history[0].doc,doc);
});

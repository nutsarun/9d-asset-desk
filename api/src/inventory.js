import { createHash } from 'node:crypto';

export const inventoryColumns = {
  computer: { host:'Title', assetNumber:'Assetorserialnumber', holder:'_x0e4a_UserName', company:'Company', remark:'Remark', assigned:'In -Use' },
  accessories: { host:'HostName', assetNumber:'AssetorserialNumber', holder:'UserName0', company:'Owner', remark:'Title', assigned:'In Use' }
};

const nativeStatuses = {
  available:'Stock', loaned:'Borrowed', unavailable:'Not available', lost:'Lost', repair:'Repair', retired:'Retired'
};
const statusByValue = new Map([...Object.entries(nativeStatuses).map(([key,value])=>[value.toLowerCase(),key]), ['in use','assigned'], ['in -use','assigned']]);
const bangkokDate = new Intl.DateTimeFormat('en-CA', {timeZone:'Asia/Bangkok',year:'numeric',month:'2-digit',day:'2-digit'});
const activeStatus = status => status==='assigned'||status==='loaned';
const clean = value => typeof value==='string' ? value.trim() : '';
const normalizedName = name => name.trim().replace(/\s+/g,' ').toLowerCase();
const dateOnly = value => {
  if (!value) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return value;
  const date = new Date(value);
  return Number.isFinite(date.valueOf()) ? bangkokDate.format(date) : '';
};
const dateField = value => value ? new Date(value+'T00:00:00+07:00').toISOString() : null;

function stableUuid(key) {
  const hex=createHash('sha256').update(key).digest('hex');
  return `${hex.slice(0,8)}-${hex.slice(8,12)}-5${hex.slice(13,16)}-8${hex.slice(17,20)}-${hex.slice(20,32)}`;
}

export function decodeInventory(entry, source) {
  const columns=inventoryColumns[source];
  const f=entry.fields;
  const spId=Number(entry.id);
  const metadata=f.AssetDeskPayload ? JSON.parse(f.AssetDeskPayload) : {};
  const history=metadata.history ?? [];
  const hostName=clean(f[columns.host]);
  const assetNumber=clean(f[columns.assetNumber]);
  const equipment={
    code:hostName || assetNumber || `${source==='computer'?'PC':'ACC'}#${spId}`,
    model:clean(f.Model) || 'ไม่ระบุรุ่น',
    category:clean(f.ItemType) || 'ไม่ระบุประเภท',
    brand:clean(f.Vendor), serial:source==='computer'?clean(f.serialnumber):'',
    site:metadata.site ?? ''
  };
  const id=metadata.id ?? stableUuid(`${source}:${spId}`);
  const holder=clean(f[columns.holder]);
  const status=statusByValue.get(clean(f.Status).toLowerCase()) ?? 'unknown';
  const receivedDate=dateOnly(f.ReceivedDate);
  const returnedDate=dateOnly(f.ReturnedDate);
  const active=activeStatus(status);
  const matchingDocument=history.findLast(event=>event.doc)?.doc;
  const assignment=active&&holder
    ? (matchingDocument?.date===receivedDate&&normalizedName(matchingDocument.person.name)===normalizedName(holder)
      ? matchingDocument
      : {id:'existing-'+id,person:{id:stableUuid('person:'+normalizedName(holder)),name:holder},asset:equipment,date:receivedDate})
    : null;
  return {
    ...equipment,id,spId,etag:entry.eTag ?? f["@odata.etag"],inventorySource:source,hostName,assetNumber,status,assignment,history,
    company:clean(f[columns.company]),previousUsers:clean(f.OldUser),receivedDate,returnedDate,
    currentHolder:holder,note:clean(f[columns.remark]),updatedAt:f.Modified ?? metadata.updatedAt
  };
}

export function inventoryFields(payload, previous) {
  const source=payload.inventorySource;
  const columns=inventoryColumns[source];
  const fields={AssetDeskPayload:JSON.stringify({id:payload.id,site:payload.site,history:payload.history,updatedAt:payload.updatedAt})};
  if (!previous || payload.code!==previous.code) fields[columns.host]=payload.code;
  if (!previous || payload.company!==previous.company) fields[columns.company]=payload.company || null;
  if (!previous || payload.assetNumber!==previous.assetNumber) fields[columns.assetNumber]=payload.assetNumber || null;
  if (!previous || payload.model!==previous.model) fields.Model=payload.model;
  if (!previous || payload.category!==previous.category) fields.ItemType=payload.category;
  if (!previous || payload.brand!==previous.brand) fields.Vendor=payload.brand || null;
  if (source==='computer'&&(!previous || payload.serial!==previous.serial)) fields.serialnumber=payload.serial || null;
  if (!previous || payload.note!==previous.note) fields[columns.remark]=payload.note || (source==='accessories'?'-':null);
  if (!previous || payload.status!==previous.status) fields.Status=payload.status==='assigned'?columns.assigned:nativeStatuses[payload.status];

  const current=previous?.assignment;
  const next=payload.assignment;
  if (next&&next.id!==current?.id) {
    fields[columns.holder]=next.person.name;
    fields.ReceivedDate=dateField(next.date);
    fields.ReturnedDate=null;
  } else if (previous&&activeStatus(previous.status)&&!next) {
    fields[columns.holder]=null;
    if (['available','repair','unavailable'].includes(payload.status)) fields.ReturnedDate=dateField(dateOnly(payload.updatedAt));
  }
  if (current&&current.person.id!==next?.person.id) {
    const oldUsers=[previous.previousUsers,current.person.name].filter(Boolean).join(' / ');
    if (oldUsers.length<=255) fields.OldUser=oldUsers;
  }
  return fields;
}

export function peopleFromInventory(assets, savedPeople) {
  const people=new Map(savedPeople.map(person=>[person.id,person]));
  for (const asset of assets) {
    const person=asset.assignment?.person;
    if (person&&!people.has(person.id)) people.set(person.id,{id:person.id,name:person.name,updatedAt:asset.updatedAt});
  }
  return [...people.values()];
}

import { isDeepStrictEqual } from 'node:util';
import { ManagedIdentityCredential } from '@azure/identity';
import { statuses } from './schema.js';
import { decodeInventory, inventoryFields, peopleFromInventory } from './inventory.js';

export class StoreError extends Error {
  constructor(status, message) { super(message); this.status = status; }
}

const graphRoot = 'https://graph.microsoft.com/v1.0/';
const graphErrors = {
  409:'ข้อมูลซ้ำ กรุณารีเฟรชแล้วตรวจสอบ',
  412:'มีผู้แก้ไขรายการนี้ กรุณารีเฟรชก่อนบันทึกใหม่',
  429:'SharePoint จำกัดคำขอชั่วคราว กรุณาลองใหม่ภายหลัง'
};

export function createStore(config, credential = new ManagedIdentityCredential(), send = fetch) {
  function collection(kind, source) {
    return 'sites/' + encodeURIComponent(config.siteId) + '/lists/' + (source ? config.inventory[source] : config.lists[kind]) + '/items';
  }

  async function request(path, options = {}) {
    const url = new URL(path, graphRoot);
    if (!url.href.startsWith(graphRoot)) throw new StoreError(502,'ที่อยู่ข้อมูล SharePoint ไม่ถูกต้อง');
    const token = await credential.getToken('https://graph.microsoft.com/.default');
    if (!token) throw new StoreError(502,'API เชื่อมต่อ SharePoint ไม่สำเร็จ');
    const response = await send(url, {
      ...options, redirect:'error', signal:AbortSignal.timeout(18000),
      headers: {...options.headers, Accept:'application/json', Authorization:'Bearer '+token.token}
    });
    if (!response.ok) {
      if (response.status === 400) {
        const error = await response.json().catch(() => null);
        if (/duplicate|unique|-2130575169/i.test(error?.error?.message ?? '')) throw new StoreError(409,graphErrors[409]);
      }
      const message = graphErrors[response.status];
      throw new StoreError(message ? response.status : 502, message ?? 'SharePoint ดำเนินการไม่สำเร็จ โปรดตรวจการตั้งค่า API และรายการ');
    }
    return response.status === 204 ? undefined : response.json();
  }

  function decode(entry, source) {
    try {
      const payload = source ? decodeInventory(entry, source) : JSON.parse(entry.fields.AssetDeskPayload);
      const spId = Number(entry.id);
      if (!payload.id || !Number.isSafeInteger(spId) || spId<1 || !(entry.eTag ?? payload.etag)) throw Error();
      return {...payload, spId, etag:entry.eTag ?? payload.etag};
    } catch { throw new StoreError(502,'ข้อมูลในรายการ SharePoint ไม่สมบูรณ์'); }
  }

  const expandFor = source => '?$select=id,eTag&$expand='+(source?'fields':'fields($select=AssetDeskPayload)');

  async function readList(kind, source) {
    let path = collection(kind,source) + expandFor(source) + '&$top=200';
    const records = [];
    while (path) {
      const result = await request(path);
      records.push(...result.value.map(entry=>decode(entry,source)));
      path = result['@odata.nextLink'];
    }
    return records;
  }

  async function save(kind, record) {
    const {spId, etag, ...payload} = record;
    const source = kind==='assets' && config.inventory ? payload.inventorySource : undefined;
    if (kind==='assets'&&config.inventory&&!Object.hasOwn(config.inventory,source)) throw new StoreError(400,'กรุณาเลือกทะเบียนทรัพย์สิน');
    const parent = collection(kind,source);
    const path = spId ? parent + '/' + spId : parent;
    const expand = expandFor(source);
    let previous;
    if (spId) {
      if (!etag || etag==='*') throw new StoreError(400,'กรุณารีเฟรชก่อนแก้ไขรายการ');
      previous = decode(await request(path + expand),source);
      if (previous.id !== payload.id || previous.etag !== etag) throw new StoreError(412,graphErrors[412]);
      if (kind === 'assets') {
        const last = payload.history.at(-1);
        if (!last || payload.history.length !== previous.history.length + 1) throw new StoreError(400,'ประวัติการเปลี่ยนแปลงไม่สมบูรณ์');
        const current = previous.assignment;
        const next = payload.assignment;
        if (source && next && next.id!==current?.id && (!next.date || !isDeepStrictEqual(next,last.doc))) {
          throw new StoreError(400,'กรุณาระบุวันที่รับสำหรับใบรับมอบใหม่');
        }
        if (current && next && current.person.id !== next.person.id) {
          if (next.previousHolder?.id !== current.person.id ||
              next.previousHolder.name !== current.person.name ||
              next.date < current.date || next.id === current.id ||
              previous.history.some(event => event.doc?.id === next.id) ||
              !isDeepStrictEqual(next, last.doc)) {
            throw new StoreError(400,'ข้อมูลการเปลี่ยนผู้ใช้งานไม่สมบูรณ์');
          }
        }
        payload.history = [...previous.history,last];
      }
    } else if (etag) throw new StoreError(400,'ข้อมูลรายการใหม่ไม่ถูกต้อง');
    if (source && ((!previous && payload.status!=='available') ||
        (payload.status==='unknown' && previous?.status!=='unknown') || (payload.note?.length??0)>255)) {
      throw new StoreError(400,'สถานะหรือหมายเหตุไม่เหมาะกับทะเบียนต้นทาง');
    }
    const serialized = JSON.stringify(payload);
    if (serialized.length > 62000) throw new StoreError(400,'ข้อมูลประวัติของรายการนี้เต็ม');
    const fields = source ? inventoryFields(payload,previous) : {
      Title:(kind==='assets' ? [payload.code,payload.model].join(' · ') : payload.name).slice(0,255),
      AssetDeskKey:payload.id,
      AssetDeskIdentity:(payload.code || payload.employeeCode || payload.id).toLowerCase(),
      AssetDeskPayload:serialized
    };
    if (kind==='assets'&&!source) Object.assign(fields,{
      AssetDeskCategory:payload.category,
      AssetDeskStatus:statuses[payload.status],
      AssetDeskHolder:payload.assignment?.person.name ?? ''
    });
    const headers = {'Content-Type':'application/json'};
    if (spId) {
      await request(path + '/fields',{method:'PATCH',headers:{...headers,'If-Match':etag},body:JSON.stringify(fields)});
      return decode(await request(path + expand),source);
    }
    const created = await request(parent,{method:'POST',headers,body:JSON.stringify({fields})});
    return decode(await request(parent + '/' + encodeURIComponent(created.id) + expand),source);
  }

  return { async read() {
    if (!config.inventory) {
      const [assets,people] = await Promise.all([readList('assets'),readList('people')]);
      return {assets,people};
    }
    const sources = Object.keys(config.inventory);
    const [people,...registries] = await Promise.all([readList('people'),...sources.map(source=>readList('assets',source))]);
    const assets = registries.flat();
    return {assets,people:peopleFromInventory(assets,people),inventorySources:sources};
  }, save };
}

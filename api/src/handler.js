import { schemas } from './schema.js';
import { StoreError } from './store.js';

function reply(status, jsonBody) {
  return {status,jsonBody,headers:{'Cache-Control':'private, no-store','X-Content-Type-Options':'nosniff'}};
}

export function createHandler({ authorize, store, origin }) {
  return async function handler(request, context, route) {
    const requestOrigin = request.headers.get('origin');
    if (requestOrigin && requestOrigin!==origin) return reply(403,{message:'แหล่งที่มาของคำขอไม่ถูกต้อง'});
    const status = await authorize(request.headers.get('authorization'));
    if (status!==200) return reply(status,{message:status===403?'ต้องมีสิทธิ์ Admin ของ 9D Asset Desk':'กรุณาเข้าสู่ระบบ'});
    try {
      if (route === 'session') return reply(200,{role:'Admin'});
      if (route === 'records') return reply(200,await store.read());
      const { kind } = request.params;
      if (!Object.hasOwn(schemas,kind)) return reply(404,{message:'ไม่พบรายการ'});
      if (!/^application\/json(?:;|$)/i.test(request.headers.get('content-type') ?? '')) return reply(415,{message:'รูปแบบข้อมูลไม่ถูกต้อง'});
      const raw = await request.text();
      if (Buffer.byteLength(raw)>128000) return reply(413,{message:'ข้อมูลมีขนาดเกินกำหนด'});
      let body;
      try { body = JSON.parse(raw); } catch { return reply(400,{message:'รูปแบบข้อมูลไม่ถูกต้อง'}); }
      const parsed = schemas[kind].safeParse(body);
      if (!parsed.success) return reply(400,{message:'กรุณาตรวจรายละเอียดก่อนบันทึก'});
      return reply(200,await store.save(kind,parsed.data));
    } catch (error) {
      if (error instanceof StoreError) return reply(error.status,{message:error.message});
      context.error('Asset Desk storage request failed');
      return reply(502,{message:'API เชื่อมต่อ SharePoint ไม่สำเร็จ'});
    }
  };
}

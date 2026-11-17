import { z } from 'zod';

export const statuses = { available:'พร้อมใช้งาน', assigned:'กำลังใช้งาน', loaned:'ให้ยืม', unavailable:'ไม่พร้อมใช้งาน', lost:'สูญหาย', unknown:'ยังไม่ได้ระบุ', repair:'ส่งซ่อม', retired:'ปลดระวาง' };
const text = z.string().max(120);
const note = z.string().max(500);
const status = z.enum(Object.keys(statuses));
const date = z.iso.date();
const optionalDate = z.union([date,z.literal('')]);
const peopleFields = {
  id: z.uuid(), name: text.trim().min(1), employeeCode: text.optional(), department: text.optional(), role: text.optional()
};
const equipmentFields = {
  code: text.trim().min(1), model: text.trim().min(1), category: text.trim().min(1),
  brand: text.optional(), serial: text.optional(), site: text.optional()
};
const row = {
  spId: z.number().int().positive().optional(), etag: z.string().min(1).max(160).optional(),
  updatedAt: z.iso.datetime()
};
const handover = z.object({
  id: z.string().max(64), person: z.object(peopleFields), asset: z.object(equipmentFields),
  previousHolder: z.object({id:peopleFields.id,name:peopleFields.name}).optional(),
  date:optionalDate, due:optionalDate.optional(), accessories: note.optional(), condition: note.optional()
}).refine(doc => !doc.due || doc.due >= doc.date);
const event = z.object({at:z.iso.datetime(),label:z.string().max(200),note:note.optional(),doc:handover.optional()});
export const schemas = {
  people: z.object({...row,...peopleFields}),
  assets: z.object({
    ...row, ...equipmentFields, id:z.uuid(), status, note:note.optional(),
    inventorySource:z.enum(['computer','accessories']).optional(), assetNumber:text.optional(), company:text.optional(),
    assignment:handover.nullish(), history:z.array(event).max(1000)
  }).refine(asset => (asset.status==='assigned'||asset.status==='loaned') === Boolean(asset.assignment))
};

import { app } from '@azure/functions';
import { createAuthorizer } from './auth.js';
import { createHandler } from './handler.js';
import { createStore } from './store.js';

const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const tenant = process.env.ENTRA_TENANT_ID;
const audience = process.env.ENTRA_API_CLIENT_ID;
const clientId = process.env.ENTRA_SPA_CLIENT_ID;
const siteId = process.env.SHAREPOINT_SITE_ID;
const origin = process.env.PAGES_ORIGIN;
const lists = {assets:process.env.SHAREPOINT_ASSETS_LIST_ID,people:process.env.SHAREPOINT_EMPLOYEES_LIST_ID};
const inventory = {computer:process.env.SHAREPOINT_COMPUTER_LIST_ID,accessories:process.env.SHAREPOINT_ACCESSORIES_LIST_ID};
const native = Boolean(inventory.computer || inventory.accessories);
const assetLists = native ? Object.values(inventory) : [lists.assets];
const siteParts = siteId?.split(',') ?? [];
const valid = [tenant,audience,clientId,lists.people,...assetLists,...siteParts.slice(1)].every(value=>guid.test(value ?? '')) &&
  new Set([...assetLists,lists.people].map(value=>value.toLowerCase())).size===assetLists.length+1 &&
  siteParts.length===3 && /^[a-z0-9-]+\.sharepoint\.com$/i.test(siteParts[0]) &&
  /^https:\/\/[a-z0-9.-]+$/i.test(origin ?? '');
const handler = valid
  ? createHandler({authorize:createAuthorizer({tenant,audience,clientId}),store:createStore({siteId,lists,inventory:native?inventory:null}),origin})
  : async () => ({status:503,headers:{'Cache-Control':'private, no-store'},jsonBody:{message:'ผู้ดูแลยังไม่ได้ตั้งค่า API'}});

for (const [suffix, methods, route] of [
  ['Session',['GET'],'session'],['Read',['GET'],'records'],['Write',['POST'],'records/{kind}']
]) {
  app.http('assetDesk'+suffix,{
    methods, route, authLevel:'anonymous',
    handler:(request, context) => handler(request,context,route)
  });
}

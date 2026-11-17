import { test } from 'node:test';
import assert from 'node:assert/strict';
import { generateKeyPair, SignJWT } from 'jose';
import { createAuthorizer } from '../src/auth.js';
import { createHandler } from '../src/handler.js';

const config={tenant:'tenant-test',audience:'audience-test',clientId:'client-test'};
const issuer='https://login.microsoftonline.com/tenant-test/v2.0';
const {publicKey,privateKey}=await generateKeyPair('RS256');
const {privateKey:otherKey}=await generateKeyPair('RS256');
const authorize=createAuthorizer(config,publicKey);
const claims={iss:issuer,aud:config.audience,tid:config.tenant,azp:config.clientId,oid:'test-account',ver:'2.0',roles:['Admin'],scp:'AssetDesk.Access',acct:0};

async function token(overrides={},key=privateKey) {
  return 'Bearer '+await new SignJWT({...claims,...overrides})
    .setProtectedHeader({alg:'RS256'})
    .setIssuedAt()
    .setExpirationTime(overrides.exp ?? '5m')
    .sign(key);
}

test('accepts only a valid member Admin access token for the app',async()=>{
  assert.equal(await authorize(await token()),200);
  for (const overrides of [
    {roles:['Reader']}, {roles:[]}, {acct:1}, {acct:undefined},
    {scp:'User.Read'}, {scp:''}
  ]) assert.equal(await authorize(await token(overrides)),403);
  for (const overrides of [
    {iss:'https://login.microsoftonline.com/other/v2.0'},
    {aud:'https://graph.microsoft.com'}, {tid:'other'}, {azp:'other'}, {ver:'1.0'},
    {exp:Math.floor(Date.now()/1000)-20}
  ]) assert.equal(await authorize(await token(overrides)),401);
  assert.equal(await authorize(await token({},otherKey)),401);
  assert.equal(await authorize(null),401);
});

test('guards every data route before touching SharePoint',async()=>{
  const calls=[];
  const origin='https://example.github.io';
  const handler=createHandler({authorize,origin,store:{
    async read(){calls.push('read');return {assets:[],people:[]};},
    async save(kind,record){calls.push(kind);return record;}
  }});
  const nonAdmin=await token({roles:['Reader']});
  const admin=await token();
  const context={error(){throw Error('Unexpected storage failure');}};
  const req=(authorization,originOverride=origin,body={})=>({
    headers:new Headers({origin:originOverride,...(authorization?{authorization}:{}),'content-type':'application/json'}),
    params:{kind:'people'},
    text:async()=>JSON.stringify(body)
  });
  assert.equal((await handler(req(),context,'session')).status,401);
  for(const path of ['session','records','records/{kind}']) assert.equal((await handler(req(nonAdmin),context,path)).status,403);
  assert.equal((await handler(req(admin,'https://unrelated.example'),context,'records')).status,403);
  assert.deepEqual(calls,[]);
  assert.equal((await handler(req(admin),context,'session')).status,200);
  assert.equal((await handler(req(admin),context,'records')).status,200);
  assert.equal((await handler(req(admin),context,'records/{kind}')).status,400);
  assert.equal((await handler(req(admin,origin,{id:crypto.randomUUID(),name:'Test employee',updatedAt:new Date().toISOString()}),context,'records/{kind}')).status,200);
  assert.deepEqual(calls,['read','people']);
});

import { createRemoteJWKSet, jwtVerify } from 'jose';

export function createAuthorizer(config, jwks) {
  const authority = `https://login.microsoftonline.com/${config.tenant}`;
  const issuer = authority + '/v2.0';
  const keys = jwks ?? createRemoteJWKSet(new URL(authority + '/discovery/v2.0/keys'));
  return async function authorize(header) {
    const match = /^Bearer ([A-Za-z0-9._-]{1,12000})$/i.exec(header ?? '');
    if (!match) return 401;
    let payload;
    try {
      ({payload} = await jwtVerify(match[1], keys, {
        issuer, audience: config.audience, algorithms: ['RS256'],
        requiredClaims: ['exp','iat','tid','azp','oid','ver']
      }));
    } catch { return 401; }
    const roles = Array.isArray(payload.roles) ? payload.roles : [];
    const scopes = typeof payload.scp === 'string' ? payload.scp.split(' ') : [];
    if (payload.tid !== config.tenant || payload.azp !== config.clientId || payload.ver !== '2.0') return 401;
    if (!scopes.includes('AssetDesk.Access') || !roles.includes('Admin') || ![0,'0'].includes(payload.acct)) return 403;
    return 200;
  };
}

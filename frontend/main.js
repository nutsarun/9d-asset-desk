import { InteractionRequiredAuthError, PublicClientApplication } from '@azure/msal-browser';
import { mountAssetDesk } from './desk.js';
import './style.css';

const gate = document.querySelector('#gate');
const root = document.querySelector('#asset-app');
const { VITE_ENTRA_TENANT_ID: tenant, VITE_ENTRA_SPA_CLIENT_ID: clientId,
  VITE_ENTRA_API_CLIENT_ID: audience, VITE_ASSET_API_URL: api } = import.meta.env;
const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const redirectUri = new URL(import.meta.env.BASE_URL, location.href).href;
const configured = [tenant, clientId, audience].every(value => guid.test(value ?? '')) && /^https:\/\/[^?#]+\/api\/?$/.test(api ?? '');

function screen(message, action, label = 'เข้าสู่ระบบด้วย Microsoft', error = false) {
  root.hidden = true;
  gate.hidden = false;
  gate.innerHTML = '<main class="login-page"><section class="login-panel"><div class="login-logo">9D</div><h1>Asset Desk</h1><p class="login-caption">สำหรับผู้ดูแลทะเบียนทรัพย์สิน Nine Distribution</p><p role="status"></p><button type="button"></button></section></main>';
  const state = gate.querySelector('[role=status]');
  state.textContent = message;
  if (error) state.className = 'login-error';
  const button = gate.querySelector('button');
  button.textContent = label;
  button.hidden = !action;
  if (action) button.onclick = async () => {
    button.disabled = true;
    try { await action(); } catch { state.textContent = 'ดำเนินการไม่สำเร็จ กรุณาลองอีกครั้ง'; button.disabled = false; }
  };
}

class ApiError extends Error {
  constructor(status, message) { super(message); this.status = status; }
}

async function start() {
  if (!configured) {
    screen('ผู้ดูแลต้องตั้งค่า Microsoft Entra และ API ก่อนเปิดใช้งาน', null);
    return;
  }
  screen('กำลังตรวจสอบบัญชี…', null);
  const msal = new PublicClientApplication({
    auth: { clientId, authority: `https://login.microsoftonline.com/${tenant}`, redirectUri, postLogoutRedirectUri: redirectUri, navigateToLoginRequestUrl: false },
    cache: { cacheLocation: 'sessionStorage' }
  });
  const scopes = [`api://${audience}/AssetDesk.Access`];
  const login = () => msal.loginRedirect({ scopes, prompt: 'select_account' });
  let account;
  const logout = () => msal.logoutRedirect({ account });
  let unmount;
  const signedOut = message => {
    unmount?.();
    root.replaceChildren();
    screen(message, login);
  };
  async function call(path, options = {}) {
    let auth;
    try { auth = await msal.acquireTokenSilent({ account, scopes }); }
    catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        signedOut('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
        throw new ApiError(401, 'กรุณาเข้าสู่ระบบใหม่');
      }
      throw error;
    }
    const response = await fetch(api.replace(/\/$/, '') + path, {
      ...options, credentials: 'omit', cache: 'no-store',
      headers: { Authorization: `Bearer ${auth.accessToken}`, ...(options.body ? {'Content-Type':'application/json'} : {}) }
    });
    if (response.status === 401 || response.status === 403) {
      signedOut(response.status === 403 ? 'บัญชีนี้ไม่มีสิทธิ์ Admin ของ 9D Asset Desk' : 'กรุณาเข้าสู่ระบบใหม่');
      throw new ApiError(response.status, 'ไม่มีสิทธิ์เข้าใช้');
    }
    const data = await response.json();
    if (!response.ok) throw new ApiError(response.status, data.message || 'ทำรายการไม่สำเร็จ');
    return data;
  }
  try {
    await msal.initialize();
    const redirect = await msal.handleRedirectPromise();
    account = redirect?.account ?? msal.getActiveAccount() ?? msal.getAllAccounts().find(candidate => candidate.tenantId === tenant);
    if (!account) return signedOut('ลงชื่อเข้าใช้ด้วยบัญชี Microsoft 365 ของบริษัท');
    if (account.tenantId !== tenant) return signedOut('กรุณาเลือกบัญชีของ Nine Distribution');
    msal.setActiveAccount(account);
    await call('/session');

    gate.hidden = true;
    root.hidden = false;
    const toolbar = document.createElement('div');
    toolbar.className = 'login-toolbar';
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'login-signout';
    button.textContent = 'ออกจากระบบ';
    button.onclick = async () => {
      button.disabled = true;
      unmount?.();
      try { await logout(); } catch { signedOut('ออกจากระบบไม่สำเร็จ กรุณาปิดแท็บนี้'); }
    };
    toolbar.append(button);
    const desk = document.createElement('div');
    root.replaceChildren(toolbar, desk);
    unmount = mountAssetDesk(desk, {
      mode: 'sharepoint',
      read: () => call('/records'),
      save: (kind, record) => call('/records/' + kind, { method: 'POST', body: JSON.stringify(record) })
    });
  } catch (error) {
    if (error instanceof ApiError && [401,403].includes(error.status)) return;
    const gated = ['access_denied','interaction_required'].includes(error.errorCode) || /AADSTS50105/.test(error.message ?? '');
    signedOut(gated ? 'บัญชีนี้ยังไม่ได้รับสิทธิ์ Admin ของ 9D Asset Desk' : 'เข้าสู่ระบบไม่สำเร็จ กรุณาลองใหม่หรือติดต่อผู้ดูแล');
  }
}
start();

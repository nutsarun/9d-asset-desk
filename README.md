# 9D Asset Desk · GitHub Pages

เว็บทะเบียนทรัพย์สินของ Nine Distribution รองรับการเพิ่มอุปกรณ์และพนักงาน การส่งมอบ เปลี่ยนผู้ใช้งาน และรับคืน การติดตามสถานะ และการพิมพ์ใบรับมอบ A4

GitHub Pages ให้บริการหน้าเว็บ ส่วน Azure Functions ตรวจ access token และสิทธิ์ Admin ทุกครั้งก่อนเรียก SharePoint ผ่าน Microsoft Graph ข้อมูลอุปกรณ์และพนักงานเก็บใน SharePoint Lists

## เปลี่ยนผู้ใช้งานทรัพย์สิน

เปิดเมนู **เปลี่ยนผู้ใช้งาน** หรือเปิด **รายละเอียด** ของทรัพย์สินที่อยู่กับพนักงาน แล้วกด **เปลี่ยนผู้ใช้งาน** เลือกผู้รับคนใหม่ ระบุวันที่รับ และกรอกกำหนดคืนของผู้รับคนใหม่หากมี เมื่อบันทึกแล้วระบบอัปเดตผู้ถือครองใน SharePoint และเปิดใบรับมอบฉบับใหม่สำหรับพิมพ์หรือบันทึก PDF ทันที ใบรับมอบเดิมและประวัติการเปลี่ยนผู้ใช้งานเปิดดูได้จากรายละเอียดทรัพย์สิน

สำหรับระบบที่ติดตั้งไว้แล้ว ให้อัปเดต Azure Functions ก่อนรัน workflow ของ GitHub Pages แพ็กเกจนี้ใช้ SharePoint Lists และคอลัมน์เดิม

## 1. เตรียม GitHub Pages

- ใช้ repository ของบริษัทที่เปิด GitHub Pages ได้ จากนั้นนำไฟล์ทั้งหมดขึ้นสาขา `main` รวมทั้ง `.github/workflows/pages.yml`
- ที่ **Settings → Pages → Build and deployment** เลือก **GitHub Actions** ใช้ URL จากหน้า Pages เป็น Redirect URI รวม path ของ repository และเครื่องหมาย `/` ท้าย URL เช่น `https://EXAMPLE.github.io/nine-asset-desk/`
- หากต้องเก็บซอร์สเป็น private ให้ตรวจว่า plan ของ GitHub รองรับ Pages สำหรับ private repository และกำหนด visibility ของ Pages ให้ตรงกับนโยบายองค์กร การตั้งค่า visibility ของ Pages กับสิทธิ์ Admin ของแอปเป็นคนละส่วน หน้าเว็บที่เผยแพร่แบบ public จะแสดงหน้าเข้าสู่ระบบและโหลดไฟล์หน้าเว็บได้ ส่วนข้อมูลและการทำรายการต้องผ่าน API
- GitHub Actions จะรันชุดทดสอบ API พร้อม build และ deploy หน้าเว็บทุกครั้งที่ push สาขา `main` หรือสั่งรัน workflow ด้วยตนเอง

[คู่มือ GitHub Pages](https://docs.github.com/en/pages/getting-started-with-github-pages/creating-a-github-pages-site) · [นโยบาย Pages ของ GitHub Organization](https://docs.github.com/en/organizations/managing-organization-settings/managing-the-publication-of-github-pages-sites-for-your-organization)

## 2. ตั้งค่า Microsoft Entra ID

สร้าง App registrations สองรายการใน tenant ของ Nine Distribution โดยเลือก **Accounts in this organizational directory only**

| รายการ | ค่า |
| --- | --- |
| API → Expose an API | Application ID URI `api://<API_CLIENT_ID>` และ delegated scope `AssetDesk.Access` |
| API → App roles | ชื่อและ value `Admin` เลือก allowed member types เป็น **Users/Groups** |
| API → Manifest | ตั้ง `api.requestedAccessTokenVersion` เป็น `2` |
| API → Token configuration | เพิ่ม optional claim `acct` ใน **Access token** เพื่อให้ API แยกสมาชิกในองค์กรจาก guest |
| SPA → Authentication | เพิ่ม platform **Single-page application** และ Redirect URI ของ Pages จากข้อ 1 |
| SPA → API Permissions | เพิ่ม delegated permission `AssetDesk.Access` จาก API registration และให้ผู้ดูแลทำ admin consent ตามนโยบาย |
| Enterprise Applications | ตั้ง **Assignment required? → Yes** ที่ API และ SPA จากนั้น assign เฉพาะผู้ใช้หรือกลุ่มที่ดูแลทะเบียน ให้ role `Admin` ที่ API และให้สิทธิ์เปิด SPA แก่คนกลุ่มเดียวกัน |

บทบาท `Admin` ในระบบนี้คือสิทธิ์ดูแลแอป Asset Desk ไม่ต้องมอบตำแหน่ง Global Administrator ให้ผู้ใช้ทั่วไปของแอป ผู้ดูแล Entra เป็นผู้กำหนดรายชื่อหรือกลุ่มที่ได้รับสิทธิ์

[Microsoft Entra App Roles](https://learn.microsoft.com/en-us/entra/identity-platform/howto-add-app-roles-in-apps) · [Optional claims](https://learn.microsoft.com/en-us/entra/identity-platform/optional-claims-reference) · [App registration สำหรับ SPA](https://learn.microsoft.com/en-us/entra/identity-platform/scenario-spa-app-registration)

## 3. เชื่อมทะเบียน Inventory เดิม

ตั้ง `SHAREPOINT_SITE_ID` เป็น Graph Site ID ของไซต์ Inventory แล้วตั้ง `SHAREPOINT_COMPUTER_LIST_ID` และ `SHAREPOINT_ACCESSORIES_LIST_ID` เป็น GUID ของทะเบียน Computer และ Accessories ตามลำดับ เมื่อใช้คู่นี้ API อ่านข้อมูลสดจากคอลัมน์เดิม และบันทึกการเปลี่ยนผู้ใช้หรือสถานะลงแถวต้นทางโดยตรวจ eTag ก่อนเขียน

1. ในทะเบียนเดิมทั้งสองรายการ ให้สร้างคอลัมน์ชื่อภายใน `AssetDeskPayload` แบบ Multiple lines of text, plain text, ไม่บังคับกรอก และปิดตัวเลือก append changes คอลัมน์นี้เก็บประวัติ ใบรับมอบ และสถานที่ที่เพิ่มผ่าน Asset Desk
2. หาก Status เป็น Choice ให้เพิ่มค่าที่ยังขาดจาก `Borrowed`, `Lost`, `Not available`, `Repair` และ `Retired` โดยคง `Stock`, `In -Use` ของ Computer, และ `In Use` ของ Accessories ไว้ ค่าว่างเดิมจะแสดงเป็น **ยังไม่ได้ระบุ**
3. ใช้หรือสร้าง **9D Asset Desk - Employees** ในไซต์เดียวกันตามคอลัมน์ในตารางด้านล่าง แล้วตั้ง `SHAREPOINT_EMPLOYEES_LIST_ID` รายชื่อผู้ถือครองปัจจุบันจากสองทะเบียนจะปรากฏให้เลือกได้ทันที ส่วนรายชื่อที่เพิ่มจากหน้าแอปจะบันทึกใน Employees
4. ให้ Managed Identity อ่านและเขียนสาม Lists ที่ใช้จริงเท่านั้น ก่อนเปิดใช้งานให้ตรวจสิทธิ์ของผู้ใช้ที่เข้าทะเบียนโดยตรงใน SharePoint ตามที่องค์กรต้องการ

ค่าเดิม `Stock`, `In -Use` / `In Use`, `Borrowed`, `Not available` และ `Lost` จะแสดงเป็น **พร้อมใช้งาน**, **กำลังใช้งาน**, **ให้ยืม**, **ไม่พร้อมใช้งาน** และ **สูญหาย** ตามลำดับ หน้าแอปรองรับ **ส่งซ่อม** และ **ปลดระวาง** เพิ่ม การเปลี่ยนป้ายแสดงผลจะไม่แก้ค่าของแถวเดิม

### การเชื่อมแบบรายการ Asset Desk (โหมดเดิม)

นำสองรายการจาก 9D Asset Desk รุ่น SharePoint มาใช้ต่อได้ หากยังไม่มีรายการ ให้สร้างรายการใหม่ตามตาราง ตั้ง **Enable versioning** และให้ **ชื่อภายในของคอลัมน์** ตรงกับตาราง

| ใช้กับรายการ | ชื่อภายใน | ชนิดและคุณสมบัติ |
| --- | --- | --- |
| ทั้งสองรายการ | `Title` | Single line of text ที่มากับรายการ |
| ทั้งสองรายการ | `AssetDeskKey` | Single line of text, Required, Enforce unique values |
| ทั้งสองรายการ | `AssetDeskIdentity` | Single line of text, Required, Enforce unique values |
| ทั้งสองรายการ | `AssetDeskPayload` | Multiple lines of text, plain text, Required |
| ทรัพย์สิน | `AssetDeskCategory` | Single line of text |
| ทรัพย์สิน | `AssetDeskStatus` | Single line of text |
| ทรัพย์สิน | `AssetDeskHolder` | Single line of text |

ใช้ `9D Asset Desk - Assets` และ `9D Asset Desk - Employees` หรือรายการเดิมที่มีคอลัมน์เดียวกัน จด List ID ของทั้งสองรายการและ Graph Site ID ของไซต์ที่เป็นเจ้าของรายการ หา Site ID ได้จาก Graph `GET /sites/{hostname}:/{site-path}` ซึ่งคืนค่ารูปแบบ `hostname,siteCollectionId,webId`

ใน SharePoint ให้ตรวจสิทธิ์เข้าถึงรายการโดยตรงด้วย หากต้องการให้ข้อมูลจำกัดเฉพาะผู้ดูแลในทุกช่องทาง ให้จำกัดสิทธิ์คนที่อ่านหรือแก้ไข Lists เดิมให้สอดคล้องกับกลุ่ม Admin ของแอป และตรวจหน้าหรือ Web Part เดิมที่มีสิทธิ์ตรงกับ Lists เหล่านี้

## 4. Deploy Azure Functions

1. ใช้ Function App ที่เปิด HTTPS และรองรับ Node.js 22 เปิด **System-assigned managed identity** จด Application (client) ID ของ service principal ของ managed identity จาก Enterprise applications
2. ให้ผู้ดูแลมอบ Microsoft Graph **Application** role `Lists.SelectedOperations.Selected` ให้ managed identity แล้วให้สิทธิ์ `write` เฉพาะ Lists ที่ตั้งค่าไว้ ตัวอย่าง payload สำหรับผู้ดูแลที่มีสิทธิ์ grant อยู่แล้ว

```http
POST https://graph.microsoft.com/v1.0/sites/{site-id}/lists/{list-id}/permissions
Content-Type: application/json

{
  "grantedToV2": {
    "application": { "id": "<MANAGED_IDENTITY_CLIENT_ID>", "displayName": "9D Asset Desk" }
  },
  "roles": ["write"]
}
```

3. ตั้ง **CORS** ของ Function App ให้ตรงกับ origin ของ Pages เช่น `https://EXAMPLE.github.io` โดยไม่ใส่ path หรือ slash ท้าย
4. ตั้งค่าแอปใน Azure ตามตาราง จากนั้นใช้ Azure Functions Core Tools และบัญชี Azure ของผู้ดูแลเพื่อ deploy จากโฟลเดอร์ `api`

```sh
npm ci
npm test
func azure functionapp publish <FUNCTION_APP_NAME>
```

| Azure application setting | ค่า |
| --- | --- |
| `ENTRA_TENANT_ID` | Directory (tenant) ID ขององค์กร |
| `ENTRA_API_CLIENT_ID` | Application (client) ID ของ API |
| `ENTRA_SPA_CLIENT_ID` | Application (client) ID ของ SPA |
| `PAGES_ORIGIN` | Origin ของ Pages ตรงกับ CORS |
| `SHAREPOINT_SITE_ID` | Graph Site ID |
| `SHAREPOINT_COMPUTER_LIST_ID` | GUID ของทะเบียน Computer สำหรับโหมด Inventory |
| `SHAREPOINT_ACCESSORIES_LIST_ID` | GUID ของทะเบียน Accessories สำหรับโหมด Inventory |
| `SHAREPOINT_ASSETS_LIST_ID` | GUID ของรายการทรัพย์สินในโหมดเดิม ใช้เมื่อไม่ได้กำหนดทะเบียน Inventory ทั้งคู่ |
| `SHAREPOINT_EMPLOYEES_LIST_ID` | GUID ของรายการพนักงาน |

[มอบ Graph permission ให้ Managed Identity](https://learn.microsoft.com/en-us/powershell/entra-powershell/grant-api-permissions-managed-identity) · [สิทธิ์เฉพาะ SharePoint Lists](https://learn.microsoft.com/en-us/graph/permissions-selected-overview) · [Grant list permission](https://learn.microsoft.com/en-us/graph/api/list-post-permissions?view=graph-rest-1.0)

## 5. ตั้งตัวแปรใน GitHub และตรวจผล

ที่ **Settings → Secrets and variables → Actions → Variables** เพิ่ม `ENTRA_TENANT_ID`, `ENTRA_SPA_CLIENT_ID`, `ENTRA_API_CLIENT_ID` และ `ASSET_API_URL` เช่น `https://<FUNCTION_HOST>/api` จากนั้นรัน workflow **Deploy 9D Asset Desk to GitHub Pages** อีกครั้ง

ตรวจด้วยบัญชี Admin ขององค์กรว่ามองเห็นทะเบียนและบันทึกได้ ตรวจด้วยบัญชีองค์กรที่ไม่มี role และบัญชี Microsoft ส่วนตัวว่าเปิดทะเบียนไม่ได้ บน API คำขอที่ไม่มี token ต้องได้ 401 และคำขอจากบัญชีที่ยืนยันตัวตนได้แต่ไม่มี role ต้องได้ 403 หาก Entra ออก token ให้บัญชีนั้น

ชุดทดสอบใน repo ใช้ test keys และ Graph จำลองสำหรับตรวจ token สิทธิ์อ่าน–เขียน การบันทึกลงรายการที่กำหนด และการป้องกันบันทึกทับเมื่อมีคนแก้ไขพร้อมกัน การตรวจด้วยบัญชีองค์กรจริงต้องทำหลัง deploy

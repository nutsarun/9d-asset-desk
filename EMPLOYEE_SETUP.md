# 9D IT Asset Pro — SharePoint Employee & Write Setup

## Current deployment

- Frontend: https://nutsarun.github.io/9d-asset-desk/
- SharePoint site: https://ninedistribution.sharepoint.com/sites/Inventory
- Lists currently read: **Inventory 9D&A9 Computer** and **Inventory 9D&A9 Accessories**
- Super Admin application UI: `nutsarun@ninedistribution.com`, `satit@ninedistribution.com`
- Other signed-in users: read-only
- This is client-side role gating. **For real access control, configure SharePoint List permissions** so only intended administrators can edit both asset lists and the employee list. No frontend JavaScript can enforce server-side authorization on its own.

## Microsoft Entra permissions

In App registrations > 9D IT Asset Pro > API permissions > Microsoft Graph > Delegated:
- `User.Read` — login
- `Sites.Read.All` — read SharePoint asset lists
- `Sites.ReadWrite.All` — required to create/update/delete SharePoint rows from this website (**admin consent required**).

After granting consent, log out and log back in. **Only grant broader permissions after confirming your organizational security policy.** `Sites.ReadWrite.All` can access more sites than this application needs. For production use, a backend with narrow site-scoped authorization is preferable.

## Employee directory — shared persistence

Create a new **SharePoint List** at the same Inventory site with display name **9D Asset Employees**.

Use the default `Title` column for the employee's full name. Add the following internal column names, each type *Single line of text*:

| Internal name | Purpose |
|---|---|
| `EmployeeCode` | Employee ID |
| `Department` | Department |
| `Position` | Job title |
| `Email` | Email |
| `Contact` | Phone/contact |

The app discovers this list by its display name at login and saves additions/edits/deletions there. Until it exists, the employee page infers names from assets and saves manual changes **only in that browser's localStorage**; other users will not see those local changes.

## Asset assignment rules

| List | Holder column (internal) | Occupied status | Available status |
|---|---|---|---|
| Computer | `_x0e4a_UserName` | `In -Use` | `Stock` |
| Accessories | `UserName0` | `Borrowed` (also recognizes `In Use`) | `Stock` |

- A non-empty user name also means **occupied**, regardless of status.
- The app re-reads each SharePoint item before checkout and uses its ETag to detect concurrent modification.
- An item with any current holder or occupied/unknown blocking status cannot be checked out.
- For multi-item checkouts, partial success is possible: each successful item is immediately written to SharePoint, and the UI reports failures individually.
- Return clears holder and resets status to `Stock`. Inspect any damaged return separately.
- Generated handover/return transactions remain in **browser session storage**, not a shared permanent audit trail. Build a dedicated SharePoint transaction list and/or immutable server-side audit service before using documents as legal acceptance records.

## Production hardening

1. Set SharePoint permissions for the two asset lists and employee list; **do not rely only on frontend email checks**.
2. Decide whether to maintain a separate immutable transaction and signature record.
3. Ensure account permissions and audit logging meet company policy.
4. Consider migrating write actions to a controlled backend for atomic multi-item operations.

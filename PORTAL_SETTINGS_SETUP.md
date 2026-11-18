# 9D IT Asset Pro — Language, User Roles and A4 Forms

## Release

- Settings > Website Language: Thai (th) or English (en). Each browser remembers its own preference. Major navigation, forms, inventory, dashboard and print controls have English labels; some informational/help messages still remain Thai.
- User Management: visible **only to the two fixed Super Admin accounts**, nutsarun@ninedistribution.com and satit@ninedistribution.com.
- Roles: Super Admin (fixed accounts), IT Staff (inventory/check-out/check-in/employee management UI), Auditor (read + asset verification), Viewer (read-only UI).
- Create the **9D Portal Users** SharePoint list by clicking **Create SharePoint User List** on the User Management tab, while signed in as Super Admin.
- Once created, the list has Title (user email), Role (text), Enabled (text). The special `__PORTAL_CONFIG__` record stores the **Allowlist Only** setting (off by default, to avoid locking out existing users).
- Existing Super Admins can add/modify/disable/delete other users. The two permanent Super Admin addresses cannot be downgraded or deleted from the UI.
- Handover (HO) and Return (RT) documents have a redesigned bilingual A4 print view with logo, employee information, asset details, notes and signature sections. The **employee possession summary form is unchanged**.
- Print HO/RT to PDF using the browser's **Print > Save as PDF**.

## Microsoft Entra and SharePoint permissions

The user's current Microsoft 365 sign-in must have permissions to read the SharePoint Inventory site. Creating a List requires delegated **Sites.Manage.All**; writing user-role rows requires delegated **Sites.ReadWrite.All** with suitable consent and underlying SharePoint permissions.

**Security limitation:** The site is a client-side GitHub Pages app. UI role checks and SharePoint-based allowlisting are convenience/access UI gates, **not a security boundary by themselves**. A user with Microsoft Graph write access outside this app may still modify SharePoint data. A malicious user could modify local JS. For secure enforcement, use SharePoint List-level ACLs or a backend using Entra ID app roles / groups and server-side authorization, with least-privilege access to the relevant Lists. Consider separate SharePoint permissions: only Super Admins can edit the 9D Portal Users list, IT Staff can edit the Asset/Employee Lists, and Viewer/Auditor have read-only access.

If SharePoint List-level restrictions are not configured, do not claim the roles can fully protect private asset data from authorized Microsoft 365 users outside this app.

## How to test safely

1. Confirm GitHub Pages deployment finishes.
2. Sign in as a Super Admin. Open Settings, switch Thai/English; verify menu/labels and refresh persistence.
3. Open User Management. Create the SharePoint user list (requires Graph scopes/tenant approval). Add a test user with `Viewer`; confirm editing is hidden. Repeat for Auditor and IT Staff, noting that actual SharePoint write permissions are controlled separately.
4. **Do not turn Allowlist Only on** until you have added all intended users; permanent Super Admins remain allowed.
5. Print one existing handover and one return form, checking A4 margins, long item descriptions, multi-page signatures and logo.
6. Verify Employee Asset Summary PDFs have not changed.

## Data hygiene

The connected SharePoint Master Asset Inventory remains the asset data source. The User Management List contains organizational login/role data only; do not commit user emails, tenant tokens, private asset rows, attachments or actual PDFs to this public repository.

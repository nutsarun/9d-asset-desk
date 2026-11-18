# Migration: 9D Master Asset Inventory

## Status — preparation complete; execution NOT performed

The existing SharePoint connector can read source lists but cannot create SharePoint lists or list items. No migration writes were made to your tenant. The site still has no `9D Master Asset Inventory` list as of the last inspection.

Source inventory confirmed in the same site:

| Source list | Items | With holder names | Has attachments | Missing HostName |
|---|---:|---:|---:|---:|
| Inventory 9D&A9 Computer | 32 | 26 | 9 | HostName column not observed |
| Inventory 9D&A9 Accessories | 38 | 27 | 0 | 15 |
| **Total** | **70** | **53** | **9** | **at least 15** |

No duplicate IDs were observed within either source using the key: Computer `Title`, Accessories populated `HostName`. Cross-list duplicates, if any, are reported by migration dry run.

## Run the migration from an authorized Windows computer

1. Install Microsoft Graph PowerShell authentication:

   ```powershell
   Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
   ```

2. Download `migrate-master-assets.ps1` from this repository into a private local working directory.
3. In PowerShell, run the script **without** `-Apply` to inspect and export data:

   ```powershell
   .\migrate-master-assets.ps1
   ```

4. Review the private `original-data.json` and `mapped-data.json` files, especially duplicate ID groups, missing values, owner information, and files with attachments.
5. Once reviewed, execute:

   ```powershell
   .\migrate-master-assets.ps1 -Apply
   ```

6. Sign into Microsoft with an account authorized for the SharePoint site and consent to the required Graph **delegated** scopes `Sites.Manage.All` and `Sites.ReadWrite.All`. These are **broad** permissions. Ask the Microsoft 365 administrator to approve them only if appropriate to organizational policy.
7. Review `migration-report.json`, `missing-items.json`, `mismatches.json`, `failed-items.json`. Validate records in the new List. Do not switch the website until:
   - all **70** original entries are represented (or differences from source rechecked),
   - source keys and company / holder / status / serial / model match,
   - the **9 records marked as having attachments** have their files handled separately,
   - version history, permissions, image columns, and other specialized metadata are reviewed,
   - there are no unresolved collisions, and
   - a test with one admin and one viewer has succeeded.

## Mapping rules

- Computer: `Title` becomes the preserved AssetID. Owner is `_x0e4a_UserName`. Company is `Company`. Serial is `Assetorserialnumber`.
- Accessories: `HostName` becomes preserved AssetID if present. Where missing, use a temporary unique `LEGACY-ACC-<source item ID>`. Original `Title` retained as `OriginalTitle`. Owner is `UserName0`. Company is `Owner`. Serial is `AssetorserialNumber`.
- Both: `SourceListId` + `SourceItemId` identifies the original row for safe reruns.
- Legacy statuses are kept verbatim, not inferred or overwritten.
- The script copies the key basic asset fields, **not every custom field, attachment or version**. The private source JSON backup retains original row fields for further mapping if needed.

The script creates only the target list and inserts target items; it does **not** edit or delete the original lists, and does **not** modify the GitHub Pages app's data source. A SharePoint migration is not atomic and concurrent updates after the read may need a delta follow-up.

## Important security note

The generated backup JSON includes employee and asset inventory data. Keep it locally in an access-controlled folder. **Do not upload it to a public GitHub repository**.

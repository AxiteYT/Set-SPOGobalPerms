# Set-SharePointOnlinePermissions

## License

This program is licensed under the GNU Affero General Public License v3.0. See the `LICENSE` file for the full license text.

## Synopsis

Set-SharePointOnlinePermissions is a PowerShell script that sets permission levels for SharePoint Online sites using PnP PowerShell. It allows administrators to easily manage permissions for users and groups across SharePoint Online sites and subsites.

## Description

The script sets permission levels for SharePoint Online sites using PnP PowerShell. The script prompts the user to confirm the action before making any changes. It includes options for specifying the SharePoint tenant URL, the SharePoint Online site URL, the identity type (user or group), the user or group identity, the permission level to set, and whether to update all document libraries in the site and subsites.

## Parameters

- **SPOTenant**: SharePoint tenant name. (Mandatory)
- **SPOSite**: SharePoint Online site name. (Optional)
- **IdentityType**: Identity type, either "User" or "Group" (Mandatory).
- **Identity**: User or group identity (Mandatory).
- **PermissionLevel**: Permission level to set (Mandatory).
- **UpdateLibraries**: Update all document libraries in the site and subsites (Optional).

## Example

```powershell
Set-SPOGobalPerms -SPOTenant 'contoso' -SPOSite 'finance' -IdentityType 'User' -Identity 'jane@contoso.com' -PermissionLevel 'Read'
```

## Author

Kyle Gary Smith, 2023

## Links
- [AGPL-3.0 License](https://www.gnu.org/licenses/agpl-3.0.en.html)

# Scoped Microsoft Graph Mail Sender

**Script:** `Create-OrUpdate-ScopedGraphMailSender.ps1`  
**Version:** `1.0.0`

## Purpose

This PowerShell script creates or reconciles a Microsoft Entra application and Exchange Online Application RBAC configuration that allows Microsoft Graph to send mail as exactly one Exchange Online mailbox.

It is intended for MFA/OTP, notifications, and other service-to-service mail scenarios where an application needs `Mail.Send` without granting tenant-wide access to every mailbox.

The script is organization-agnostic. It prompts for the Microsoft 365 administrator account, sender mailbox, and application display name, and it can either create a configuration from scratch or update an existing one.

> **Security:** The script intentionally does not assign tenant-wide Microsoft Graph `Mail.Send` Application permission. Mailbox restriction is enforced with Exchange Online RBAC for Applications.

## Capabilities

- Create a new Entra app registration and Enterprise Application.
- Find and update an existing app by exact display name.
- Update a specific existing app by Application (client) ID.
- Create or reconcile a mailbox-specific Exchange management scope.
- Create or reconcile the Exchange `Application Mail.Send` role assignment.
- Create the Exchange service-principal pointer if it is missing.
- Detect broad Microsoft Graph `Mail.Send` application permission.
- Optionally remove a broad `Mail.Send` app-role assignment.
- Create or rotate a client secret.
- Prompt for 3, 6, 12, 18, 24, or custom 1-24 month secret lifetime.
- Validate the intended mailbox is authorized and another mailbox is not.
- Run a read-only preflight.
- Report the script version.
- Check a GitHub raw file for a newer version.
- Update the local script from a newer GitHub copy.

## Security model

```text
Application
  -> Microsoft Graph
  -> Exchange Online Application RBAC
  -> Application Mail.Send
  -> Custom recipient scope
  -> One sender mailbox
```

Do not separately grant Microsoft Graph `Mail.Send` under Entra **API permissions > Application permissions**. Entra application permissions and Exchange Application RBAC permissions are additive. A broad Graph `Mail.Send` grant can defeat the mailbox-only restriction.

## Prerequisites

- PowerShell 7.0.3 or later.
- Internet access to Microsoft 365, Microsoft Graph, Exchange Online, PowerShell Gallery, and Microsoft device login.
- An existing Exchange Online sender mailbox. The address entered must be its primary SMTP address.
- A Microsoft 365 administrator account that is different from the sender mailbox.
- Sufficient Microsoft Graph rights for delegated `Application.ReadWrite.All`.
- Sufficient Exchange Online rights to manage Application RBAC scopes and role assignments.
- Required PowerShell modules:
  - `Microsoft.Graph.Authentication`
  - `Microsoft.Graph.Applications`
  - `ExchangeOnlineManagement`

The script installs or updates the required modules from PSGallery when necessary.

## Quick start

Run a read-only preflight first:

```powershell
.\Create-OrUpdate-ScopedGraphMailSender.ps1 -PreflightOnly
```

Then run normally:

```powershell
.\Create-OrUpdate-ScopedGraphMailSender.ps1
```

The script prompts for:

1. Microsoft 365 administrator UPN, for example `admin@contoso.com`.
2. Sender mailbox, for example `login@contoso.com`.
3. Entra application display name, for example `Contoso MFA Mail 2FA`.

Microsoft Graph authentication uses device-code login. Authenticate with the same administrator account entered at the beginning of the script.

## Create and update behavior

`-Mode Auto` is the default.

- **Auto**: Creates a new app when no exact display-name match exists. Updates the existing app when exactly one match exists. Stops if duplicate exact-name apps exist.
- **Create**: Requires a new app and stops if an app with the requested display name already exists.
- **Update**: Requires an existing app. `-ExistingAppId` is recommended for deterministic production updates.

Example:

```powershell
.\Create-OrUpdate-ScopedGraphMailSender.ps1 `
  -Mode Update `
  -ExistingAppId "00000000-0000-0000-0000-000000000000" `
  -AdminUPN "admin@contoso.com" `
  -Mailbox "login@contoso.com" `
  -AppDisplayName "Contoso MFA Mail 2FA"
```

## Client secrets

Default secret behavior is `-SecretAction Auto`:

- New app: creates a secret.
- Existing app with an unexpired secret: preserves it.
- Existing app with no unexpired secret: creates a new one.

To deliberately issue a new secret:

```powershell
.\Create-OrUpdate-ScopedGraphMailSender.ps1 -IssueNewSecret
```

If `-SecretValidityMonths` is omitted, the script prompts for:

- 3 months
- 6 months
- 12 months (default)
- 18 months
- 24 months
- Custom 1-24 months

To issue a 24-month secret without prompting:

```powershell
.\Create-OrUpdate-ScopedGraphMailSender.ps1 `
  -IssueNewSecret `
  -SecretValidityMonths 24
```

Existing secret values cannot be recovered from Entra. Copy a newly created secret immediately and store it in an approved secrets-management system.

## Broad Microsoft Graph Mail.Send

The script checks whether the Enterprise Application already has tenant-wide Microsoft Graph `Mail.Send` application permission.

If detected, a full run stops or prompts for remediation unless removal is explicitly allowed:

```powershell
.\Create-OrUpdate-ScopedGraphMailSender.ps1 -RemoveBroadGraphMailSend
```

Review this carefully on shared applications because removing a broad grant can affect other workloads.

## Versioning

The public filename remains stable:

```text
Create-OrUpdate-ScopedGraphMailSender.ps1
```

The semantic version is stored inside the script. The initial GitHub release is:

```text
1.0.0
```

Check the local version:

```powershell
.\Create-OrUpdate-ScopedGraphMailSender.ps1 -Version
```

Short alias:

```powershell
.\Create-OrUpdate-ScopedGraphMailSender.ps1 -v
```

## GitHub update checking

The script can compare itself to a raw GitHub copy.

```powershell
.\Create-OrUpdate-ScopedGraphMailSender.ps1 `
  -CheckForUpdate `
  -GitHubRawUrl "https://raw.githubusercontent.com/OWNER/REPOSITORY/main/Create-OrUpdate-ScopedGraphMailSender.ps1"
```

To install a newer GitHub version:

```powershell
.\Create-OrUpdate-ScopedGraphMailSender.ps1 `
  -UpdateFromGitHub `
  -GitHubRawUrl "https://raw.githubusercontent.com/OWNER/REPOSITORY/main/Create-OrUpdate-ScopedGraphMailSender.ps1"
```

You can also set `$DefaultGitHubRawUrl` near the top of the script after publishing it. Then `-CheckForUpdate` and `-UpdateFromGitHub` can be used without supplying the URL each time.

The self-update process validates the script identity and semantic version, creates a `.bak` backup of the local script, and only replaces the local copy when the GitHub copy has a newer version.

## Parameter reference

| Parameter | Purpose | Default / Notes |
| --- | --- | --- |
| `-Version` / `-v` | Show the internal script version and exit. | `1.0.0` in this release. |
| `-CheckForUpdate` | Compare local version to a GitHub raw file. | Requires configured or supplied raw URL. |
| `-UpdateFromGitHub` | Install a newer GitHub copy. | Creates a `.bak` backup. |
| `-GitHubRawUrl` | Raw GitHub URL used for version checks and updates. | Optional if `$DefaultGitHubRawUrl` is set. |
| `-Mode` | `Auto`, `Create`, or `Update`. | `Auto` |
| `-AppDisplayName` | Entra application display name. | Prompted if omitted. |
| `-Mailbox` | Single sender mailbox to authorize. | Prompted; primary SMTP required. |
| `-AdminUPN` | Administrator used for setup. | Prompted; must differ from sender. |
| `-ExistingAppId` | Existing Application/client ID. | Recommended for production updates. |
| `-SecretValidityMonths` | Lifetime for a newly created secret. | 1-24 months. |
| `-SecretAction` | `Auto`, `Always`, or `Never`. | `Auto` |
| `-IssueNewSecret` | Force a fresh secret value. | Switch. |
| `-RemoveBroadGraphMailSend` | Permit removal of broad Graph `Mail.Send`. | Switch. |
| `-RenameExistingApp` | Rename an existing app to `-AppDisplayName`. | Switch. |
| `-PreflightOnly` | Run validation without tenant changes. | Switch. |

## Successful-run checks

A successful run should show:

- The Exchange management scope resolves to exactly one recipient.
- The intended sender passes the positive `Application Mail.Send` authorization test.
- A different mailbox fails the scope test.
- No tenant-wide Graph `Mail.Send` permission is assigned by the script.

The OAuth token scope is:

```text
https://graph.microsoft.com/.default
```

The application sends with:

```text
POST https://graph.microsoft.com/v1.0/users/{mailbox-object-id}/sendMail
```

For client-credentials authentication, use `/users/{id}/sendMail`, not `/me/sendMail`.

## Troubleshooting

| Symptom | Action |
| --- | --- |
| Exchange RBAC cmdlets are missing | Confirm Exchange authenticated as the administrator account, not the sender mailbox. |
| Graph authentication is canceled or opens unexpectedly | Use the device-code URL/code printed by the script and sign in with the requested admin UPN. |
| Multiple apps share the same display name | Rerun with `-ExistingAppId`. |
| Existing app has broad Graph `Mail.Send` | Review impact, then allow removal only if appropriate. |
| Mailbox is not the primary SMTP address | Use the mailbox primary SMTP address. |
| Scope matches more than one recipient | Stop and review Exchange addresses/filtering before continuing. |
| Need an old secret value | Existing values cannot be recovered; use `-IssueNewSecret`. |
| Local script may be old | Run `-Version`, then `-CheckForUpdate`. |

## Operational recommendations

- Store client secrets in an approved secrets/password-management system.
- Never commit client secret values to GitHub.
- Never place real tenant IDs, Application IDs, service-principal IDs, mailbox addresses, admin UPNs, access tokens, or device-login codes in documentation or examples.
- Run `-PreflightOnly` before planned changes.
- Prefer `-ExistingAppId` for production updates.
- After secret rotation, validate the consuming application before deleting the previous secret.
- Review the configuration when the sender mailbox changes or the application is retired.

## Public repository privacy checklist

Before each GitHub release:

- Search the script and documentation for real email addresses and UPNs.
- Search for tenant IDs, Application IDs, Object IDs, and service-principal IDs.
- Search for client secret values, bearer tokens, device-login codes, and passwords.
- Use only fictional example domains such as `contoso.com`.
- Keep real console output and screenshots out of the public repository.
- Scrub document author/last-modified metadata before publishing Office documents.
- Review Git history as well as the current files. Deleting a secret from the latest commit does not remove it from previous commits.

## Microsoft documentation

- Exchange Online RBAC for Applications: https://learn.microsoft.com/exchange/permissions-exo/application-rbac
- Connect to Exchange Online PowerShell: https://learn.microsoft.com/powershell/exchange/connect-to-exchange-online-powershell
- Microsoft Graph PowerShell `Connect-MgGraph`: https://learn.microsoft.com/powershell/module/microsoft.graph.authentication/connect-mggraph
- Microsoft Graph `sendMail`: https://learn.microsoft.com/graph/api/user-sendmail
- Microsoft Graph permissions reference: https://learn.microsoft.com/graph/permissions-reference

## Change record

| Version | Notes |
| --- | --- |
| `1.0.0` | Initial public GitHub release. Create/update workflow, preflight, mailbox-scoped Exchange Application RBAC, secret rotation/lifetime selection, broad `Mail.Send` detection/remediation, authorization verification, internal version reporting, and GitHub update checking. |

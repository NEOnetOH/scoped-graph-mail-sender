<#
.SYNOPSIS
Creates OR updates a Microsoft Entra / Exchange Online application that can send
mail as exactly one Exchange Online mailbox using Exchange Online RBAC for Applications.

.DESCRIPTION
This is an organization-agnostic, idempotent setup/update script with a stable filename and an internal semantic version for GitHub distribution and synchronization.

It can:
  - create a brand-new Entra app registration and Enterprise Application
  - automatically find an existing app by exact display name and update it
  - update a specific existing app by Application (client) ID
  - create or update the Exchange management scope for a single sender mailbox
  - create or update the Exchange "Application Mail.Send" role assignment
  - create the Exchange pointer to the Entra service principal if it is missing
  - rotate/create a client secret when requested
  - use -IssueNewSecret during an update to force a fresh secret
  - prompt for a 3/6/12/18/24-month or custom secret lifetime when needed
  - report its internal version with -Version
  - compare itself with a GitHub raw file using -CheckForUpdate
  - update itself from a GitHub raw file using -UpdateFromGitHub
  - detect tenant-wide Microsoft Graph Mail.Send application permission
  - optionally remove the broad Microsoft Graph Mail.Send grant
  - validate the Exchange scope resolves to exactly one recipient
  - verify the intended mailbox is in scope and a second mailbox is out of scope

Microsoft Graph work runs in a separate PowerShell 7 child process to avoid
Microsoft.Graph / ExchangeOnlineManagement MSAL/WAM module conflicts.

IMPORTANT
Do not leave a tenant-wide Microsoft Graph Mail.Send APPLICATION permission on
this app. Exchange RBAC application permissions and Entra application permissions
are additive. A broad Graph Mail.Send grant would defeat the mailbox-only scope.

.EXAMPLE
# Interactive create-or-update mode:
.\Create-OrUpdate-ScopedGraphMailSender.ps1

.EXAMPLE
# Read-only preflight:
.\Create-OrUpdate-ScopedGraphMailSender.ps1 -PreflightOnly

.EXAMPLE
# Update a specific existing app and issue a new client secret.
# If -SecretValidityMonths is omitted, an interactive lifetime menu is shown:
.\Create-OrUpdate-ScopedGraphMailSender.ps1 `
  -ExistingAppId "00000000-0000-0000-0000-000000000000" `
  -AdminUPN "admin@contoso.com" `
  -Mailbox "login@contoso.com" `
  -AppDisplayName "Contoso MFA Mail" `
  -IssueNewSecret

.EXAMPLE
# Automatically remove a tenant-wide Graph Mail.Send grant if detected:
.\Create-OrUpdate-ScopedGraphMailSender.ps1 -RemoveBroadGraphMailSend

.EXAMPLE
# Show the script version:
.\Create-OrUpdate-ScopedGraphMailSender.ps1 -Version

.EXAMPLE
# Check the GitHub copy for a newer version:
.\Create-OrUpdate-ScopedGraphMailSender.ps1 `
  -CheckForUpdate `
  -GitHubRawUrl "https://raw.githubusercontent.com/OWNER/REPOSITORY/main/Create-OrUpdate-ScopedGraphMailSender.ps1"

.EXAMPLE
# Replace the local script with a newer GitHub copy:
.\Create-OrUpdate-ScopedGraphMailSender.ps1 `
  -UpdateFromGitHub `
  -GitHubRawUrl "https://raw.githubusercontent.com/OWNER/REPOSITORY/main/Create-OrUpdate-ScopedGraphMailSender.ps1"
#>

[CmdletBinding()]
param(
    [Alias('v')]
    [switch]$Version,

    [switch]$CheckForUpdate,

    [switch]$UpdateFromGitHub,

    [string]$GitHubRawUrl = "",

    [ValidateSet('Auto','Create','Update')]
    [string]$Mode = 'Auto',

    [string]$AppDisplayName = "",

    [string]$Mailbox = "",

    [string]$AdminUPN = "",

    [string]$ExistingAppId = "",

    [ValidateRange(1,24)]
    [int]$SecretValidityMonths = 12,

    [ValidateSet('Auto','Always','Never')]
    [string]$SecretAction = 'Auto',

    [switch]$IssueNewSecret,

    [switch]$RemoveBroadGraphMailSend,

    [switch]$RenameExistingApp,

    [switch]$PreflightOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Stable release metadata. Keep the filename unchanged in GitHub and bump only
# this version for releases.
$ScriptIdentity = 'ScopedGraphMailSender'
$ScriptName = 'Create-OrUpdate-ScopedGraphMailSender.ps1'
$ScriptVersion = [version]'1.0.0'

# Optional: set this once after publishing the script to GitHub. If left blank,
# callers can provide -GitHubRawUrl when checking/updating.
$DefaultGitHubRawUrl = ''

# Preserve whether the caller explicitly supplied a lifetime. If not, the script
# can prompt with friendly lifetime choices only when it actually needs to issue
# a new secret.
$SecretValidityMonthsWasSpecified = $PSBoundParameters.ContainsKey('SecretValidityMonths')

# ---------------------------------------------------------------------------
# UI helpers
# ---------------------------------------------------------------------------

function Section([string]$Text) {
    Write-Host ""
    Write-Host ('=' * 78) -ForegroundColor DarkCyan
    Write-Host " $Text" -ForegroundColor Cyan
    Write-Host ('=' * 78) -ForegroundColor DarkCyan
}

function Pass([string]$Text) { Write-Host "[PASS] $Text" -ForegroundColor Green }
function Info([string]$Text) { Write-Host "[INFO] $Text" -ForegroundColor Gray }
function Warn([string]$Text) { Write-Host "[WARN] $Text" -ForegroundColor Yellow }

function Test-MailAddress([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    return $Value.Trim() -match '^[^@\s]+@[^@\s]+\.[^@\s]+$'
}

function Test-GuidString([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $g = [guid]::Empty
    return [guid]::TryParse($Value, [ref]$g)
}

function Test-InScopeTrue($Value) {
    if ($Value -is [bool]) { return [bool]$Value }

    return [string]::Equals(
        ([string]$Value).Trim(),
        'True',
        [System.StringComparison]::OrdinalIgnoreCase
    )
}

function Read-YesNo {
    param(
        [Parameter(Mandatory=$true)][string]$Prompt,
        [bool]$Default = $false
    )

    $suffix = if ($Default) { '[Y/n]' } else { '[y/N]' }

    while ($true) {
        $answer = Read-Host "$Prompt $suffix"

        if ([string]::IsNullOrWhiteSpace($answer)) {
            return $Default
        }

        switch ($answer.Trim().ToLowerInvariant()) {
            'y'   { return $true }
            'yes' { return $true }
            'n'   { return $false }
            'no'  { return $false }
        }

        Write-Host "Please enter Y or N." -ForegroundColor Yellow
    }
}

function Read-SecretLifetimeMonths {
    Write-Host ""
    Write-Host "Choose the lifetime for the NEW client secret:" -ForegroundColor Cyan
    Write-Host "  1) 3 months"
    Write-Host "  2) 6 months"
    Write-Host "  3) 12 months (default)"
    Write-Host "  4) 18 months"
    Write-Host "  5) 24 months"
    Write-Host "  6) Custom (1-24 months)"

    while ($true) {
        $choice = Read-Host "Secret lifetime [3]"

        if ([string]::IsNullOrWhiteSpace($choice)) {
            return 12
        }

        switch ($choice.Trim()) {
            '1' { return 3 }
            '2' { return 6 }
            '3' { return 12 }
            '4' { return 18 }
            '5' { return 24 }
            '6' {
                while ($true) {
                    $custom = Read-Host "Enter secret lifetime in months (1-24)"
                    $months = 0

                    if ([int]::TryParse($custom, [ref]$months) -and
                        $months -ge 1 -and
                        $months -le 24) {
                        return $months
                    }

                    Write-Host "Enter a whole number from 1 through 24." -ForegroundColor Yellow
                }
            }
            default {
                Write-Host "Choose 1, 2, 3, 4, 5, or 6." -ForegroundColor Yellow
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Script version / GitHub synchronization
# ---------------------------------------------------------------------------

function Resolve-GitHubRawUrl {
    if (-not [string]::IsNullOrWhiteSpace($GitHubRawUrl)) {
        return $GitHubRawUrl.Trim()
    }

    if (-not [string]::IsNullOrWhiteSpace($DefaultGitHubRawUrl)) {
        return $DefaultGitHubRawUrl.Trim()
    }

    throw @"
No GitHub raw URL is configured.

Either:
  1. Pass -GitHubRawUrl with the raw GitHub URL, or
  2. Set `$DefaultGitHubRawUrl near the top of this script.

Expected form:
https://raw.githubusercontent.com/OWNER/REPOSITORY/main/Create-OrUpdate-ScopedGraphMailSender.ps1
"@
}

function Get-RemoteScriptRelease {
    param(
        [Parameter(Mandatory=$true)]
        [string]$RawUrl
    )

    Info "Reading release metadata from GitHub..."

    $response = Invoke-WebRequest `
        -Uri $RawUrl `
        -UseBasicParsing `
        -ErrorAction Stop

    $remoteText = [string]$response.Content

    if ([string]::IsNullOrWhiteSpace($remoteText)) {
        throw "GitHub returned an empty script."
    }

    $identityPattern = [regex]::Escape('$ScriptIdentity') +
        "\s*=\s*['""]ScopedGraphMailSender['""]"

    if ($remoteText -notmatch $identityPattern) {
        throw "The GitHub file does not identify itself as '$ScriptIdentity'. Update aborted."
    }

    $versionPattern = [regex]::Escape('$ScriptVersion') +
        "\s*=\s*\[version\]\s*['""](?<Version>\d+\.\d+\.\d+)['""]"

    $match = [regex]::Match($remoteText, $versionPattern)

    if (-not $match.Success) {
        throw "Could not read the remote script version."
    }

    [pscustomobject]@{
        Version = [version]$match.Groups['Version'].Value
        Content = $remoteText
        Url     = $RawUrl
    }
}

function Show-ScriptVersion {
    Write-Output "$ScriptName $ScriptVersion"
}

function Invoke-GitHubVersionCheck {
    param(
        [switch]$InstallUpdate
    )

    $rawUrl = Resolve-GitHubRawUrl
    $remote = Get-RemoteScriptRelease -RawUrl $rawUrl

    Write-Host ""
    Write-Host "Script              : $ScriptName"
    Write-Host "Local version       : $ScriptVersion"
    Write-Host "GitHub version      : $($remote.Version)"
    Write-Host "GitHub source       : $rawUrl"

    if ($remote.Version -lt $ScriptVersion) {
        Warn "The GitHub copy is older than this local copy."
        return
    }

    if ($remote.Version -eq $ScriptVersion) {
        Pass "This script is already current."
        return
    }

    Warn "A newer version is available: $($remote.Version)"

    if (-not $InstallUpdate) {
        Info "Run with -UpdateFromGitHub to install it."
        return
    }

    if ([string]::IsNullOrWhiteSpace($PSCommandPath) -or
        -not (Test-Path -LiteralPath $PSCommandPath)) {
        throw "The running script path could not be determined, so self-update is unavailable."
    }

    $currentPath = [IO.Path]::GetFullPath($PSCommandPath)
    $directory = Split-Path -Parent $currentPath
    $fileName = Split-Path -Leaf $currentPath
    $backupPath = "$currentPath.bak"
    $stagingPath = Join-Path $directory ".$fileName.update"

    # Stage and validate before touching the running copy.
    [IO.File]::WriteAllText(
        $stagingPath,
        $remote.Content,
        [Text.UTF8Encoding]::new($false)
    )

    try {
        Copy-Item `
            -LiteralPath $currentPath `
            -Destination $backupPath `
            -Force `
            -ErrorAction Stop

        Copy-Item `
            -LiteralPath $stagingPath `
            -Destination $currentPath `
            -Force `
            -ErrorAction Stop

        Remove-Item `
            -LiteralPath $stagingPath `
            -Force `
            -ErrorAction SilentlyContinue

        Pass "Updated $fileName from $ScriptVersion to $($remote.Version)."
        Info "Backup created: $backupPath"
        Info "Rerun the script to use the new version."
    }
    catch {
        Warn "Automatic replacement failed: $($_.Exception.Message)"
        Warn "The downloaded update was left here: $stagingPath"
        throw
    }
}

# Version/update operations intentionally happen before module installation,
# authentication, or tenant changes.
if ($Version) {
    Show-ScriptVersion
    exit 0
}

if ($CheckForUpdate -and $UpdateFromGitHub) {
    throw "Use either -CheckForUpdate or -UpdateFromGitHub, not both."
}

if ($CheckForUpdate) {
    Invoke-GitHubVersionCheck
    exit 0
}

if ($UpdateFromGitHub) {
    Invoke-GitHubVersionCheck -InstallUpdate
    exit 0
}

# ---------------------------------------------------------------------------
# PowerShell / module bootstrap
# ---------------------------------------------------------------------------

function Find-Pwsh7 {
    $candidates = @()

    $cmd = Get-Command pwsh.exe -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Source) {
        $candidates += $cmd.Source
    }

    if ($env:LOCALAPPDATA) {
        $candidates += "$env:LOCALAPPDATA\Microsoft\WindowsApps\pwsh.exe"
        $candidates += "$env:LOCALAPPDATA\Microsoft\WindowsApps\Microsoft.PowerShell_8wekyb3d8bbwe\pwsh.exe"
    }

    if ($env:ProgramFiles) {
        $candidates += "$env:ProgramFiles\PowerShell\7\pwsh.exe"
    }

    if (${env:ProgramFiles(x86)}) {
        $candidates += "${env:ProgramFiles(x86)}\PowerShell\7\pwsh.exe"
    }

    foreach ($path in ($candidates | Select-Object -Unique)) {
        if (-not (Test-Path $path)) { continue }

        try {
            $versionText = & $path `
                -NoLogo `
                -NoProfile `
                -Command '$PSVersionTable.PSVersion.ToString()' 2>$null |
                Select-Object -First 1

            $version = [version]$versionText

            if ($version -ge [version]'7.0.3') {
                return $path
            }
        }
        catch {}
    }

    return $null
}

function Ensure-PowerShell7 {
    $pwsh = Find-Pwsh7

    if ($pwsh) {
        return $pwsh
    }

    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue

    if (-not $winget) {
        throw "PowerShell 7.0.3+ is required and WinGet is unavailable."
    }

    Info "PowerShell 7 not found. Installing it with WinGet..."

    Start-Process `
        -FilePath $winget.Source `
        -ArgumentList @(
            'install',
            '--id','Microsoft.PowerShell',
            '--source','winget',
            '--exact',
            '--accept-source-agreements',
            '--accept-package-agreements',
            '--silent'
        ) `
        -Wait | Out-Null

    Start-Sleep -Seconds 2

    $pwsh = Find-Pwsh7

    if (-not $pwsh) {
        throw "PowerShell 7 could not be found after installation."
    }

    return $pwsh
}

function Ensure-ModuleInstalled([string]$Name) {
    $gallery = Find-Module `
        -Name $Name `
        -Repository PSGallery `
        -ErrorAction Stop

    $installed = Get-Module `
        -ListAvailable `
        -Name $Name |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if (-not $installed -or $installed.Version -lt $gallery.Version) {
        Info "Installing/updating $Name to $($gallery.Version)..."

        Install-Module `
            -Name $Name `
            -RequiredVersion $gallery.Version `
            -Scope CurrentUser `
            -Repository PSGallery `
            -Force `
            -AllowClobber `
            -ErrorAction Stop

        $installed = Get-Module `
            -ListAvailable `
            -Name $Name |
            Where-Object Version -eq $gallery.Version |
            Select-Object -First 1
    }

    if (-not $installed) {
        throw "Unable to install $Name."
    }

    Pass "$Name installed: $($installed.Version)"
    return $installed.Version
}

function Invoke-GraphChild {
    param(
        [Parameter(Mandatory=$true)][string]$PwshPath,
        [Parameter(Mandatory=$true)][string]$ScriptPath,
        [Parameter()][string[]]$ChildArguments = @()
    )

    & $PwshPath `
        -NoLogo `
        -NoProfile `
        -ExecutionPolicy Bypass `
        -File $ScriptPath `
        @ChildArguments 2>&1 |
        ForEach-Object { Write-Host $_ }

    $code = $LASTEXITCODE

    if ($null -eq $code) {
        $code = 1
    }

    Info "Graph child process exit code: $code"
    return [int]$code
}

# ---------------------------------------------------------------------------
# Bootstrap into PowerShell 7 if necessary
# ---------------------------------------------------------------------------

$pwsh = Ensure-PowerShell7

if ($PSVersionTable.PSEdition -ne 'Core' -or
    $PSVersionTable.PSVersion -lt [version]'7.0.3') {

    $forward = @(
        '-Mode', $Mode,
        '-SecretAction', $SecretAction
    )

    if ($SecretValidityMonthsWasSpecified) {
        $forward += '-SecretValidityMonths'
        $forward += [string]$SecretValidityMonths
    }

    if ($IssueNewSecret) {
        $forward += '-IssueNewSecret'
    }

    if (-not [string]::IsNullOrWhiteSpace($GitHubRawUrl)) {
        $forward += '-GitHubRawUrl'
        $forward += $GitHubRawUrl
    }

    if (-not [string]::IsNullOrWhiteSpace($AppDisplayName)) {
        $forward += '-AppDisplayName'
        $forward += $AppDisplayName
    }

    if (-not [string]::IsNullOrWhiteSpace($Mailbox)) {
        $forward += '-Mailbox'
        $forward += $Mailbox
    }

    if (-not [string]::IsNullOrWhiteSpace($AdminUPN)) {
        $forward += '-AdminUPN'
        $forward += $AdminUPN
    }

    if (-not [string]::IsNullOrWhiteSpace($ExistingAppId)) {
        $forward += '-ExistingAppId'
        $forward += $ExistingAppId
    }

    if ($RemoveBroadGraphMailSend) {
        $forward += '-RemoveBroadGraphMailSend'
    }

    if ($RenameExistingApp) {
        $forward += '-RenameExistingApp'
    }

    if ($PreflightOnly) {
        $forward += '-PreflightOnly'
    }

    & $pwsh `
        -NoLogo `
        -NoProfile `
        -ExecutionPolicy Bypass `
        -File $PSCommandPath `
        @forward

    exit $LASTEXITCODE
}

# ---------------------------------------------------------------------------
# Local prerequisites
# ---------------------------------------------------------------------------

Section "LOCAL PREREQUISITES"

Info "Script name: $ScriptName"
Info "Script version: $ScriptVersion"
Info "Running file: $PSCommandPath"

$null = Get-PSRepository -Name PSGallery -ErrorAction Stop

$graphAuthVersion = Ensure-ModuleInstalled 'Microsoft.Graph.Authentication'
$graphAppsVersion = Ensure-ModuleInstalled 'Microsoft.Graph.Applications'
$exoVersion       = Ensure-ModuleInstalled 'ExchangeOnlineManagement'

if ($graphAuthVersion -ne $graphAppsVersion) {
    throw "Microsoft Graph module versions do not match: Authentication=$graphAuthVersion Applications=$graphAppsVersion."
}

Import-Module ExchangeOnlineManagement -ErrorAction Stop

Pass "Loaded ExchangeOnlineManagement $exoVersion in the current process."
Info "Microsoft Graph modules will run only in a separate PowerShell 7 process."

# ---------------------------------------------------------------------------
# Interactive / parameter inputs
# ---------------------------------------------------------------------------

if ([string]::IsNullOrWhiteSpace($AdminUPN)) {
    $AdminUPN = Read-Host "Enter the Microsoft 365 administrator UPN for setup"
}

$AdminUPN = $AdminUPN.Trim()

if (-not (Test-MailAddress $AdminUPN)) {
    throw "Invalid administrator UPN: '$AdminUPN'."
}

if ([string]::IsNullOrWhiteSpace($Mailbox)) {
    $Mailbox = Read-Host "Enter the sender mailbox address the application should be allowed to send as"
}

$Mailbox = $Mailbox.Trim()

if (-not (Test-MailAddress $Mailbox)) {
    throw "Invalid sender mailbox address: '$Mailbox'."
}

if ($AdminUPN.ToLowerInvariant() -eq $Mailbox.ToLowerInvariant()) {
    throw "The setup administrator cannot be the sender mailbox ($Mailbox)."
}

if ([string]::IsNullOrWhiteSpace($AppDisplayName)) {
    $exampleAppName = 'Contoso MFA Mail 2FA'
    $genericDefault = "Scoped Mail Sender - $Mailbox"

    Write-Host ""
    Write-Host "Enter the Entra application display name." -ForegroundColor Cyan
    Write-Host "Example: $exampleAppName" -ForegroundColor DarkGray
    Write-Host "Press Enter to use: $genericDefault" -ForegroundColor DarkGray

    $enteredName = Read-Host "Application name"

    if ([string]::IsNullOrWhiteSpace($enteredName)) {
        $AppDisplayName = $genericDefault
    }
    else {
        $AppDisplayName = $enteredName.Trim()
    }
}
else {
    $AppDisplayName = $AppDisplayName.Trim()
}

if ([string]::IsNullOrWhiteSpace($AppDisplayName)) {
    throw "Application display name is required."
}

if (-not [string]::IsNullOrWhiteSpace($ExistingAppId) -and
    -not (Test-GuidString $ExistingAppId)) {
    throw "ExistingAppId is not a valid GUID: $ExistingAppId"
}

if (-not [string]::IsNullOrWhiteSpace($ExistingAppId) -and
    $Mode -eq 'Create') {
    throw "Mode=Create cannot be combined with ExistingAppId."
}

Info "Mode: $Mode"
Info "Setup administrator: $AdminUPN"
Info "Sender mailbox: $Mailbox"
if ($IssueNewSecret -and $SecretAction -eq 'Never') {
    throw "-IssueNewSecret cannot be combined with -SecretAction Never."
}

Info "Requested application name: $AppDisplayName"
Info "Secret action: $SecretAction"
Info "Issue new secret flag: $([bool]$IssueNewSecret)"

# ---------------------------------------------------------------------------
# Temporary Graph child scripts
# ---------------------------------------------------------------------------

$work = Join-Path `
    ([IO.Path]::GetTempPath()) `
    ('ScopedMailSender-' + [guid]::NewGuid().ToString('N'))

New-Item -ItemType Directory -Path $work -Force | Out-Null

$graphPrePath    = Join-Path $work 'graph-preflight.ps1'
$graphBuildPath  = Join-Path $work 'graph-build.ps1'
$graphPreResult  = Join-Path $work 'graph-preflight-result.json'
$graphBuildResult = Join-Path $work 'graph-build-result.json'

# ---------------------------------------------------------------------------
# Graph preflight child
# ---------------------------------------------------------------------------

@'
param(
    [string]$Mode,
    [string]$AppDisplayName,
    [string]$ExistingAppId,
    [string]$AdminUPN,
    [string]$ResultPath
)

$ErrorActionPreference = 'Stop'

function Pass($x) { Write-Host "[PASS] $x" -ForegroundColor Green }
function Info($x) { Write-Host "[INFO] $x" -ForegroundColor Gray }
function Warn($x) { Write-Host "[WARN] $x" -ForegroundColor Yellow }
function Fail($x) { Write-Host "[FAIL] $x" -ForegroundColor Red }

try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Applications -ErrorAction Stop

    Info "Using Microsoft Graph device-code authentication."

    Connect-MgGraph `
        -Scopes 'Application.ReadWrite.All' `
        -UseDeviceCode `
        -NoWelcome `
        -ErrorAction Stop

    $ctx = Get-MgContext

    if (-not $ctx -or [string]::IsNullOrWhiteSpace([string]$ctx.Account)) {
        throw "Microsoft Graph returned no authenticated account."
    }

    Pass "Microsoft Graph authenticated as: $($ctx.Account)"

    if ($ctx.Account.Trim().ToLowerInvariant() -ne
        $AdminUPN.Trim().ToLowerInvariant()) {
        throw "Microsoft Graph authenticated as '$($ctx.Account)' instead of '$AdminUPN'."
    }

    $graphResource = Get-MgServicePrincipal `
        -Filter "appId eq '00000003-0000-0000-c000-000000000000'" `
        -ErrorAction Stop |
        Select-Object -First 1

    if (-not $graphResource) {
        throw "Could not locate the Microsoft Graph service principal."
    }

    $mailSendRole = $graphResource.AppRoles |
        Where-Object {
            $_.Value -eq 'Mail.Send' -and
            $_.AllowedMemberTypes -contains 'Application'
        } |
        Select-Object -First 1

    if (-not $mailSendRole) {
        throw "Could not locate the Microsoft Graph Mail.Send application role."
    }

    $app = $null
    $sp = $null
    $action = ''

    if (-not [string]::IsNullOrWhiteSpace($ExistingAppId)) {
        $app = Get-MgApplication `
            -Filter "appId eq '$ExistingAppId'" `
            -ErrorAction Stop |
            Select-Object -First 1

        if (-not $app) {
            throw "Existing Application ID $ExistingAppId was not found."
        }

        $action = 'Update'
        Pass "Existing app selected by Application ID: $($app.DisplayName) [$($app.AppId)]"
    }
    else {
        $sameNameApps = @(
            Get-MgApplication `
                -All `
                -Property Id,AppId,DisplayName,PasswordCredentials `
                -ErrorAction Stop |
            Where-Object {
                $_.DisplayName -eq $AppDisplayName
            }
        )

        if ($Mode -eq 'Create') {
            if ($sameNameApps.Count -gt 0) {
                $ids = ($sameNameApps | ForEach-Object AppId) -join ', '
                throw "Mode=Create was requested, but '$AppDisplayName' already exists. Application ID(s): $ids"
            }

            $action = 'Create'
        }
        elseif ($Mode -eq 'Update') {
            if ($sameNameApps.Count -eq 0) {
                throw "Mode=Update was requested, but no app named '$AppDisplayName' exists."
            }

            if ($sameNameApps.Count -gt 1) {
                $ids = ($sameNameApps | ForEach-Object AppId) -join ', '
                throw "Multiple apps named '$AppDisplayName' exist. Rerun with -ExistingAppId. IDs: $ids"
            }

            $app = $sameNameApps[0]
            $action = 'Update'
        }
        else {
            if ($sameNameApps.Count -eq 0) {
                $action = 'Create'
            }
            elseif ($sameNameApps.Count -eq 1) {
                $app = $sameNameApps[0]
                $action = 'Update'
            }
            else {
                $ids = ($sameNameApps | ForEach-Object AppId) -join ', '
                throw "Multiple apps named '$AppDisplayName' exist. Rerun with -ExistingAppId. IDs: $ids"
            }
        }
    }

    $hasBroadMailSend = $false
    $validSecretCount = 0
    $latestSecretExpiry = $null

    if ($app) {
        $sp = Get-MgServicePrincipal `
            -Filter "appId eq '$($app.AppId)'" `
            -ErrorAction Stop |
            Select-Object -First 1

        if ($sp) {
            $assignments = Invoke-MgGraphRequest `
                -Method GET `
                -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.Id)/appRoleAssignments"

            $broadAssignments = @(
                $assignments.value |
                Where-Object {
                    $_.resourceId -eq $graphResource.Id -and
                    $_.appRoleId -eq $mailSendRole.Id
                }
            )

            $hasBroadMailSend = ($broadAssignments.Count -gt 0)
        }

        $now = [DateTime]::UtcNow

        $validSecrets = @(
            $app.PasswordCredentials |
            Where-Object {
                $_.EndDateTime -and
                ([datetime]$_.EndDateTime).ToUniversalTime() -gt $now
            }
        )

        $validSecretCount = $validSecrets.Count

        if ($validSecretCount -gt 0) {
            $latestSecretExpiry = (
                $validSecrets |
                Sort-Object EndDateTime -Descending |
                Select-Object -First 1
            ).EndDateTime
        }
    }

    if ($action -eq 'Create') {
        Pass "Planned action: CREATE new application '$AppDisplayName'."
    }
    else {
        Pass "Planned action: UPDATE existing application '$($app.DisplayName)' [$($app.AppId)]."
    }

    if ($hasBroadMailSend) {
        Warn "Existing app has tenant-wide Microsoft Graph Mail.Send APPLICATION permission."
    }
    elseif ($action -eq 'Update') {
        Pass "Existing app does not have tenant-wide Microsoft Graph Mail.Send application permission."
    }

    [pscustomobject]@{
        TenantId                = $ctx.TenantId
        PlannedAction           = $action
        ResolvedAppId           = if ($app) { $app.AppId } else { '' }
        ResolvedAppObjectId     = if ($app) { $app.Id } else { '' }
        ResolvedAppDisplayName  = if ($app) { $app.DisplayName } else { $AppDisplayName }
        ServicePrincipalObjectId = if ($sp) { $sp.Id } else { '' }
        HasBroadGraphMailSend   = $hasBroadMailSend
        ValidSecretCount        = $validSecretCount
        LatestSecretExpiry      = $latestSecretExpiry
    } |
        ConvertTo-Json -Depth 5 |
        Set-Content -Path $ResultPath -Encoding utf8

    exit 0
}
catch {
    Fail $_.Exception.Message

    exit 1
}
'@ | Set-Content -Path $graphPrePath -Encoding utf8

# ---------------------------------------------------------------------------
# Graph build child
# ---------------------------------------------------------------------------

@'
param(
    [string]$AppDisplayName,
    [string]$ResolvedAppId,
    [string]$AdminUPN,
    [int]$SecretValidityMonths,
    [switch]$CreateSecret,
    [switch]$RemoveBroadGraphMailSend,
    [switch]$RenameExistingApp,
    [string]$ResultPath
)

$ErrorActionPreference = 'Stop'

function Pass($x) { Write-Host "[PASS] $x" -ForegroundColor Green }
function Info($x) { Write-Host "[INFO] $x" -ForegroundColor Gray }
function Warn($x) { Write-Host "[WARN] $x" -ForegroundColor Yellow }
function Fail($x) { Write-Host "[FAIL] $x" -ForegroundColor Red }

try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Applications -ErrorAction Stop

    Info "Using Microsoft Graph device-code authentication."

    Connect-MgGraph `
        -Scopes 'Application.ReadWrite.All' `
        -UseDeviceCode `
        -NoWelcome `
        -ErrorAction Stop

    $ctx = Get-MgContext

    if (-not $ctx -or [string]::IsNullOrWhiteSpace([string]$ctx.Account)) {
        throw "Microsoft Graph returned no authenticated account."
    }

    Pass "Microsoft Graph authenticated as: $($ctx.Account)"

    if ($ctx.Account.Trim().ToLowerInvariant() -ne
        $AdminUPN.Trim().ToLowerInvariant()) {
        throw "Microsoft Graph authenticated as '$($ctx.Account)' instead of '$AdminUPN'."
    }

    $graphResource = Get-MgServicePrincipal `
        -Filter "appId eq '00000003-0000-0000-c000-000000000000'" `
        -ErrorAction Stop |
        Select-Object -First 1

    $mailSendRole = $graphResource.AppRoles |
        Where-Object {
            $_.Value -eq 'Mail.Send' -and
            $_.AllowedMemberTypes -contains 'Application'
        } |
        Select-Object -First 1

    if (-not $mailSendRole) {
        throw "Could not locate Microsoft Graph Mail.Send application role."
    }

    $createdNewApp = $false

    if ([string]::IsNullOrWhiteSpace($ResolvedAppId)) {
        $app = New-MgApplication `
            -DisplayName $AppDisplayName `
            -SignInAudience 'AzureADMyOrg' `
            -ErrorAction Stop

        $createdNewApp = $true
        Pass "Created Entra application $($app.AppId)."
    }
    else {
        $app = Get-MgApplication `
            -Filter "appId eq '$ResolvedAppId'" `
            -ErrorAction Stop |
            Select-Object -First 1

        if (-not $app) {
            throw "Existing application $ResolvedAppId was not found during build."
        }

        Pass "Updating existing Entra application $($app.AppId)."

        if ($RenameExistingApp -and $app.DisplayName -ne $AppDisplayName) {
            Update-MgApplication `
                -ApplicationId $app.Id `
                -DisplayName $AppDisplayName `
                -ErrorAction Stop

            $app = Get-MgApplication `
                -ApplicationId $app.Id `
                -ErrorAction Stop

            Pass "Renamed existing app to '$AppDisplayName'."
        }
    }

    $sp = Get-MgServicePrincipal `
        -Filter "appId eq '$($app.AppId)'" `
        -ErrorAction Stop |
        Select-Object -First 1

    if (-not $sp) {
        $sp = New-MgServicePrincipal `
            -AppId $app.AppId `
            -ErrorAction Stop

        Pass "Created Enterprise Application/service principal $($sp.Id)."
    }
    else {
        Pass "Enterprise Application/service principal already exists: $($sp.Id)."
    }

    $assignments = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.Id)/appRoleAssignments"

    $broadAssignments = @(
        $assignments.value |
        Where-Object {
            $_.resourceId -eq $graphResource.Id -and
            $_.appRoleId -eq $mailSendRole.Id
        }
    )

    if ($broadAssignments.Count -gt 0) {
        if (-not $RemoveBroadGraphMailSend) {
            throw "SAFETY STOP: tenant-wide Microsoft Graph Mail.Send is assigned. Rerun with -RemoveBroadGraphMailSend or remove the grant manually."
        }

        foreach ($assignment in $broadAssignments) {
            Remove-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $sp.Id `
                -AppRoleAssignmentId $assignment.id `
                -Confirm:$false `
                -ErrorAction Stop
        }

        Pass "Removed tenant-wide Microsoft Graph Mail.Send application permission."
    }
    else {
        Pass "No tenant-wide Microsoft Graph Mail.Send application permission is assigned."
    }

    $secretText = ''
    $secretExpiry = $null

    if ($CreateSecret) {
        $credential = @{
            displayName   = 'Scoped Mail Sender client secret'
            startDateTime = [DateTime]::UtcNow
            endDateTime   = [DateTime]::UtcNow.AddMonths($SecretValidityMonths)
        }

        $secret = Add-MgApplicationPassword `
            -ApplicationId $app.Id `
            -PasswordCredential $credential `
            -ErrorAction Stop

        if (-not $secret.SecretText) {
            throw "Microsoft Graph did not return the new client secret value."
        }

        $secretText = $secret.SecretText
        $secretExpiry = $secret.EndDateTime

        Pass "Created a new client secret."
    }
    else {
        Info "No new client secret was created."
    }

    [pscustomobject]@{
        TenantId                 = $ctx.TenantId
        CreatedNewApp            = $createdNewApp
        AppDisplayName           = $app.DisplayName
        AppId                    = $app.AppId
        AppObjectId              = $app.Id
        ServicePrincipalObjectId = $sp.Id
        SecretCreated            = [bool]$CreateSecret
        SecretValue              = $secretText
        SecretExpires            = $secretExpiry
    } |
        ConvertTo-Json -Depth 5 |
        Set-Content -Path $ResultPath -Encoding utf8

    exit 0
}
catch {
    Fail $_.Exception.Message

    exit 1
}
'@ | Set-Content -Path $graphBuildPath -Encoding utf8

# ---------------------------------------------------------------------------
# Main workflow
# ---------------------------------------------------------------------------

try {
    Section "PREFLIGHT 1 OF 2 - MICROSOFT GRAPH"

    $preArgs = @(
        '-Mode', $Mode,
        '-AppDisplayName', $AppDisplayName,
        '-AdminUPN', $AdminUPN,
        '-ResultPath', $graphPreResult
    )

    if (-not [string]::IsNullOrWhiteSpace($ExistingAppId)) {
        $preArgs += '-ExistingAppId'
        $preArgs += $ExistingAppId
    }

    $preCode = Invoke-GraphChild `
        -PwshPath $pwsh `
        -ScriptPath $graphPrePath `
        -ChildArguments $preArgs

    if ($preCode -ne 0) {
        throw "Microsoft Graph preflight failed."
    }

    $graphPre = Get-Content `
        -Path $graphPreResult `
        -Raw |
        ConvertFrom-Json

    Section "PREFLIGHT 2 OF 2 - EXCHANGE ONLINE"

    # Clean stale Exchange sessions/modules before connecting.
    try {
        Disconnect-ExchangeOnline `
            -Confirm:$false `
            -ErrorAction SilentlyContinue
    }
    catch {}

    Get-Module `
        -Name 'tmpEXO_*' `
        -ErrorAction SilentlyContinue |
        Remove-Module `
            -Force `
            -ErrorAction SilentlyContinue

    Info "Connecting to Exchange Online as $AdminUPN..."

    Connect-ExchangeOnline `
        -UserPrincipalName $AdminUPN `
        -DisableWAM `
        -ShowBanner:$false `
        -ErrorAction Stop

    Pass "Connected to Exchange Online."

    $connectionInfo = Get-ConnectionInformation `
        -ErrorAction SilentlyContinue |
        Select-Object -First 1

    if ($connectionInfo) {
        $connectedUPN = [string]$connectionInfo.UserPrincipalName
        Info "Exchange connected as: $connectedUPN"

        if ($connectedUPN.Trim().ToLowerInvariant() -ne
            $AdminUPN.Trim().ToLowerInvariant()) {
            throw "Exchange authenticated as '$connectedUPN' instead of '$AdminUPN'."
        }

        Pass "Exchange authenticated with the requested administrator."
    }

    $requiredExchangeCommands = @(
        'Get-EXOMailbox',
        'Get-Recipient',
        'Get-ManagementRole',
        'Get-ManagementScope',
        'New-ManagementScope',
        'Set-ManagementScope',
        'Get-ManagementRoleAssignment',
        'New-ManagementRoleAssignment',
        'Set-ManagementRoleAssignment',
        'Get-ServicePrincipal',
        'New-ServicePrincipal',
        'Set-ServicePrincipal',
        'Test-ServicePrincipalAuthorization'
    )

    $allCommandsReady = $false

    for ($attempt = 1; $attempt -le 12; $attempt++) {
        $missing = @(
            $requiredExchangeCommands |
            Where-Object {
                -not (Get-Command $_ -ErrorAction SilentlyContinue)
            }
        )

        if ($missing.Count -eq 0) {
            $allCommandsReady = $true
            break
        }

        Info "Waiting for Exchange RBAC cmdlets ($attempt/12). Missing: $($missing -join ', ')"
        Start-Sleep -Seconds 5
    }

    if (-not $allCommandsReady) {
        throw "Required Exchange cmdlets did not load. Missing: $($missing -join ', ')"
    }

    foreach ($commandName in $requiredExchangeCommands) {
        Pass "Exchange command available: $commandName"
    }

    $mailboxObject = Get-EXOMailbox `
        -Identity $Mailbox `
        -Properties PrimarySmtpAddress,ExternalDirectoryObjectId,RecipientTypeDetails `
        -ErrorAction Stop

    if (-not $mailboxObject) {
        throw "Mailbox '$Mailbox' was not found."
    }

    if ($mailboxObject.PrimarySmtpAddress.ToString().ToLowerInvariant() -ne
        $Mailbox.ToLowerInvariant()) {
        throw "'$Mailbox' is not the mailbox's primary SMTP address. Primary SMTP is '$($mailboxObject.PrimarySmtpAddress)'."
    }

    if ([string]::IsNullOrWhiteSpace([string]$mailboxObject.ExternalDirectoryObjectId)) {
        throw "Mailbox '$Mailbox' has no ExternalDirectoryObjectId."
    }

    Pass "Mailbox exists: $($mailboxObject.DisplayName) [$($mailboxObject.RecipientTypeDetails)]."
    Pass "$Mailbox is the primary SMTP address."

    $mailSendRole = Get-ManagementRole `
        -Identity 'Application Mail.Send' `
        -ErrorAction Stop

    if (-not $mailSendRole) {
        throw "Exchange role 'Application Mail.Send' was not found."
    }

    Pass "Exchange role 'Application Mail.Send' exists."

    Section "PREFLIGHT RESULT"

    Write-Host "Planned action              : $($graphPre.PlannedAction)"
    Write-Host "Resolved app name           : $($graphPre.ResolvedAppDisplayName)"

    if ($graphPre.ResolvedAppId) {
        Write-Host "Resolved Application ID     : $($graphPre.ResolvedAppId)"
    }

    Write-Host "Broad Graph Mail.Send       : $($graphPre.HasBroadGraphMailSend)"
    Write-Host "Valid existing secrets      : $($graphPre.ValidSecretCount)"

    if ($graphPre.LatestSecretExpiry) {
        Write-Host "Latest secret expiry        : $($graphPre.LatestSecretExpiry)"
    }

    Pass "ALL PREFLIGHT CHECKS PASSED."

    if ($PreflightOnly) {
        if ($graphPre.HasBroadGraphMailSend) {
            Warn "Full execution must remove the broad Graph Mail.Send grant for mailbox-only scoping to be effective."
        }

        Disconnect-ExchangeOnline `
            -Confirm:$false `
            -ErrorAction SilentlyContinue

        exit 0
    }

    # -----------------------------------------------------------------------
    # Resolve broad Graph Mail.Send remediation
    # -----------------------------------------------------------------------

    $removeBroad = [bool]$RemoveBroadGraphMailSend

    if ($graphPre.HasBroadGraphMailSend -and -not $removeBroad) {
        Warn "The existing application has tenant-wide Microsoft Graph Mail.Send."
        Warn "Mailbox-only Exchange RBAC is NOT sufficient while that broad grant remains."

        $removeBroad = Read-YesNo `
            -Prompt "Remove the tenant-wide Graph Mail.Send permission now?" `
            -Default $false

        if (-not $removeBroad) {
            throw "Stopped because broad Microsoft Graph Mail.Send remains assigned."
        }
    }

    # -----------------------------------------------------------------------
    # Decide whether to create/rotate client secret
    # -----------------------------------------------------------------------

    $createSecret = $false

    if ($IssueNewSecret) {
        $createSecret = $true
        Info "-IssueNewSecret was specified; a fresh client secret will be issued."
    }
    elseif ($SecretAction -eq 'Always') {
        $createSecret = $true
    }
    elseif ($SecretAction -eq 'Never') {
        $createSecret = $false
    }
    else {
        if ($graphPre.PlannedAction -eq 'Create') {
            $createSecret = $true
        }
        elseif ([int]$graphPre.ValidSecretCount -eq 0) {
            $createSecret = $true
            Info "No unexpired client secrets were found; Auto mode will create one."
        }
        else {
            $createSecret = $false
            Info "An unexpired client secret already exists; Auto mode will preserve it."
            Info "Use -IssueNewSecret when updating and you want a fresh secret value."
        }
    }

    if ($createSecret) {
        if (-not $SecretValidityMonthsWasSpecified) {
            $SecretValidityMonths = Read-SecretLifetimeMonths
        }

        Info "New secret lifetime: $SecretValidityMonths month(s)"
    }

    # -----------------------------------------------------------------------
    # Graph create/update
    # -----------------------------------------------------------------------

    Section "BUILD 1 OF 2 - CREATE OR UPDATE ENTRA APPLICATION"

    $resolvedId = [string]$graphPre.ResolvedAppId

    $buildArgs = @(
        '-AppDisplayName', $AppDisplayName,
        '-AdminUPN', $AdminUPN,
        '-SecretValidityMonths', [string]$SecretValidityMonths,
        '-ResultPath', $graphBuildResult
    )

    # Switch parameters are added only when true. This avoids PowerShell's
    # external-process parameter-binding problem where the literal strings
    # "True"/"False" cannot reliably bind to [bool] parameters under -File.
    if ($createSecret) {
        $buildArgs += '-CreateSecret'
    }

    if ($removeBroad) {
        $buildArgs += '-RemoveBroadGraphMailSend'
    }

    if ($RenameExistingApp) {
        $buildArgs += '-RenameExistingApp'
    }

    if (-not [string]::IsNullOrWhiteSpace($resolvedId)) {
        $buildArgs += '-ResolvedAppId'
        $buildArgs += $resolvedId
    }

    $buildCode = Invoke-GraphChild `
        -PwshPath $pwsh `
        -ScriptPath $graphBuildPath `
        -ChildArguments $buildArgs

    if ($buildCode -ne 0) {
        throw "Microsoft Graph create/update failed."
    }

    $build = Get-Content `
        -Path $graphBuildResult `
        -Raw |
        ConvertFrom-Json

    if ($build.SecretCreated) {
        Section "CLIENT SECRET CREATED - COPY THIS NOW"

        Write-Host "Application (Client) ID      : $($build.AppId)"
        Write-Host "Service Principal Object ID  : $($build.ServicePrincipalObjectId)"
        Write-Host "CLIENT SECRET VALUE          :" -ForegroundColor Yellow
        Write-Host $build.SecretValue -ForegroundColor Yellow
        Write-Host "Secret expires               : $($build.SecretExpires)"

        Warn "Microsoft only returns the secret VALUE at creation time."
    }
    else {
        Info "No new client secret was created."
    }

    # -----------------------------------------------------------------------
    # Exchange create/update
    # -----------------------------------------------------------------------

    Section "BUILD 2 OF 2 - CREATE OR UPDATE EXCHANGE RBAC"

    $shortAppId = $build.AppId.Substring(0,8)
    $scopeName = "Scoped Mailbox - $shortAppId"
    $assignmentName = "Scoped Mail.Send - $shortAppId"
    $escapedMailbox = $Mailbox.Replace("'","''")
    $expectedScopeFilter = "EmailAddresses -eq '$escapedMailbox'"

    $scope = Get-ManagementScope `
        -Identity $scopeName `
        -ErrorAction SilentlyContinue

    if (-not $scope) {
        $scope = New-ManagementScope `
            -Name $scopeName `
            -RecipientRestrictionFilter $expectedScopeFilter `
            -ErrorAction Stop

        Pass "Created Exchange scope '$scopeName'."
    }
    else {
        # Exchange Online's REST-backed Get-ManagementScope object does not
        # consistently expose RecipientRestrictionFilter as a property even
        # though New-/Set-ManagementScope accept that parameter. To make the
        # script deterministic and idempotent, simply reconcile the existing
        # scope to the desired filter every run instead of reading that property.
        Set-ManagementScope `
            -Identity $scopeName `
            -RecipientRestrictionFilter $expectedScopeFilter `
            -ErrorAction Stop

        Pass "Reconciled existing Exchange scope '$scopeName' to target $Mailbox."

        $scope = Get-ManagementScope `
            -Identity $scopeName `
            -ErrorAction Stop
    }

    # Validate the DESIRED OPATH filter directly instead of depending on the
    # shape of the object returned by Get-ManagementScope.
    $scopeMatches = @(
        Get-Recipient `
            -Filter $expectedScopeFilter `
            -ResultSize Unlimited `
            -ErrorAction Stop
    )

    if ($scopeMatches.Count -ne 1) {
        $matches = (
            $scopeMatches |
            ForEach-Object {
                if ($_.PrimarySmtpAddress) {
                    $_.PrimarySmtpAddress.ToString()
                }
                else {
                    $_.Name
                }
            }
        ) -join ', '

        throw "SAFETY STOP: scope '$scopeName' matched $($scopeMatches.Count) recipients instead of exactly one. Matches: $matches"
    }

    if (([string]$scopeMatches[0].ExternalDirectoryObjectId).ToLowerInvariant() -ne
        ([string]$mailboxObject.ExternalDirectoryObjectId).ToLowerInvariant()) {
        throw "SAFETY STOP: scope '$scopeName' matched '$($scopeMatches[0].PrimarySmtpAddress)' instead of '$Mailbox'."
    }

    Pass "SCOPE TEST: '$scopeName' matches only $Mailbox."

    # Create/update Exchange pointer to Entra service principal.
    $exchangeServicePrincipal = Get-ServicePrincipal `
        -Identity $build.ServicePrincipalObjectId `
        -ErrorAction SilentlyContinue

    if (-not $exchangeServicePrincipal) {
        for ($attempt = 1; $attempt -le 30; $attempt++) {
            try {
                $exchangeServicePrincipal = New-ServicePrincipal `
                    -AppId $build.AppId `
                    -ObjectId $build.ServicePrincipalObjectId `
                    -DisplayName $build.AppDisplayName `
                    -ErrorAction Stop

                break
            }
            catch {
                if ($_.Exception.Message -match 'AADServicePrincipalNotFound') {
                    Info "Waiting for Entra service principal replication ($attempt/30)..."
                    Start-Sleep -Seconds 10
                }
                else {
                    throw
                }
            }
        }

        if (-not $exchangeServicePrincipal) {
            throw "Exchange still cannot see the Entra service principal after five minutes."
        }

        Pass "Created Exchange service-principal pointer."
    }
    else {
        if ($exchangeServicePrincipal.DisplayName -ne $build.AppDisplayName) {
            Set-ServicePrincipal `
                -Identity $build.ServicePrincipalObjectId `
                -DisplayName $build.AppDisplayName `
                -ErrorAction Stop

            Pass "Updated Exchange service-principal display name."
        }
        else {
            Pass "Exchange service-principal pointer already exists."
        }
    }

    # Create/update application role assignment.
    $roleAssignment = Get-ManagementRoleAssignment `
        -Identity $assignmentName `
        -ErrorAction SilentlyContinue

    if (-not $roleAssignment) {
        New-ManagementRoleAssignment `
            -Name $assignmentName `
            -App $build.ServicePrincipalObjectId `
            -Role 'Application Mail.Send' `
            -CustomResourceScope $scopeName `
            -ErrorAction Stop |
            Out-Null

        Pass "Created mailbox-scoped Application Mail.Send role assignment."
    }
    else {
        Set-ManagementRoleAssignment `
            -Identity $assignmentName `
            -CustomResourceScope $scopeName `
            -Enabled $true `
            -ErrorAction Stop

        Pass "Updated existing Application Mail.Send role assignment to the desired scope."
    }

    # -----------------------------------------------------------------------
    # Authorization verification
    # -----------------------------------------------------------------------

    Section "AUTHORIZATION VERIFICATION"

    $positive = Test-ServicePrincipalAuthorization `
        -Identity $build.ServicePrincipalObjectId `
        -Resource $Mailbox `
        -ErrorAction Stop

    $positiveRows = @(
        $positive |
        Where-Object {
            $_.RoleName -eq 'Application Mail.Send'
        }
    )

    $positiveRows |
        Format-Table `
            RoleName,
            GrantedPermissions,
            AllowedResourceScope,
            ScopeType,
            InScope `
            -AutoSize

    $positiveInScope = @(
        $positiveRows |
        Where-Object {
            Test-InScopeTrue $_.InScope
        }
    )

    if ($positiveInScope.Count -eq 0) {
        throw "Authorization failed: $Mailbox is not in scope for Application Mail.Send."
    }

    Pass "POSITIVE TEST: $Mailbox is authorized."

    $otherMailbox = Get-EXOMailbox `
        -ResultSize 25 `
        -Properties PrimarySmtpAddress |
        Where-Object {
            $_.PrimarySmtpAddress.ToString().ToLowerInvariant() -ne
            $Mailbox.ToLowerInvariant()
        } |
        Select-Object -First 1

    if ($otherMailbox) {
        $otherAddress = $otherMailbox.PrimarySmtpAddress.ToString()

        $negative = Test-ServicePrincipalAuthorization `
            -Identity $build.ServicePrincipalObjectId `
            -Resource $otherAddress `
            -ErrorAction Stop

        $negativeRows = @(
            $negative |
            Where-Object {
                $_.RoleName -eq 'Application Mail.Send'
            }
        )

        $negativeRows |
            Format-Table `
                RoleName,
                GrantedPermissions,
                AllowedResourceScope,
                ScopeType,
                InScope `
                -AutoSize

        $negativeInScope = @(
            $negativeRows |
            Where-Object {
                Test-InScopeTrue $_.InScope
            }
        )

        if ($negativeInScope.Count -gt 0) {
            throw "SAFETY FAILURE: $otherAddress is unexpectedly in scope. Another Exchange RBAC assignment or broad permission may exist for this service principal."
        }

        Pass "NEGATIVE TEST: $otherAddress is NOT authorized."
    }
    else {
        Warn "No second mailbox was available for a negative authorization test."
    }

    # -----------------------------------------------------------------------
    # Final output
    # -----------------------------------------------------------------------

    Section "SUCCESS - FINAL CONFIGURATION"

    Write-Host "Operation                     : $($graphPre.PlannedAction)"
    Write-Host "Tenant ID                     : $($build.TenantId)"
    Write-Host "Application name              : $($build.AppDisplayName)"
    Write-Host "Application (Client) ID       : $($build.AppId)"
    Write-Host "App Registration Object ID    : $($build.AppObjectId)"
    Write-Host "Service Principal Object ID   : $($build.ServicePrincipalObjectId)"
    Write-Host "Authorized sender             : $Mailbox"
    Write-Host "Mailbox Entra Object ID       : $($mailboxObject.ExternalDirectoryObjectId)"
    Write-Host "Exchange management scope     : $scopeName"
    Write-Host "Exchange role assignment      : $assignmentName"
    Write-Host "OAuth token scope             : https://graph.microsoft.com/.default"
    Write-Host "Graph send endpoint           : https://graph.microsoft.com/v1.0/users/$($mailboxObject.ExternalDirectoryObjectId)/sendMail"

    if ($build.SecretCreated) {
        Write-Host ""
        Write-Host "CLIENT SECRET VALUE           :" -ForegroundColor Yellow
        Write-Host $build.SecretValue -ForegroundColor Yellow
        Write-Host "Secret lifetime               : $SecretValidityMonths month(s)"
        Write-Host "Secret expires                : $($build.SecretExpires)"
    }
    else {
        Write-Host ""
        Write-Host "Client secret                 : Existing secret retained; no new value created."
        Info "Existing secret values cannot be recovered from Entra. Use -SecretAction Always to create a new value."
    }

    Write-Host ""
    Pass "The requested mailbox is authorized."
    Pass "The Exchange scope resolves to exactly one recipient."
    Pass "Tenant-wide Microsoft Graph Mail.Send is not assigned by this script."
}
finally {
    try {
        Disconnect-ExchangeOnline `
            -Confirm:$false `
            -ErrorAction SilentlyContinue
    }
    catch {}

    if (Test-Path $work) {
        Remove-Item `
            -Path $work `
            -Recurse `
            -Force `
            -ErrorAction SilentlyContinue
    }
}

<#
.SYNOPSIS
    Audits SharePoint Online site collection permissions across web, library, list, folder,
    and item levels with broken inheritance detection, group expansion, and export capabilities.

.DESCRIPTION
    Audit-SharePoint-Site is an enterprise-grade PowerShell automation utility engineered for
    Microsoft 365 and SharePoint Online administrators and security auditors.

    The script performs deep inspection of permission architectures within SharePoint Online sites:
    - Extracts site-level role assignments including Owners, Members, Visitors, custom SharePoint groups,
      and Microsoft Entra ID (Azure AD) security groups.
    - Resolves SharePoint group memberships down to individual user principals, clearly distinguishing
      between direct user permissions and group-inherited access.
    - Identifies all Site Collection Administrators holding unconditional access across the site.
    - Inspects all document libraries and lists, evaluating inheritance status (HasUniqueRoleAssignments).
    - Recursively audits folders and items possessing broken permission inheritance.
    - Emits strongly-typed custom objects to the PowerShell pipeline for seamless integration.
    - Renders an interactive formatted console summary with color-coded severity tracking.
    - Exports audit datasets to standardized CSV or JSON formats.
    - Maintains enterprise-grade audit logging in C:\Logs\ with timestamped telemetry.
    - Features automatic host elevation: seamlessly relaunches in PowerShell 7+ (pwsh.exe) if invoked
      from Windows PowerShell 5.1.

.PARAMETER SiteUrl
    The absolute URL of the target SharePoint Online site collection (e.g. 'https://contoso.sharepoint.com/sites/finance').
    Accepts pipeline input by value.
    Mandatory when running in audit mode.

.PARAMETER ClientId
    The Application (Client) ID of the Microsoft Entra ID App Registration used for App-Only non-interactive authentication.

.PARAMETER Tenant
    The Microsoft 365 Tenant domain (e.g. 'contoso.onmicrosoft.com') or Tenant ID (GUID) used for App-Only authentication.

.PARAMETER CertificateThumbprint
    The thumbprint of an X.509 certificate installed in the local certificate store (Cert:\CurrentUser\My or Cert:\LocalMachine\My)
    used for secure certificate-based daemon authentication.

.PARAMETER Interactive
    Switch parameter. Forces an interactive modern web browser authentication prompt with MFA support via PnP.PowerShell.

.PARAMETER IncludeInherited
    Switch parameter. When specified, the audit captures all lists, libraries, folders, and items regardless of
    whether they inherit permissions or have broken inheritance. By default, only unique role assignments are audited
    for subordinate resources to optimize scan performance.

.PARAMETER ExportPath
    Optional file system path to export the audit records (e.g. 'C:\Reports\FinanceSitePermissions.csv').
    Parent directories will be created automatically if they do not exist.

.PARAMETER ExportFormat
    The file export format when -ExportPath is supplied.
    Valid options:
      - 'CSV'  (Comma-Separated Values, UTF-8 encoded)
      - 'JSON' (JavaScript Object Notation, indented, UTF-8 encoded)
    Default: 'CSV'.

.PARAMETER Force
    Switch parameter. Bypasses interactive confirmation prompts (such as module installation) for automated execution.

.PARAMETER Help
    Displays the built-in help documentation, parameter guidance, and usage examples, then terminates execution.
    Aliases: -h.

.EXAMPLE
    # 1. Display built-in help and practical examples (default when executed without parameters)
    .\Audit-SharePoint-Site.ps1
    .\Audit-SharePoint-Site.ps1 -Help
    .\Audit-SharePoint-Site.ps1 -h

.EXAMPLE
    # 2. Interactive site permission audit with console output and pipeline binding
    .\Audit-SharePoint-Site.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/finance" -Interactive

.EXAMPLE
    # 3. Non-interactive daemon audit using Entra ID App Registration and certificate authentication, exporting to CSV
    .\Audit-SharePoint-Site.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/finance" `
        -ClientId "11111111-2222-3333-4444-555555555555" `
        -Tenant "contoso.onmicrosoft.com" `
        -CertificateThumbprint "A1B2C3D4E5F67890123456789ABCDEF012345678" `
        -ExportPath "C:\Reports\FinancePermissionsAudit.csv" `
        -ExportFormat CSV

.EXAMPLE
    # 4. Deep comprehensive audit including all inherited lists, folders, and items exported to JSON
    .\Audit-SharePoint-Site.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/marketing" `
        -IncludeInherited `
        -ExportPath "C:\Reports\MarketingFullPermissions.json" `
        -ExportFormat JSON

.EXAMPLE
    # 5. Pipeline processing with Out-GridView
    "https://contoso.sharepoint.com/sites/hr" | .\Audit-SharePoint-Site.ps1 -Interactive | Out-GridView

.NOTES
    Author: Roman Pindela
    Email: roman.pindela@gmail.com
    GitHub: https://github.com/romanpindela
    Version: 1.0.0
    Last Modified: 2026-10-23
    Standards: Adheres to CleanCode and PowerShell Script Standards (PowerShell Script Standards.txt)
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$Tenant,

    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory = $false)]
    [switch]$Interactive,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeInherited,

    [Parameter(Mandatory = $false)]
    [string]$ExportPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet('CSV', 'JSON')]
    [string]$ExportFormat = 'CSV',

    [Parameter(Mandatory = $false)]
    [switch]$Force,

    [Parameter(Mandatory = $false)]
    [Alias('h')]
    [switch]$Help
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    # ==============================================================================
    # SCRIPT METADATA & TELEMETRY DECLARATION
    # ==============================================================================
    $Script:Metadata = @{
        Title        = "Audit-SharePoint-Site.ps1"
        Version      = "1.0.0"
        Author       = "Roman Pindela"
        Email        = "roman.pindela@gmail.com"
        GitHub       = "https://github.com/romanpindela"
        LastModified = "2026-10-23"
        RequiredMod  = "PnP.PowerShell"
    }

    $Script:StartTime        = Get-Date
    $Script:LogFilePath      = $null
    $Script:ActiveConnection = $null
    $Script:ScriptFilePath   = if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
        $PSCommandPath
    }
    elseif ($MyInvocation.MyCommand -and -not [string]::IsNullOrWhiteSpace($MyInvocation.MyCommand.Path)) {
        $MyInvocation.MyCommand.Path
    }
    else {
        $null
    }

    # ==============================================================================
    # POWERSHELL VERSION CHECK & AUTOMATIC RELAUNCH IN PWSH (PowerShell 7+)
    # ==============================================================================
    function Assert-PowerShellVersion {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param(
            [Parameter(Mandatory = $false)]
            [System.Collections.IDictionary]$BoundParameters
        )

        if ($PSVersionTable.PSVersion.Major -lt 7) {
            $pwsh = Get-Command pwsh.exe -ErrorAction SilentlyContinue
            if ($pwsh -and (-not [string]::IsNullOrWhiteSpace($Script:ScriptFilePath))) {
                Write-Host "[*] PnP.PowerShell requires PowerShell 7+ (Core). Current host: Windows PowerShell $($PSVersionTable.PSVersion)" -ForegroundColor Cyan
                Write-Host "[*] Automatically relaunching script in PowerShell 7 (pwsh.exe)...`n" -ForegroundColor Yellow

                $forwardArgs = [System.Collections.Generic.List[string]]::new()
                $escapedScriptPath = $Script:ScriptFilePath.Replace('"', '`"')
                $forwardArgs.Add("& `"$escapedScriptPath`"")

                if ($BoundParameters) {
                    foreach ($key in $BoundParameters.Keys) {
                        $val = $BoundParameters[$key]
                        if ($val -is [switch] -or $val -is [System.Management.Automation.SwitchParameter]) {
                            if ($val.IsPresent) {
                                $forwardArgs.Add("-$key")
                            }
                        }
                        elseif ($val -is [bool]) {
                            $forwardArgs.Add("-$key")
                            $forwardArgs.Add("`$$val")
                        }
                        else {
                            $forwardArgs.Add("-$key")
                            $escapedVal = $val.ToString().Replace('"', '`"')
                            $forwardArgs.Add("`"$escapedVal`"")
                        }
                    }
                }

                $cmdBlock = $forwardArgs -join " "
                $bytes = [System.Text.Encoding]::Unicode.GetBytes($cmdBlock)
                $encoded = [System.Convert]::ToBase64String($bytes)
                $process = Start-Process -FilePath $pwsh.Source -ArgumentList "-NoProfile -EncodedCommand $encoded" -Wait -PassThru -NoNewWindow
                exit $process.ExitCode
            }
            elseif (-not $pwsh) {
                $ps7Msg = @"
PnP.PowerShell requires PowerShell 7.4.0+ (pwsh.exe).
Current environment is Windows PowerShell $($PSVersionTable.PSVersion).
Please install PowerShell 7 from https://aka.ms/powershell or via:
    winget install --id Microsoft.PowerShell --source winget
"@
                Write-Host "[-] $ps7Msg" -ForegroundColor Red
                throw $ps7Msg
            }
        }
    }

    # ==============================================================================
    # LOGGING SYSTEM (CleanCode & Enterprise Standard)
    # ==============================================================================
    function Initialize-AuditLogging {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            [string]$TargetDirectory = 'C:\Logs'
        )

        $timestamp    = (Get-Date).ToString("yyyyMMdd-HHmmss")
        $computerName = $env:COMPUTERNAME
        $userName     = $env:USERNAME
        $logFileName  = "${timestamp}-Audit-SharePoint-Site-${computerName}-${userName}.logs"

        $resolvedDir = $TargetDirectory
        try {
            if (-not (Test-Path -Path $resolvedDir)) {
                $null = New-Item -Path $resolvedDir -ItemType Directory -Force -ErrorAction Stop
            }
        }
        catch {
            # Fallback to LocalAppData\Logs if C:\Logs is inaccessible or lacks permissions
            $resolvedDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath "Logs"
            if (-not (Test-Path -Path $resolvedDir)) {
                $null = New-Item -Path $resolvedDir -ItemType Directory -Force -ErrorAction SilentlyContinue
            }
        }

        $Script:LogFilePath = Join-Path -Path $resolvedDir -ChildPath $logFileName
        $header = @"
================================================================================
Audit Log Initialized : $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Script                : $($Script:Metadata.Title) v$($Script:Metadata.Version)
Author                : $($Script:Metadata.Author) ($($Script:Metadata.Email))
GitHub                : $($Script:Metadata.GitHub)
Host Machine          : $computerName
User Context          : $userName
Target Directory      : $resolvedDir
================================================================================
"@
        Add-Content -Path $Script:LogFilePath -Value $header -Encoding UTF8 -ErrorAction SilentlyContinue
    }

    function Write-Log {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '')]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Message,

            [Parameter(Mandatory = $true)]
            [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
            [string]$Level,

            [Parameter(Mandatory = $false)]
            [switch]$NoConsole
        )

        $timeStr = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $line    = "[$timeStr] [$($Level.PadRight(7))] $Message"

        if ($Script:LogFilePath) {
            try {
                Add-Content -Path $Script:LogFilePath -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
            }
            catch {
                Write-Verbose -Message "Unable to write log entry to file: $($_.Exception.Message)"
            }
        }

        if (-not $NoConsole) {
            $color = switch ($Level) {
                'INFO'    { 'Cyan' }
                'SUCCESS' { 'Green' }
                'WARN'    { 'Yellow' }
                'ERROR'   { 'Red' }
                Default   { 'White' }
            }

            $tag = switch ($Level) {
                'INFO'    { '[*]' }
                'SUCCESS' { '[+]' }
                'WARN'    { '[!]' }
                'ERROR'   { '[-]' }
                Default   { '   ' }
            }

            Write-Host "$tag $Message" -ForegroundColor $color
        }
    }

    # ==============================================================================
    # HELP & DOCUMENTATION (Standard 1, 7 & 14)
    # ==============================================================================
    function Show-ScriptHelp {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param()

        $helpLines = @(
            "================================================================================",
            "SCRIPT  : $($Script:Metadata.Title)",
            "VERSION : $($Script:Metadata.Version)",
            "AUTHOR  : $($Script:Metadata.Author)",
            "CONTACT : $($Script:Metadata.Email) | $($Script:Metadata.GitHub)",
            "UPDATED : $($Script:Metadata.LastModified)",
            "================================================================================",
            "",
            "DESCRIPTION:",
            "    Audit-SharePoint-Site is an enterprise permission auditing utility designed for",
            "    SharePoint Online and Microsoft 365. It scans site-level role assignments, resolves",
            "    SharePoint group memberships, enumerates Site Collection Administrators, and",
            "    recursively audits document libraries, lists, folders, and items with broken",
            "    permission inheritance (HasUniqueRoleAssignments).",
            "",
            "AUTHENTICATION MODES:",
            "    1. Interactive Modern Authentication:",
            "       - Requires an account with SharePoint Administrator or Site Admin rights.",
            "       - Uses web browser interactive sign-in with full MFA support:",
            "         .\Audit-SharePoint-Site.ps1 -SiteUrl 'https://contoso.sharepoint.com/sites/hr' -Interactive",
            "",
            "    2. Microsoft Entra ID App-Only Daemon Authentication (Certificate-Based):",
            "       - Register an application in Microsoft Entra ID.",
            "       - Grant SharePoint Online Application permissions: 'Sites.FullControl.All' or 'Sites.Read.All'.",
            "       - Upload X.509 Certificate and install it into local Cert store.",
            "       - Execute unattended audit:",
            "         .\Audit-SharePoint-Site.ps1 -SiteUrl 'https://contoso.sharepoint.com/sites/hr' ``",
            "             -ClientId '<AppId>' -Tenant 'contoso.onmicrosoft.com' ``",
            "             -CertificateThumbprint '<Thumbprint>' -ExportPath 'C:\Reports\Audit.csv'",
            "",
            "PREREQUISITES & ENVIRONMENT:",
            "    - Required Module : 'PnP.PowerShell' (Version 2.x+ or 3.x+ in PowerShell 7+)",
            "    - Execution Policy: Unblock-File -Path .\Audit-SharePoint-Site.ps1",
            "    - Log Directory   : C:\Logs (fallback to $env:LOCALAPPDATA\Logs)",
            "",
            "PARAMETERS:",
            "    -SiteUrl                [String]  Absolute URL of the SharePoint Online site collection (Mandatory).",
            "    -ClientId               [String]  Microsoft Entra ID Application (Client) ID.",
            "    -Tenant                 [String]  Tenant domain (e.g. contoso.onmicrosoft.com) or Tenant GUID.",
            "    -CertificateThumbprint  [String]  Certificate thumbprint for certificate-based daemon auth.",
            "    -Interactive            [Switch]  Prompts for interactive browser sign-in.",
            "    -IncludeInherited       [Switch]  Audits inherited lists/items in addition to broken inheritance.",
            "    -ExportPath             [String]  File path to export results (.csv or .json).",
            "    -ExportFormat           [String]  Export format: 'CSV' (default) or 'JSON'.",
            "    -Force                  [Switch]  Bypasses confirmation prompts (auto-installs missing modules).",
            "    -Help, -h               [Switch]  Displays this comprehensive help document.",
            "",
            "USAGE EXAMPLES:",
            "    1. Interactive scan of HR site with console output and pipeline binding:",
            "       .\Audit-SharePoint-Site.ps1 -SiteUrl 'https://contoso.sharepoint.com/sites/hr' -Interactive",
            "",
            "    2. Unattended certificate scan exporting unique permissions to CSV:",
            "       .\Audit-SharePoint-Site.ps1 -SiteUrl 'https://contoso.sharepoint.com/sites/finance' ``",
            "           -ClientId '00000000-0000-0000-0000-000000000000' ``",
            "           -Tenant 'contoso.onmicrosoft.com' ``",
            "           -CertificateThumbprint '1234567890ABCDEF1234567890ABCDEF12345678' ``",
            "           -ExportPath 'C:\Reports\FinancePermissions.csv' -ExportFormat CSV",
            "",
            "    3. Pipeline stream to Out-GridView:",
            "       'https://contoso.sharepoint.com/sites/marketing' | .\Audit-SharePoint-Site.ps1 -Interactive | Out-GridView",
            "================================================================================"
        )
        Write-Host ($helpLines -join "`n") -ForegroundColor Cyan
    }

    # ==============================================================================
    # INPUT DEFENSE & SANITIZATION (Standard 7)
    # ==============================================================================
    function Assert-SafeSiteUrl {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Url
        )

        if ([string]::IsNullOrWhiteSpace($Url)) {
            throw "Site URL cannot be empty or whitespace."
        }

        $cleanUrl = $Url.Trim()
        if ($cleanUrl.Length -gt 512) {
            throw "Site URL exceeds maximum allowed length of 512 characters."
        }

        # Defense against command injection and shell meta-characters
        if ($cleanUrl -match '[`$<>|;&\0]') {
            throw "Site URL contains forbidden characters or potential injection vectors."
        }

        # Strict SharePoint Online site URL structure validation
        if ($cleanUrl -notmatch '^https:\/\/[a-zA-Z0-9.-]+\.sharepoint\.com(\/sites\/[a-zA-Z0-9._-]+|\/teams\/[a-zA-Z0-9._-]+)?\/?$') {
            throw "Invalid SharePoint Online site URL: '$cleanUrl'. Expected format: 'https://<tenant>.sharepoint.com/sites/<site-name>'"
        }
    }

    # ==============================================================================
    # PREREQUISITE & MODULE VERIFICATION (WITH AUTO-INSTALLATION SUPPORT)
    # ==============================================================================
    function Assert-Prerequisite {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param(
            [switch]$ForceInstall
        )

        Write-Log -Level INFO -Message "Validating prerequisites and required module: $($Script:Metadata.RequiredMod)"

        $moduleInstalled = Get-Module -ListAvailable -Name $Script:Metadata.RequiredMod
        if (-not $moduleInstalled) {
            Write-Log -Level WARN -Message "Required PowerShell module '$($Script:Metadata.RequiredMod)' is not installed."

            $shouldInstall = $false
            if ($ForceInstall) {
                $shouldInstall = $true
            }
            elseif ([Environment]::UserInteractive) {
                Write-Host "Would you like to install '$($Script:Metadata.RequiredMod)' from PSGallery now? (Y/N) [T/N]: " -NoNewline -ForegroundColor Yellow
                $response = Read-Host
                if ($response -match '^[ytT1]$') {
                    $shouldInstall = $true
                }
            }

            if ($shouldInstall) {
                Write-Log -Level INFO -Message "Installing '$($Script:Metadata.RequiredMod)' from PSGallery for CurrentUser..."
                try {
                    Install-Module -Name $Script:Metadata.RequiredMod -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
                    Write-Log -Level SUCCESS -Message "Module '$($Script:Metadata.RequiredMod)' installed successfully."
                }
                catch {
                    $installErr = "Automatic installation of '$($Script:Metadata.RequiredMod)' failed: $($_.Exception.Message)"
                    Write-Log -Level ERROR -Message $installErr
                    throw $installErr
                }
            }
            else {
                $errorMessage = @"
Required PowerShell module '$($Script:Metadata.RequiredMod)' is not installed.
Please install it by executing:
    Install-Module -Name $($Script:Metadata.RequiredMod) -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
"@
                Write-Log -Level ERROR -Message $errorMessage
                throw $errorMessage
            }
        }

        $loaded = Get-Module -Name $Script:Metadata.RequiredMod
        if (-not $loaded) {
            Write-Log -Level INFO -Message "Importing $($Script:Metadata.RequiredMod)..."
            Import-Module -Name $Script:Metadata.RequiredMod -ErrorAction Stop
        }

        $currentMod = Get-Module -Name $Script:Metadata.RequiredMod
        Write-Log -Level SUCCESS -Message "Module '$($Script:Metadata.RequiredMod)' (v$($currentMod.Version)) verified successfully."
    }

    # ==============================================================================
    # CONNECTION MANAGEMENT
    # ==============================================================================
    function Connect-TargetSite {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$TargetSiteUrl,

            [Parameter(Mandatory = $false)]
            [string]$AppId,

            [Parameter(Mandatory = $false)]
            [string]$TenantDomain,

            [Parameter(Mandatory = $false)]
            [string]$CertThumbprint,

            [Parameter(Mandatory = $false)]
            [switch]$InteractiveAuth
        )

        Write-Log -Level INFO -Message "Connecting to SharePoint Online site: $TargetSiteUrl"

        try {
            # Scenario A: Certificate-based non-interactive authentication (Entra ID App-Only)
            if (-not [string]::IsNullOrWhiteSpace($AppId) -and
                -not [string]::IsNullOrWhiteSpace($TenantDomain) -and
                -not [string]::IsNullOrWhiteSpace($CertThumbprint)) {

                Write-Log -Level INFO -Message "Authenticating using Entra ID App-Only Certificate (AppId: $AppId, Tenant: $TenantDomain)..."
                $conn = Connect-PnPOnline -Url $TargetSiteUrl `
                    -ClientId $AppId `
                    -Tenant $TenantDomain `
                    -Thumbprint $CertThumbprint `
                    -ReturnConnection -ErrorAction Stop

                Write-Log -Level SUCCESS -Message "Connected successfully via App-Only Certificate."
                return $conn
            }

            # Scenario B & C: Modern Interactive Authentication (Explicit or Default)
            # In PnP.PowerShell 2.x and 3.x, -Interactive requires a ClientId for Entra ID app registration.
            # If not explicitly specified via -ClientId, fall back to the well-known PnP Management Shell AppId.
            $effectiveClientId = if (-not [string]::IsNullOrWhiteSpace($AppId)) {
                $AppId
            }
            else {
                "31359c7f-bd7e-475c-86db-f1e93e6075b9"
            }

            $interactiveParams = @{
                Url              = $TargetSiteUrl
                Interactive      = $true
                ClientId         = $effectiveClientId
                ReturnConnection = $true
                ErrorAction      = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace($TenantDomain)) {
                $interactiveParams['Tenant'] = $TenantDomain
            }

            $authLabel = if ($InteractiveAuth) { "Interactive browser" } else { "Default interactive" }
            Write-Log -Level INFO -Message "Initiating $authLabel authentication (ClientId: $effectiveClientId)..."
            $conn = Connect-PnPOnline @interactiveParams
            Write-Log -Level SUCCESS -Message "Connected successfully to $TargetSiteUrl via interactive session."
            return $conn
        }
        catch {
            $connError = "Failed to connect to SharePoint Online site '$TargetSiteUrl'. Details: $($_.Exception.Message)"
            Write-Log -Level ERROR -Message $connError
            throw $connError
        }
    }

    function Disconnect-TargetSite {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            $Connection
        )

        try {
            if ($null -ne $Connection) {
                Write-Log -Level INFO -Message "Disconnecting active SharePoint Online PnP session..."
                Disconnect-PnPOnline -Connection $Connection -ErrorAction SilentlyContinue
            }
            else {
                Disconnect-PnPOnline -ErrorAction SilentlyContinue
            }
            Write-Log -Level SUCCESS -Message "SharePoint Online session terminated cleanly."
        }
        catch {
            Write-Log -Level WARN -Message "Session disconnect warning: $($_.Exception.Message)"
        }
    }

    # ==============================================================================
    # PERMISSION AUDIT ENGINE
    # ==============================================================================
    function Get-SitePermissionsAudit {
        [CmdletBinding()]
        [OutputType([System.Collections.Generic.List[PSCustomObject]])]
        param(
            [Parameter(Mandatory = $true)]
            [string]$TargetSiteUrl,

            [Parameter(Mandatory = $false)]
            [switch]$AuditInherited,

            [Parameter(Mandatory = $false)]
            $Connection
        )

        $results = [System.Collections.Generic.List[PSCustomObject]]::new()
        $scanTimestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

        Write-Log -Level INFO -Message "Initiating deep permission scan on: $TargetSiteUrl"

        # --------------------------------------------------------------------------
        # 1. SITE-LEVEL WEB PERMISSIONS & ADMINISTRATORS
        # --------------------------------------------------------------------------
        Write-Progress -Activity "Auditing SharePoint Site Permissions" -Status "Auditing Site-Level Role Assignments..." -PercentComplete 10
        Write-Log -Level INFO -Message "Auditing site web role assignments and site collection administrators..."

        $web = if ($null -ne $Connection) {
            Get-PnPWeb -Connection $Connection -Includes Title, Url, HasUniqueRoleAssignments
        } else {
            Get-PnPWeb -Includes Title, Url, HasUniqueRoleAssignments
        }

        # Audit Site Collection Administrators
        try {
            $siteAdmins = if ($null -ne $Connection) {
                Get-PnPSiteCollectionAdmin -Connection $Connection -ErrorAction SilentlyContinue
            } else {
                Get-PnPSiteCollectionAdmin -ErrorAction SilentlyContinue
            }

            if ($siteAdmins) {
                foreach ($admin in $siteAdmins) {
                    $upn = if (-not [string]::IsNullOrWhiteSpace($admin.Email)) { $admin.Email } else { $admin.LoginName }
                    $adminRecord = [PSCustomObject]@{
                        Timestamp             = $scanTimestamp
                        SiteUrl               = $TargetSiteUrl
                        Scope                 = "Site"
                        ResourceType          = "Web"
                        ResourceName          = $web.Title
                        ResourceUrl           = $web.Url
                        HasUniquePermissions = $true
                        InheritanceStatus     = "Unique (Site Collection)"
                        PrincipalTitle        = $admin.Title
                        PrincipalType         = "SiteCollectionAdmin"
                        PrincipalLoginName    = $admin.LoginName
                        AssignmentType        = "Site Collection Administrator"
                        GroupName             = "N/A"
                        UserDisplayName       = $admin.Title
                        UserPrincipalName     = $upn
                        PermissionLevels      = "Full Control (Site Collection Administrator)"
                    }
                    $results.Add($adminRecord)
                }
                Write-Log -Level INFO -Message "Found $($siteAdmins.Count) Site Collection Administrator(s)."
            }
        }
        catch {
            Write-Log -Level WARN -Message "Unable to query Site Collection Administrators: $($_.Exception.Message)"
        }

        # Audit Web Role Assignments
        try {
            $null = Get-PnPProperty -ClientObject $web -Property RoleAssignments
            if ($web.RoleAssignments) {
                foreach ($assignment in $web.RoleAssignments) {
                    $null = Get-PnPProperty -ClientObject $assignment -Property Member, RoleDefinitionBindings
                    $member = $assignment.Member
                    $roleNames = ($assignment.RoleDefinitionBindings | ForEach-Object { $_.Name }) -join ", "
                    $pType = $member.PrincipalType.ToString()
                    $isGroup = ($pType -eq 'SharePointGroup')

                    if ($isGroup) {
                        # Group-level entry
                        $groupRecord = [PSCustomObject]@{
                            Timestamp             = $scanTimestamp
                            SiteUrl               = $TargetSiteUrl
                            Scope                 = "Site"
                            ResourceType          = "Web"
                            ResourceName          = $web.Title
                            ResourceUrl           = $web.Url
                            HasUniquePermissions = $web.HasUniqueRoleAssignments
                            InheritanceStatus     = if ($web.HasUniqueRoleAssignments) { "Unique" } else { "Inherited" }
                            PrincipalTitle        = $member.Title
                            PrincipalType         = "SharePointGroup"
                            PrincipalLoginName    = $member.LoginName
                            AssignmentType        = "Group Assignment"
                            GroupName             = $member.Title
                            UserDisplayName       = "N/A (Group Entity)"
                            UserPrincipalName     = "N/A"
                            PermissionLevels      = $roleNames
                        }
                        $results.Add($groupRecord)

                        # Expand members within the SharePoint Group
                        try {
                            $groupUsers = if ($null -ne $Connection) {
                                Get-PnPGroupMember -Group $member.Id -Connection $Connection -ErrorAction SilentlyContinue
                            } else {
                                Get-PnPGroupMember -Group $member.Id -ErrorAction SilentlyContinue
                            }

                            if ($groupUsers) {
                                foreach ($gUser in $groupUsers) {
                                    $userUpn = if (-not [string]::IsNullOrWhiteSpace($gUser.Email)) { $gUser.Email } else { $gUser.LoginName }
                                    $userRecord = [PSCustomObject]@{
                                        Timestamp             = $scanTimestamp
                                        SiteUrl               = $TargetSiteUrl
                                        Scope                 = "Site"
                                        ResourceType          = "Web"
                                        ResourceName          = $web.Title
                                        ResourceUrl           = $web.Url
                                        HasUniquePermissions = $web.HasUniqueRoleAssignments
                                        InheritanceStatus     = if ($web.HasUniqueRoleAssignments) { "Unique" } else { "Inherited" }
                                        PrincipalTitle        = $gUser.Title
                                        PrincipalType         = $gUser.PrincipalType.ToString()
                                        PrincipalLoginName    = $gUser.LoginName
                                        AssignmentType        = "Group Membership"
                                        GroupName             = $member.Title
                                        UserDisplayName       = $gUser.Title
                                        UserPrincipalName     = $userUpn
                                        PermissionLevels      = $roleNames
                                    }
                                    $results.Add($userRecord)
                                }
                            }
                        }
                        catch {
                            Write-Log -Level WARN -Message "Could not expand group members for '$($member.Title)': $($_.Exception.Message)"
                        }
                    }
                    else {
                        # Direct User or Security Group Assignment
                        $assignType = if ($pType -eq 'User') { "Direct Assignment" } else { "Security Group Assignment" }
                        $upn = if (-not [string]::IsNullOrWhiteSpace($member.Email)) { $member.Email } else { $member.LoginName }

                        $directRecord = [PSCustomObject]@{
                            Timestamp             = $scanTimestamp
                            SiteUrl               = $TargetSiteUrl
                            Scope                 = "Site"
                            ResourceType          = "Web"
                            ResourceName          = $web.Title
                            ResourceUrl           = $web.Url
                            HasUniquePermissions = $web.HasUniqueRoleAssignments
                            InheritanceStatus     = if ($web.HasUniqueRoleAssignments) { "Unique" } else { "Inherited" }
                            PrincipalTitle        = $member.Title
                            PrincipalType         = $pType
                            PrincipalLoginName    = $member.LoginName
                            AssignmentType        = $assignType
                            GroupName             = "N/A"
                            UserDisplayName       = $member.Title
                            UserPrincipalName     = $upn
                            PermissionLevels      = $roleNames
                        }
                        $results.Add($directRecord)
                    }
                }
            }
        }
        catch {
            Write-Log -Level WARN -Message "Failed to audit web role assignments: $($_.Exception.Message)"
        }

        # --------------------------------------------------------------------------
        # 2. DOCUMENT LIBRARIES AND LISTS
        # --------------------------------------------------------------------------
        Write-Progress -Activity "Auditing SharePoint Site Permissions" -Status "Discovering Lists and Document Libraries..." -PercentComplete 30
        Write-Log -Level INFO -Message "Discovering document libraries and lists..."

        $lists = if ($null -ne $Connection) {
            Get-PnPList -Connection $Connection -ErrorAction Stop | Where-Object {
                -not $_.Hidden -and $_.Title -ne "TaxonomyHiddenList" -and $_.Title -ne "Style Library"
            }
        } else {
            Get-PnPList -ErrorAction Stop | Where-Object {
                -not $_.Hidden -and $_.Title -ne "TaxonomyHiddenList" -and $_.Title -ne "Style Library"
            }
        }

        $listCount = 0
        $totalLists = @($lists).Count
        Write-Log -Level INFO -Message "Discovered $totalLists list(s) and library(ies) to inspect."

        foreach ($list in $lists) {
            $listCount++
            $percentLists = 30 + [int](($listCount / [Math]::Max($totalLists, 1)) * 30)
            Write-Progress -Activity "Auditing SharePoint Site Permissions" `
                -Status "Auditing List $($listCount) of $($totalLists) - $($list.Title)" `
                -PercentComplete $percentLists

            $isUniqueList = $list.HasUniqueRoleAssignments
            $resType = if ($list.BaseTemplate -eq 101) { "DocumentLibrary" } else { "List" }

            if ($isUniqueList -or $AuditInherited) {
                $inheritStatus = if ($isUniqueList) { "Broken (Unique)" } else { "Inherited" }
                if ($isUniqueList) {
                    Write-Log -Level WARN -Message "Unique permissions detected on ${resType} - '$($list.Title)'"
                }

                try {
                    $null = Get-PnPProperty -ClientObject $list -Property RoleAssignments -ErrorAction SilentlyContinue
                    if ($list.RoleAssignments) {
                        foreach ($lAssignment in $list.RoleAssignments) {
                            $null = Get-PnPProperty -ClientObject $lAssignment -Property Member, RoleDefinitionBindings
                            $lMember = $lAssignment.Member
                            $lRoleNames = ($lAssignment.RoleDefinitionBindings | ForEach-Object { $_.Name }) -join ", "
                            $pType = $lMember.PrincipalType.ToString()
                            $assignType = if ($pType -eq 'SharePointGroup') { "Group Assignment" } elseif ($pType -eq 'User') { "Direct Assignment" } else { "Security Group Assignment" }
                            $upn = if (-not [string]::IsNullOrWhiteSpace($lMember.Email)) { $lMember.Email } else { $lMember.LoginName }

                            $listRecord = [PSCustomObject]@{
                                Timestamp             = $scanTimestamp
                                SiteUrl               = $TargetSiteUrl
                                Scope                 = "List"
                                ResourceType          = $resType
                                ResourceName          = $list.Title
                                ResourceUrl           = $list.DefaultViewUrl
                                HasUniquePermissions = $isUniqueList
                                InheritanceStatus     = $inheritStatus
                                PrincipalTitle        = $lMember.Title
                                PrincipalType         = $pType
                                PrincipalLoginName    = $lMember.LoginName
                                AssignmentType        = $assignType
                                GroupName             = if ($pType -eq 'SharePointGroup') { $lMember.Title } else { "N/A" }
                                UserDisplayName       = $lMember.Title
                                UserPrincipalName     = $upn
                                PermissionLevels      = $lRoleNames
                            }
                            $results.Add($listRecord)
                        }
                    }
                }
                catch {
                    Write-Log -Level WARN -Message "Could not retrieve role assignments for list '$($list.Title)': $($_.Exception.Message)"
                }
            }

            # ----------------------------------------------------------------------
            # 3. FOLDERS AND LIST ITEMS WITH BROKEN INHERITANCE
            # ----------------------------------------------------------------------
            try {
                $items = if ($null -ne $Connection) {
                    Get-PnPListItem -List $list -PageSize 500 -Fields "Id", "Title", "FileRef", "FileSystemObjectType", "HasUniqueRoleAssignments" -Connection $Connection -ErrorAction SilentlyContinue
                } else {
                    Get-PnPListItem -List $list -PageSize 500 -Fields "Id", "Title", "FileRef", "FileSystemObjectType", "HasUniqueRoleAssignments" -ErrorAction SilentlyContinue
                }

                if ($items) {
                    $targetItems = if ($AuditInherited) {
                        $items
                    } else {
                        $items | Where-Object { $_.HasUniqueRoleAssignments -eq $true }
                    }

                    $uniqueItemCount = @($targetItems).Count
                    if ($uniqueItemCount -gt 0) {
                        Write-Log -Level INFO -Message "Auditing $uniqueItemCount item(s)/folder(s) in '$($list.Title)'..."
                    }

                    foreach ($item in $targetItems) {
                        $itemUnique = $item.HasUniqueRoleAssignments
                        $itemType = if ($item.FileSystemObjectType -eq 'Folder') { "Folder" } else { "Item/File" }
                        $itemRef = if (-not [string]::IsNullOrWhiteSpace($item["FileRef"])) { $item["FileRef"] } else { "$($list.Title) - Item #$($item.Id)" }
                        $itemInheritStatus = if ($itemUnique) { "Broken (Unique)" } else { "Inherited" }

                        try {
                            $itemPerms = if ($null -ne $Connection) {
                                Get-PnPListItemPermission -List $list -Identity $item.Id -Connection $Connection -ErrorAction SilentlyContinue
                            } else {
                                Get-PnPListItemPermission -List $list -Identity $item.Id -ErrorAction SilentlyContinue
                            }

                            if ($itemPerms) {
                                foreach ($perm in $itemPerms) {
                                    $pType = if ($perm.PrincipalType) { $perm.PrincipalType.ToString() } else { "User" }
                                    $iRoleNames = ($perm.PermissionLevels | ForEach-Object { $_.Name }) -join ", "
                                    $assignType = if ($pType -eq 'SharePointGroup') { "Group Assignment" } elseif ($pType -eq 'User') { "Direct Assignment" } else { "Security Group Assignment" }

                                    $itemRecord = [PSCustomObject]@{
                                        Timestamp             = $scanTimestamp
                                        SiteUrl               = $TargetSiteUrl
                                        Scope                 = "Item"
                                        ResourceType          = $itemType
                                        ResourceName          = $itemRef
                                        ResourceUrl           = $itemRef
                                        HasUniquePermissions = $itemUnique
                                        InheritanceStatus     = $itemInheritStatus
                                        PrincipalTitle        = $perm.PrincipalName
                                        PrincipalType         = $pType
                                        PrincipalLoginName    = $perm.PrincipalName
                                        AssignmentType        = $assignType
                                        GroupName             = if ($pType -eq 'SharePointGroup') { $perm.PrincipalName } else { "N/A" }
                                        UserDisplayName       = $perm.PrincipalName
                                        UserPrincipalName     = $perm.PrincipalName
                                        PermissionLevels      = $iRoleNames
                                    }
                                    $results.Add($itemRecord)
                                }
                            }
                        }
                        catch {
                            Write-Log -Level WARN -Message "Could not retrieve role assignments for item $($item.Id) in '$($list.Title)': $($_.Exception.Message)"
                        }
                    }
                }
            }
            catch {
                Write-Log -Level WARN -Message "Failed to query items in list '$($list.Title)': $($_.Exception.Message)"
            }
        }

        Write-Progress -Activity "Auditing SharePoint Site Permissions" -Completed
        Write-Log -Level SUCCESS -Message "Permission audit completed. Discovered $($results.Count) permission record(s)."
        return $results
    }

    # ==============================================================================
    # AUDIT SUMMARY & CONSOLE REPORTING
    # ==============================================================================
    function Show-AuditSummary {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param(
            [Parameter(Mandatory = $true)]
            [System.Collections.Generic.List[PSCustomObject]]$AuditRecords
        )

        $totalRecords = $AuditRecords.Count
        $siteAdmins   = @($AuditRecords | Where-Object { $_.AssignmentType -eq "Site Collection Administrator" }).Count
        $uniqueSites  = @($AuditRecords | Where-Object { $_.Scope -eq "Site" }).Count
        $uniqueLists  = @($AuditRecords | Where-Object { $_.Scope -eq "List" -and $_.HasUniquePermissions }).Count
        $uniqueUnique = @($AuditRecords | Where-Object { $_.Scope -eq "Item" -and $_.HasUniquePermissions }).Count
        $directAssign = @($AuditRecords | Where-Object { $_.AssignmentType -eq "Direct Assignment" }).Count
        $groupAssign  = @($AuditRecords | Where-Object { $_.AssignmentType -like "*Group*" }).Count

        Write-Host "`n================================================================================" -ForegroundColor Cyan
        Write-Host " SHAREPOINT ONLINE PERMISSION AUDIT SUMMARY" -ForegroundColor Cyan
        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host " Total Permission Entries Extracted   : $totalRecords" -ForegroundColor White
        Write-Host " Site Collection Administrators       : $siteAdmins" -ForegroundColor White
        Write-Host " Site-Level Entitlements Evaluated    : $uniqueSites" -ForegroundColor White
        Write-Host " Lists/Libraries with Broken Perms    : $uniqueLists" -ForegroundColor $(if ($uniqueLists -gt 0) { 'Yellow' } else { 'White' })
        Write-Host " Folders/Items with Broken Perms      : $uniqueUnique" -ForegroundColor $(if ($uniqueUnique -gt 0) { 'Yellow' } else { 'White' })
        Write-Host " Direct User Assignments (High Risk)  : $directAssign" -ForegroundColor $(if ($directAssign -gt 0) { 'Yellow' } else { 'White' })
        Write-Host " Group-Based Access Assignments       : $groupAssign" -ForegroundColor White
        Write-Host "================================================================================`n" -ForegroundColor Cyan
    }

    # ==============================================================================
    # EXPORT DATA ENGINE
    # ==============================================================================
    function Export-AuditData {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [System.Collections.Generic.List[PSCustomObject]]$Records,

            [Parameter(Mandatory = $true)]
            [string]$Path,

            [Parameter(Mandatory = $true)]
            [ValidateSet('CSV', 'JSON')]
            [string]$Format
        )

        Write-Log -Level INFO -Message "Exporting $($Records.Count) audit record(s) to $Format file: $Path"

        try {
            $parentDir = Split-Path -Path $Path -Parent
            if (-not [string]::IsNullOrWhiteSpace($parentDir) -and -not (Test-Path -Path $parentDir)) {
                $null = New-Item -Path $parentDir -ItemType Directory -Force -ErrorAction Stop
            }

            if ($Format -eq 'CSV') {
                $Records | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Force
            }
            else {
                $jsonContent = $Records | ConvertTo-Json -Depth 6
                [System.IO.File]::WriteAllText($Path, $jsonContent, [System.Text.UTF8Encoding]::new($false))
            }

            $fileSize = (Get-Item -Path $Path).Length
            Write-Log -Level SUCCESS -Message "Export completed successfully. File size: $fileSize bytes ($Path)"
        }
        catch {
            $exportErr = "Failed to export audit data to '$Path'. Details: $($_.Exception.Message)"
            Write-Log -Level ERROR -Message $exportErr
            throw $exportErr
        }
    }

    # --------------------------------------------------------------------------
    # SCRIPT INITIALIZATION CHECKS & HELP ROUTING
    # --------------------------------------------------------------------------
    $isPiped = $MyInvocation.ExpectingInput
    if ($Help -or (($PSBoundParameters.Count -eq 0) -and -not $isPiped)) {
        Show-ScriptHelp
        exit 0
    }

    # Auto-elevate to PowerShell 7+ if running under legacy Windows PowerShell 5.1
    Assert-PowerShellVersion -BoundParameters $PSBoundParameters

    Initialize-AuditLogging
    Write-Log -Level INFO -Message "Starting $($Script:Metadata.Title) v$($Script:Metadata.Version)"
    Write-Log -Level INFO -Message "Log initialized at: $Script:LogFilePath"

    Assert-Prerequisite -ForceInstall:$Force
}

process {
    if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
        Write-Log -Level ERROR -Message "SiteUrl parameter is required for audit execution."
        Show-ScriptHelp
        exit 0
    }

    Assert-SafeSiteUrl -Url $SiteUrl

    if ($PSCmdlet.ShouldProcess($SiteUrl, "Audit SharePoint Online Permissions")) {
        $auditData = $null
        try {
            $Script:ActiveConnection = Connect-TargetSite `
                -TargetSiteUrl $SiteUrl `
                -AppId $ClientId `
                -TenantDomain $Tenant `
                -CertThumbprint $CertificateThumbprint `
                -InteractiveAuth:$Interactive

            $auditData = Get-SitePermissionsAudit `
                -TargetSiteUrl $SiteUrl `
                -AuditInherited:$IncludeInherited `
                -Connection $Script:ActiveConnection

            Show-AuditSummary -AuditRecords $auditData

            if (-not [string]::IsNullOrWhiteSpace($ExportPath)) {
                Export-AuditData -Records $auditData -Path $ExportPath -Format $ExportFormat
            }

            # Emit structured objects to the PowerShell pipeline
            if ($auditData) {
                foreach ($record in $auditData) {
                    Write-Output $record
                }
            }
        }
        catch {
            Write-Log -Level ERROR -Message "Audit execution encountered a critical error: $($_.Exception.Message)"
            throw
        }
        finally {
            Disconnect-TargetSite -Connection $Script:ActiveConnection
            $Script:ActiveConnection = $null
        }
    }
}

end {
    $elapsed = (Get-Date) - $Script:StartTime
    Write-Log -Level INFO -Message "Execution completed in $($elapsed.ToString('hh\:mm\:ss\.fff'))."
    Write-Log -Level INFO -Message "Audit log persisted at: $Script:LogFilePath"
}

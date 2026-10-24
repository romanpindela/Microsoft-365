<#
.SYNOPSIS
    Enterprise-grade audit utility that evaluates SharePoint Online tenant site collections
    for external sharing configurations, site administrators, access governance, and security posture.

.DESCRIPTION
    Audit-ExchangeSecuritySummary is an enterprise security assessment automation utility engineered for
    Microsoft 365 and SharePoint Online administrators, security engineers, and compliance auditors.

    The script performs deep inspection across tenant site collections to evaluate security configurations:
    - Enumerates SharePoint Online site collections (with optional filtering for personal/OneDrive and system sites).
    - Audits external sharing policies (Disabled, ExistingGuests, ExternalUserSharingOnly, ExternalUserAndGuestSharing).
    - Checks access governance controls including anonymous 'Anyone' links, default sharing link types, and
      whether non-owners are permitted to invite external collaborators (MembersCanShare / DisableSharingForNonOwnersStatus).
    - Inventories Site Collection Administrators (Primary, Secondary, and elevated admins) to identify admin sprawl.
    - Inspects Conditional Access device policies (AllowFullAccess, AllowLimitedAccess, BlockAccess) and unmanaged device downloads.
    - Analyzes customization restrictions (DenyAddAndCustomizePages) to prevent malicious script injection.
    - Computes a weighted, multi-factor Security Posture Score (0-100) and assigns an actionable Risk Tier.
    - Renders a rich, color-coded executive console dashboard and KPI matrix.
    - Exports comprehensive audit datasets to standardized CSV or indented JSON artifacts.
    - Features automated throttling resilience (HTTP 429 / 503 backoff logic with exponential jitter).
    - Maintains enterprise-grade audit logging in C:\Logs\ with timestamped telemetry.
    - Enforces automatic PowerShell 7+ host elevation (relaunching from Windows PowerShell 5.1).

.PARAMETER SiteUrl
    The absolute URL of a specific SharePoint Online site collection to audit
    (e.g., 'https://contoso.sharepoint.com/sites/finance').
    Accepts pipeline input by value. If omitted, the script audits all site collections across the tenant.

.PARAMETER AdminUrl
    The SharePoint Online Admin Center URL (e.g., 'https://contoso-admin.sharepoint.com').
    If omitted, auto-derived from -TenantId, -SiteUrl, or discovered during interactive sign-in.

.PARAMETER TenantId
    The Microsoft 365 tenant domain (e.g., 'contoso.onmicrosoft.com' or 'contoso') or Tenant GUID.
    Aliases: -Tenant.

.PARAMETER ClientId
    The Application (Client) ID of the Microsoft Entra ID App Registration used for authentication.

.PARAMETER CertificateThumbprint
    The thumbprint of an X.509 certificate installed in the local certificate store (Cert:\CurrentUser\My or Cert:\LocalMachine\My)
    used for secure certificate-based daemon authentication.

.PARAMETER ClientSecret
    SecureString containing the client secret for Microsoft Entra ID App Registration authentication.

.PARAMETER Interactive
    Switch parameter. Forces an interactive modern web browser authentication prompt with MFA support via PnP.PowerShell.

.PARAMETER Search
    Optional keyword to filter site collections by URL or Title.

.PARAMETER IncludePersonalSites
    Switch parameter. Includes personal sites (OneDrive for Business site collections) in the tenant audit.
    By default, personal sites are excluded to focus on collaborative team and communication sites.

.PARAMETER IncludeSystemSites
    Switch parameter. Includes internal system sites (such as App Catalog, PointPublishing Hub, Search Center) in the audit.
    By default, internal system templates are excluded to prevent audit skew.

.PARAMETER ExportPath
    Optional file system path to export the audit records (e.g., 'C:\Reports\SharePointSecurityAssessment.csv').
    Parent directories will be created automatically if they do not exist.

.PARAMETER ExportFormat
    The file export format when -ExportPath is supplied.
    Valid options:
      - 'CSV'  (Comma-Separated Values, UTF-8 encoded)
      - 'JSON' (JavaScript Object Notation, indented, UTF-8 encoded)
    Default: 'CSV'.

.PARAMETER PassThru
    Switch parameter. Returns strongly-typed [PSCustomObject] records to the PowerShell pipeline for downstream processing.

.PARAMETER Force
    Switch parameter. Bypasses interactive confirmation prompts (such as module installation) for automated execution.

.PARAMETER Help
    Displays the built-in help documentation, parameter guidance, and usage examples, then terminates execution.
    Aliases: -h.

.EXAMPLE
    # 1. Display built-in help and practical examples (default when executed without parameters)
    .\Audit-ExchangeSecuritySummary.ps1
    .\Audit-ExchangeSecuritySummary.ps1 -Help
    .\Audit-ExchangeSecuritySummary.ps1 -h

.EXAMPLE
    # 2. Interactive full tenant security posture audit with rich console metrics
    .\Audit-ExchangeSecuritySummary.ps1 -TenantId "contoso.onmicrosoft.com" -Interactive

.EXAMPLE
    # 3. Targeted audit of a single high-impact site collection
    .\Audit-ExchangeSecuritySummary.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/finance" -Interactive

.EXAMPLE
    # 4. Search and audit sites matching a keyword, exporting results to CSV
    .\Audit-ExchangeSecuritySummary.ps1 -TenantId "contoso.onmicrosoft.com" -Search "Project" `
        -ExportPath "C:\Reports\ProjectSitesSecurity.csv" -ExportFormat CSV -Interactive

.EXAMPLE
    # 5. Unattended daemon audit using Entra ID App Registration with certificate authentication
    .\Audit-ExchangeSecuritySummary.ps1 -AdminUrl "https://contoso-admin.sharepoint.com" `
        -ClientId "11111111-2222-3333-4444-555555555555" `
        -TenantId "contoso.onmicrosoft.com" `
        -CertificateThumbprint "A1B2C3D4E5F67890123456789ABCDEF012345678" `
        -ExportPath "C:\Reports\TenantSecuritySummary.json" `
        -ExportFormat JSON

.EXAMPLE
    # 6. Unattended daemon audit using Entra ID App Registration with Client Secret
    $secret = Read-Host -AsSecureString "Enter Client Secret"
    .\Audit-ExchangeSecuritySummary.ps1 -AdminUrl "https://contoso-admin.sharepoint.com" `
        -ClientId "11111111-2222-3333-4444-555555555555" `
        -TenantId "contoso.onmicrosoft.com" `
        -ClientSecret $secret `
        -ExportPath "C:\Reports\TenantSecuritySummary.csv"

.EXAMPLE
    # 7. Pipeline stream to Out-GridView for interactive filtering
    .\Audit-ExchangeSecuritySummary.ps1 -TenantId "contoso.onmicrosoft.com" -Interactive -PassThru | Out-GridView

.NOTES
    Script Name   : Audit-ExchangeSecuritySummary.ps1
    Author        : Roman Pindela
    Email         : roman.pindela@gmail.com
    GitHub        : https://github.com/romanpindela
    Version       : 1.0.0
    Last Modified : 2026-10-24
    Standards     : Adheres to CleanCode and PowerShell Script Standards (PowerShell Script Standards.txt)
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$AdminUrl,

    [Parameter(Mandatory = $false)]
    [Alias('Tenant')]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory = $false)]
    [System.Security.SecureString]$ClientSecret,

    [Parameter(Mandatory = $false)]
    [switch]$Interactive,

    [Parameter(Mandatory = $false)]
    [string]$Search,

    [Parameter(Mandatory = $false)]
    [switch]$IncludePersonalSites,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeSystemSites,

    [Parameter(Mandatory = $false)]
    [string]$ExportPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet('CSV', 'JSON')]
    [string]$ExportFormat = 'CSV',

    [Parameter(Mandatory = $false)]
    [switch]$PassThru,

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
        Title        = "Audit-ExchangeSecuritySummary.ps1"
        Description  = "SharePoint Online Tenant Security Posture Assessment & Sharing Governance Audit"
        Version      = "1.0.0"
        Author       = "Roman Pindela"
        Email        = "roman.pindela@gmail.com"
        GitHub       = "https://github.com/romanpindela"
        LastModified = "2026-10-24"
        RequiredMod  = "PnP.PowerShell"
    }

    $Script:StartTime        = Get-Date
    $Script:LogFilePath      = $null
    $Script:ActiveConnection = $null
    $Script:ActiveClientId   = $null
    $Script:ActiveTenant     = $null
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
    # HELP & DOCUMENTATION (Invoked on -Help or Zero Parameters)
    # ==============================================================================
    function Show-ScriptHelp {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param()

        $helpLines = @(
            "================================================================================",
            "SCRIPT  : $($Script:Metadata.Title)",
            "DESC    : $($Script:Metadata.Description)",
            "VERSION : $($Script:Metadata.Version)",
            "AUTHOR  : $($Script:Metadata.Author)",
            "CONTACT : $($Script:Metadata.Email) | $($Script:Metadata.GitHub)",
            "UPDATED : $($Script:Metadata.LastModified)",
            "================================================================================",
            "",
            "DESCRIPTION:",
            "    Audit-ExchangeSecuritySummary performs an enterprise-grade security audit across",
            "    SharePoint Online site collections. It extracts external sharing configurations,",
            "    access governance controls (anonymous links, MembersCanShare), catalogs Site Collection",
            "    Administrators, evaluates Conditional Access restrictions, and computes a multi-factor",
            "    weighted Security Posture Score (0-100) with actionable risk tiers.",
            "",
            "AUTHENTICATION MODES & SUPPORTED CREDENTIALS:",
            "    1. Interactive Modern Authentication (Browser + MFA):",
            "       .\Audit-ExchangeSecuritySummary.ps1 -TenantId 'contoso.onmicrosoft.com' -Interactive",
            "",
            "    2. Microsoft Entra ID App-Only Daemon (X.509 Certificate):",
            "       .\Audit-ExchangeSecuritySummary.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com' ``",
            "           -ClientId '<AppId>' -TenantId 'contoso.onmicrosoft.com' ``",
            "           -CertificateThumbprint '<Thumbprint>' -ExportPath 'C:\Reports\Audit.csv'",
            "",
            "    3. Microsoft Entra ID App-Only Daemon (Client Secret):",
            "       .\Audit-ExchangeSecuritySummary.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com' ``",
            "           -ClientId '<AppId>' -TenantId 'contoso.onmicrosoft.com' ``",
            "           -ClientSecret `$secureSecret -ExportPath 'C:\Reports\Audit.json' -ExportFormat JSON",
            "",
            "PREREQUISITES & ENVIRONMENT:",
            "    - Required Role   : SharePoint Administrator or Global Administrator in Microsoft 365",
            "    - Required Module : 'PnP.PowerShell' (Version 2.x+ or 3.x+ in PowerShell 7+)",
            "    - Execution Policy: Unblock-File -Path .\Audit-ExchangeSecuritySummary.ps1",
            "    - Log Directory   : C:\Logs\ (fallback to $env:LOCALAPPDATA\Logs)",
            "",
            "PARAMETERS:",
            "    -SiteUrl                [String]       URL of a single site to audit (audits tenant if omitted).",
            "    -AdminUrl               [String]       SharePoint Online Admin Center URL.",
            "    -TenantId, -Tenant      [String]       Tenant domain (e.g. contoso.onmicrosoft.com) or GUID.",
            "    -ClientId               [String]       Microsoft Entra ID Application (Client) ID.",
            "    -CertificateThumbprint  [String]       Certificate thumbprint for certificate-based daemon auth.",
            "    -ClientSecret           [SecureString] Client secret for application authentication.",
            "    -Interactive            [Switch]       Forces interactive browser authentication prompt.",
            "    -Search                 [String]       Keyword filter matching site URL or Title.",
            "    -IncludePersonalSites   [Switch]       Includes OneDrive for Business personal sites in audit.",
            "    -IncludeSystemSites     [Switch]       Includes system templates (AppCatalog, Hubs, Search).",
            "    -ExportPath             [String]       File path to export results (.csv or .json).",
            "    -ExportFormat           [String]       Export format: 'CSV' (default) or 'JSON'.",
            "    -PassThru               [Switch]       Emits structured PSCustomObject items to the pipeline.",
            "    -Force                  [Switch]       Bypasses confirmation prompts (auto-installs module).",
            "    -Help, -h               [Switch]       Displays this comprehensive help document.",
            "",
            "USAGE EXAMPLES:",
            "    1. Interactive full tenant security posture scan:",
            "       .\Audit-ExchangeSecuritySummary.ps1 -TenantId 'contoso.onmicrosoft.com' -Interactive",
            "",
            "    2. Audit a specific site collection:",
            "       .\Audit-ExchangeSecuritySummary.ps1 -SiteUrl 'https://contoso.sharepoint.com/sites/legal' -Interactive",
            "",
            "    3. Search sites matching 'Finance' and export report to CSV:",
            "       .\Audit-ExchangeSecuritySummary.ps1 -TenantId 'contoso.onmicrosoft.com' -Search 'Finance' ``",
            "           -ExportPath 'C:\Reports\FinanceSitesSecurity.csv' -Interactive",
            "",
            "    4. Automated unattended daemon execution with certificate:",
            "       .\Audit-ExchangeSecuritySummary.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com' ``",
            "           -ClientId '00000000-0000-0000-0000-000000000000' ``",
            "           -TenantId 'contoso.onmicrosoft.com' ``",
            "           -CertificateThumbprint '1234567890ABCDEF1234567890ABCDEF12345678' ``",
            "           -ExportPath 'C:\Reports\TenantSecurityReport.json' -ExportFormat JSON",
            "",
            "    5. Stream audit objects to Out-GridView:",
            "       .\Audit-ExchangeSecuritySummary.ps1 -TenantId 'contoso.onmicrosoft.com' -Interactive -PassThru | Out-GridView",
            "================================================================================"
        )
        Write-Host ($helpLines -join "`n") -ForegroundColor Cyan
    }

    # Intercept zero parameters or explicit -Help switch
    if ($Help -or ($PSBoundParameters.Count -eq 0)) {
        Show-ScriptHelp
        exit 0
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
                        elseif ($val -is [System.Security.SecureString]) {
                            # SecureString cannot be forwarded via command line arguments safely
                            Write-Host "[!] Cannot forward SecureString parameter -$key across process boundary. Please execute directly in pwsh." -ForegroundColor Red
                            throw "SecureString parameter cannot be passed across process elevation boundary."
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
        $logFileName  = "${timestamp}-Audit-ExchangeSecuritySummary-${computerName}-${userName}.logs"

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
Log File Location     : $($Script:LogFilePath)
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
            [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS', 'DEBUG')]
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
                'DEBUG'   { 'DarkGray' }
                Default   { 'White' }
            }

            $tag = switch ($Level) {
                'INFO'    { '[*]' }
                'SUCCESS' { '[+]' }
                'WARN'    { '[!]' }
                'ERROR'   { '[-]' }
                'DEBUG'   { '[D]' }
                Default   { '   ' }
            }

            Write-Host "$tag $Message" -ForegroundColor $color
        }
    }

    # ==============================================================================
    # INPUT DEFENSE & SANITIZATION
    # ==============================================================================
    function Assert-SafeUrl {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Url,

            [Parameter(Mandatory = $false)]
            [switch]$IsAdminUrl
        )

        if ([string]::IsNullOrWhiteSpace($Url)) {
            throw "Target URL cannot be null, empty, or whitespace."
        }

        $clean = $Url.Trim()
        if ($clean.Length -gt 512) {
            throw "Target URL exceeds maximum allowed length of 512 characters."
        }

        # Defense against shell injection vectors
        if ($clean -match '[`$<>|;&\0]') {
            throw "Target URL contains forbidden characters or potential injection vectors."
        }

        if ($IsAdminUrl) {
            if ($clean -notmatch '^https:\/\/[a-zA-Z0-9.-]+-admin\.sharepoint\.com\/?$') {
                throw "Invalid SharePoint Online Admin URL: '$clean'. Expected format: 'https://<tenant>-admin.sharepoint.com'"
            }
        }
        else {
            if ($clean -notmatch '^https:\/\/[a-zA-Z0-9.-]+\.sharepoint\.com(\/sites\/[a-zA-Z0-9._-]+|\/teams\/[a-zA-Z0-9._-]+)?\/?$') {
                throw "Invalid SharePoint Online site URL: '$clean'. Expected format: 'https://<tenant>.sharepoint.com/sites/<name>'"
            }
        }
    }

    function Assert-SafeThumbprint {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Thumbprint
        )

        $clean = ($Thumbprint -replace '\s+', '').Trim()
        if ($clean -notmatch '^[0-9a-fA-F]{40}$') {
            throw "Invalid Certificate Thumbprint: '$Thumbprint'. Expected 40 hexadecimal characters."
        }
        return $clean
    }

    # ==============================================================================
    # PREREQUISITE & MODULE VERIFICATION
    # ==============================================================================
    function Assert-Prerequisite {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param(
            [switch]$ForceInstall
        )

        Write-Log -Level INFO -Message "Verifying required PowerShell module: $($Script:Metadata.RequiredMod)"

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
        Write-Log -Level SUCCESS -Message "Module '$($Script:Metadata.RequiredMod)' (v$($currentMod.Version)) loaded and verified."
    }

    # ==============================================================================
    # AUTOMATED ENTRA ID APP REGISTRATION HELPER
    # ==============================================================================
    function Invoke-EntraIDAppRegistration {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Tenant,

            [Parameter(Mandatory = $false)]
            [string]$ApplicationName = "SharePoint-Security-Audit"
        )

        $targetAppName = $ApplicationName
        $maxAttempts   = 3
        $attempt       = 0

        while ($attempt -lt $maxAttempts) {
            $attempt++
            Write-Log -Level INFO -Message "Registering Microsoft Entra ID App ('$targetAppName') in tenant '$Tenant'..."
            Write-Host "`n[*] Microsoft Entra ID registration requires modern administrator consent flow." -ForegroundColor Yellow
            Write-Host "[*] A browser window will open. Please sign in with a Global/SharePoint Admin account.`n" -ForegroundColor Yellow

            try {
                $regResult = Register-PnPEntraIDAppForInteractiveLogin `
                    -ApplicationName $targetAppName `
                    -Tenant $Tenant `
                    -SharePointDelegatePermissions "AllSites.FullControl" `
                    -GraphDelegatePermissions "Group.Read.All", "User.Read.All" `
                    -ErrorAction Stop

                $newAppId = $null
                if ($null -ne $regResult) {
                    foreach ($item in @($regResult)) {
                        if ($null -eq $item) { continue }
                        $itemStr = $item.ToString()
                        if ($itemStr -match '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})') {
                            $newAppId = $Matches[1]
                            break
                        }
                    }
                }

                if (-not [string]::IsNullOrWhiteSpace($newAppId)) {
                    Write-Log -Level SUCCESS -Message "Entra ID App '$targetAppName' registered. Application (Client) ID: $newAppId"

                    try {
                        $cacheDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath "SharePoint-Audit"
                        if (-not (Test-Path -Path $cacheDir)) {
                            $null = New-Item -Path $cacheDir -ItemType Directory -Force
                        }
                        $cacheFile = Join-Path -Path $cacheDir -ChildPath "${Tenant}_client_id.txt"
                        [System.IO.File]::WriteAllText($cacheFile, $newAppId)
                        Write-Log -Level SUCCESS -Message "Application ID cached locally at: $cacheFile"
                    }
                    catch {
                        Write-Log -Level WARN -Message "Unable to persist Application ID to cache: $($_.Exception.Message)"
                    }

                    return $newAppId
                }
                else {
                    throw "Could not extract Application ID from registration output."
                }
            }
            catch {
                $errMsg = $_.Exception.Message
                Write-Log -Level WARN -Message "Entra ID app registration attempt failed: $errMsg"

                if ([Environment]::UserInteractive) {
                    $suggestedName = "SharePoint-Security-Audit-" + (Get-Date).ToString("yyyyMMdd-HHmm")
                    Write-Host "`n[!] Application registration for '$targetAppName' failed: $errMsg" -ForegroundColor Yellow
                    Write-Host "[?] Press ENTER to retry as [$suggestedName] or paste an existing Client ID (GUID): " -NoNewline -ForegroundColor Yellow
                    $inputVal = Read-Host

                    if ([string]::IsNullOrWhiteSpace($inputVal)) {
                        $targetAppName = $suggestedName
                        continue
                    }
                    elseif ($inputVal.Trim() -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
                        return $inputVal.Trim()
                    }
                    else {
                        $targetAppName = $inputVal.Trim()
                        continue
                    }
                }
                throw "Exceeded attempts for Entra ID app registration: $errMsg"
            }
        }
        throw "Failed to complete Microsoft Entra ID application registration."
    }

    # ==============================================================================
    # CONNECTION MANAGEMENT (Interactive, Certificate, Client Secret)
    # ==============================================================================
    function Connect-SharePointSession {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param(
            [Parameter(Mandatory = $true)]
            [string]$TargetUrl,

            [Parameter(Mandatory = $false)]
            [string]$TenantDomain,

            [Parameter(Mandatory = $false)]
            [string]$AppId,

            [Parameter(Mandatory = $false)]
            [string]$CertThumb,

            [Parameter(Mandatory = $false)]
            [System.Security.SecureString]$Secret,

            [Parameter(Mandatory = $false)]
            [switch]$InteractiveAuth
        )

        Write-Log -Level INFO -Message "Establishing SharePoint Online session for: $TargetUrl"

        # Determine effective tenant domain
        $inferredTenant = try {
            $uri = [System.Uri]$TargetUrl
            $hostParts = $uri.Host.Split('.')
            $prefix = $hostParts[0] -replace '-admin$', ''
            "${prefix}.onmicrosoft.com"
        }
        catch {
            "yourtenant.onmicrosoft.com"
        }

        $effectiveTenant = if (-not [string]::IsNullOrWhiteSpace($TenantDomain)) {
            $TenantDomain
        }
        else {
            $inferredTenant
        }

        try {
            # Method 1: App-Only Certificate Authentication
            if (-not [string]::IsNullOrWhiteSpace($AppId) -and
                -not [string]::IsNullOrWhiteSpace($CertThumb)) {

                $cleanThumb = Assert-SafeThumbprint -Thumbprint $CertThumb
                Write-Log -Level INFO -Message "Authenticating via Entra ID App-Only Certificate (AppId: $AppId, Tenant: $effectiveTenant, Thumb: $cleanThumb)..."

                $conn = Connect-PnPOnline -Url $TargetUrl `
                    -ClientId $AppId `
                    -Tenant $effectiveTenant `
                    -Thumbprint $cleanThumb `
                    -ReturnConnection -ErrorAction Stop

                Write-Log -Level SUCCESS -Message "Connected successfully using App-Only Certificate."
                return $conn
            }

            # Method 2: App-Only Client Secret Authentication
            if (-not [string]::IsNullOrWhiteSpace($AppId) -and $null -ne $Secret) {
                Write-Log -Level INFO -Message "Authenticating via Entra ID App-Only Client Secret (AppId: $AppId, Tenant: $effectiveTenant)..."

                $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secret)
                $plainSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
                [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)

                $conn = Connect-PnPOnline -Url $TargetUrl `
                    -ClientId $AppId `
                    -ClientSecret $plainSecret `
                    -ReturnConnection -ErrorAction Stop

                Write-Log -Level SUCCESS -Message "Connected successfully using App-Only Client Secret."
                return $conn
            }

            # Method 3: Interactive Modern Browser Authentication
            $effectiveClientId = $AppId

            # Check environment variables
            if ([string]::IsNullOrWhiteSpace($effectiveClientId)) {
                if (-not [string]::IsNullOrWhiteSpace($env:ENTRAID_APP_ID)) {
                    $effectiveClientId = $env:ENTRAID_APP_ID
                }
                elseif (-not [string]::IsNullOrWhiteSpace($env:PNP_APP_ID)) {
                    $effectiveClientId = $env:PNP_APP_ID
                }
            }

            # Check local disk cache
            if ([string]::IsNullOrWhiteSpace($effectiveClientId)) {
                $cacheDir  = Join-Path -Path $env:LOCALAPPDATA -ChildPath "SharePoint-Audit"
                $cacheFile = Join-Path -Path $cacheDir -ChildPath "${effectiveTenant}_client_id.txt"
                if (Test-Path -Path $cacheFile) {
                    try {
                        $cachedId = [System.IO.File]::ReadAllText($cacheFile).Trim()
                        if ($cachedId -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
                            Write-Log -Level INFO -Message "Loaded cached Entra ID Application ID: $cachedId"
                            $effectiveClientId = $cachedId
                        }
                    }
                    catch {
                        Write-Log -Level WARN -Message "Unable to read cached ClientId: $($_.Exception.Message)"
                    }
                }
            }

            # If still missing, prompt user or auto-register
            if ([string]::IsNullOrWhiteSpace($effectiveClientId)) {
                Write-Log -Level WARN -Message "No Entra ID Application ID configured for tenant '$effectiveTenant'."

                $shouldRegister = $false
                if ([Environment]::UserInteractive) {
                    Write-Host "`n[?] Microsoft Entra ID requires an Application Client ID for modern authentication." -ForegroundColor Yellow
                    Write-Host "[?] Would you like to automatically register and configure an Entra ID App in '$effectiveTenant' now? (Y/N): " -NoNewline -ForegroundColor Yellow
                    $answer = Read-Host
                    if ($answer -match '^[yY1tT]$') {
                        $shouldRegister = $true
                    }
                }

                if ($shouldRegister) {
                    $effectiveClientId = Invoke-EntraIDAppRegistration -Tenant $effectiveTenant
                }
                else {
                    $guidance = @"
Microsoft Entra ID interactive authentication requires an Application (Client) ID.
Microsoft permanently retired the legacy shared multi-tenant PnP app on September 9, 2024.

To register a dedicated Entra ID App for '$effectiveTenant', run once in PowerShell 7:
    Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "SharePoint-Security-Audit" -Tenant "$effectiveTenant"

Then re-run this script with the generated Application ID:
    .\Audit-ExchangeSecuritySummary.ps1 -TenantId "$effectiveTenant" -ClientId "<ApplicationId>" -Interactive
"@
                    Write-Log -Level ERROR -Message $guidance
                    throw $guidance
                }
            }

            $Script:ActiveClientId = $effectiveClientId
            $Script:ActiveTenant   = $effectiveTenant

            $interactiveParams = @{
                Url              = $TargetUrl
                Interactive      = $true
                ClientId         = $effectiveClientId
                Tenant           = $effectiveTenant
                ReturnConnection = $true
                ErrorAction      = 'Stop'
            }

            $authLabel = if ($InteractiveAuth) { "Interactive browser" } else { "Default interactive" }
            Write-Log -Level INFO -Message "Initiating $authLabel sign-in (ClientId: $effectiveClientId, Tenant: $effectiveTenant)..."

            $conn = Connect-PnPOnline @interactiveParams
            Write-Log -Level SUCCESS -Message "Connected successfully to $TargetUrl via interactive session."
            return $conn
        }
        catch {
            $connErr = "Failed to establish connection to '$TargetUrl'. Details: $($_.Exception.Message)"
            Write-Log -Level ERROR -Message $connErr
            throw $connErr
        }
    }

    function Disconnect-SharePointSession {
        [CmdletBinding()]
        param()

        try {
            Write-Log -Level INFO -Message "Terminating active SharePoint Online session..."
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
            Write-Log -Level SUCCESS -Message "SharePoint Online session disconnected cleanly."
        }
        catch {
            Write-Log -Level WARN -Message "Disconnect warning: $($_.Exception.Message)"
        }
    }

    # ==============================================================================
    # RESILIENCE & THROTTLING RETRY ENGINE (HTTP 429 / 503 Exponential Backoff)
    # ==============================================================================
    function Invoke-PnPActionWithRetry {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [scriptblock]$Action,

            [Parameter(Mandatory = $false)]
            [string]$ActionDescription = "PnP Operation",

            [Parameter(Mandatory = $false)]
            [int]$MaxRetries = 4,

            [Parameter(Mandatory = $false)]
            [int]$InitialDelaySeconds = 3
        )

        $attempt = 0
        $delay   = $InitialDelaySeconds

        while ($attempt -le $MaxRetries) {
            $attempt++
            try {
                return (& $Action)
            }
            catch {
                $ex = $_.Exception
                $isThrottled = $false
                $retryAfterSeconds = 0

                # Inspect for HTTP 429 (Too Many Requests) or 503 (Service Unavailable)
                if ($ex.Message -match '429' -or
                    $ex.Message -match '503' -or
                    $ex.Message -match 'throttled' -or
                    $ex.Message -match 'temporarily unavailable' -or
                    $ex.Message -match 'server busy') {
                    $isThrottled = $true
                }

                # Check WebException headers for Retry-After if available
                if ($ex.InnerException -and $ex.InnerException -is [System.Net.WebException]) {
                    $webEx = [System.Net.WebException]$ex.InnerException
                    if ($null -ne $webEx.Response) {
                        $headers = $webEx.Response.Headers
                        if ($headers -and $headers["Retry-After"]) {
                            $retryVal = [int]0
                            if ([int]::TryParse($headers["Retry-After"], [ref]$retryVal)) {
                                $retryAfterSeconds = $retryVal
                            }
                        }
                    }
                }

                if ($isThrottled -and ($attempt -le $MaxRetries)) {
                    $sleepTime = if ($retryAfterSeconds -gt 0) {
                        $retryAfterSeconds
                    }
                    else {
                        # Exponential backoff + randomized jitter (0.5 to 1.5 multiplier)
                        $jitter = [double](Get-Random -Minimum 50 -Maximum 150) / 100.0
                        [int][Math]::Ceiling($delay * $jitter)
                    }

                    Write-Log -Level WARN -Message "[$ActionDescription] Throttled by SharePoint Online (HTTP 429/503). Retrying in ${sleepTime}s (Attempt $attempt of $MaxRetries)..."
                    Start-Sleep -Seconds $sleepTime
                    $delay = [Math]::Min($delay * 2, 60)
                    continue
                }

                # Non-throttling error or retries exhausted
                Write-Log -Level ERROR -Message "[$ActionDescription] Action failed permanently on attempt ${attempt}: $($ex.Message)"
                throw
            }
        }
    }

    # ==============================================================================
    # SECURITY POSTURE SCORING ALGORITHM
    # ==============================================================================
    function Get-SecurityRiskScore {
        [CmdletBinding()]
        [OutputType([PSCustomObject])]
        param(
            [Parameter(Mandatory = $true)]
            [string]$SharingCapability,

            [Parameter(Mandatory = $true)]
            [bool]$MembersCanShare,

            [Parameter(Mandatory = $true)]
            [int]$AdminCount,

            [Parameter(Mandatory = $true)]
            [string]$ConditionalAccessPolicy,

            [Parameter(Mandatory = $true)]
            [bool]$DenyAddAndCustomizePages,

            [Parameter(Mandatory = $false)]
            [string]$DefaultSharingLinkType = 'None'
        )

        $score = 100
        $riskFactors = [System.Collections.Generic.List[string]]::new()

        # 1. External Sharing Capability Assessment (Max deduction: -35)
        switch ($SharingCapability) {
            'ExternalUserAndGuestSharing' {
                $score -= 30
                $riskFactors.Add("Anonymous 'Anyone' links and guest sharing enabled (ExternalUserAndGuestSharing)")
            }
            'ExternalUserSharingOnly' {
                $score -= 15
                $riskFactors.Add("Authenticated external guests permitted (ExternalUserSharingOnly)")
            }
            'ExistingGuests' {
                $score -= 5
                $riskFactors.Add("Only existing directory guests permitted (ExistingGuests)")
            }
            'Disabled' {
                # Optimal posture: 0 deduction
            }
            Default {
                $score -= 10
                $riskFactors.Add("Unknown or non-standard sharing configuration: '$SharingCapability'")
            }
        }

        # 2. Access Governance: Non-Owner Sharing / MembersCanShare (Max deduction: -20)
        if ($MembersCanShare) {
            $score -= 20
            $riskFactors.Add("Non-owners/members are permitted to share site and invite collaborators (MembersCanShare=True)")
        }

        # 3. Default Sharing Link Type (Max deduction: -15)
        if ($DefaultSharingLinkType -eq 'AnonymousAccess') {
            $score -= 15
            $riskFactors.Add("Default sharing link generates anonymous Anyone links")
        }
        elseif ($DefaultSharingLinkType -eq 'Internal') {
            $score -= 5
            $riskFactors.Add("Default sharing link defaults to company-wide access")
        }

        # 4. Site Collection Administrator Sprawl (Max deduction: -15)
        if ($AdminCount -gt 5) {
            $score -= 15
            $riskFactors.Add("High administrative sprawl ($AdminCount Site Collection Admins; recommended <= 3)")
        }
        elseif ($AdminCount -gt 3) {
            $score -= 10
            $riskFactors.Add("Elevated administrative footprint ($AdminCount Site Collection Admins)")
        }
        elseif ($AdminCount -eq 0) {
            $score -= 5
            $riskFactors.Add("No explicit Site Collection Administrators detected")
        }

        # 5. Device Restrictions & Conditional Access (Max deduction: -10)
        if ($ConditionalAccessPolicy -eq 'AllowFullAccess' -or [string]::IsNullOrWhiteSpace($ConditionalAccessPolicy)) {
            $score -= 10
            $riskFactors.Add("Unmanaged devices hold full download access (No Conditional Access restrictions)")
        }

        # 6. Custom Script Security (Max deduction: -10)
        if (-not $DenyAddAndCustomizePages) {
            $score -= 10
            $riskFactors.Add("Custom scripting enabled (DenyAddAndCustomizePages=False); elevated script injection risk")
        }

        # Normalize score bounds
        $score = [Math]::Max(0, [Math]::Min(100, $score))

        # Classify Risk Tier
        $riskTier = if ($score -ge 85) {
            'Low'
        }
        elseif ($score -ge 65) {
            'Medium'
        }
        elseif ($score -ge 45) {
            'High'
        }
        else {
            'Critical'
        }

        return [PSCustomObject]@{
            Score       = $score
            RiskTier    = $riskTier
            RiskFactors = ($riskFactors -join "; ")
        }
    }

    # ==============================================================================
    # RICH CONSOLE EXECUTIVE DASHBOARD & MATRIX RENDERER
    # ==============================================================================
    function Show-SecuritySummaryMatrix {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param(
            [Parameter(Mandatory = $true)]
            [System.Collections.Generic.List[PSCustomObject]]$SiteRecords
        )

        $totalSites     = $SiteRecords.Count
        $criticalRisk   = ($SiteRecords | Where-Object { $_.RiskTier -eq 'Critical' }).Count
        $highRisk       = ($SiteRecords | Where-Object { $_.RiskTier -eq 'High' }).Count
        $mediumRisk     = ($SiteRecords | Where-Object { $_.RiskTier -eq 'Medium' }).Count
        $lowRisk        = ($SiteRecords | Where-Object { $_.RiskTier -eq 'Low' }).Count

        $extAndGuest    = ($SiteRecords | Where-Object { $_.SharingCapability -eq 'ExternalUserAndGuestSharing' }).Count
        $extOnly        = ($SiteRecords | Where-Object { $_.SharingCapability -eq 'ExternalUserSharingOnly' }).Count
        $existingOnly   = ($SiteRecords | Where-Object { $_.SharingCapability -eq 'ExistingGuests' }).Count
        $disabledShare  = ($SiteRecords | Where-Object { $_.SharingCapability -eq 'Disabled' }).Count

        $membersCanShareCount = ($SiteRecords | Where-Object { $_.MembersCanShare -eq $true }).Count
        $unmanagedAccessCount = ($SiteRecords | Where-Object { $_.ConditionalAccessPolicy -eq 'AllowFullAccess' }).Count
        $avgScore = if ($totalSites -gt 0) {
            [Math]::Round(($SiteRecords | Measure-Object -Property SecurityScore -Average).Average, 1)
        }
        else { 0 }

        Write-Host "`n" -NoNewline
        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host "         SHAREPOINT ONLINE TENANT SECURITY POSTURE ASSESSMENT MATRIX            " -ForegroundColor Yellow
        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host " Total Sites Audited     : $totalSites" -ForegroundColor White
        Write-Host " Average Security Score  : " -NoNewline
        $scoreColor = if ($avgScore -ge 85) { 'Green' } elseif ($avgScore -ge 65) { 'Yellow' } else { 'Red' }
        Write-Host "$avgScore / 100" -ForegroundColor $scoreColor

        Write-Host "--------------------------------------------------------------------------------" -ForegroundColor DarkGray
        Write-Host " POSTURE DISTRIBUTION   :" -ForegroundColor Cyan
        Write-Host "   - Low Risk (Secure)   : " -NoNewline; Write-Host $lowRisk -ForegroundColor Green
        Write-Host "   - Medium Risk         : " -NoNewline; Write-Host $mediumRisk -ForegroundColor Yellow
        Write-Host "   - High Risk           : " -NoNewline; Write-Host $highRisk -ForegroundColor Red
        Write-Host "   - Critical Risk       : " -NoNewline; Write-Host $criticalRisk -ForegroundColor Magenta

        Write-Host " SHARING CAPABILITIES   :" -ForegroundColor Cyan
        Write-Host "   - Anyone / Anonymous  : " -NoNewline; Write-Host $extAndGuest -ForegroundColor $(if ($extAndGuest -gt 0) { 'Red' } else { 'Green' })
        Write-Host "   - Authenticated Guests: " -NoNewline; Write-Host $extOnly -ForegroundColor Yellow
        Write-Host "   - Existing Guests Only: " -NoNewline; Write-Host $existingOnly -ForegroundColor Cyan
        Write-Host "   - Completely Disabled : " -NoNewline; Write-Host $disabledShare -ForegroundColor Green

        Write-Host " ACCESS GOVERNANCE      :" -ForegroundColor Cyan
        Write-Host "   - Members Can Share   : " -NoNewline; Write-Host "$membersCanShareCount sites" -ForegroundColor $(if ($membersCanShareCount -gt 0) { 'Yellow' } else { 'Green' })
        Write-Host "   - Unrestricted Devices: " -NoNewline; Write-Host "$unmanagedAccessCount sites" -ForegroundColor $(if ($unmanagedAccessCount -gt 0) { 'Yellow' } else { 'Green' })
        Write-Host "================================================================================`n" -ForegroundColor Cyan

        # Render Formatted Table of Evaluated Sites
        Write-Host "DETAILED SITE SECURITY SUMMARY:" -ForegroundColor Cyan
        $formattedTable = $SiteRecords | Select-Object `
            @{Name = "Site Title"; Expression = { if ($_.Title.Length -gt 28) { $_.Title.Substring(0, 25) + "..." } else { $_.Title } }},
            @{Name = "Sharing Policy"; Expression = { $_.SharingCapability }},
            @{Name = "Members Share"; Expression = { if ($_.MembersCanShare) { "YES" } else { "No" } }},
            @{Name = "Admins"; Expression = { $_.AdminCount }},
            @{Name = "Cond Access"; Expression = { $_.ConditionalAccessPolicy }},
            @{Name = "Score"; Expression = { $_.SecurityScore }},
            @{Name = "Risk Tier"; Expression = { $_.RiskTier }}

        $formattedTable | Format-Table -AutoSize | Out-String | Write-Host -ForegroundColor White

        # Display Top Risk Remediation Guidance if high/critical sites exist
        $urgentSites = $SiteRecords | Where-Object { $_.RiskTier -in @('Critical', 'High') }
        if ($urgentSites.Count -gt 0) {
            Write-Host "[!] HIGH & CRITICAL RISK ACTION ITEMS IDENTIFIED ($($urgentSites.Count) Sites):" -ForegroundColor Red
            foreach ($site in ($urgentSites | Select-Object -First 5)) {
                Write-Host "    * [$($site.RiskTier)] $($site.Url)" -ForegroundColor Yellow
                Write-Host "      Issues: $($site.RiskFactors)" -ForegroundColor DarkGray
            }
            if ($urgentSites.Count -gt 5) {
                Write-Host "    * ... and $($urgentSites.Count - 5) more high-risk site(s)." -ForegroundColor DarkGray
            }
            Write-Host "`n"
        }
    }

    # ==============================================================================
    # EXPORT ENGINE (CSV / JSON)
    # ==============================================================================
    function Export-AuditArtifact {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [System.Collections.Generic.List[PSCustomObject]]$Records,

            [Parameter(Mandatory = $true)]
            [string]$DestinationPath,

            [Parameter(Mandatory = $true)]
            [ValidateSet('CSV', 'JSON')]
            [string]$Format
        )

        Write-Log -Level INFO -Message "Preparing export artifact at: $DestinationPath (Format: $Format)"

        try {
            $resolvedPath = [System.IO.Path]::GetFullPath($DestinationPath)
            $parentDir    = [System.IO.Path]::GetDirectoryName($resolvedPath)

            if (-not [string]::IsNullOrWhiteSpace($parentDir) -and -not (Test-Path -Path $parentDir)) {
                $null = New-Item -Path $parentDir -ItemType Directory -Force -ErrorAction Stop
                Write-Log -Level INFO -Message "Created destination directory: $parentDir"
            }

            if ($Format -eq 'CSV') {
                $Records | Export-Csv -Path $resolvedPath -NoTypeInformation -Encoding UTF8 -Force -ErrorAction Stop
            }
            else {
                $jsonContent = $Records | ConvertTo-Json -Depth 5
                [System.IO.File]::WriteAllText($resolvedPath, $jsonContent, [System.Text.Encoding]::UTF8)
            }

            $itemInfo = Get-Item -Path $resolvedPath
            $fileSizeKb = [Math]::Round($itemInfo.Length / 1KB, 2)
            Write-Log -Level SUCCESS -Message "Export completed successfully: $resolvedPath ($fileSizeKb KB, $($Records.Count) records)"
        }
        catch {
            $exportErr = "Failed to export audit artifact to '$DestinationPath'. Details: $($_.Exception.Message)"
            Write-Log -Level ERROR -Message $exportErr
            throw $exportErr
        }
    }

    # Run initial checks
    Assert-PowerShellVersion -BoundParameters $PSBoundParameters
    Initialize-AuditLogging
    Assert-Prerequisite -ForceInstall:$Force
}

process {
    $auditRecords = [System.Collections.Generic.List[PSCustomObject]]::new()
    $targetAdminUrl = $null

    try {
        # 1. Resolve Target Admin URL
        if (-not [string]::IsNullOrWhiteSpace($AdminUrl)) {
            Assert-SafeUrl -Url $AdminUrl -IsAdminUrl
            $targetAdminUrl = $AdminUrl.TrimEnd('/')
        }
        elseif (-not [string]::IsNullOrWhiteSpace($TenantId)) {
            $cleanTenant = $TenantId.Trim().ToLower()
            $prefix = if ($cleanTenant -match '^([a-z0-9-]+)\.onmicrosoft\.com$') {
                $Matches[1]
            }
            else {
                $cleanTenant
            }
            $targetAdminUrl = "https://${prefix}-admin.sharepoint.com"
        }
        elseif (-not [string]::IsNullOrWhiteSpace($SiteUrl)) {
            Assert-SafeUrl -Url $SiteUrl
            $uri = [System.Uri]$SiteUrl
            $prefix = $uri.Host.Split('.')[0]
            $targetAdminUrl = "https://${prefix}-admin.sharepoint.com"
        }
        else {
            $targetAdminUrl = "https://yourtenant-admin.sharepoint.com"
        }

        # 2. Connect to SharePoint Online Admin Session
        $Script:ActiveConnection = Connect-SharePointSession `
            -TargetUrl $targetAdminUrl `
            -TenantDomain $TenantId `
            -AppId $ClientId `
            -CertThumb $CertificateThumbprint `
            -Secret $ClientSecret `
            -InteractiveAuth:$Interactive

        # 3. Discovery: Enumerate Site Collections
        Write-Log -Level INFO -Message "Enumerating SharePoint Online site collections..."

        $allSites = $null
        if (-not [string]::IsNullOrWhiteSpace($SiteUrl)) {
            # Single Site Audit Mode
            Assert-SafeUrl -Url $SiteUrl
            Write-Log -Level INFO -Message "Targeting specific site collection: $SiteUrl"

            $singleSite = Invoke-PnPActionWithRetry -ActionDescription "Query Site $SiteUrl" -Action {
                Get-PnPTenantSite -Identity $SiteUrl -Detailed -Connection $Script:ActiveConnection -ErrorAction Stop
            }
            if ($null -ne $singleSite) {
                $allSites = @($singleSite)
            }
            else {
                $allSites = @()
            }
        }
        else {
            # Tenant-wide Audit Mode
            $allSites = Invoke-PnPActionWithRetry -ActionDescription "Enumerate Tenant Sites" -Action {
                Get-PnPTenantSite -Detailed `
                    -IncludeOneDriveSites:$IncludePersonalSites `
                    -Connection $Script:ActiveConnection `
                    -ErrorAction Stop
            }
        }

        if ($null -eq $allSites -or $allSites.Count -eq 0) {
            Write-Log -Level WARN -Message "No site collections were discovered in tenant."
            return
        }

        Write-Log -Level SUCCESS -Message "Discovered $($allSites.Count) raw site collection(s)."

        # 4. Filter System Sites and Keyword Matches
        $filteredSites = [System.Collections.Generic.List[object]]::new()
        $systemTemplates = @(
            'APPCATALOG#0',
            'POINTPUBLISHINGHUB#0',
            'POINTPUBLISHINGTOPIC#0',
            'EDISC#0',
            'SEARCHCENTER#0',
            'BLANKINTERNETCONTAINER#0'
        )

        foreach ($s in $allSites) {
            if ($null -eq $s) { continue }

            # Exclude system templates unless -IncludeSystemSites is specified
            if (-not $IncludeSystemSites -and ($s.Template -in $systemTemplates)) {
                continue
            }

            # Filter by search keyword if provided
            if (-not [string]::IsNullOrWhiteSpace($Search)) {
                $kw = $Search.Trim()
                $matchUrl   = ($s.Url -and $s.Url -match [regex]::Escape($kw))
                $matchTitle = ($s.Title -and $s.Title -match [regex]::Escape($kw))
                if (-not ($matchUrl -or $matchTitle)) {
                    continue
                }
            }

            $filteredSites.Add($s)
        }

        Write-Log -Level INFO -Message "Evaluating security posture for $($filteredSites.Count) site collection(s)..."

        $index = 0
        $total = $filteredSites.Count

        foreach ($site in $filteredSites) {
            $index++
            $siteUrl   = $site.Url
            $siteTitle = if ([string]::IsNullOrWhiteSpace($site.Title)) { "(Untitled Site)" } else { $site.Title }

            Write-Progress -Activity "Auditing SharePoint Security Posture" `
                -Status "Evaluating [$index/$total]: $siteTitle" `
                -PercentComplete ([int](($index / $total) * 100))

            # Audit Properties Extraction
            $sharingCapability  = if ($site.SharingCapability) { $site.SharingCapability.ToString() } else { "Disabled" }
            $defaultLinkType    = if ($site.DefaultSharingLinkType) { $site.DefaultSharingLinkType.ToString() } else { "None" }
            $defaultPermission  = if ($site.DefaultLinkPermission) { $site.DefaultLinkPermission.ToString() } else { "None" }
            $condAccessPolicy   = if ($site.ConditionalAccessPolicy) { $site.ConditionalAccessPolicy.ToString() } else { "AllowFullAccess" }
            $allowDownloadNonWeb= if ($site.AllowDownloadingNonWebViewableFiles) { $true } else { $false }
            $denyCustomPages    = if ($site.DenyAddAndCustomizePages) {
                ($site.DenyAddAndCustomizePages.ToString() -eq 'Enabled')
            } else { $true }

            # MembersCanShare: Evaluate DisableSharingForNonOwnersStatus
            # If DisableSharingForNonOwnersStatus is Disabled/True -> Non-owners cannot share -> MembersCanShare is False
            $membersCanShare = $true
            if ($site.DisableSharingForNonOwnersStatus) {
                $statusStr = $site.DisableSharingForNonOwnersStatus.ToString()
                if ($statusStr -eq 'Disabled' -or $statusStr -eq 'True') {
                    $membersCanShare = $false
                }
            }

            # Site Collection Administrators Enumeration
            $siteAdmins = [System.Collections.Generic.List[string]]::new()
            if (-not [string]::IsNullOrWhiteSpace($site.Owner)) {
                $siteAdmins.Add($site.Owner)
            }
            if (-not [string]::IsNullOrWhiteSpace($site.SecondaryOwner)) {
                if (-not $siteAdmins.Contains($site.SecondaryOwner)) {
                    $siteAdmins.Add($site.SecondaryOwner)
                }
            }

            $adminCount = [Math]::Max(1, $siteAdmins.Count)

            # Evaluate Security Risk Score
            $scoreObj = Get-SecurityRiskScore `
                -SharingCapability $sharingCapability `
                -MembersCanShare $membersCanShare `
                -AdminCount $adminCount `
                -ConditionalAccessPolicy $condAccessPolicy `
                -DenyAddAndCustomizePages $denyCustomPages `
                -DefaultSharingLinkType $defaultLinkType

            # Build Strongly-Typed Audit Record
            $record = [PSCustomObject]@{
                SiteUrl                       = $siteUrl
                Title                         = $siteTitle
                Template                      = $site.Template
                Owner                         = $site.Owner
                SiteAdmins                    = ($siteAdmins -join "; ")
                AdminCount                    = $adminCount
                SharingCapability             = $sharingCapability
                MembersCanShare               = $membersCanShare
                DefaultSharingLinkType        = $defaultLinkType
                DefaultLinkPermission         = $defaultPermission
                ConditionalAccessPolicy       = $condAccessPolicy
                AllowDownloadingNonWebViewable= $allowDownloadNonWeb
                DenyAddAndCustomizePages      = $denyCustomPages
                StorageUsageMB                = if ($site.StorageUsageCurrent) { $site.StorageUsageCurrent } else { 0 }
                LockState                     = if ($site.LockState) { $site.LockState.ToString() } else { "Unlock" }
                LastContentModifiedDate       = if ($site.LastContentModifiedDate) { $site.LastContentModifiedDate } else { $null }
                SecurityScore                 = $scoreObj.Score
                RiskTier                      = $scoreObj.RiskTier
                RiskFactors                   = $scoreObj.RiskFactors
                AuditedAt                     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            }

            $auditRecords.Add($record)
        }

        Write-Progress -Activity "Auditing SharePoint Security Posture" -Completed

        # 5. Render Console Matrix
        Show-SecuritySummaryMatrix -SiteRecords $auditRecords

        # 6. Export Artifact if Requested
        if (-not [string]::IsNullOrWhiteSpace($ExportPath)) {
            Export-AuditArtifact -Records $auditRecords -DestinationPath $ExportPath -Format $ExportFormat
        }

        # 7. Pipeline Emission
        if ($PassThru) {
            foreach ($rec in $auditRecords) {
                Write-Output $rec
            }
        }
    }
    catch {
        Write-Log -Level ERROR -Message "Execution terminated unexpectedly: $($_.Exception.Message)"
        throw
    }
}

end {
    $duration = (Get-Date) - $Script:StartTime
    $durStr   = "{0:D2}m {1:D2}s" -f $duration.Minutes, $duration.Seconds
    Write-Log -Level INFO -Message "Execution completed in $durStr."

    # Resource Cleanup & Session Termination
    Disconnect-SharePointSession
}

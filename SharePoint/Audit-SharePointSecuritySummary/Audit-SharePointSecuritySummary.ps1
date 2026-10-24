<#
.SYNOPSIS
    Audit-SharePointSecuritySummary - Audits SharePoint Online security posture and generates executive reports.

.DESCRIPTION
    Audit-SharePointSecuritySummary is an enterprise-grade PowerShell automation tool engineered for
    Microsoft 365 administrators, Information Security Officers, and compliance auditors.

    The script performs comprehensive security posture inspection across SharePoint Online site collections:
    - External Sharing Governance: Audits site sharing capabilities (Disabled, Existing Guests, Anyone/Anonymous links).
    - Anonymous & Guest Link Analysis: Identifies unrestricted anyone links and lack of expiration policies.
    - Site Ownership Footprint: Flags orphaned site collections (0 active owners/admins) and excessive administrators (>5).
    - Permission Inheritance & Unique Access: Detects broken inheritance on site webs, lists, and document libraries.
    - Sensitivity Labels & Group Governance: Captures Microsoft 365 Group privacy (Public vs Private) and sensitivity labels.
    - Multi-Tiered Risk Engine: Computes mathematical risk scores and categorizes sites into Critical, High, Medium, and Low risk.
    - Executive Reporting: Generates an interactive colorized console dashboard, raw CSV remediation dataset,
      and a standalone, responsive HTML executive report with interactive filtering and remediation guidance.
    - Pipeline Support: Emits strongly-typed [PSCustomObject] records to the PowerShell pipeline for downstream automation.
    - Post-2024 Entra ID Architecture: Automated Entra ID App registration, local credential/client ID caching,
      and seamless modern authentication compliant with Microsoft's retired multi-tenant PnP app policy.
    - Dual Authentication: Supports interactive browser OAuth with MFA and non-interactive Entra ID App-Only certificate authentication.
    - Enterprise Logging: Maintains timestamped execution logs in C:\Logs\ with automatic fallback.

.PARAMETER AdminUrl
    The SharePoint Online Admin URL (e.g. 'https://contoso-admin.sharepoint.com').
    Can also accept the root tenant URL (e.g. 'https://contoso.sharepoint.com'), which will be automatically normalized.

.PARAMETER ClientId
    Microsoft Entra ID Application (Client) ID for unattended App-Only daemon authentication or interactive modern login.
    If omitted in interactive mode, the script checks local cache or guides automated on-the-fly registration.

.PARAMETER TenantId
    Microsoft Entra ID Directory (Tenant) ID or tenant domain (e.g. 'contoso.onmicrosoft.com').
    If omitted, the tenant domain is inferred from the SharePoint Admin URL.

.PARAMETER CertificateThumbprint
    The thumbprint of an X.509 certificate installed in Cert:\CurrentUser\My or Cert:\LocalMachine\My
    used for unattended App-Only authentication.

.PARAMETER CertificatePath
    File system path to an X.509 certificate file (.pfx or .cer) for App-Only authentication.

.PARAMETER CertificatePassword
    Optional SecureString password for the certificate file if password-protected.

.PARAMETER Interactive
    Switch parameter. Initiates interactive modern web browser authentication prompt with MFA support.

.PARAMETER OutputPath
    Directory path where CSV and HTML reports will be saved. Defaults to 'C:\Logs'.

.PARAMETER GenerateHtmlReport
    Switch parameter. Generates a standalone, responsive HTML executive summary dashboard.

.PARAMETER LogPath
    Directory path for transcript and execution logs. Defaults to 'C:\Logs'.

.PARAMETER SiteFilter
    Optional wildcard or exact string filter to audit specific site collections by URL or Title (e.g. '*finance*', 'https://contoso.sharepoint.com/sites/hr').

.PARAMETER IncludePersonalSites
    Switch parameter. When specified, personal OneDrive for Business site collections ('/personal/') are included.
    By default, personal sites are excluded to focus on organizational SharePoint sites.

.PARAMETER DetailedScan
    Switch parameter. Performs deeper sub-inspection of individual site document libraries and lists for broken inheritance.

.PARAMETER PassThru
    Switch parameter. Emits the custom audit PSCustomObject collection to the PowerShell pipeline for downstream processing.

.PARAMETER Force
    Switch parameter. Bypasses interactive prompts and confirmations for automated execution.

.PARAMETER Help
    Displays the built-in help screen, parameter guidance, and usage examples, then terminates execution.
    Aliases: -h.

.EXAMPLE
    # 1. Interactive audit of all tenant sites with HTML executive dashboard
    .\Audit-SharePointSecuritySummary.ps1 -AdminUrl "https://contoso-admin.sharepoint.com" -Interactive -GenerateHtmlReport

.EXAMPLE
    # 2. Automated daemon audit using Entra ID App Registration and certificate authentication
    .\Audit-SharePointSecuritySummary.ps1 -AdminUrl "https://contoso-admin.sharepoint.com" `
        -ClientId "00000000-0000-0000-0000-000000000000" `
        -TenantId "11111111-1111-1111-1111-111111111111" `
        -CertificateThumbprint "A1B2C3D4E5F67890123456789ABCDEF012345678" `
        -OutputPath "C:\Reports" -GenerateHtmlReport

.EXAMPLE
    # 3. Targeted audit filtering specific site collections, piping results to Out-GridView
    .\Audit-SharePointSecuritySummary.ps1 -AdminUrl "https://contoso-admin.sharepoint.com" -Interactive `
        -SiteFilter "*teams*" -PassThru | Out-GridView

.EXAMPLE
    # 4. Pipeline filtering for high-risk sites only
    $riskySites = .\Audit-SharePointSecuritySummary.ps1 -AdminUrl "https://contoso-admin.sharepoint.com" -Interactive -PassThru |
        Where-Object { $_.RiskLevel -in @('Critical', 'High') }

.NOTES
    Author: Roman Pindela
    Email: roman.pindela@gmail.com
    GitHub: romanpindela
    Version: 1.0.1
    Last Modified: 2026-10-24
    Standards: CleanCode & PSScriptAnalyzer Zero-Warning Standard
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 0)]
    [Alias('AdminSiteUrl', 'Url')]
    [string]$AdminUrl,

    [Parameter(Mandatory = $false)]
    [Alias('AppId')]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [Alias('Tenant')]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [Alias('Thumbprint')]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory = $false)]
    [string]$CertificatePath,

    [Parameter(Mandatory = $false)]
    [System.Security.SecureString]$CertificatePassword,

    [Parameter(Mandatory = $false)]
    [switch]$Interactive,

    [Parameter(Mandatory = $false)]
    [Alias('ReportDir', 'OutDir')]
    [string]$OutputPath = 'C:\Logs',

    [Parameter(Mandatory = $false)]
    [Alias('Html')]
    [switch]$GenerateHtmlReport,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = 'C:\Logs',

    [Parameter(Mandatory = $false)]
    [string]$SiteFilter,

    [Parameter(Mandatory = $false)]
    [switch]$IncludePersonalSites,

    [Parameter(Mandatory = $false)]
    [switch]$DetailedScan,

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
    # SCRIPT METADATA & TELEMETRY
    # ==============================================================================
    $Script:Metadata = @{
        Title        = "Audit-SharePointSecuritySummary.ps1"
        Version      = "1.0.1"
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
    $Script:AutoSiteFilter   = $null
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
    # ZERO-PARAMETER & -HELP SAFETY GUARD (Specification Section 3.3)
    # ==============================================================================
    function Show-ScriptHelp {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param()

        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host " SCRIPT  : $($Script:Metadata.Title) (v$($Script:Metadata.Version))" -ForegroundColor Cyan
        Write-Host " AUTHOR  : $($Script:Metadata.Author) <$($Script:Metadata.Email)>" -ForegroundColor Cyan
        Write-Host " GITHUB  : $($Script:Metadata.GitHub)" -ForegroundColor Cyan
        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host "`nDESCRIPTION:" -ForegroundColor White
        Write-Host "    Audits SharePoint Online security posture across tenant site collections." -ForegroundColor Gray
        Write-Host "    Evaluates external sharing governance, orphaned site collections, excessive admin" -ForegroundColor Gray
        Write-Host "    assignments, anonymous sharing links, and broken permission inheritance." -ForegroundColor Gray
        Write-Host "    Generates an executive console dashboard, CSV dataset, and standalone HTML report." -ForegroundColor Gray
        Write-Host "`nPOST-2024 ENTRA ID COMPLIANCE:" -ForegroundColor Yellow
        Write-Host "    Microsoft permanently retired the multi-tenant PnP App on September 9, 2024." -ForegroundColor Gray
        Write-Host "    The script automatically resolves or registers a dedicated Entra ID App in your" -ForegroundColor Gray
        Write-Host "    tenant and caches the Client ID locally under AppData\Local\SharePoint-Audit." -ForegroundColor Gray
        Write-Host "`nPARAMETERS:" -ForegroundColor Yellow
        Write-Host "    -AdminUrl               [String]       SharePoint Online Admin URL (e.g. 'https://contoso-admin.sharepoint.com')." -ForegroundColor Gray
        Write-Host "    -Interactive            [Switch]       Initiate interactive browser OAuth authentication with MFA." -ForegroundColor Gray
        Write-Host "    -ClientId               [String]       Microsoft Entra ID Application (Client) ID for App-Only or interactive auth." -ForegroundColor Gray
        Write-Host "    -TenantId               [String]       Microsoft Entra ID Tenant ID or domain (contoso.onmicrosoft.com)." -ForegroundColor Gray
        Write-Host "    -CertificateThumbprint  [String]       X.509 Certificate thumbprint for unattended daemon auth." -ForegroundColor Gray
        Write-Host "    -CertificatePath        [String]       Path to certificate (.pfx/.cer) file." -ForegroundColor Gray
        Write-Host "    -OutputPath             [String]       Target directory for CSV and HTML reports (Default: 'C:\Logs')." -ForegroundColor Gray
        Write-Host "    -GenerateHtmlReport     [Switch]       Generate standalone HTML executive summary report." -ForegroundColor Gray
        Write-Host "    -LogPath                [String]       Target directory for transcript logs (Default: 'C:\Logs')." -ForegroundColor Gray
        Write-Host "    -SiteFilter             [String]       Filter sites by URL or Title wildcard (e.g. '*finance*')." -ForegroundColor Gray
        Write-Host "    -IncludePersonalSites   [Switch]       Include OneDrive for Business (/personal/) sites." -ForegroundColor Gray
        Write-Host "    -DetailedScan           [Switch]       Perform deeper inspection of document library inheritance." -ForegroundColor Gray
        Write-Host "    -PassThru               [Switch]       Emit [PSCustomObject] collection to pipeline." -ForegroundColor Gray
        Write-Host "    -Help (-h)              [Switch]       Display this help documentation and exit." -ForegroundColor Gray
        Write-Host "`nEXAMPLES:" -ForegroundColor Yellow
        Write-Host "    .\Audit-SharePointSecuritySummary.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com' -Interactive -GenerateHtmlReport" -ForegroundColor Green
        Write-Host "    .\Audit-SharePointSecuritySummary.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com' -Interactive -SiteFilter '*hr*' -PassThru | Out-GridView" -ForegroundColor Green
        Write-Host "    .\Audit-SharePointSecuritySummary.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com' -ClientId '<AppId>' -TenantId '<TenantId>' -CertificateThumbprint '<Thumbprint>' -OutputPath 'C:\Reports' -GenerateHtmlReport" -ForegroundColor Green
        Write-Host "`n================================================================================`n" -ForegroundColor Cyan
    }

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
            [System.Collections.IDictionary]$BoundParams
        )

        if ($PSVersionTable.PSVersion.Major -lt 7) {
            $pwsh = Get-Command pwsh.exe -ErrorAction SilentlyContinue
            if ($pwsh -and (-not [string]::IsNullOrWhiteSpace($Script:ScriptFilePath))) {
                Write-Host "[*] PnP.PowerShell requires PowerShell 7+ (Core). Current host: Windows PowerShell $($PSVersionTable.PSVersion)" -ForegroundColor Cyan
                Write-Host "[*] Automatically relaunching script in PowerShell 7 (pwsh.exe)...`n" -ForegroundColor Yellow

                $forwardArgs = [System.Collections.Generic.List[string]]::new()
                $escapedScriptPath = $Script:ScriptFilePath.Replace('"', '`"')
                $forwardArgs.Add("& `"$escapedScriptPath`"")

                if ($BoundParams) {
                    foreach ($key in $BoundParams.Keys) {
                        $val = $BoundParams[$key]
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
    # LOGGING INFRASTRUCTURE (Specification Section 3.4)
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
        $logFileName  = "${timestamp}-Audit-SharePointSecuritySummary-${computerName}-${userName}.log"

        $resolvedDir = $TargetDirectory
        try {
            if (-not (Test-Path -Path $resolvedDir)) {
                $null = New-Item -Path $resolvedDir -ItemType Directory -Force -ErrorAction Stop
            }
        }
        catch {
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
Log File Location     : $Script:LogFilePath
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

            [Parameter(Mandatory = $false)]
            [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
            [string]$Level = 'INFO'
        )

        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry  = "[$timestamp] [$Level] $Message"

        if ($Script:LogFilePath) {
            try {
                Add-Content -Path $Script:LogFilePath -Value $logEntry -Encoding UTF8 -ErrorAction SilentlyContinue
            }
            catch {
                Write-Verbose -Message "Unable to write log entry to file: $($_.Exception.Message)"
            }
        }

        switch ($Level) {
            'INFO'    { Write-Host $logEntry -ForegroundColor Cyan }
            'WARN'    { Write-Host $logEntry -ForegroundColor Yellow }
            'ERROR'   { Write-Host $logEntry -ForegroundColor Red }
            'SUCCESS' { Write-Host $logEntry -ForegroundColor Green }
        }
    }

    # ==============================================================================
    # HELPER: SAFE PROPERTY VALUE RETRIEVAL (StrictMode Resilient)
    # ==============================================================================
    function Get-SafePropertyValue {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            $InputObject,

            [Parameter(Mandatory = $true)]
            [string]$PropertyName,

            [Parameter(Mandatory = $false)]
            $DefaultValue = $null
        )

        if ($null -eq $InputObject) {
            return $DefaultValue
        }

        if ($InputObject -is [System.Collections.IDictionary]) {
            if ($InputObject.Contains($PropertyName)) {
                return $InputObject[$PropertyName]
            }
            return $DefaultValue
        }

        if ($InputObject.PSObject -and $InputObject.PSObject.Properties) {
            $prop = $InputObject.PSObject.Properties[$PropertyName]
            if ($null -ne $prop) {
                return $prop.Value
            }
        }

        return $DefaultValue
    }

    # ==============================================================================
    # HELPER: INTERACTIVE INPUT PROMPT
    # ==============================================================================
    function Read-AdminUrlInteractive {
        [CmdletBinding()]
        [OutputType([string])]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param()

        Write-Host "`n[?] Enter SharePoint Online Admin URL (e.g. 'https://contoso-admin.sharepoint.com'): " -NoNewline -ForegroundColor Yellow
        $entered = Read-Host
        return $entered
    }

    # ==============================================================================
    # HELPER: RESOLVE ADMIN URL
    # ==============================================================================
    function Resolve-SharePointAdminUrl {
        [CmdletBinding()]
        [OutputType([string])]
        param(
            [Parameter(Mandatory = $false)]
            [string]$Url
        )

        if ([string]::IsNullOrWhiteSpace($Url)) {
            return ""
        }

        $trimmed = $Url.Trim().TrimEnd('/')
        if ($trimmed -notmatch '^https?://') {
            $trimmed = "https://$trimmed"
        }

        try {
            $uri = [System.Uri]$trimmed
            $hostName = $uri.Host

            # Detect if user passed a specific site collection URL (e.g. /sites/sales or /teams/it)
            if ($uri.AbsolutePath -match '^/(sites|teams)/([^/?#]+)') {
                $subSite = $Matches[2]
                if ([string]::IsNullOrWhiteSpace($Script:AutoSiteFilter)) {
                    $Script:AutoSiteFilter = "*$subSite*"
                    Write-Log -Level INFO -Message "Specific site collection detected in target URL ('$subSite'). Setting auto-filter: $Script:AutoSiteFilter"
                }
            }

            # Normalize root tenant url to admin url
            # e.g., 'https://contoso.sharepoint.com' -> 'https://contoso-admin.sharepoint.com'
            if ($hostName -notmatch '-admin\.sharepoint\.com$' -and $hostName -match '^([a-zA-Z0-9\-]+)\.sharepoint\.com$') {
                $tenantName = $Matches[1].Replace('-my', '')
                $adminHost = "${tenantName}-admin.sharepoint.com"
                $trimmed = "https://$adminHost"
                Write-Log -Level INFO -Message "Normalized tenant URL to SharePoint Admin URL: $trimmed"
            }
        }
        catch {
            Write-Log -Level WARN -Message "Unable to parse URL '$Url'. Using as provided: $trimmed"
        }

        return $trimmed
    }

    # ==============================================================================
    # HELPER: MODULE DEPENDENCY VALIDATION
    # ==============================================================================
    function Assert-PnPModule {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param(
            [Parameter(Mandatory = $false)]
            [switch]$ForceInstall
        )

        $modName = $Script:Metadata.RequiredMod
        $moduleInstalled = Get-Module -ListAvailable -Name $modName
        if (-not $moduleInstalled) {
            Write-Log -Level WARN -Message "Required PowerShell module '$modName' is not installed."

            $shouldInstall = $false
            if ($ForceInstall) {
                $shouldInstall = $true
            }
            elseif ([Environment]::UserInteractive) {
                Write-Host "Would you like to install '$modName' from PSGallery now? (Y/N): " -NoNewline -ForegroundColor Yellow
                $response = Read-Host
                if ($response -match '^[ytT1]$') {
                    $shouldInstall = $true
                }
            }

            if ($shouldInstall) {
                Write-Log -Level INFO -Message "Installing '$modName' from PSGallery for CurrentUser..."
                try {
                    Install-Module -Name $modName -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
                    Write-Log -Level SUCCESS -Message "Module '$modName' installed successfully."
                }
                catch {
                    $installErr = "Automatic installation of '$modName' failed: $($_.Exception.Message)"
                    Write-Log -Level ERROR -Message $installErr
                    throw $installErr
                }
            }
            else {
                $errorMessage = @"
Required PowerShell module '$modName' is not installed.
Please install it by running:
    Install-Module -Name modName -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
"@
                Write-Log -Level ERROR -Message $errorMessage
                throw $errorMessage
            }
        }

        $loaded = Get-Module -Name $modName
        if (-not $loaded) {
            Write-Log -Level INFO -Message "Importing module '$modName'..."
            Import-Module -Name $modName -ErrorAction Stop
        }

        $currentMod = Get-Module -Name $modName
        Write-Log -Level SUCCESS -Message "Module '$modName' (v$($currentMod.Version)) verified."
    }

    # ==============================================================================
    # AUTOMATED ENTRA ID APP REGISTRATION (Post-2024 Modern PnP Compliance)
    # ==============================================================================
    function Invoke-EntraIDAppRegistration {
        [CmdletBinding()]
        [OutputType([string])]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Tenant,

            [Parameter(Mandatory = $false)]
            [string]$ApplicationName = "SharePoint-Audit"
        )

        $targetAppName = $ApplicationName
        $maxAttempts   = 3
        $attempt       = 0

        while ($attempt -lt $maxAttempts) {
            $attempt++

            Write-Log -Level INFO -Message "Initiating automated Microsoft Entra ID App registration ('$targetAppName') in tenant '$Tenant'..."
            Write-Host "`n[*] Microsoft Entra ID registration requires a one-time administrator consent flow." -ForegroundColor Yellow
            Write-Host "[*] A browser window will now open for modern interactive sign-in." -ForegroundColor Yellow
            Write-Host "[*] Please sign in with your Microsoft 365 account and grant consent to the requested SharePoint permissions.`n" -ForegroundColor Yellow

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
                        if ($item.PSObject) {
                            foreach ($prop in $item.PSObject.Properties) {
                                if ($prop.Value -and $prop.Value.ToString() -match '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})') {
                                    $newAppId = $Matches[1]
                                    break
                                }
                            }
                            if ($newAppId) { break }
                        }
                    }

                    if ([string]::IsNullOrWhiteSpace($newAppId)) {
                        $allText = ($regResult | Out-String)
                        if ($allText -match '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})') {
                            $newAppId = $Matches[1]
                        }
                    }
                }

                if (-not [string]::IsNullOrWhiteSpace($newAppId)) {
                    Write-Log -Level SUCCESS -Message "Entra ID App '$targetAppName' registered successfully. Application (Client) ID: $newAppId"

                    # Persist to local cache so subsequent runs never need prompt or re-registration
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
                        Write-Log -Level WARN -Message "Unable to persist Application ID to local cache: $($_.Exception.Message)"
                    }

                    return $newAppId
                }
                else {
                    throw "Could not extract Application (Client) ID from registration output."
                }
            }
            catch {
                $errMsg = $_.Exception.Message
                Write-Log -Level WARN -Message "Error during Microsoft Entra ID app registration: $errMsg"

                if ([Environment]::UserInteractive) {
                    $suggestedName = "SharePoint-Audit-" + (Get-Date).ToString("yyyyMMdd-HHmm")
                    Write-Host "`n[!] Application registration for '$targetAppName' failed: $errMsg" -ForegroundColor Yellow
                    Write-Host "[*] Suggested unique application name: $suggestedName" -ForegroundColor Cyan
                    Write-Host "[?] Press ENTER to try [$suggestedName], enter a different name, or paste an existing Application ID (GUID): " -NoNewline -ForegroundColor Yellow
                    $inputVal = Read-Host

                    if ([string]::IsNullOrWhiteSpace($inputVal)) {
                        $targetAppName = $suggestedName
                        Write-Log -Level INFO -Message "Retrying registration with suggested name: $targetAppName"
                        continue
                    }
                    elseif ($inputVal.Trim() -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
                        $pastedId = $inputVal.Trim()
                        Write-Log -Level SUCCESS -Message "Using provided existing Application ID: $pastedId"
                        try {
                            $cacheDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath "SharePoint-Audit"
                            if (-not (Test-Path -Path $cacheDir)) {
                                $null = New-Item -Path $cacheDir -ItemType Directory -Force
                            }
                            $cacheFile = Join-Path -Path $cacheDir -ChildPath "${Tenant}_client_id.txt"
                            [System.IO.File]::WriteAllText($cacheFile, $pastedId)
                            Write-Log -Level SUCCESS -Message "Application ID cached locally at: $cacheFile"
                        }
                        catch {
                            Write-Log -Level WARN -Message "Unable to persist ClientId to cache: $($_.Exception.Message)"
                        }
                        return $pastedId
                    }
                    else {
                        $targetAppName = $inputVal.Trim()
                        Write-Log -Level INFO -Message "Retrying registration with new name: $targetAppName"
                        continue
                    }
                }
                else {
                    if ($errMsg -match 'already exists') {
                        $targetAppName = "SharePoint-Audit-" + (Get-Date).ToString("yyyyMMdd-HHmm")
                        continue
                    }
                    $regErr = "Failed to register Microsoft Entra ID application: $errMsg"
                    Write-Log -Level ERROR -Message $regErr
                    throw $regErr
                }
            }
        }

        throw "Exceeded maximum attempts for Microsoft Entra ID application registration."
    }

    # ==============================================================================
    # AUTHENTICATION ENGINE (Interactive & App-Only Daemon)
    # ==============================================================================
    function Connect-AuditSession {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '')]
        param(
            [Parameter(Mandatory = $true)]
            [string]$TargetAdminUrl,

            [Parameter(Mandatory = $false)]
            [string]$AppId,

            [Parameter(Mandatory = $false)]
            [string]$Tenant,

            [Parameter(Mandatory = $false)]
            [string]$Thumbprint,

            [Parameter(Mandatory = $false)]
            [string]$CertFilePath,

            [Parameter(Mandatory = $false)]
            [System.Security.SecureString]$CertPass,

            [Parameter(Mandatory = $false)]
            [switch]$UseInteractive
        )

        Write-Log -Level INFO -Message "Establishing connection to SharePoint Online Admin: $TargetAdminUrl"

        # Determine Tenant Domain hint from AdminUrl if not explicitly provided
        $inferredTenant = try {
            $uri = [System.Uri]$TargetAdminUrl
            $tenantSubdomain = $uri.Host.Split('.')[0].Replace('-admin', '')
            $tenantSubdomain + ".onmicrosoft.com"
        }
        catch {
            "yourtenant.onmicrosoft.com"
        }

        $effectiveTenant = if (-not [string]::IsNullOrWhiteSpace($Tenant)) {
            $Tenant
        }
        else {
            $inferredTenant
        }

        try {
            # Mode 1: Certificate Thumbprint
            if (-not [string]::IsNullOrWhiteSpace($Thumbprint)) {
                if ([string]::IsNullOrWhiteSpace($AppId)) {
                    throw "App-Only authentication via -CertificateThumbprint requires -ClientId."
                }
                Write-Log -Level INFO -Message "Authenticating via X.509 Certificate Thumbprint '$Thumbprint' (ClientId: $AppId, Tenant: $effectiveTenant)..."
                $conn = Connect-PnPOnline -Url $TargetAdminUrl `
                    -ClientId $AppId `
                    -Tenant $effectiveTenant `
                    -Thumbprint $Thumbprint `
                    -ReturnConnection `
                    -ErrorAction Stop

                $Script:ActiveClientId = $AppId
                $Script:ActiveTenant   = $effectiveTenant
                Write-Log -Level SUCCESS -Message "Connected successfully using Certificate Thumbprint."
                return $conn
            }

            # Mode 2: Certificate File Path
            if (-not [string]::IsNullOrWhiteSpace($CertFilePath)) {
                if ([string]::IsNullOrWhiteSpace($AppId)) {
                    throw "App-Only authentication via -CertificatePath requires -ClientId."
                }
                if (-not (Test-Path -Path $CertFilePath)) {
                    throw "Certificate file not found at path: $CertFilePath"
                }
                Write-Log -Level INFO -Message "Authenticating via Certificate file '$CertFilePath' (ClientId: $AppId, Tenant: $effectiveTenant)..."
                $connectParams = @{
                    Url              = $TargetAdminUrl
                    ClientId         = $AppId
                    Tenant           = $effectiveTenant
                    CertificatePath  = $CertFilePath
                    ReturnConnection = $true
                    ErrorAction      = 'Stop'
                }
                if ($CertPass) {
                    $connectParams['CertificatePassword'] = $CertPass
                }
                $conn = Connect-PnPOnline @connectParams

                $Script:ActiveClientId = $AppId
                $Script:ActiveTenant   = $effectiveTenant
                Write-Log -Level SUCCESS -Message "Connected successfully using Certificate Path."
                return $conn
            }

            # Mode 3: Modern Interactive Browser Authentication
            # 1. Check explicit parameter -ClientId
            $effectiveClientId = $AppId

            # 2. Check environment variables
            if ([string]::IsNullOrWhiteSpace($effectiveClientId)) {
                if (-not [string]::IsNullOrWhiteSpace($env:ENTRAID_APP_ID)) {
                    $effectiveClientId = $env:ENTRAID_APP_ID
                }
                elseif (-not [string]::IsNullOrWhiteSpace($env:PNP_APP_ID)) {
                    $effectiveClientId = $env:PNP_APP_ID
                }
            }

            # 3. Check persistent tenant cache on local disk (AppData\Local\SharePoint-Audit)
            if ([string]::IsNullOrWhiteSpace($effectiveClientId)) {
                $cacheDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath "SharePoint-Audit"
                $cacheFile = Join-Path -Path $cacheDir -ChildPath "${effectiveTenant}_client_id.txt"
                if (Test-Path -Path $cacheFile) {
                    try {
                        $cachedId = [System.IO.File]::ReadAllText($cacheFile).Trim()
                        if ($cachedId -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
                            Write-Log -Level INFO -Message "Loaded cached Entra ID Application ID for '$effectiveTenant': $cachedId"
                            $effectiveClientId = $cachedId
                        }
                    }
                    catch {
                        Write-Log -Level WARN -Message "Unable to read cached ClientId: $($_.Exception.Message)"
                    }
                }
            }

            # 4. If ClientId is still missing, guide registration on-the-fly
            if ([string]::IsNullOrWhiteSpace($effectiveClientId)) {
                Write-Log -Level WARN -Message "No Entra ID Application ID configured for tenant '$effectiveTenant'."

                $shouldRegister = $false
                if ([Environment]::UserInteractive) {
                    Write-Host "`n[?] Microsoft Entra ID interactive authentication requires an Application (Client) ID." -ForegroundColor Yellow
                    Write-Host "[?] Microsoft permanently retired the legacy shared multi-tenant PnP app on Sept 9, 2024." -ForegroundColor Yellow
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

To register a dedicated Entra ID App for your tenant ('$effectiveTenant'), run once in PowerShell 7:
    Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "SharePoint-Audit" -Tenant "$effectiveTenant"

Then re-run this script with the generated Application ID:
    .\Audit-SharePointSecuritySummary.ps1 -AdminUrl "$TargetAdminUrl" -ClientId "<ApplicationId>" -Interactive

Tip: You can also set an environment variable or paste the ID into:
    `$env:ENTRAID_APP_ID = "<ApplicationId>"
"@
                    Write-Log -Level ERROR -Message $guidance
                    throw $guidance
                }
            }

            # Attempt connection with the effective ClientId and Tenant
            $Script:ActiveClientId = $effectiveClientId
            $Script:ActiveTenant   = $effectiveTenant

            $interactiveParams = @{
                Url              = $TargetAdminUrl
                Interactive      = $true
                ClientId         = $effectiveClientId
                Tenant           = $effectiveTenant
                ReturnConnection = $true
                ErrorAction      = 'Stop'
            }

            Write-Log -Level INFO -Message "Initiating modern interactive browser authentication (ClientId: $effectiveClientId, Tenant: $effectiveTenant)..."

            try {
                $conn = Connect-PnPOnline @interactiveParams
                Write-Log -Level SUCCESS -Message "Connected successfully to $TargetAdminUrl via interactive session."
                return $conn
            }
            catch {
                if ($_.Exception.Message -match 'AADSTS700016' -or $_.Exception.Message -match 'was not found in the directory') {
                    Write-Log -Level WARN -Message "Application ID '$effectiveClientId' was not found or not consented in directory '$effectiveTenant'."

                    if ([Environment]::UserInteractive) {
                        Write-Host "`n[!] Application ID '$effectiveClientId' was not found in directory '$effectiveTenant' (AADSTS700016)." -ForegroundColor Yellow
                        Write-Host "[?] Would you like to register a new Entra ID App in '$effectiveTenant' now? (Y/N): " -NoNewline -ForegroundColor Yellow
                        $reAns = Read-Host
                        if ($reAns -match '^[yY1tT]$') {
                            $newId = Invoke-EntraIDAppRegistration -Tenant $effectiveTenant
                            $interactiveParams['ClientId'] = $newId
                            $Script:ActiveClientId = $newId
                            Write-Log -Level INFO -Message "Retrying connection with new Application ID: $newId..."
                            $conn = Connect-PnPOnline @interactiveParams
                            Write-Log -Level SUCCESS -Message "Connected successfully to $TargetAdminUrl via interactive session."
                            return $conn
                        }
                    }
                }
                throw
            }
        }
        catch {
            $authErr = "Failed to establish connection to SharePoint Admin ($TargetAdminUrl): $($_.Exception.Message)"
            Write-Log -Level ERROR -Message $authErr
            throw $authErr
        }
    }

    # ==============================================================================
    # HELPER: TEMPLATE IDENTIFIER RESOLUTION
    # ==============================================================================
    function Get-FriendlyTemplateName {
        [CmdletBinding()]
        [OutputType([string])]
        param([string]$TemplateCode)

        if ([string]::IsNullOrWhiteSpace($TemplateCode)) {
            return "Unknown"
        }

        switch -Wildcard ($TemplateCode) {
            "GROUP#0*"                { return "Modern Team Site (M365 Group)" }
            "SITEPAGE-PUBLISHING#0*"  { return "Communication Site" }
            "STS#3*"                  { return "Modern Team Site (No Group)" }
            "STS#0*"                  { return "Classic Team Site" }
            "BDR#0*"                  { return "Document Center" }
            "BICenterSite#0*"         { return "Business Intelligence Center" }
            "COMMUNITY#0*"            { return "Community Site" }
            "ENTERWIKI#0*"            { return "Enterprise Wiki" }
            "POINTPUBLISHINGHUB#0*"   { return "Hub Site" }
            "POINTPUBLISHINGTOPIC#0*" { return "Topic Site" }
            "SRCHCEN#0*"              { return "Search Center" }
            "VISUSVR#0*"              { return "Visio Process Repository" }
            "APPCATALOG#0*"           { return "App Catalog Site" }
            "SPSPERS#*"               { return "OneDrive / Personal Site" }
            Default                   { return $TemplateCode }
        }
    }

    # ==============================================================================
    # RISK ASSESSMENT ENGINE (Specification Section 3.5 & 2.2)
    # ==============================================================================
    function Measure-SiteSecurityRisk {
        [CmdletBinding()]
        [OutputType([PSCustomObject])]
        param(
            [Parameter(Mandatory = $true)]
            [PSCustomObject]$SiteData
        )

        $riskScore = 0
        $riskFactors = [System.Collections.Generic.List[string]]::new()
        $remediations = [System.Collections.Generic.List[string]]::new()

        $siteUrl          = [string](Get-SafePropertyValue -InputObject $SiteData -PropertyName "Url" -DefaultValue "")
        $siteOwner        = [string](Get-SafePropertyValue -InputObject $SiteData -PropertyName "Owner" -DefaultValue "")
        $siteTemplate     = [string](Get-SafePropertyValue -InputObject $SiteData -PropertyName "Template" -DefaultValue "")
        $adminsCount      = [int](Get-SafePropertyValue -InputObject $SiteData -PropertyName "SiteAdminsCount" -DefaultValue 0)
        $sharing          = [string](Get-SafePropertyValue -InputObject $SiteData -PropertyName "SharingCapability" -DefaultValue "Unknown")
        $anonExpDays      = [int](Get-SafePropertyValue -InputObject $SiteData -PropertyName "AnonymousLinkExpirationInDays" -DefaultValue 0)
        $groupPrivacy     = [string](Get-SafePropertyValue -InputObject $SiteData -PropertyName "GroupPrivacy" -DefaultValue "N/A")
        $sensitivity      = [string](Get-SafePropertyValue -InputObject $SiteData -PropertyName "SensitivityLabel" -DefaultValue "None")
        $hasBrokenInherit = [bool](Get-SafePropertyValue -InputObject $SiteData -PropertyName "HasBrokenInheritance" -DefaultValue $false)
        $storageUsed      = [int](Get-SafePropertyValue -InputObject $SiteData -PropertyName "StorageUsageMB" -DefaultValue 0)
        $storageLimit     = [int](Get-SafePropertyValue -InputObject $SiteData -PropertyName "StorageLimitMB" -DefaultValue 0)

        # Vector 1: Ownership & Orphaned Sites
        $isSystemSite = ($siteOwner -eq 'SharePoint System' -or $siteTemplate -in @('SRCHCEN#0', 'APPCATALOG#0', 'POINTPUBLISHINGHUB#0', 'COMMUNITYPORTAL#0', 'BLANKINTERNETCONTAINER#0', 'EDISC#0'))
        $isOrphaned = ($adminsCount -eq 0 -or [string]::IsNullOrWhiteSpace($siteOwner) -or $siteOwner -eq 'N/A')

        if ($isSystemSite) {
            # System infrastructure sites (Search, Hub, AppCatalog) do not require user owners
        }
        elseif ($isOrphaned) {
            $riskScore += 45
            $riskFactors.Add("ORPHANED_SITE: 0 active Site Collection Administrators or missing owner")
            $remediations.Add("Assign at least two designated site administrators: Set-PnPTenantSite -Url '$siteUrl' -Owners 'admin@domain.com'")
        }
        elseif ($adminsCount -gt 5) {
            $riskScore += 25
            $riskFactors.Add("EXCESSIVE_ADMINS: $adminsCount Site Collection Administrators detected (Threshold: >5)")
            $remediations.Add("Review and reduce excessive site collection administrators down to 2-4 primary owners.")
        }
        elseif ($adminsCount -eq 1 -and $siteOwner -notin @('M365 Group Owners', 'M365 Group-Managed')) {
            $riskScore += 10
            $riskFactors.Add("SINGLE_ADMIN: Single administrator represents an administrative single point of failure")
            $remediations.Add("Assign a secondary site administrator to prevent orphaned state if current owner leaves.")
        }

        # Vector 2: External Sharing & Anonymous Access
        if ($sharing -eq 'ExternalUserAndGuestSharing') {
            $riskScore += 30
            $riskFactors.Add("ANYONE_LINKS: Site permits Anonymous/Anyone link sharing without authentication")

            if ($anonExpDays -le 0) {
                $riskScore += 20
                $riskFactors.Add("NO_EXPIRATION: Anonymous links do not enforce expiration policies")
                $remediations.Add("Enforce anonymous link expiration: Set-PnPTenantSite -Url '$siteUrl' -AnonymousLinkExpirationInDays 14")
            }
            $remediations.Add("Restrict sharing to authenticated guests: Set-PnPTenantSite -Url '$siteUrl' -SharingCapability ExistingExternalUserSharingOnly")
        }
        elseif ($sharing -eq 'ExistingExternalUserSharingOnly') {
            $riskScore += 15
            $riskFactors.Add("GUEST_SHARING: External sharing enabled for existing directory guests")
        }
        elseif ($sharing -eq 'ExternalUserSharingDisabled' -or $sharing -eq 'Disabled') {
            # Compliant / Secure sharing posture
        }

        # Vector 3: M365 Group Privacy & Public Exposure
        if ($groupPrivacy -eq 'Public') {
            $riskScore += 15
            $riskFactors.Add("PUBLIC_GROUP: Connected M365 Group is Public; accessible to all internal employees")
            if ($sharing -eq 'ExternalUserAndGuestSharing') {
                $riskScore += 20
                $riskFactors.Add("PUBLIC_WITH_ANONYMOUS: Public group allows unrestricted anonymous file sharing")
            }
            $remediations.Add("Review group privacy settings and convert to Private if confidential assets are stored.")
        }

        # Vector 4: Sensitivity Labels
        if ([string]::IsNullOrWhiteSpace($sensitivity) -or $sensitivity -eq 'None') {
            if ($sharing -eq 'ExternalUserAndGuestSharing' -or $sharing -eq 'ExistingExternalUserSharingOnly') {
                $riskScore += 15
                $riskFactors.Add("UNLABELED_SHARING: External sharing is enabled on an unlabeled site collection")
                $remediations.Add("Classify site collection by assigning an appropriate Microsoft Purview Sensitivity Label.")
            }
        }
        elseif ($sensitivity -like "*Confidential*" -or $sensitivity -like "*Secret*") {
            if ($sharing -eq 'ExternalUserAndGuestSharing') {
                $riskScore += 35
                $riskFactors.Add("CONFIDENTIAL_LEAK_RISK: Confidential/Secret labeled site allows Anonymous/Anyone link sharing")
                $remediations.Add("Immediately disable anyone links: Set-PnPTenantSite -Url '$siteUrl' -SharingCapability ExternalUserSharingDisabled")
            }
        }

        # Vector 5: Broken Permission Inheritance
        if ($hasBrokenInherit) {
            $riskScore += 15
            $riskFactors.Add("BROKEN_INHERITANCE: Site or document libraries contain unique role assignments")
            $remediations.Add("Audit subordinate library permissions and restore inheritance where custom access is stale.")
        }

        # Vector 6: Storage Consumption
        if ($storageUsed -gt 0 -and $storageLimit -gt 0) {
            $storagePct = [Math]::Round(($storageUsed / $storageLimit) * 100, 1)
            if ($storagePct -ge 90) {
                $riskScore += 10
                $riskFactors.Add("STORAGE_CAPACITY: Storage utilization at $storagePct% of maximum allocated quota")
                $remediations.Add("Increase storage quota or purge unneeded versions/recycle bin content.")
            }
        }

        # Determine Final Risk Level
        $riskLevel = if ($riskScore -ge 50 -or $riskFactors -match "ORPHANED_SITE" -or $riskFactors -match "CONFIDENTIAL_LEAK_RISK") {
            'Critical'
        }
        elseif ($riskScore -ge 30 -or $riskFactors -match "ANYONE_LINKS" -or $riskFactors -match "EXCESSIVE_ADMINS") {
            'High'
        }
        elseif ($riskScore -ge 15 -or $riskFactors -match "GUEST_SHARING" -or $riskFactors -match "BROKEN_INHERITANCE") {
            'Medium'
        }
        else {
            'Low'
        }

        if ($riskFactors.Count -eq 0) {
            $riskFactors.Add("COMPLIANT: Standard governance posture; no immediate risks identified")
        }
        if ($remediations.Count -eq 0) {
            $remediations.Add("Maintain regular access reviews and audit logging.")
        }

        return [PSCustomObject]@{
            RiskLevel    = $riskLevel
            RiskScore    = $riskScore
            RiskFactors  = ($riskFactors -join " | ")
            Remediations = ($remediations -join " | ")
        }
    }

    # ==============================================================================
    # CONSOLE DASHBOARD RENDERING (Specification Section 2.2 & 3.5)
    # ==============================================================================
    function Show-ConsoleAuditDashboard {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param(
            [Parameter(Mandatory = $true)]
            [System.Collections.Generic.List[PSCustomObject]]$Records
        )

        $totalSites     = $Records.Count
        $criticalCount  = @($Records | Where-Object { $_.RiskLevel -eq 'Critical' }).Count
        $highCount      = @($Records | Where-Object { $_.RiskLevel -eq 'High' }).Count
        $mediumCount    = @($Records | Where-Object { $_.RiskLevel -eq 'Medium' }).Count
        $lowCount       = @($Records | Where-Object { $_.RiskLevel -eq 'Low' }).Count

        $extSharing     = @($Records | Where-Object { $_.SharingCapability -in @('ExternalUserAndGuestSharing', 'ExistingExternalUserSharingOnly') }).Count
        $anyoneCount    = @($Records | Where-Object { $_.SharingCapability -eq 'ExternalUserAndGuestSharing' }).Count
        $orphanedCount  = @($Records | Where-Object { $_.RiskFactors -like "*ORPHANED_SITE*" }).Count
        $excessiveAdmin = @($Records | Where-Object { $_.SiteAdminsCount -gt 5 }).Count
        $brokenInherit  = @($Records | Where-Object { $_.HasBrokenInheritance -eq $true }).Count

        Write-Host "`n" -NoNewline
        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host "           SHAREPOINT ONLINE SECURITY AUDIT EXECUTIVE SUMMARY" -ForegroundColor Cyan
        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host " Total Site Collections Scanned     : $totalSites" -ForegroundColor White
        Write-Host " [CRITICAL] Immediate Action Sites   : $criticalCount" -ForegroundColor $(if ($criticalCount -gt 0) { 'Red' } else { 'Green' })
        Write-Host " [HIGH]     High Risk Posture Sites  : $highCount" -ForegroundColor $(if ($highCount -gt 0) { 'Red' } else { 'Green' })
        Write-Host " [MEDIUM]   Elevated Warning Sites   : $mediumCount" -ForegroundColor $(if ($mediumCount -gt 0) { 'Yellow' } else { 'White' })
        Write-Host " [LOW]      Compliant / Secure Sites : $lowCount" -ForegroundColor Green
        Write-Host "--------------------------------------------------------------------------------" -ForegroundColor DarkGray
        Write-Host " External Sharing Enabled (Any Mode) : $extSharing site(s)" -ForegroundColor $(if ($extSharing -gt 0) { 'Yellow' } else { 'Green' })
        Write-Host " Anonymous / Anyone Links Permitted  : $anyoneCount site(s)" -ForegroundColor $(if ($anyoneCount -gt 0) { 'Red' } else { 'Green' })
        Write-Host " Orphaned Sites (0 Active Owners)    : $orphanedCount site(s)" -ForegroundColor $(if ($orphanedCount -gt 0) { 'Red' } else { 'Green' })
        Write-Host " Excessive Admins (>5 Administrators): $excessiveAdmin site(s)" -ForegroundColor $(if ($excessiveAdmin -gt 0) { 'Yellow' } else { 'Green' })
        Write-Host " Broken Permission Inheritance       : $brokenInherit site(s)" -ForegroundColor $(if ($brokenInherit -gt 0) { 'Yellow' } else { 'White' })
        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host "`nDETAILED SITE SECURITY POSTURE AUDIT TABLE:" -ForegroundColor Yellow
        Write-Host "--------------------------------------------------------------------------------" -ForegroundColor DarkGray

        $tableRows = foreach ($rec in $Records) {
            $shortUrl = if ($rec.SiteUrl.Length -gt 45) { $rec.SiteUrl.Substring(0, 42) + "..." } else { $rec.SiteUrl }
            $shortTitle = if ($rec.SiteTitle.Length -gt 25) { $rec.SiteTitle.Substring(0, 22) + "..." } else { $rec.SiteTitle }
            $shortSharing = switch ($rec.SharingCapability) {
                'ExternalUserAndGuestSharing'     { 'Anyone (Anon)' }
                'ExistingExternalUserSharingOnly' { 'Existing Guests' }
                'Disabled'                        { 'Disabled' }
                'ExternalUserSharingDisabled'     { 'Disabled' }
                Default                           { $rec.SharingCapability }
            }

            [PSCustomObject]@{
                "Risk Level" = $rec.RiskLevel
                "Score"      = $rec.RiskScore
                "Site Title" = $shortTitle
                "Sharing"    = $shortSharing
                "Admins"     = $rec.SiteAdminsCount
                "Privacy"    = if ($rec.GroupPrivacy) { $rec.GroupPrivacy } else { "N/A" }
                "Label"      = if ($rec.SensitivityLabel) { $rec.SensitivityLabel } else { "-" }
                "Broken"     = if ($rec.HasBrokenInheritance) { "Yes" } else { "No" }
                "Site URL"   = $shortUrl
            }
        }

        # Print sorted by risk priority: Critical -> High -> Medium -> Low
        $priorityOrder = @{ 'Critical' = 1; 'High' = 2; 'Medium' = 3; 'Low' = 4 }
        $sortedRows = $tableRows | Sort-Object -Property @{ Expression = { $priorityOrder[$_."Risk Level"] } }, @{ Expression = "Score"; Descending = $true }

        $sortedRows | Format-Table -AutoSize | Out-Host

        Write-Host "--------------------------------------------------------------------------------" -ForegroundColor DarkGray
        Write-Host "[*] Complete raw datasets exported to CSV and HTML report." -ForegroundColor Gray
        Write-Host "[*] Use -PassThru to stream [PSCustomObject] records into pipeline or Out-GridView." -ForegroundColor Gray
        Write-Host "================================================================================`n" -ForegroundColor DarkGray
    }

    # ==============================================================================
    # CSV EXPORT ENGINE
    # ==============================================================================
    function Export-AuditCsv {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [System.Collections.Generic.List[PSCustomObject]]$Records,

            [Parameter(Mandatory = $true)]
            [string]$FilePath
        )

        Write-Log -Level INFO -Message "Exporting $($Records.Count) audit records to CSV: $FilePath"

        try {
            $parentDir = Split-Path -Path $FilePath -Parent
            if (-not [string]::IsNullOrWhiteSpace($parentDir) -and -not (Test-Path -Path $parentDir)) {
                $null = New-Item -Path $parentDir -ItemType Directory -Force -ErrorAction Stop
            }

            $Records | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8 -Force
            Write-Log -Level SUCCESS -Message "CSV export completed successfully: $FilePath"
        }
        catch {
            Write-Log -Level ERROR -Message "Failed to export CSV: $($_.Exception.Message)"
        }
    }

    # ==============================================================================
    # HTML EXECUTIVE DASHBOARD GENERATOR (Specification Section 2.2)
    # ==============================================================================
    function Export-AuditHtmlReport {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [System.Collections.Generic.List[PSCustomObject]]$Records,

            [Parameter(Mandatory = $true)]
            [string]$FilePath,

            [Parameter(Mandatory = $false)]
            [string]$TenantAdminEndpoint
        )

        Write-Log -Level INFO -Message "Generating standalone HTML executive summary report: $FilePath"

        $scanDate       = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $totalSites     = $Records.Count
        $criticalCount  = @($Records | Where-Object { $_.RiskLevel -eq 'Critical' }).Count
        $highCount      = @($Records | Where-Object { $_.RiskLevel -eq 'High' }).Count
        $mediumCount    = @($Records | Where-Object { $_.RiskLevel -eq 'Medium' }).Count
        $lowCount       = @($Records | Where-Object { $_.RiskLevel -eq 'Low' }).Count

        $extSharing     = @($Records | Where-Object { $_.SharingCapability -in @('ExternalUserAndGuestSharing', 'ExistingExternalUserSharingOnly') }).Count
        $anyoneCount    = @($Records | Where-Object { $_.SharingCapability -eq 'ExternalUserAndGuestSharing' }).Count
        $orphanedCount  = @($Records | Where-Object { $_.RiskFactors -like "*ORPHANED_SITE*" }).Count
        $excessiveAdmin = @($Records | Where-Object { $_.SiteAdminsCount -gt 5 }).Count

        # Build table rows HTML
        $tableHtmlBuilder = [System.Text.StringBuilder]::new()
        foreach ($r in $Records) {
            $badgeClass = switch ($r.RiskLevel) {
                'Critical' { 'badge-critical' }
                'High'     { 'badge-high' }
                'Medium'   { 'badge-medium' }
                'Low'      { 'badge-low' }
                Default    { 'badge-low' }
            }

            $sharingBadge = switch ($r.SharingCapability) {
                'ExternalUserAndGuestSharing'     { '<span class="badge badge-critical">Anyone / Anonymous</span>' }
                'ExistingExternalUserSharingOnly' { '<span class="badge badge-medium">Existing Guests</span>' }
                'Disabled'                        { '<span class="badge badge-low">Disabled</span>' }
                'ExternalUserSharingDisabled'     { '<span class="badge badge-low">Disabled</span>' }
                Default                           { "<span class=`"badge badge-low`">$($r.SharingCapability)</span>" }
            }

            $privacyBadge = if ($r.GroupPrivacy -eq 'Public') {
                '<span class="badge badge-medium">Public</span>'
            }
            elseif ($r.GroupPrivacy -eq 'Private') {
                '<span class="badge badge-low">Private</span>'
            }
            else {
                '<span class="text-muted">N/A</span>'
            }

            $labelDisplay = if (-not [string]::IsNullOrWhiteSpace($r.SensitivityLabel) -and $r.SensitivityLabel -ne 'None') {
                "<span class=`"badge badge-info`">$([System.Web.HttpUtility]::HtmlEncode($r.SensitivityLabel))</span>"
            }
            else {
                '<span class="text-muted">Unlabeled</span>'
            }

            $brokenDisplay = if ($r.HasBrokenInheritance) {
                '<span class="badge badge-medium">Unique Perms</span>'
            }
            else {
                '<span class="text-muted">Inherited</span>'
            }

            $encodedTitle = [System.Web.HttpUtility]::HtmlEncode($r.SiteTitle)
            $encodedUrl   = [System.Web.HttpUtility]::HtmlEncode($r.SiteUrl)
            $encodedOwner = [System.Web.HttpUtility]::HtmlEncode($r.Owner)
            $encodedFactors = [System.Web.HttpUtility]::HtmlEncode($r.RiskFactors)
            $encodedRemediation = [System.Web.HttpUtility]::HtmlEncode($r.Remediations)

            $rowHtml = @"
            <tr class="audit-row" data-risk="$($r.RiskLevel)" data-sharing="$($r.SharingCapability)" data-orphaned="$(if ($r.RiskFactors -like '*ORPHANED_SITE*') { 'true' } else { 'false' })">
                <td><span class="badge $badgeClass">$($r.RiskLevel)</span></td>
                <td><strong>$($r.RiskScore)</strong></td>
                <td>
                    <div class="site-title">$encodedTitle</div>
                    <div class="site-url"><a href="$encodedUrl" target="_blank" rel="noopener noreferrer">$encodedUrl</a></div>
                </td>
                <td>$sharingBadge</td>
                <td>
                    <span class="admin-count $(if ($r.SiteAdminsCount -gt 5) { 'text-danger font-bold' } elseif ($r.SiteAdminsCount -eq 0) { 'text-critical font-bold' })">
                        $($r.SiteAdminsCount)
                    </span>
                    <div class="admin-owner text-muted" title="$encodedOwner">$encodedOwner</div>
                </td>
                <td>$privacyBadge</td>
                <td>$labelDisplay</td>
                <td>$brokenDisplay</td>
                <td>
                    <div class="risk-factors">$encodedFactors</div>
                    <div class="remediation-hint"><em>$encodedRemediation</em></div>
                </td>
            </tr>
"@
            $null = $tableHtmlBuilder.AppendLine($rowHtml)
        }

        $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SharePoint Online Security Posture Audit - Executive Report</title>
    <style>
        :root {
            --bg-primary: #0f172a;
            --bg-secondary: #1e293b;
            --bg-card: #1e293b;
            --bg-hover: #334155;
            --text-primary: #f8fafc;
            --text-secondary: #94a3b8;
            --text-muted: #64748b;
            --border-color: #334155;
            --color-critical: #ef4444;
            --color-high: #f97316;
            --color-medium: #eab308;
            --color-low: #10b981;
            --color-info: #3b82f6;
            --color-purple: #a855f7;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background-color: var(--bg-primary);
            color: var(--text-primary);
            line-height: 1.5;
            padding: 24px;
        }

        .container {
            max-width: 1540px;
            margin: 0 auto;
        }

        /* HEADER */
        .header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding-bottom: 24px;
            border-bottom: 1px solid var(--border-color);
            margin-bottom: 24px;
            flex-wrap: wrap;
            gap: 16px;
        }

        .header-title h1 {
            font-size: 1.85rem;
            font-weight: 700;
            color: var(--text-primary);
            display: flex;
            align-items: center;
            gap: 12px;
        }

        .header-title p {
            color: var(--text-secondary);
            font-size: 0.95rem;
            margin-top: 4px;
        }

        .header-meta {
            text-align: right;
            font-size: 0.85rem;
            color: var(--text-muted);
        }

        .header-meta strong {
            color: var(--text-secondary);
        }

        /* METRIC CARDS */
        .metrics-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }

        .metric-card {
            background-color: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 10px;
            padding: 18px;
            position: relative;
            overflow: hidden;
            transition: transform 0.15s ease;
        }

        .metric-card:hover {
            transform: translateY(-2px);
        }

        .metric-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 4px;
            background: var(--text-muted);
        }

        .card-total::before { background: var(--color-info); }
        .card-critical::before { background: var(--color-critical); }
        .card-high::before { background: var(--color-high); }
        .card-medium::before { background: var(--color-medium); }
        .card-low::before { background: var(--color-low); }
        .card-sharing::before { background: var(--color-purple); }

        .metric-label {
            font-size: 0.82rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: var(--text-secondary);
            margin-bottom: 6px;
        }

        .metric-value {
            font-size: 2.1rem;
            font-weight: 700;
        }

        .text-critical { color: var(--color-critical); }
        .text-high { color: var(--color-high); }
        .text-medium { color: var(--color-medium); }
        .text-low { color: var(--color-low); }
        .text-info { color: var(--color-info); }
        .text-purple { color: var(--color-purple); }
        .text-danger { color: #f87171; }
        .font-bold { font-weight: 700; }

        /* CONTROLS & FILTER BAR */
        .controls-bar {
            background-color: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 10px;
            padding: 16px;
            margin-bottom: 24px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 16px;
        }

        .search-box {
            flex: 1;
            min-width: 280px;
            position: relative;
        }

        .search-input {
            width: 100%;
            padding: 10px 14px;
            background-color: var(--bg-primary);
            border: 1px solid var(--border-color);
            border-radius: 6px;
            color: var(--text-primary);
            font-size: 0.9rem;
            outline: none;
        }

        .search-input:focus {
            border-color: var(--color-info);
        }

        .filter-buttons {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
        }

        .filter-btn {
            background-color: var(--bg-primary);
            border: 1px solid var(--border-color);
            color: var(--text-secondary);
            padding: 8px 14px;
            border-radius: 6px;
            font-size: 0.82rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.15s ease;
        }

        .filter-btn:hover {
            background-color: var(--bg-hover);
            color: var(--text-primary);
        }

        .filter-btn.active {
            background-color: var(--color-info);
            border-color: var(--color-info);
            color: #ffffff;
        }

        /* TABLE */
        .table-container {
            background-color: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2);
        }

        table {
            width: 100%;
            border-collapse: collapse;
            text-align: left;
            font-size: 0.88rem;
        }

        th {
            background-color: #172033;
            color: var(--text-secondary);
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.75rem;
            letter-spacing: 0.05em;
            padding: 14px 16px;
            border-bottom: 1px solid var(--border-color);
            white-space: nowrap;
        }

        td {
            padding: 14px 16px;
            border-bottom: 1px solid var(--border-color);
            vertical-align: top;
        }

        tr:hover {
            background-color: rgba(255, 255, 255, 0.02);
        }

        .site-title {
            font-weight: 600;
            color: var(--text-primary);
        }

        .site-url a {
            color: var(--color-info);
            text-decoration: none;
            font-size: 0.78rem;
            word-break: break-all;
        }

        .site-url a:hover {
            text-decoration: underline;
        }

        .admin-owner {
            font-size: 0.75rem;
            max-width: 160px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }

        .risk-factors {
            font-size: 0.8rem;
            color: #cbd5e1;
            margin-bottom: 4px;
        }

        .remediation-hint {
            font-size: 0.75rem;
            color: var(--text-muted);
        }

        /* BADGES */
        .badge {
            display: inline-block;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 0.75rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.04em;
        }

        .badge-critical {
            background-color: rgba(239, 68, 68, 0.15);
            color: #fca5a5;
            border: 1px solid rgba(239, 68, 68, 0.3);
        }

        .badge-high {
            background-color: rgba(249, 115, 22, 0.15);
            color: #fdba74;
            border: 1px solid rgba(249, 115, 22, 0.3);
        }

        .badge-medium {
            background-color: rgba(234, 179, 8, 0.15);
            color: #fde047;
            border: 1px solid rgba(234, 179, 8, 0.3);
        }

        .badge-low {
            background-color: rgba(16, 185, 129, 0.15);
            color: #6ee7b7;
            border: 1px solid rgba(16, 185, 129, 0.3);
        }

        .badge-info {
            background-color: rgba(59, 130, 246, 0.15);
            color: #93c5fd;
            border: 1px solid rgba(59, 130, 246, 0.3);
        }

        /* REMEDIATION PLAYBOOK */
        .remediation-section {
            margin-top: 32px;
            background-color: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 10px;
            padding: 24px;
        }

        .remediation-section h2 {
            font-size: 1.25rem;
            margin-bottom: 16px;
            color: var(--text-primary);
        }

        .playbook-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
            gap: 16px;
        }

        .playbook-card {
            background-color: var(--bg-primary);
            border: 1px solid var(--border-color);
            border-radius: 8px;
            padding: 16px;
        }

        .playbook-card h3 {
            font-size: 0.95rem;
            margin-bottom: 8px;
            color: var(--color-info);
        }

        .playbook-card p {
            font-size: 0.82rem;
            color: var(--text-secondary);
            margin-bottom: 8px;
        }

        .code-block {
            background-color: #020617;
            border: 1px solid #1e293b;
            border-radius: 4px;
            padding: 8px 10px;
            font-family: monospace;
            font-size: 0.78rem;
            color: #38bdf8;
            overflow-x: auto;
        }

        /* FOOTER */
        .footer {
            margin-top: 32px;
            padding-top: 20px;
            border-top: 1px solid var(--border-color);
            text-align: center;
            font-size: 0.82rem;
            color: var(--text-muted);
        }

        .footer a {
            color: var(--text-secondary);
            text-decoration: none;
        }

        .footer a:hover {
            text-decoration: underline;
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- HEADER -->
        <header class="header">
            <div class="header-title">
                <h1>
                    <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#38bdf8" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>
                    </svg>
                    SharePoint Online Security Posture Audit
                </h1>
                <p>Tenant-wide Governance, External Sharing, Orphaned Site & Inheritance Exposure Assessment</p>
            </div>
            <div class="header-meta">
                <div>Scan Date: <strong>$scanDate</strong></div>
                <div>Admin Endpoint: <strong>$TenantAdminEndpoint</strong></div>
                <div>Tool: <strong>Audit-SharePointSecuritySummary v1.0.1</strong></div>
            </div>
        </header>

        <!-- METRIC CARDS -->
        <section class="metrics-grid">
            <div class="metric-card card-total">
                <div class="metric-label">Total Sites Audited</div>
                <div class="metric-value text-info">$totalSites</div>
            </div>
            <div class="metric-card card-critical">
                <div class="metric-label">Critical Risk Sites</div>
                <div class="metric-value text-critical">$criticalCount</div>
            </div>
            <div class="metric-card card-high">
                <div class="metric-label">High Risk Sites</div>
                <div class="metric-value text-high">$highCount</div>
            </div>
            <div class="metric-card card-medium">
                <div class="metric-label">Medium Risk Sites</div>
                <div class="metric-value text-medium">$mediumCount</div>
            </div>
            <div class="metric-card card-low">
                <div class="metric-label">Compliant / Low Risk</div>
                <div class="metric-value text-low">$lowCount</div>
            </div>
            <div class="metric-card card-sharing">
                <div class="metric-label">Anonymous Sharing</div>
                <div class="metric-value text-purple">$anyoneCount</div>
            </div>
            <div class="metric-card card-critical">
                <div class="metric-label">Orphaned Sites</div>
                <div class="metric-value text-critical">$orphanedCount</div>
            </div>
            <div class="metric-card card-high">
                <div class="metric-label">Excessive Admins (>5)</div>
                <div class="metric-value text-high">$excessiveAdmin</div>
            </div>
        </section>

        <!-- CONTROLS & FILTER BAR -->
        <section class="controls-bar">
            <div class="search-box">
                <input type="text" id="searchInput" class="search-input" placeholder="Search by site title, URL, owner, or risk factor..." onkeyup="filterTable()">
            </div>
            <div class="filter-buttons">
                <button class="filter-btn active" onclick="setFilter('all', this)">All ($totalSites)</button>
                <button class="filter-btn" onclick="setFilter('Critical', this)">Critical ($criticalCount)</button>
                <button class="filter-btn" onclick="setFilter('High', this)">High ($highCount)</button>
                <button class="filter-btn" onclick="setFilter('Medium', this)">Medium ($mediumCount)</button>
                <button class="filter-btn" onclick="setFilter('Low', this)">Low ($lowCount)</button>
                <button class="filter-btn" onclick="setFilter('sharing', this)">External Sharing ($extSharing)</button>
                <button class="filter-btn" onclick="setFilter('orphaned', this)">Orphaned ($orphanedCount)</button>
            </div>
        </section>

        <!-- AUDIT TABLE -->
        <section class="table-container">
            <table>
                <thead>
                    <tr>
                        <th>Risk Level</th>
                        <th>Score</th>
                        <th>Site Collection</th>
                        <th>Sharing Governance</th>
                        <th>Admins / Owner</th>
                        <th>M365 Privacy</th>
                        <th>Sensitivity Label</th>
                        <th>Inheritance</th>
                        <th>Risk Drivers & Remediation</th>
                    </tr>
                </thead>
                <tbody id="auditTableBody">
$($tableHtmlBuilder.ToString())
                </tbody>
            </table>
        </section>

        <!-- REMEDIATION PLAYBOOK -->
        <section class="remediation-section">
            <h2>Recommended Governance Remediation Playbook</h2>
            <div class="playbook-grid">
                <div class="playbook-card">
                    <h3>1. Disable Anonymous/Anyone Links</h3>
                    <p>Restrict site sharing to authenticated organization members or invited directory guests.</p>
                    <div class="code-block">Set-PnPTenantSite -Url "&lt;SiteUrl&gt;" -SharingCapability ExistingExternalUserSharingOnly</div>
                </div>
                <div class="playbook-card">
                    <h3>2. Enforce Link Expiration Policy</h3>
                    <p>Enforce mandatory expiration on anyone links to prevent stale public file access.</p>
                    <div class="code-block">Set-PnPTenantSite -Url "&lt;SiteUrl&gt;" -AnonymousLinkExpirationInDays 14</div>
                </div>
                <div class="playbook-card">
                    <h3>3. Remediate Orphaned Sites</h3>
                    <p>Assign designated organizational owners to orphaned site collections with 0 administrators.</p>
                    <div class="code-block">Set-PnPTenantSite -Url "&lt;SiteUrl&gt;" -Owners "sec-admin@contoso.com"</div>
                </div>
                <div class="playbook-card">
                    <h3>4. Reduce Excessive Administrators</h3>
                    <p>Review site collections with more than 5 administrators and apply least-privilege RBAC.</p>
                    <div class="code-block">Remove-PnPSiteCollectionAdmin -Owners "user@contoso.com"</div>
                </div>
            </div>
        </section>

        <!-- FOOTER -->
        <footer class="footer">
            <p>Report generated by <strong>Audit-SharePointSecuritySummary.ps1</strong> &bull; Author: <a href="mailto:roman.pindela@gmail.com">Roman Pindela</a> &bull; GitHub: <a href="https://github.com/romanpindela" target="_blank" rel="noopener noreferrer">romanpindela</a></p>
        </footer>
    </div>

    <!-- CLIENT-SIDE FILTERING JAVASCRIPT -->
    <script>
        let currentFilter = 'all';

        function setFilter(filterType, btnElement) {
            currentFilter = filterType;
            document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
            if (btnElement) btnElement.classList.add('active');
            filterTable();
        }

        function filterTable() {
            const searchVal = document.getElementById('searchInput').value.toLowerCase();
            const rows = document.querySelectorAll('#auditTableBody tr.audit-row');

            rows.forEach(row => {
                const text = row.innerText.toLowerCase();
                const risk = row.getAttribute('data-risk');
                const sharing = row.getAttribute('data-sharing');
                const isOrphaned = (row.getAttribute('data-orphaned') || '').toLowerCase() === 'true';

                let matchesFilter = false;
                if (currentFilter === 'all') {
                    matchesFilter = true;
                } else if (currentFilter === 'sharing') {
                    matchesFilter = (sharing === 'ExternalUserAndGuestSharing' || sharing === 'ExistingExternalUserSharingOnly');
                } else if (currentFilter === 'orphaned') {
                    matchesFilter = isOrphaned;
                } else if (risk === currentFilter) {
                    matchesFilter = true;
                }

                const matchesSearch = text.includes(searchVal);

                if (matchesFilter && matchesSearch) {
                    row.style.display = '';
                } else {
                    row.style.display = 'none';
                }
            });
        }
    </script>
</body>
</html>
"@

        try {
            $parentDir = Split-Path -Path $FilePath -Parent
            if (-not [string]::IsNullOrWhiteSpace($parentDir) -and -not (Test-Path -Path $parentDir)) {
                $null = New-Item -Path $parentDir -ItemType Directory -Force -ErrorAction Stop
            }

            [System.IO.File]::WriteAllText($FilePath, $htmlContent, [System.Text.UTF8Encoding]::new($false))
            Write-Log -Level SUCCESS -Message "HTML report generated successfully: $FilePath"
        }
        catch {
            Write-Log -Level ERROR -Message "Failed to generate HTML report: $($_.Exception.Message)"
        }
    }
}

process {
    try {
        # 1. Check PowerShell Core compatibility & auto-relaunch
        Assert-PowerShellVersion -BoundParams $PSBoundParameters

        # 2. Initialize enterprise transcript logging
        Initialize-AuditLogging -TargetDirectory $LogPath
        Write-Log -Level INFO -Message "Starting SharePoint Online Security Posture Audit ($($Script:Metadata.Title) v$($Script:Metadata.Version))..."

        # 3. Verify PnP Module
        Assert-PnPModule -ForceInstall:$Force

        # 4. Resolve SharePoint Online Admin URL
        if ([string]::IsNullOrWhiteSpace($AdminUrl)) {
            if ([Environment]::UserInteractive) {
                $AdminUrl = Read-AdminUrlInteractive
            }
            if ([string]::IsNullOrWhiteSpace($AdminUrl)) {
                throw "SharePoint Online Admin URL is mandatory. Pass -AdminUrl 'https://contoso-admin.sharepoint.com'."
            }
        }

        $resolvedAdminUrl = Resolve-SharePointAdminUrl -Url $AdminUrl
        if ([string]::IsNullOrWhiteSpace($resolvedAdminUrl)) {
            throw "Invalid SharePoint Online Admin URL: $AdminUrl"
        }

        # 5. Connect to SharePoint Online Admin Center
        $Script:ActiveConnection = Connect-AuditSession `
            -TargetAdminUrl $resolvedAdminUrl `
            -AppId $ClientId `
            -Tenant $TenantId `
            -Thumbprint $CertificateThumbprint `
            -CertFilePath $CertificatePath `
            -CertPass $CertificatePassword `
            -UseInteractive:$Interactive

        # 6. Discover Site Collections
        Write-Log -Level INFO -Message "Discovering site collections from tenant admin endpoint: $resolvedAdminUrl"
        Write-Progress -Activity "Auditing SharePoint Security Posture" -Status "Discovering site collections..." -PercentComplete 10

        $allSites = try {
            Get-PnPTenantSite -Detailed -Connection $Script:ActiveConnection -ErrorAction Stop
        }
        catch {
            Write-Log -Level WARN -Message "Get-PnPTenantSite -Detailed failed: $($_.Exception.Message). Retrying standard query..."
            Get-PnPTenantSite -Connection $Script:ActiveConnection -ErrorAction Stop
        }

        if (-not $allSites -or (@($allSites).Count -eq 0)) {
            Write-Log -Level WARN -Message "No site collections returned from tenant admin endpoint."
            return
        }

        # Auto-apply AutoSiteFilter if SiteFilter was not explicitly passed
        if ([string]::IsNullOrWhiteSpace($SiteFilter) -and -not [string]::IsNullOrWhiteSpace($Script:AutoSiteFilter)) {
            $SiteFilter = $Script:AutoSiteFilter
            Write-Log -Level INFO -Message "Targeting specific site collection filter: $SiteFilter"
        }

        # Filter out personal sites (/personal/) unless explicitly requested
        $filteredSites = [System.Collections.Generic.List[object]]::new()
        foreach ($s in @($allSites)) {
            $sUrl = [string](Get-SafePropertyValue -InputObject $s -PropertyName "Url" -DefaultValue "")
            $sTitle = [string](Get-SafePropertyValue -InputObject $s -PropertyName "Title" -DefaultValue "")
            if (-not $IncludePersonalSites -and ($sUrl -like "*/personal/*" -or $sUrl -like "*-my.sharepoint.com*")) {
                continue
            }
            if (-not [string]::IsNullOrWhiteSpace($SiteFilter)) {
                if ($sUrl -notlike $SiteFilter -and $sTitle -notlike $SiteFilter) {
                    continue
                }
            }
            $filteredSites.Add($s)
        }

        $totalFound = $filteredSites.Count
        Write-Log -Level SUCCESS -Message "Discovered $totalFound site collection(s) to evaluate (Personal sites excluded: $(-not $IncludePersonalSites))."

        if ($totalFound -eq 0) {
            Write-Log -Level WARN -Message "No site collections matched the filter '$SiteFilter'."
            return
        }

        # 7. Audit Each Site Collection
        $auditRecords = [System.Collections.Generic.List[PSCustomObject]]::new()
        $counter = 0

        foreach ($site in $filteredSites) {
            $counter++
            $siteUrl   = [string](Get-SafePropertyValue -InputObject $site -PropertyName "Url" -DefaultValue "")
            $siteTitle = [string](Get-SafePropertyValue -InputObject $site -PropertyName "Title" -DefaultValue "Untitled Site")

            $pct = [int](15 + (($counter / [Math]::Max($totalFound, 1)) * 70))
            Write-Progress -Activity "Auditing SharePoint Security Posture" `
                -Status "Inspecting site $counter of $totalFound : $siteTitle" `
                -PercentComplete $pct

            Write-Log -Level INFO -Message "[$counter/$totalFound] Auditing site collection: $siteUrl"

            # Extract template & friendly name
            $templateCode = [string](Get-SafePropertyValue -InputObject $site -PropertyName "Template" -DefaultValue "Unknown")
            $friendlyTemplate = Get-FriendlyTemplateName -TemplateCode $templateCode

            # Resolve Site Owner & Admin Count
            $ownerVal = Get-SafePropertyValue -InputObject $site -PropertyName "Owner" -DefaultValue ""
            if ([string]::IsNullOrWhiteSpace($ownerVal)) {
                $ownerVal = Get-SafePropertyValue -InputObject $site -PropertyName "OwnerEmail" -DefaultValue ""
            }
            if ([string]::IsNullOrWhiteSpace($ownerVal)) {
                $ownerVal = Get-SafePropertyValue -InputObject $site -PropertyName "OwnerLoginName" -DefaultValue ""
            }
            if ([string]::IsNullOrWhiteSpace($ownerVal)) {
                $ownerVal = Get-SafePropertyValue -InputObject $site -PropertyName "OwnerName" -DefaultValue ""
            }
            if ([string]::IsNullOrWhiteSpace($ownerVal)) {
                $ownerVal = Get-SafePropertyValue -InputObject $site -PropertyName "SecondaryContact" -DefaultValue ""
            }

            # Extract sensitivity label
            $sensitivity = "None"
            $labelRaw = Get-SafePropertyValue -InputObject $site -PropertyName "SensitivityLabel" -DefaultValue $null
            if ($labelRaw) {
                $sensitivity = $labelRaw.ToString()
            }
            else {
                $labelInfo = Get-SafePropertyValue -InputObject $site -PropertyName "SensitivityLabelInfo" -DefaultValue $null
                if ($labelInfo -and $labelInfo.PSObject -and $labelInfo.PSObject.Properties['DisplayName']) {
                    $sensitivity = $labelInfo.DisplayName
                }
            }

            # Extract M365 Group Privacy & Association
            $groupPrivacy = "N/A"
            $groupIdVal = Get-SafePropertyValue -InputObject $site -PropertyName "GroupId" -DefaultValue $null
            $groupId = if ($groupIdVal) { $groupIdVal.ToString() } else { $null }
            $isGroupConnected = (-not [string]::IsNullOrWhiteSpace($groupId) -and $groupId -ne [Guid]::Empty.ToString() -and $groupId -ne "00000000-0000-0000-0000-000000000000")
            if ($isGroupConnected) {
                $channelType = Get-SafePropertyValue -InputObject $site -PropertyName "ChannelType" -DefaultValue $null
                if ($channelType) {
                    $groupPrivacy = [string]$channelType
                }
                else {
                    $groupPrivacy = "Group-Connected"
                }
                if ([string]::IsNullOrWhiteSpace($ownerVal)) {
                    $ownerVal = "M365 Group Owners"
                }
            }

            # Detect system / tenant infrastructure sites
            $isSystemSite = ($templateCode -in @('SRCHCEN#0', 'APPCATALOG#0', 'POINTPUBLISHINGHUB#0', 'COMMUNITYPORTAL#0', 'BLANKINTERNETCONTAINER#0', 'EDISC#0') -or $siteUrl -like "*/search" -or $siteUrl -like "*/portals/*" -or $siteUrl -like "*/sites/AppCatalog*")
            if ($isSystemSite -and [string]::IsNullOrWhiteSpace($ownerVal)) {
                $ownerVal = "SharePoint System"
            }

            $ownerUpn = if (-not [string]::IsNullOrWhiteSpace($ownerVal)) { [string]$ownerVal } else { "N/A" }
            $adminsCount = 0
            $adminList = [System.Collections.Generic.List[string]]::new()
            $hasUniquePerms = $false

            # Evaluate Site Collection Administrators and unique inheritance
            if ($DetailedScan) {
                try {
                    $siteConn = Connect-PnPOnline -Url $siteUrl -Interactive:$Interactive `
                        -ClientId $Script:ActiveClientId -Tenant $Script:ActiveTenant -Thumbprint $CertificateThumbprint `
                        -ReturnConnection -ErrorAction SilentlyContinue

                    if ($siteConn) {
                        $siteAdmins = Get-PnPSiteCollectionAdmin -Connection $siteConn -ErrorAction SilentlyContinue
                        if ($siteAdmins) {
                            $adminsCount = @($siteAdmins).Count
                            foreach ($adm in $siteAdmins) {
                                $admEmail = [string](Get-SafePropertyValue -InputObject $adm -PropertyName "Email" -DefaultValue "")
                                $admLogin = [string](Get-SafePropertyValue -InputObject $adm -PropertyName "LoginName" -DefaultValue "")
                                $admTitle = [string](Get-SafePropertyValue -InputObject $adm -PropertyName "Title" -DefaultValue "")
                                $admName  = if (-not [string]::IsNullOrWhiteSpace($admEmail)) { $admEmail } elseif (-not [string]::IsNullOrWhiteSpace($admLogin)) { $admLogin } else { $admTitle }
                                if (-not [string]::IsNullOrWhiteSpace($admName)) {
                                    $adminList.Add($admName)
                                }
                            }
                        }

                        # Check broken inheritance on root web or primary document library
                        $rootWeb = Get-PnPWeb -Connection $siteConn -ErrorAction SilentlyContinue
                        $webInherit = Get-SafePropertyValue -InputObject $rootWeb -PropertyName "HasUniqueRoleAssignments" -DefaultValue $false
                        if ($webInherit) {
                            $hasUniquePerms = $true
                        }

                        $docLib = Get-PnPList -Identity "Documents" -Connection $siteConn -ErrorAction SilentlyContinue
                        $docInherit = Get-SafePropertyValue -InputObject $docLib -PropertyName "HasUniqueRoleAssignments" -DefaultValue $false
                        if ($docInherit) {
                            $hasUniquePerms = $true
                        }

                        Disconnect-PnPOnline -Connection $siteConn -ErrorAction SilentlyContinue
                    }
                }
                catch {
                    Write-Log -Level WARN -Message "Could not inspect site-level admins for '$siteUrl': $($_.Exception.Message)"
                }
            }

            # Fallback if admin inspection was not performed or returned 0
            if ($adminsCount -eq 0 -and -not [string]::IsNullOrWhiteSpace($ownerUpn) -and $ownerUpn -notin @("N/A", "Unknown")) {
                $adminsCount = 1
                $adminList.Add($ownerUpn)
            }

            # Storage metrics
            $storageUsedRaw  = Get-SafePropertyValue -InputObject $site -PropertyName "StorageUsageCurrent" -DefaultValue 0
            $storageLimitRaw = Get-SafePropertyValue -InputObject $site -PropertyName "StorageMaximumLevel" -DefaultValue 0
            $storageUsedMB   = if ($null -ne $storageUsedRaw) { [int]$storageUsedRaw } else { 0 }
            $storageLimitMB  = if ($null -ne $storageLimitRaw) { [int]$storageLimitRaw } else { 0 }

            # Anonymous link expiration
            $anonExpRaw = Get-SafePropertyValue -InputObject $site -PropertyName "AnonymousLinkExpirationInDays" -DefaultValue 0
            $anonExpirationDays = if ($null -ne $anonExpRaw) { [int]$anonExpRaw } else { 0 }

            # Sharing capability
            $sharingRaw = Get-SafePropertyValue -InputObject $site -PropertyName "SharingCapability" -DefaultValue "Unknown"
            $sharingCapability = if ($sharingRaw) { $sharingRaw.ToString() } else { "Unknown" }

            # LockState
            $lockStateRaw = Get-SafePropertyValue -InputObject $site -PropertyName "LockState" -DefaultValue "Unlock"
            $lockState = if ($lockStateRaw) { $lockStateRaw.ToString() } else { "Unlock" }

            # Default sharing link
            $defLinkTypeRaw = Get-SafePropertyValue -InputObject $site -PropertyName "DefaultSharingLinkType" -DefaultValue "N/A"
            $defaultSharingLinkType = if ($defLinkTypeRaw) { $defLinkTypeRaw.ToString() } else { "N/A" }

            $defLinkPermRaw = Get-SafePropertyValue -InputObject $site -PropertyName "DefaultLinkPermission" -DefaultValue "N/A"
            $defaultLinkPermission = if ($defLinkPermRaw) { $defLinkPermRaw.ToString() } else { "N/A" }

            # Build intermediate record for risk evaluation
            $siteEvalData = [PSCustomObject]@{
                Url                           = $siteUrl
                Title                         = $siteTitle
                Template                      = $templateCode
                Owner                         = $ownerUpn
                SiteAdminsCount               = $adminsCount
                SharingCapability             = $sharingCapability
                AnonymousLinkExpirationInDays = $anonExpirationDays
                GroupPrivacy                  = $groupPrivacy
                SensitivityLabel              = $sensitivity
                HasBrokenInheritance          = $hasUniquePerms
                StorageUsageMB                = $storageUsedMB
                StorageLimitMB                = $storageLimitMB
            }

            # Evaluate Security Risk
            $riskAssessment = Measure-SiteSecurityRisk -SiteData $siteEvalData

            # Compile Full Audit Object
            $record = [PSCustomObject]@{
                ScanTimestamp                 = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                SiteTitle                     = $siteTitle
                SiteUrl                       = $siteUrl
                Template                      = $templateCode
                TemplateFriendlyName          = $friendlyTemplate
                Owner                         = $ownerUpn
                SiteAdminsCount               = $adminsCount
                SiteAdminsList                = ($adminList -join "; ")
                SharingCapability             = $sharingCapability
                AnonymousLinksPermitted       = ($sharingCapability -eq 'ExternalUserAndGuestSharing')
                AnonymousLinkExpirationInDays = $anonExpirationDays
                DefaultSharingLinkType        = $defaultSharingLinkType
                DefaultLinkPermission         = $defaultLinkPermission
                M365GroupId                   = $groupId
                GroupPrivacy                  = $groupPrivacy
                SensitivityLabel              = $sensitivity
                HasBrokenInheritance          = $hasUniquePerms
                StorageUsageMB                = $storageUsedMB
                StorageLimitMB                = $storageLimitMB
                LockState                     = $lockState
                RiskLevel                     = $riskAssessment.RiskLevel
                RiskScore                     = $riskAssessment.RiskScore
                RiskFactors                   = $riskAssessment.RiskFactors
                Remediations                  = $riskAssessment.Remediations
            }

            $auditRecords.Add($record)
        }

        Write-Progress -Activity "Auditing SharePoint Security Posture" -Completed
        Write-Log -Level SUCCESS -Message "Completed evaluation of $($auditRecords.Count) site collection(s)."

        # 8. Render Console Dashboard
        Show-ConsoleAuditDashboard -Records $auditRecords

        # 9. Export to CSV
        $timestamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
        $csvFileName = "Audit-SharePointSecuritySummary-${timestamp}.csv"
        $csvPath = Join-Path -Path $OutputPath -ChildPath $csvFileName
        Export-AuditCsv -Records $auditRecords -FilePath $csvPath

        # 10. Export HTML Executive Summary
        if ($GenerateHtmlReport) {
            $htmlFileName = "Audit-SharePointSecuritySummary-${timestamp}.html"
            $htmlPath = Join-Path -Path $OutputPath -ChildPath $htmlFileName
            Export-AuditHtmlReport -Records $auditRecords -FilePath $htmlPath -TenantAdminEndpoint $resolvedAdminUrl
        }

        # 11. Pipeline Emission
        if ($PassThru) {
            $auditRecords
        }
    }
    catch {
        Write-Log -Level ERROR -Message "Critical error during SharePoint security audit: $($_.Exception.Message)"
        throw
    }
    finally {
        # Session Teardown
        if ($Script:ActiveConnection) {
            Write-Log -Level INFO -Message "Terminating active SharePoint Online connection..."
            try {
                Disconnect-PnPOnline -Connection $Script:ActiveConnection -ErrorAction SilentlyContinue
            }
            catch {
                Write-Verbose -Message "Connection teardown notice: $($_.Exception.Message)"
            }
        }
    }
}

end {
    $duration = (Get-Date) - $Script:StartTime
    Write-Log -Level SUCCESS -Message "Audit execution finished in $($duration.Minutes)m $($duration.Seconds)s."
    if ($Script:LogFilePath) {
        Write-Log -Level INFO -Message "Audit execution log saved at: $Script:LogFilePath"
    }
}

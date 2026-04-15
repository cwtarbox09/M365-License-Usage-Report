[CmdletBinding()]
param(
    [ValidateSet('D7','D30','D90','D180')]
    [string]$Period = 'D30',

    [ValidateSet('Html','Csv','Both')]
    [string]$OutputFormat = 'Html',

    [string]$OutputFolder = (Join-Path -Path $PSScriptRoot -ChildPath 'output'),

    [switch]$SkipModuleInstall
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Mapping of SKU part numbers to human-readable Microsoft 365 product display names.
# Falls back to the raw SkuPartNumber for any unlisted SKU.
$script:LicenseDisplayNames = @{
    # Microsoft 365 Enterprise
    'SPE_E3'                                    = 'Microsoft 365 E3'
    'SPE_E5'                                    = 'Microsoft 365 E5'
    'SPE_E3_RPA1'                               = 'Microsoft 365 E3 (with Power Automate)'
    'SPE_E5_CALLINGMINUTES'                     = 'Microsoft 365 E5 (with Calling Minutes)'

    # Microsoft 365 Business
    'SPB'                                       = 'Microsoft 365 Business Premium'
    'SMB_BUSINESS_PREMIUM'                      = 'Microsoft 365 Business Premium'
    'O365_BUSINESS_PREMIUM'                     = 'Microsoft 365 Business Premium'
    'O365_BUSINESS_ESSENTIALS'                  = 'Microsoft 365 Business Basic'
    'SMB_BUSINESS_ESSENTIALS'                   = 'Microsoft 365 Business Basic'
    'O365_BUSINESS'                             = 'Microsoft 365 Apps for Business'
    'SMB_BUSINESS'                              = 'Microsoft 365 Apps for Business'

    # Microsoft 365 Frontline Workers
    'SPE_F1'                                    = 'Microsoft 365 F1'
    'M365_F1'                                   = 'Microsoft 365 F1'
    'M365_F3'                                   = 'Microsoft 365 F3'

    # Microsoft 365 Education
    'M365EDU_A1'                                = 'Microsoft 365 A1'
    'M365EDU_A3_FACULTY'                        = 'Microsoft 365 A3 for Faculty'
    'M365EDU_A3_STUDENT'                        = 'Microsoft 365 A3 for Students'
    'M365EDU_A5_FACULTY'                        = 'Microsoft 365 A5 for Faculty'
    'M365EDU_A5_STUDENT'                        = 'Microsoft 365 A5 for Students'

    # Office 365
    'STANDARDPACK'                              = 'Office 365 E1'
    'STANDARDWOFFPACK'                          = 'Office 365 E2'
    'ENTERPRISEPACK'                            = 'Office 365 E3'
    'ENTERPRISEWITHSCAL'                        = 'Office 365 E4'
    'ENTERPRISEPREMIUM'                         = 'Office 365 E5'
    'ENTERPRISEPREMIUM_NOPSTNCONF'              = 'Office 365 E5 (without Audio Conferencing)'
    'DESKLESSPACK'                              = 'Office 365 F3'
    'DESKLESSWOFFPACK'                          = 'Office 365 F1'
    'ENTERPRISEPACKWITHOUTPROPLUS'              = 'Office 365 E3 (No Apps)'

    # Office 365 Education
    'STANDARDWOFFPACK_FACULTY'                  = 'Office 365 Education A1 for Faculty'
    'STANDARDWOFFPACK_STUDENT'                  = 'Office 365 Education A1 for Students'
    'STANDARDPACK_FACULTY'                      = 'Office 365 Education A1 for Faculty'
    'STANDARDPACK_STUDENT'                      = 'Office 365 Education A1 for Students'
    'O365_EDUCATION_E1'                         = 'Office 365 Education E1'
    'O365_EDUCATION_E3_FACULTY'                 = 'Office 365 Education E3 for Faculty'

    # Microsoft 365 Apps
    'OFFICESUBSCRIPTION'                        = 'Microsoft 365 Apps for Enterprise'
    'OFFICESUBSCRIPTION_FACULTY'                = 'Microsoft 365 Apps for Faculty'
    'OFFICESUBSCRIPTION_STUDENT'                = 'Microsoft 365 Apps for Students'

    # Exchange Online
    'EXCHANGESTANDARD'                          = 'Exchange Online Plan 1'
    'EXCHANGEENTERPRISE'                        = 'Exchange Online Plan 2'
    'EXCHANGE_S_DESKLESS'                       = 'Exchange Online Kiosk'
    'EXCHANGEDESKLESS'                          = 'Exchange Online Kiosk'
    'EXCHANGEARCHIVE_ADDON'                     = 'Exchange Online Archiving'
    'EXCHANGEESSENTIALS'                        = 'Exchange Online Essentials'
    'EXCHANGE_B_STANDARD'                       = 'Exchange Online Plan 1'

    # SharePoint Online
    'SHAREPOINTSTANDARD'                        = 'SharePoint Online Plan 1'
    'SHAREPOINTENTERPRISE'                      = 'SharePoint Online Plan 2'

    # OneDrive for Business
    'ONEDRIVE_BASIC'                            = 'OneDrive for Business Plan 1'
    'WACONEDRIVESTANDARD'                       = 'OneDrive for Business with Office Online'

    # Microsoft Teams / Skype
    'MCOSTANDARD'                               = 'Skype for Business Online Plan 2'
    'MCOMEETADV'                                = 'Microsoft 365 Audio Conferencing'
    'TEAMS_EXPLORATORY'                         = 'Microsoft Teams Exploratory'
    'TEAMS_FREE'                                = 'Microsoft Teams (Free)'
    'Teams_Room_Standard'                       = 'Microsoft Teams Rooms Standard'
    'Teams_Room_Pro'                            = 'Microsoft Teams Rooms Pro'
    'MCOEV'                                     = 'Microsoft 365 Phone System'
    'MCOEV_DOD'                                 = 'Microsoft 365 Phone System for DoD'
    'MCOPSTN1'                                  = 'Microsoft 365 Domestic Calling Plan'
    'MCOPSTN2'                                  = 'Microsoft 365 International Calling Plan'
    'MCOPSTNCAP'                                = 'Microsoft 365 Communications Credits'
    'PHONESYSTEM_VIRTUALUSER'                   = 'Microsoft Teams Phone Resource Account'

    # Microsoft Defender / Security
    'ATP_ENTERPRISE'                            = 'Microsoft Defender for Office 365 Plan 1'
    'THREAT_INTELLIGENCE'                       = 'Microsoft Defender for Office 365 Plan 2'
    'IDENTITY_THREAT_PROTECTION'                = 'Microsoft 365 E5 Security'
    'IDENTITY_THREAT_PROTECTION_FOR_EMS_E5'     = 'Microsoft 365 E5 Security for EMS E5'
    'INFORMATION_PROTECTION_COMPLIANCE'         = 'Microsoft 365 E5 Compliance'
    'ATA'                                       = 'Microsoft Defender for Identity'
    'MDATP'                                     = 'Microsoft Defender for Endpoint P2'
    'WIN_DEF_ATP'                               = 'Microsoft Defender for Endpoint P1'

    # Enterprise Mobility + Security
    'EMS'                                       = 'Enterprise Mobility + Security E3'
    'EMSPREMIUM'                                = 'Enterprise Mobility + Security E5'

    # Microsoft Intune
    'INTUNE_A'                                  = 'Microsoft Intune Plan 1'
    'INTUNE_A_D'                                = 'Microsoft Intune Plan 1 for Education'
    'INTUNE_SMB'                                = 'Microsoft Intune SMB'

    # Microsoft Entra ID (formerly Azure AD)
    'AAD_PREMIUM'                               = 'Microsoft Entra ID P1'
    'AAD_PREMIUM_P2'                            = 'Microsoft Entra ID P2'
    'AAD_SMB'                                   = 'Microsoft Entra ID P1 (SMB)'

    # Power Platform
    'POWER_BI_STANDARD'                         = 'Power BI (Free)'
    'POWER_BI_PRO'                              = 'Power BI Pro'
    'POWER_BI_PREMIUM_PER_USER'                 = 'Power BI Premium Per User'
    'POWERAPPS_INDIVIDUAL_USER'                 = 'Microsoft Power Apps Plan 2 Trial'
    'POWERAPPS_VIRAL'                           = 'Microsoft Power Apps Plan 2 Trial'
    'FLOW_FREE'                                 = 'Power Automate Free'
    'POWERFLOW_P1'                              = 'Power Automate Premium'
    'POWERFLOW_P2'                              = 'Power Automate Process'
    'POWER_VIRTUAL_AGENTS_VIRAL_TRIAL'          = 'Power Virtual Agents Viral Trial'

    # Microsoft Project
    'PROJECT_P1'                                = 'Project Plan 1'
    'PROJECT_P2'                                = 'Project Plan 3'
    'PROJECT_P3'                                = 'Project Plan 5'
    'PROJECTESSENTIALS'                         = 'Project Plan 1'
    'PROJECTPROFESSIONAL'                       = 'Project Plan 3'
    'PROJECTPREMIUM'                            = 'Project Plan 5'

    # Microsoft Visio
    'VISIO_PLAN1_DEP'                           = 'Visio Plan 1'
    'VISIO_PLAN2_DEP'                           = 'Visio Plan 2'
    'VISIOCLIENT'                               = 'Visio Plan 2'

    # Microsoft 365 Copilot
    'Copilot_Pro'                               = 'Microsoft Copilot Pro'
    'Microsoft_365_Copilot'                     = 'Microsoft 365 Copilot'

    # Developer
    'DEVELOPERPACK'                             = 'Microsoft 365 E3 Developer'
    'DEVELOPERPACK_E5'                          = 'Microsoft 365 E5 Developer'

    # Yammer
    'YAMMER_ENTERPRISE'                         = 'Yammer Enterprise'

    # Microsoft Stream
    'STREAM'                                    = 'Microsoft Stream (Classic)'

    # Windows
    'WIN10_PRO_ENT_SUB'                         = 'Windows 10/11 Enterprise E3'
    'WIN10_ENT_A3_FAC'                          = 'Windows 10/11 Enterprise A3 for Faculty'
    'WIN10_ENT_A3_STU'                          = 'Windows 10/11 Enterprise A3 for Students'

    # Dynamics 365
    'DYN365_ENTERPRISE_PLAN1'                   = 'Dynamics 365 Customer Engagement Plan'
    'DYN365_ENTERPRISE_SALES'                   = 'Dynamics 365 Sales Enterprise'
    'DYN365_ENTERPRISE_CUSTOMER_SERVICE'        = 'Dynamics 365 Customer Service Enterprise'
    'DYN365_FINANCIALS_BUSINESS_SKU'            = 'Dynamics 365 Business Central Essential'

    # Microsoft Viva
    'VIVA'                                      = 'Microsoft Viva Suite'
    'VIVA_INSIGHTS_P1'                          = 'Microsoft Viva Insights'
}

# SKU part numbers for licenses that include Microsoft Intune.
# Only these SKUs trigger Intune signal detection.
$script:IntuneBearingSkus = [System.Collections.Generic.HashSet[string]]@(
    # Microsoft 365 Business Premium
    'SPB', 'SMB_BUSINESS_PREMIUM', 'O365_BUSINESS_PREMIUM',
    # Microsoft 365 E3 / E5
    'SPE_E3', 'SPE_E3_RPA1', 'SPE_E5', 'SPE_E5_CALLINGMINUTES',
    # Enterprise Mobility + Security E3 / E5
    'EMS', 'EMSPREMIUM',
    # Microsoft Intune standalone
    'INTUNE_A', 'INTUNE_A_D', 'INTUNE_SMB'
)

function Get-LicenseFriendlyName {
    param([Parameter(Mandatory)][string]$SkuPartNumber)

    if ($script:LicenseDisplayNames.ContainsKey($SkuPartNumber)) {
        return $script:LicenseDisplayNames[$SkuPartNumber]
    }

    return $SkuPartNumber
}

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "[$timestamp] [$Level] $Message"
}

function Ensure-Modules {
    param([switch]$SkipInstall)

    $requiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'Microsoft.Graph.DeviceManagement',
        'Microsoft.Graph.Reports'
    )

    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            if ($SkipInstall) {
                throw "Required module '$module' is not installed and -SkipModuleInstall was specified."
            }

            Write-Log "Installing module: $module"
            Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }

        Import-Module $module -ErrorAction Stop
    }
}

function Connect-M365Graph {
    $scopes = @(
        'User.Read.All',
        'Directory.Read.All',
        'DeviceManagementManagedDevices.Read.All',
        'Reports.Read.All',
        'AuditLog.Read.All'
    )

    Write-Log 'Connecting to Microsoft Graph. Please sign in with an admin account when prompted.'
    Connect-MgGraph -Scopes $scopes -NoWelcome

    $ctx = Get-MgContext
    if (-not $ctx) {
        throw 'Unable to establish Microsoft Graph context.'
    }

    Write-Log "Connected to tenant $($ctx.TenantId) as $($ctx.Account)"
}

function Get-GraphReportData {
    param(
        [Parameter(Mandatory)][string]$Endpoint,
        [Parameter(Mandatory)][string]$TempFile
    )

    Invoke-MgGraphRequest -Method GET -Uri $Endpoint -OutputFilePath $TempFile

    if (-not (Test-Path -Path $TempFile)) {
        throw "Report download failed for endpoint: $Endpoint"
    }

    $raw = Get-Content -Path $TempFile -Raw
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return @()
    }

    return $raw | ConvertFrom-Csv
}

function Build-LicenseLookup {
    $skuLookup = @{}
    $skus = Get-MgSubscribedSku -All

    foreach ($sku in $skus) {
        $skuLookup[$sku.SkuId.ToString()] = [PSCustomObject]@{
            SkuPartNumber = $sku.SkuPartNumber
            DisplayName   = Get-LicenseFriendlyName -SkuPartNumber $sku.SkuPartNumber
            ServicePlans  = $sku.ServicePlans
        }
    }

    return $skuLookup
}

function Add-ReportKeyedHashtable {
    param(
        [Parameter(Mandatory)]$Rows,
        [Parameter(Mandatory)][string]$KeyProperty
    )

    $result = @{}
    foreach ($row in $Rows) {
        $key = Get-ReportValue -Row $row -CandidateProperties @($KeyProperty)
        if (-not [string]::IsNullOrWhiteSpace($key)) {
            $result[$key.ToLowerInvariant()] = $row
        }
    }

    return $result
}

function Test-RecentActivity {
    param(
        [AllowNull()]$DateValue,
        [int]$LookbackDays = 30
    )

    if (-not $DateValue) {
        return $false
    }

    $parsedDate = $DateValue -as [datetime]
    if ($null -eq $parsedDate) {
        return $false
    }

    return $parsedDate -ge (Get-Date).AddDays(-1 * $LookbackDays)
}

function Get-ReportValue {
    param(
        [AllowNull()]$Row,
        [Parameter(Mandatory)][string[]]$CandidateProperties
    )

    if (-not $Row) {
        return $null
    }

    $allProperties = @($Row.PSObject.Properties)

    foreach ($name in $CandidateProperties) {
        $property = $allProperties | Where-Object { $_.Name -eq $name } | Select-Object -First 1
        if ($property) {
            return $property.Value
        }

        $normalizedTarget = ($name -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
        $fallback = $allProperties |
            Where-Object { (($_.Name -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()) -eq $normalizedTarget } |
            Select-Object -First 1

        if ($fallback) {
            return $fallback.Value
        }
    }

    return $null
}

function Get-UserIntuneDeviceCounts {
    Write-Log 'Collecting Intune managed device inventory...'

    $devices = Get-MgDeviceManagementManagedDevice -All -Property userPrincipalName,deviceName,operatingSystem,lastSyncDateTime
    $counts = @{}

    foreach ($device in $devices) {
        if ([string]::IsNullOrWhiteSpace($device.UserPrincipalName)) {
            continue
        }

        $upn = $device.UserPrincipalName.ToLowerInvariant()
        if (-not $counts.ContainsKey($upn)) {
            $counts[$upn] = 0
        }

        $counts[$upn]++
    }

    return $counts
}

function New-LicenseUtilizationRows {
    param(
        [Parameter(Mandatory)]$Users,
        [Parameter(Mandatory)]$SkuLookup,
        [Parameter(Mandatory)]$MailboxUsage,
        [Parameter(Mandatory)]$OneDriveUsage,
        [Parameter(Mandatory)]$TeamsUsage,
        [Parameter(Mandatory)]$IntuneDeviceCounts,
        [string]$Period
    )

    $outputRows = New-Object System.Collections.Generic.List[object]

    foreach ($user in $Users) {
        $upn = $user.UserPrincipalName
        $upnLower = $upn.ToLowerInvariant()

        $mailboxRow = $MailboxUsage[$upnLower]
        $oneDriveRow = $OneDriveUsage[$upnLower]
        $teamsRow = $TeamsUsage[$upnLower]
        $intuneDevices = if ($IntuneDeviceCounts.ContainsKey($upnLower)) { $IntuneDeviceCounts[$upnLower] } else { 0 }

        $assignedSkuNames = New-Object System.Collections.Generic.List[string]
        $assignedSkuPartNumbers = New-Object System.Collections.Generic.List[string]
        $allServicePlanNames = New-Object System.Collections.Generic.List[string]

        foreach ($license in $user.AssignedLicenses) {
            $skuId = $license.SkuId.ToString()
            if (-not $SkuLookup.ContainsKey($skuId)) {
                continue
            }

            $sku = $SkuLookup[$skuId]
            $assignedSkuNames.Add($sku.DisplayName)
            $assignedSkuPartNumbers.Add($sku.SkuPartNumber)
            foreach ($servicePlan in $sku.ServicePlans) {
                $allServicePlanNames.Add($servicePlan.ServicePlanName)
            }
        }

        $hasExchangePlan = $allServicePlanNames -match 'EXCHANGE_S_STANDARD|EXCHANGE_S_ENTERPRISE|EXCHANGE_S_DESKLESS|EXCHANGE_S_ESSENTIALS'
        $hasSharePointPlan = $allServicePlanNames -match 'SHAREPOINT|ONEDRIVE'
        $hasTeamsPlan = $allServicePlanNames -match 'TEAMS'
        $hasIntunePlan = $assignedSkuPartNumbers | Where-Object { $script:IntuneBearingSkus.Contains($_) }

        $signals = New-Object System.Collections.Generic.List[string]
        $evidence = New-Object System.Collections.Generic.List[string]

        if ($hasExchangePlan) {
            $signals.Add('Exchange')
            $mailboxLastActivity = Get-ReportValue -Row $mailboxRow -CandidateProperties @('Last Activity Date','LastActivityDate')
            $mailboxActive = Test-RecentActivity -DateValue $mailboxLastActivity -LookbackDays 30
            $evidence.Add("ExchangeLastActivity=$mailboxLastActivity")
        }
        else {
            $mailboxActive = $false
        }

        if ($hasSharePointPlan) {
            $signals.Add('OneDrive/SharePoint')
            $oneDriveLastActivity = Get-ReportValue -Row $oneDriveRow -CandidateProperties @('Last Activity Date','LastActivityDate')
            $oneDriveActive = Test-RecentActivity -DateValue $oneDriveLastActivity -LookbackDays 30
            $evidence.Add("OneDriveLastActivity=$oneDriveLastActivity")
        }
        else {
            $oneDriveActive = $false
        }

        if ($hasTeamsPlan) {
            $signals.Add('Teams')
            $teamsLastActivity = Get-ReportValue -Row $teamsRow -CandidateProperties @('Last Activity Date','LastActivityDate')
            $teamsActive = Test-RecentActivity -DateValue $teamsLastActivity -LookbackDays 30
            $evidence.Add("TeamsLastActivity=$teamsLastActivity")
        }
        else {
            $teamsActive = $false
        }

        if ($hasIntunePlan) {
            $signals.Add('Intune')
            $intuneActive = $intuneDevices -gt 0
            $evidence.Add("IntuneDevices=$intuneDevices")
        }
        else {
            $intuneActive = $false
        }

        $checkedSignals = @($mailboxActive, $oneDriveActive, $teamsActive, $intuneActive)
        $activeSignalCount = @($checkedSignals | Where-Object { $_ }).Count
        $availableSignalCount = $signals.Count
        $distinctSkuNames = @($assignedSkuNames | Sort-Object -Unique)

        $utilizationState = if ($availableSignalCount -eq 0) {
            'NoTrackedWorkload'
        }
        elseif ($activeSignalCount -eq 0) {
            'Unused'
        }
        elseif ($activeSignalCount -lt $availableSignalCount) {
            'PartiallyUsed'
        }
        else {
            'Used'
        }

        $outputRows.Add([PSCustomObject]@{
            DisplayName          = $user.DisplayName
            UserPrincipalName    = $upn
            AccountEnabled       = $user.AccountEnabled
            AssignedSkuCount     = @($distinctSkuNames).Count
            AssignedSkus         = ($distinctSkuNames -join '; ')
            TrackedWorkloads     = ($signals -join '; ')
            WorkloadSignalsFound = $activeSignalCount
            WorkloadSignalsTotal = $availableSignalCount
            UtilizationState     = $utilizationState
            ExchangeActive       = $mailboxActive
            OneDriveActive       = $oneDriveActive
            TeamsActive          = $teamsActive
            IntuneDeviceCount    = $intuneDevices
            EvaluationPeriod     = $Period
            Evidence             = ($evidence -join '; ')
        })
    }

    return $outputRows
}

function Export-ReportFiles {
    param(
        [Parameter(Mandatory)]$Rows,
        [Parameter(Mandatory)][string]$Format,
        [Parameter(Mandatory)][string]$Folder
    )

    if (-not (Test-Path -Path $Folder)) {
        New-Item -Path $Folder -ItemType Directory -Force | Out-Null
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $csvPath = Join-Path -Path $Folder -ChildPath "M365_License_Utilization_$timestamp.csv"
    $htmlPath = Join-Path -Path $Folder -ChildPath "M365_License_Utilization_$timestamp.html"

    if ($Format -in @('Csv','Both')) {
        $Rows | Sort-Object UtilizationState,UserPrincipalName | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Log "CSV report written: $csvPath"
    }

    if ($Format -in @('Html','Both')) {
        $summaryByState = $Rows | Group-Object UtilizationState | Sort-Object Name
        $summaryBySku = $Rows |
            ForEach-Object { ($_.AssignedSkus -split '; ') } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Group-Object |
            Sort-Object Count -Descending

        $rowsForExport = $Rows | Sort-Object UtilizationState,UserPrincipalName
        $rowsJson = $rowsForExport | ConvertTo-Json -Depth 5 -Compress
        $summaryStateRows = ($summaryByState | Select-Object Name,Count | ConvertTo-Html -Fragment)
        $summarySkuRows = ($summaryBySku | Select-Object Name,Count | ConvertTo-Html -Fragment)

        $style = @'
<style>
* { box-sizing: border-box; }
body {
    font-family: "Segoe UI", "Inter", Arial, sans-serif;
    margin: 0;
    background: linear-gradient(180deg, #f5f7fb 0%, #eef3fb 100%);
    color: #162238;
}
.container {
    max-width: 1300px;
    margin: 0 auto;
    padding: 24px;
}
.hero {
    background: linear-gradient(135deg, #0b3d91 0%, #245abf 100%);
    color: #fff;
    border-radius: 16px;
    padding: 24px;
    box-shadow: 0 16px 40px rgba(11, 61, 145, 0.25);
}
.hero h1 { margin: 0 0 8px; font-size: 28px; }
.hero p { margin: 0; opacity: 0.95; }
.stats {
    margin-top: 16px;
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    gap: 12px;
}
.stat-card {
    background: rgba(255,255,255,0.14);
    border: 1px solid rgba(255,255,255,0.25);
    border-radius: 12px;
    padding: 14px;
}
.stat-label { font-size: 12px; opacity: 0.9; text-transform: uppercase; letter-spacing: 0.4px; }
.stat-value { font-size: 28px; font-weight: 700; margin-top: 4px; }
.section-card {
    margin-top: 20px;
    background: #ffffff;
    border-radius: 14px;
    box-shadow: 0 8px 24px rgba(15, 23, 42, 0.08);
    border: 1px solid #e4e9f4;
    padding: 18px;
}
.section-title {
    margin: 0 0 14px;
    color: #0f2f66;
    font-size: 20px;
}
.toolbar {
    display: flex;
    justify-content: flex-end;
    margin-bottom: 10px;
}
.btn {
    border: none;
    border-radius: 10px;
    padding: 10px 14px;
    font-weight: 600;
    cursor: pointer;
    background: #2563eb;
    color: #fff;
}
.btn:hover { background: #1d4ed8; }
.table-wrap {
    overflow: auto;
    border-radius: 12px;
    border: 1px solid #e5eaf5;
}
table {
    width: 100%;
    border-collapse: collapse;
    background: #fff;
}
th, td {
    padding: 10px;
    border-bottom: 1px solid #edf1f9;
    text-align: left;
    vertical-align: top;
    font-size: 13px;
}
th {
    position: sticky;
    top: 0;
    background: #f2f6ff;
    color: #1b376d;
    font-weight: 700;
}
tr:hover td { background: #f9fbff; }
.badge {
    padding: 4px 8px;
    border-radius: 999px;
    font-weight: 700;
    display: inline-block;
    font-size: 12px;
}
.state-used { background: #dcfce7; color: #166534; }
.state-partiallyused { background: #fef3c7; color: #92400e; }
.state-unused { background: #fee2e2; color: #991b1b; }
.state-notrackedworkload { background: #e5e7eb; color: #374151; }
</style>
'@

        $scriptTemplate = @'
<script>
const reportRows = __ROWS_JSON__;
const exportTimestamp = "__EXPORT_TIMESTAMP__";

function getStateClass(state) {
  const normalized = (state || '').toLowerCase();
  return `state-${normalized}`;
}

function renderDetailTable() {
  const columns = [
    'DisplayName','UserPrincipalName','AccountEnabled','AssignedSkuCount','AssignedSkus','TrackedWorkloads',
    'WorkloadSignalsFound','WorkloadSignalsTotal','UtilizationState','ExchangeActive','OneDriveActive',
    'TeamsActive','IntuneDeviceCount','EvaluationPeriod','Evidence'
  ];

  const header = columns.map(c => `<th>${c}</th>`).join('');
  const rows = reportRows.map(row => {
    return `<tr>${columns.map(col => {
      let value = row[col];
      if (col === 'UtilizationState') {
        return `<td><span class="badge ${getStateClass(value)}">${value ?? ''}</span></td>`;
      }
      if (value === null || value === undefined) value = '';
      return `<td>${String(value)}</td>`;
    }).join('')}</tr>`;
  }).join('');

  document.getElementById('detailTable').innerHTML =
    `<thead><tr>${header}</tr></thead><tbody>${rows}</tbody>`;
}

function downloadCsv() {
  if (!reportRows.length) return;
  const columns = Object.keys(reportRows[0]);
  const escapeCsv = (value) => {
    const str = value === null || value === undefined ? '' : String(value);
    return /[",\n]/.test(str) ? '"' + str.replace(/"/g, '""') + '"' : str;
  };
  const csv = [
    columns.join(','),
    ...reportRows.map(r => columns.map(c => escapeCsv(r[c])).join(','))
  ].join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `M365_License_Utilization_${exportTimestamp}.csv`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

renderDetailTable();
</script>
'@
        $script = $scriptTemplate.
            Replace('__ROWS_JSON__', $rowsJson).
            Replace('__EXPORT_TIMESTAMP__', (Get-Date -Format 'yyyyMMdd_HHmmss'))

        $bodySections = @(
            '<div class="container">',
            '<section class="hero">',
            '<h1>M365 License Utilization Report</h1>',
            "<p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>",
            '<div class="stats">',
            '<div class="stat-card"><div class="stat-label">Total Licensed Users</div>',
            "<div class='stat-value'>$($Rows.Count)</div></div>",
            '<div class="stat-card"><div class="stat-label">States Tracked</div>',
            "<div class='stat-value'>$($summaryByState.Count)</div></div>",
            '</div>',
            '</section>',
            '<section class="section-card">',
            '<h2 class="section-title">Summary by Utilization State</h2>',
            ($summaryStateRows -join "`n"),
            '</section>',
            '<section class="section-card">',
            '<h2 class="section-title">Top License SKUs by Assignment Count</h2>',
            ($summarySkuRows -join "`n"),
            '</section>',
            '<section class="section-card">',
            '<div class="toolbar"><button class="btn" onclick="downloadCsv()">Export details to CSV</button></div>',
            '<h2 class="section-title">Detailed Results</h2>',
            '<div class="table-wrap"><table id="detailTable"></table></div>',
            '</section>',
            $script,
            '</div>'
        )
        $fullHtml = ConvertTo-Html -Title 'M365 License Utilization Report' -Head $style -Body ($bodySections -join "`n")

        $fullHtml | Out-File -Path $htmlPath -Encoding UTF8
        Write-Log "HTML report written: $htmlPath"
    }
}

try {
    Ensure-Modules -SkipInstall:$SkipModuleInstall
    Connect-M365Graph

    Write-Log 'Loading user and license inventory...'
    $users = Get-MgUser -All -Property id,displayName,userPrincipalName,assignedLicenses,accountEnabled |
        Where-Object { $_.AssignedLicenses.Count -gt 0 }

    $skuLookup = Build-LicenseLookup

    $tempRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("m365-usage-{0}" -f ([guid]::NewGuid().Guid))
    New-Item -Path $tempRoot -ItemType Directory -Force | Out-Null

    try {
        Write-Log 'Downloading usage reports from Microsoft Graph...'
        $mailboxData = Get-GraphReportData -Endpoint "/v1.0/reports/getMailboxUsageDetail(period='$Period')" -TempFile (Join-Path $tempRoot 'mailbox.csv')
        $oneDriveData = Get-GraphReportData -Endpoint "/v1.0/reports/getOneDriveUsageAccountDetail(period='$Period')" -TempFile (Join-Path $tempRoot 'onedrive.csv')
        $teamsData = Get-GraphReportData -Endpoint "/v1.0/reports/getTeamsUserActivityUserDetail(period='$Period')" -TempFile (Join-Path $tempRoot 'teams.csv')
    }
    finally {
        Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }

    $mailboxLookup = Add-ReportKeyedHashtable -Rows $mailboxData -KeyProperty 'User Principal Name'
    $oneDriveLookup = Add-ReportKeyedHashtable -Rows $oneDriveData -KeyProperty 'Owner Principal Name'
    $teamsLookup = Add-ReportKeyedHashtable -Rows $teamsData -KeyProperty 'User Principal Name'

    $intuneDeviceCounts = Get-UserIntuneDeviceCounts

    Write-Log 'Evaluating license utilization signals...'
    $rows = New-LicenseUtilizationRows -Users $users -SkuLookup $skuLookup -MailboxUsage $mailboxLookup -OneDriveUsage $oneDriveLookup -TeamsUsage $teamsLookup -IntuneDeviceCounts $intuneDeviceCounts -Period $Period

    Export-ReportFiles -Rows $rows -Format $OutputFormat -Folder $OutputFolder

    Write-Log 'Report generation completed successfully.'
}
catch {
    Write-Log -Level 'ERROR' -Message $_.Exception.Message
    throw
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}

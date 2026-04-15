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

    Write-Log 'Connecting to Microsoft Graph. Sign in with a Microsoft 365 admin account when prompted.'
    Connect-MgGraph -Scopes $scopes -NoWelcome

    $context = Get-MgContext
    if (-not $context) {
        throw 'Microsoft Graph connection failed.'
    }

    Write-Log "Connected to tenant: $($context.TenantId) as account: $($context.Account)"
}

function Get-GraphReportData {
    param(
        [Parameter(Mandatory)]
        [string]$Endpoint,
        [Parameter(Mandatory)]
        [string]$TempFile
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
        $key = $row.$KeyProperty
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

    $parsedDate = $null
    if (-not [DateTime]::TryParse($DateValue.ToString(), [ref]$parsedDate)) {
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
    foreach ($name in $CandidateProperties) {
        $property = $Row.PSObject.Properties[$name]
        if ($property) {
            return $property.Value
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

        foreach ($license in $user.AssignedLicenses) {
            $skuId = $license.SkuId.ToString()
            if (-not $SkuLookup.ContainsKey($skuId)) {
                continue
            }

            $sku = $SkuLookup[$skuId]
            $signals = New-Object System.Collections.Generic.List[string]
            $evidence = New-Object System.Collections.Generic.List[string]

            $hasExchangePlan = $sku.ServicePlans.ServicePlanName -match 'EXCHANGE'
            $hasSharePointPlan = $sku.ServicePlans.ServicePlanName -match 'SHAREPOINT|ONEDRIVE'
            $hasTeamsPlan = $sku.ServicePlans.ServicePlanName -match 'TEAMS'
            $hasIntunePlan = $sku.ServicePlans.ServicePlanName -match 'INTUNE|EMS|AAD_PREMIUM'

            if ($hasExchangePlan) {
                $signals.Add('Exchange')
                $mailboxLastActivity = Get-ReportValue -Row $mailboxRow -CandidateProperties @('Last Activity Date','LastActivityDate')
                $mailboxActive = Test-RecentActivity -DateValue $mailboxLastActivity -LookbackDays 30
                $evidence.Add("ExchangeLastActivity=$mailboxLastActivity")
                $mailboxActive = Test-RecentActivity -DateValue $mailboxRow.'Last Activity Date' -LookbackDays 30
                $evidence.Add("ExchangeLastActivity=$($mailboxRow.'Last Activity Date')")
            }
            else {
                $mailboxActive = $false
            }

            if ($hasSharePointPlan) {
                $signals.Add('OneDrive/SharePoint')
                $oneDriveLastActivity = Get-ReportValue -Row $oneDriveRow -CandidateProperties @('Last Activity Date','LastActivityDate')
                $oneDriveActive = Test-RecentActivity -DateValue $oneDriveLastActivity -LookbackDays 30
                $evidence.Add("OneDriveLastActivity=$oneDriveLastActivity")
                $oneDriveActive = Test-RecentActivity -DateValue $oneDriveRow.'Last Activity Date' -LookbackDays 30
                $evidence.Add("OneDriveLastActivity=$($oneDriveRow.'Last Activity Date')")
            }
            else {
                $oneDriveActive = $false
            }

            if ($hasTeamsPlan) {
                $signals.Add('Teams')
                $teamsLastActivity = Get-ReportValue -Row $teamsRow -CandidateProperties @('Last Activity Date','LastActivityDate')
                $teamsActive = Test-RecentActivity -DateValue $teamsLastActivity -LookbackDays 30
                $evidence.Add("TeamsLastActivity=$teamsLastActivity")
                $teamsActive = Test-RecentActivity -DateValue $teamsRow.'Last Activity Date' -LookbackDays 30
                $evidence.Add("TeamsLastActivity=$($teamsRow.'Last Activity Date')")
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
            $activeSignalCount = ($checkedSignals | Where-Object { $_ }).Count
            $availableSignalCount = $signals.Count

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
                LicenseSku           = $sku.SkuPartNumber
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
        $Rows | Sort-Object UtilizationState,LicenseSku,UserPrincipalName | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Log "CSV report written: $csvPath"
    }

    if ($Format -in @('Html','Both')) {
        $summaryByState = $Rows | Group-Object UtilizationState | Sort-Object Name
        $summaryBySku = $Rows | Group-Object LicenseSku | Sort-Object Count -Descending

        $style = @'
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 20px; }
h1, h2 { color: #0b3d91; }
table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
th, td { border: 1px solid #ddd; padding: 8px; }
th { background-color: #f4f6fa; text-align: left; }
tr:nth-child(even) { background-color: #f9fbff; }
</style>
'@

        $summaryHtml = @()
        $summaryHtml += '<h2>Summary by Utilization State</h2>'
        $summaryHtml += ($summaryByState | Select-Object Name,Count | ConvertTo-Html -Fragment)
        $summaryHtml += '<h2>Top License SKUs by Assignment Count</h2>'
        $summaryHtml += ($summaryBySku | Select-Object Name,Count | ConvertTo-Html -Fragment)

        $detailHtml = $Rows |
            Sort-Object UtilizationState,LicenseSku,UserPrincipalName |
            ConvertTo-Html -Fragment

        $fullHtml = ConvertTo-Html -Title 'M365 License Utilization Report' -Head $style -Body @(
            "<h1>M365 License Utilization Report</h1>",
            "<p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>",
            "<p>Total License Assignments Evaluated: $($Rows.Count)</p>",
            ($summaryHtml -join "`n"),
            '<h2>Detailed Results</h2>',
            $detailHtml
        )

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

[CmdletBinding()]
param(
    [ValidateSet('D7','D30','D90','D180')]
    [string]$Period = 'D30',

    [ValidateSet('Html','Csv','Both')]
    [string]$OutputFormat = 'Html',

    [string]$OutputFolder = (Join-Path -Path $PSScriptRoot -ChildPath 'output'),

    [switch]$SkipModuleInstall
)

$ErrorActionPreference = 'Stop'

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
}

function Ensure-Modules {
    param([switch]$SkipInstall)

    $modules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'Microsoft.Graph.DeviceManagement',
        'Microsoft.Graph.Reports'
    )

    foreach ($moduleName in $modules) {
        $installed = Get-Module -ListAvailable -Name $moduleName
        if (-not $installed) {
            if ($SkipInstall) {
                throw "Required module '$moduleName' is missing and -SkipModuleInstall was used."
            }

            Write-Log "Installing module $moduleName"
            Install-Module -Name $moduleName -Scope CurrentUser -AllowClobber -Force
        }

        Import-Module -Name $moduleName -ErrorAction Stop
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
        throw "Failed to download report from endpoint $Endpoint"
    }

    $content = Get-Content -Path $TempFile -Raw
    if ([string]::IsNullOrWhiteSpace($content)) {
        return @()
    }

    return ($content | ConvertFrom-Csv)
}

function Get-ReportValue {
    param(
        [AllowNull()]$Row,
        [Parameter(Mandatory)][string[]]$CandidateProperties
    )

    if (-not $Row) {
        return $null
    }

    $properties = @($Row.PSObject.Properties)
    foreach ($candidate in $CandidateProperties) {
        $exact = $properties | Where-Object { $_.Name -eq $candidate } | Select-Object -First 1
        if ($exact) {
            return $exact.Value
        }

        $target = ($candidate -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
        $normalized = $properties |
            Where-Object { (($_.Name -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()) -eq $target } |
            Select-Object -First 1

        if ($normalized) {
            return $normalized.Value
        }
    }

    return $null
}

function Add-ReportKeyedHashtable {
    param(
        [Parameter(Mandatory)]$Rows,
        [Parameter(Mandatory)][string[]]$KeyProperties
    )

    $map = @{}
    foreach ($row in $Rows) {
        $key = Get-ReportValue -Row $row -CandidateProperties $KeyProperties
        if ([string]::IsNullOrWhiteSpace($key)) {
            continue
        }

        $map[$key.ToLowerInvariant()] = $row
    }

    return $map
}

function Test-RecentActivity {
    param(
        [AllowNull()]$DateValue,
        [Parameter(Mandatory)][int]$LookbackDays
    )

    if (-not $DateValue) {
        return $false
    }

    $parsed = $null
    if (-not [DateTime]::TryParse($DateValue.ToString(), [ref]$parsed)) {
        return $false
    }

    return $parsed -ge (Get-Date).AddDays(-1 * $LookbackDays)
}

function Get-LookbackDaysFromPeriod {
    param([Parameter(Mandatory)][string]$PeriodValue)

    switch ($PeriodValue) {
        'D7' { return 7 }
        'D30' { return 30 }
        'D90' { return 90 }
        'D180' { return 180 }
        default { return 30 }
    }
}

function Build-LicenseLookup {
    $lookup = @{}
    $skus = Get-MgSubscribedSku -All

    foreach ($sku in $skus) {
        $lookup[$sku.SkuId.ToString()] = [PSCustomObject]@{
            SkuPartNumber = $sku.SkuPartNumber
            ServicePlans  = $sku.ServicePlans
        }
    }

    return $lookup
}

function Get-UserIntuneDeviceCounts {
    Write-Log 'Collecting Intune managed device counts...'
    $counts = @{}

    $devices = Get-MgDeviceManagementManagedDevice -All -Property userPrincipalName
    foreach ($device in $devices) {
        $upn = $device.UserPrincipalName
        if ([string]::IsNullOrWhiteSpace($upn)) {
            continue
        }

        $k = $upn.ToLowerInvariant()
        if (-not $counts.ContainsKey($k)) {
            $counts[$k] = 0
        }

        $counts[$k]++
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
        [Parameter(Mandatory)][int]$LookbackDays,
        [Parameter(Mandatory)][string]$Period
    )

    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($user in $Users) {
        if ([string]::IsNullOrWhiteSpace($user.UserPrincipalName)) {
            continue
        }

        $upn = $user.UserPrincipalName
        $key = $upn.ToLowerInvariant()

        $mailboxRow = if ($MailboxUsage.ContainsKey($key)) { $MailboxUsage[$key] } else { $null }
        $oneDriveRow = if ($OneDriveUsage.ContainsKey($key)) { $OneDriveUsage[$key] } else { $null }
        $teamsRow = if ($TeamsUsage.ContainsKey($key)) { $TeamsUsage[$key] } else { $null }
        $intuneCount = if ($IntuneDeviceCounts.ContainsKey($key)) { $IntuneDeviceCounts[$key] } else { 0 }

        foreach ($assigned in $user.AssignedLicenses) {
            $skuId = $assigned.SkuId.ToString()
            if (-not $SkuLookup.ContainsKey($skuId)) {
                continue
            }

            $sku = $SkuLookup[$skuId]
            $planNames = @($sku.ServicePlans | ForEach-Object { $_.ServicePlanName })

            $tracksExchange = ($planNames -match 'EXCHANGE').Count -gt 0
            $tracksSharePoint = ($planNames -match 'SHAREPOINT|ONEDRIVE').Count -gt 0
            $tracksTeams = ($planNames -match 'TEAMS').Count -gt 0
            $tracksIntune = ($planNames -match 'INTUNE|EMS|AAD_PREMIUM').Count -gt 0

            $mailboxLast = Get-ReportValue -Row $mailboxRow -CandidateProperties @('Last Activity Date','LastActivityDate')
            $oneDriveLast = Get-ReportValue -Row $oneDriveRow -CandidateProperties @('Last Activity Date','LastActivityDate')
            $teamsLast = Get-ReportValue -Row $teamsRow -CandidateProperties @('Last Activity Date','LastActivityDate')

            $exchangeActive = if ($tracksExchange) { Test-RecentActivity -DateValue $mailboxLast -LookbackDays $LookbackDays } else { $false }
            $oneDriveActive = if ($tracksSharePoint) { Test-RecentActivity -DateValue $oneDriveLast -LookbackDays $LookbackDays } else { $false }
            $teamsActive = if ($tracksTeams) { Test-RecentActivity -DateValue $teamsLast -LookbackDays $LookbackDays } else { $false }
            $intuneActive = if ($tracksIntune) { $intuneCount -gt 0 } else { $false }

            $workloads = New-Object System.Collections.Generic.List[string]
            if ($tracksExchange) { $workloads.Add('Exchange') }
            if ($tracksSharePoint) { $workloads.Add('OneDrive/SharePoint') }
            if ($tracksTeams) { $workloads.Add('Teams') }
            if ($tracksIntune) { $workloads.Add('Intune') }

            $signalCount = $workloads.Count
            $activeCount = @($exchangeActive, $oneDriveActive, $teamsActive, $intuneActive | Where-Object { $_ }).Count

            $state = if ($signalCount -eq 0) {
                'NoTrackedWorkload'
            }
            elseif ($activeCount -eq 0) {
                'Unused'
            }
            elseif ($activeCount -lt $signalCount) {
                'PartiallyUsed'
            }
            else {
                'Used'
            }

            $rows.Add([PSCustomObject]@{
                DisplayName          = $user.DisplayName
                UserPrincipalName    = $upn
                AccountEnabled       = $user.AccountEnabled
                LicenseSku           = $sku.SkuPartNumber
                TrackedWorkloads     = ($workloads -join '; ')
                WorkloadSignalsFound = $activeCount
                WorkloadSignalsTotal = $signalCount
                UtilizationState     = $state
                ExchangeActive       = $exchangeActive
                OneDriveActive       = $oneDriveActive
                TeamsActive          = $teamsActive
                IntuneDeviceCount    = $intuneCount
                EvaluationPeriod     = $Period
                Evidence             = "ExchangeLastActivity=$mailboxLast; OneDriveLastActivity=$oneDriveLast; TeamsLastActivity=$teamsLast; IntuneDevices=$intuneCount"
            })
        }
    }

    return $rows
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

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $csvPath = Join-Path -Path $Folder -ChildPath "M365_License_Utilization_$stamp.csv"
    $htmlPath = Join-Path -Path $Folder -ChildPath "M365_License_Utilization_$stamp.html"

    if ($Format -in @('Csv','Both')) {
        $Rows | Sort-Object UtilizationState, LicenseSku, UserPrincipalName | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csvPath
        Write-Log "CSV report saved to $csvPath"
    }

    if ($Format -in @('Html','Both')) {
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

        $summaryState = $Rows | Group-Object UtilizationState | Sort-Object Name | Select-Object Name, Count
        $summarySku = $Rows | Group-Object LicenseSku | Sort-Object Count -Descending | Select-Object Name, Count

        $html = ConvertTo-Html -Title 'M365 License Utilization Report' -Head $style -Body @(
            '<h1>M365 License Utilization Report</h1>',
            "<p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>",
            "<p>Total License Assignments Evaluated: $($Rows.Count)</p>",
            '<h2>Summary by Utilization State</h2>',
            ($summaryState | ConvertTo-Html -Fragment),
            '<h2>Top License SKUs by Assignment Count</h2>',
            ($summarySku | ConvertTo-Html -Fragment),
            '<h2>Detailed Results</h2>',
            ($Rows | Sort-Object UtilizationState, LicenseSku, UserPrincipalName | ConvertTo-Html -Fragment)
        )

        $html | Out-File -Path $htmlPath -Encoding UTF8
        Write-Log "HTML report saved to $htmlPath"
    }
}

try {
    Ensure-Modules -SkipInstall:$SkipModuleInstall
    Connect-M365Graph

    Write-Log 'Loading users and license assignments...'
    $users = Get-MgUser -All -Property id,displayName,userPrincipalName,assignedLicenses,accountEnabled |
        Where-Object { $_.AssignedLicenses.Count -gt 0 }

    $skuLookup = Build-LicenseLookup
    $lookback = Get-LookbackDaysFromPeriod -PeriodValue $Period

    $tempPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("m365-usage-{0}" -f ([guid]::NewGuid().Guid))
    New-Item -Path $tempPath -ItemType Directory -Force | Out-Null

    try {
        Write-Log 'Downloading Graph usage reports...'
        $mailboxData = Get-GraphReportData -Endpoint "/v1.0/reports/getMailboxUsageDetail(period='$Period')" -TempFile (Join-Path -Path $tempPath -ChildPath 'mailbox.csv')
        $oneDriveData = Get-GraphReportData -Endpoint "/v1.0/reports/getOneDriveUsageAccountDetail(period='$Period')" -TempFile (Join-Path -Path $tempPath -ChildPath 'onedrive.csv')
        $teamsData = Get-GraphReportData -Endpoint "/v1.0/reports/getTeamsUserActivityUserDetail(period='$Period')" -TempFile (Join-Path -Path $tempPath -ChildPath 'teams.csv')
    }
    finally {
        Remove-Item -Path $tempPath -Recurse -Force -ErrorAction SilentlyContinue
    }

    $mailboxLookup = Add-ReportKeyedHashtable -Rows $mailboxData -KeyProperties @('User Principal Name','UserPrincipalName')
    $oneDriveLookup = Add-ReportKeyedHashtable -Rows $oneDriveData -KeyProperties @('Owner Principal Name','OwnerPrincipalName','User Principal Name','UserPrincipalName')
    $teamsLookup = Add-ReportKeyedHashtable -Rows $teamsData -KeyProperties @('User Principal Name','UserPrincipalName')

    $intuneCounts = Get-UserIntuneDeviceCounts

    Write-Log 'Evaluating license utilization...'
    $resultRows = New-LicenseUtilizationRows -Users $users -SkuLookup $skuLookup -MailboxUsage $mailboxLookup -OneDriveUsage $oneDriveLookup -TeamsUsage $teamsLookup -IntuneDeviceCounts $intuneCounts -LookbackDays $lookback -Period $Period

    Export-ReportFiles -Rows $resultRows -Format $OutputFormat -Folder $OutputFolder
    Write-Log 'Done.'
}
catch {
    Write-Log -Level 'ERROR' -Message $_.Exception.Message
    throw
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}

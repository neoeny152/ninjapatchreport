# ====================================================================================
# Generate-NinjaPatchReport.ps1
# ------------------------------------------------------------------------------------
# v12.13 - Modified by Gemini
#        - Fixed a bug where historical reports always showed the current month in the header.
#        - Improved historical accuracy by only querying for installed/failed patches for past months.
#        - Added disclaimers for the experimental historical reporting feature.
#
# v12.12 - Modified by Gemini
#        - Added logic for historical reporting and a "Patch Tuesday" warning banner.
# ====================================================================================

# ★★★ SCRIPT CONFIGURATION ★★★
$timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$outputCsvFile = "C:\admin\PatchReport_$timestamp.csv"
$outputHtmlFile = "C:\admin\PatchReport_$timestamp.html"
$logFile = "C:\admin\PatchReportLog_$timestamp.log"
$inactiveDeviceThresholdDays = 30

# Set to 1 to generate a separate HTML report for each organization, 0 to disable.
$generateOrgReports = 0

# --- (EXPERIMENTAL) Historical Reporting ---
# To run a report for a specific historical month, uncomment and set the variables below.
# NOTE: This feature is experimental and may not be fully accurate. It relies on
# historical patch installation data available via the API.
# $reportMonth = 6 # Example: 6 for June
# $reportYear = 2025 # Example: 2025

# To save a copy of the reports to a network share, uncomment the line below
# and replace the path with your desired network location.
# $fileSharePath = "\\YourServer\YourShare\PatchReports"

# ★★★ END CONFIGURATION ★★★


# ================================================================
# SCRIPT BODY
# ================================================================

# Ensure the C:\admin folder exists for logs and reports
if (!(Test-Path -Path "C:\admin")) {
    New-Item -ItemType Directory -Path "C:\admin" | Out-Null
}

# Start logging all actions to a transcript file
Start-Transcript -Path $logFile

try {
    # --- Announce Start ---
    Write-Host "================================================="
    Write-Host "Starting NinjaOne Patch Report Generation"
    Write-Host "Report CSV will be saved to: $outputCsvFile"
    Write-Host "Report HTML will be saved to: $outputHtmlFile"
    Write-Host "Full log will be saved to: $logFile"
    Write-Host "================================================="

    # (1) NinjaOne credentials (pull from your NinjaOne “Secret Custom Fields”)
    $API_INSTANCE        = Ninja-Property-Get ninjaoneInstance
    
    # --- Smart URL Handling ---
    if ($API_INSTANCE.StartsWith("http")) {
        $API_BASE_URL = $API_INSTANCE
    } else {
        $API_BASE_URL = "https://$($API_INSTANCE)"
    }

    $NINJA_CLIENT_ID     = Ninja-Property-Get ninjaoneClientId
    $NINJA_CLIENT_SECRET = Ninja-Property-Get ninjaoneClientSecret
    $NINJA_SCOPE         = "monitoring management"

    #================================================================
    # FUNCTIONS
    #================================================================

    function Get-AccessToken_Ninja {
        param($ApiBaseUrl, $ClientId, $ClientSecret, $Scope)
        Write-Host "Attempting to get NinjaOne Access Token..."
        try {
            $uri = "$ApiBaseUrl/oauth/token"
            $body = "grant_type=client_credentials&client_id=$ClientId&client_secret=$ClientSecret&scope=$Scope"
            $response = Invoke-RestMethod -Uri $uri -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded'
            if ($response.access_token) {
                Write-Host "✔ NinjaOne access token retrieved successfully." -ForegroundColor Green
                return $response.access_token
            }
        } catch {
            Write-Error "Failed to retrieve NinjaOne access token: $($_.Exception.Message)"
            throw
        }
    }

    function Get-AllNinjaApiData {
        param($AccessToken, $ApiBaseUrl, $Endpoint)
        $allData = [System.Collections.Generic.List[object]]::new()
        $uri = "$ApiBaseUrl/v2/$Endpoint`?pageSize=1000"
        $lastId = 0
        Write-Host "Fetching all records from '$Endpoint'..."
        do {
            $pageUri = "$uri&after=$lastId"
            try {
                $headers = @{"Authorization" = "Bearer $AccessToken"}
                $response = Invoke-RestMethod -Method Get -Uri $pageUri -Headers $headers
                if ($response) {
                    $allData.AddRange($response)
                    $lastId = $response[-1].id
                }
            } catch {
                Write-Error "Error fetching page for '$Endpoint': $($_.Exception.Message)"
                break
            }
        } while ($response -and $response.Count -gt 0)
        Write-Host "✔ Total '$Endpoint' records fetched: $($allData.Count)." -ForegroundColor Green
        return $allData
    }

    function Query-PatchData {
        param(
            [string]$AccessToken,
            [string]$ApiBaseUrl,
            [string]$Endpoint,
            [hashtable]$Parameters
        )
        $allResults = [System.Collections.Generic.List[object]]::new()
        $queryString = ($Parameters.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '&'
        try {
            $uriBuilder = [System.UriBuilder]::new($ApiBaseUrl)
            $uriBuilder.Path = $Endpoint
            $uriBuilder.Query = $queryString
            $fullUri = $uriBuilder.ToString()
        } catch {
            Write-Error "Failed to construct a valid URI. Error: $($_.Exception.Message)"
            return $null
        }
        do {
            try {
                $headers = @{"Authorization" = "Bearer $AccessToken"}
                $response = Invoke-RestMethod -Method Get -Uri $fullUri -Headers $headers
                if ($response.results) {
                    $allResults.AddRange($response.results)
                }
                if ($response.cursor -and $response.results.Count -gt 0) {
                    $cursorName = $response.cursor.name
                    $uriBuilder.Query = "$queryString&cursor=$cursorName"
                    $fullUri = $uriBuilder.ToString()
                } else {
                    $cursorName = $null
                }
            } catch {
                Write-Error "Error querying patch data: $($_.Exception.Message)"
                break 
            }
        } while ($cursorName)
        Write-Host "   ...fetched $($allResults.Count) total patch records for status '$($Parameters.status)'."
        return $allResults
    }
    
    function Convert-FromUnixTime {
        param([double]$UnixTime)
        if ($UnixTime -eq 0) { return "N/A" }
        return [System.DateTimeOffset]::FromUnixTimeSeconds($UnixTime).DateTime.ToString("yyyy-MM-dd HH:mm")
    }

    function ConvertTo-HtmlReport {
        param(
            [array]$ReportData,
            [string]$OutputFile,
            [string]$ApiBaseUrl,
            [hashtable]$WorkstationDeviceStats,
            [hashtable]$ServerDeviceStats,
            [string]$ReportTitle = "Monthly Patch Compliance Report",
            [int]$TotalDeviceCount,
            [datetime]$ReportDate,
            [bool]$IsCurrentMonthReport
        )

        Write-Host "`nGenerating HTML Report for '$($ReportTitle)'..."
        
        function Get-ComplianceColor {
            param($compliance)
            if ($compliance -ge 95) { return "#D5F5E3" } # Green
            if ($compliance -ge 90) { return "#FEF9E7" } # Yellow
            return "#FADBD8" # Red
        }

        # Helper function to generate the HTML for a table of devices
        function Generate-DeviceTable {
            param(
                [System.Text.StringBuilder]$htmlBuilder,
                [array]$DeviceData,
                [string]$ApiBaseUrl
            )
            $htmlBuilder.AppendLine('<table>')
            $htmlBuilder.AppendLine('<thead><tr><th>Device Name</th><th>Last User</th><th>OS</th><th>Location</th><th>Status</th><th>Installed KBs (This Month)</th><th>Pending KBs</th><th>Last Contact</th><th>Uptime</th><th>Days Offline</th></tr></thead>')
            $htmlBuilder.AppendLine('<tbody>')

            foreach($device in $DeviceData){
                $statusClass = "status-" + ($device.PatchStatus -replace ' ', '-')
                # Add a data-org attribute to each row for JavaScript filtering
                $orgNameAttribute = $device.Organization -replace "'", "&apos;"
                [void]$htmlBuilder.AppendLine("<tr class='${statusClass}' data-org='${orgNameAttribute}'>")
                $deviceUrl = "$ApiBaseUrl/#/deviceDashboard/$($device.DeviceId)/overview"
                [void]$htmlBuilder.Append("<td><a href='$deviceUrl' target='_blank'>$($device.DeviceName)</a></td>")
                $uptimeText = if ($device.UptimeDays -eq "N/A") { "N/A" } else { "$($device.UptimeDays) days" }
                $offlineText = if ($device.DaysOffline -eq 0) { "0" } else { "$($device.DaysOffline) days" }
                [void]$htmlBuilder.Append("<td>$($device.LastUser)</td><td>$($device.OS_Version)</td><td>$($device.Location)</td><td>$($device.PatchStatus)</td>")
                $installedKbLinks = $device.InstalledKBs | ForEach-Object { if (-not [string]::IsNullOrEmpty($_)) { $kb = $_.Trim(); "<a href='https://support.microsoft.com/kb/$($kb -replace 'KB')' target='_blank'>$kb</a>" } }
                [void]$htmlBuilder.Append("<td>$($installedKbLinks -join ', ')</td>")
                $pendingKbLinks = $device.PendingKBs | ForEach-Object { if (-not [string]::IsNullOrEmpty($_)) { $kb = $_.Trim(); "<a href='https://support.microsoft.com/kb/$($kb -replace 'KB')' target='_blank'>$kb</a>" } }
                [void]$htmlBuilder.Append("<td>$($pendingKbLinks -join ', ')</td>")
                [void]$htmlBuilder.Append("<td>$($device.LastContact)</td><td>$uptimeText</td><td>$offlineText</td></tr>")
            }

            $htmlBuilder.AppendLine('</tbody></table>')
        }

        $wsComplianceColor = Get-ComplianceColor -compliance $WorkstationDeviceStats.Compliance
        $svrComplianceColor = Get-ComplianceColor -compliance $ServerDeviceStats.Compliance

        $htmlBuilder = [System.Text.StringBuilder]::new()
        [void]$htmlBuilder.AppendLine('<!DOCTYPE html><html><head>')
        [void]$htmlBuilder.AppendLine("<meta charset='UTF-8'><title>$ReportTitle</title>")
        [void]$htmlBuilder.AppendLine("<style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f4f7f6; }
    .header-container { display: flex; justify-content: space-between; align-items: center; }
    h1, h2 { color: #2E4053; border-bottom: 2px solid #ddd; padding-bottom: 10px; }
    .warning-banner { background-color: #fff3cd; color: #856404; border: 1px solid #ffeeba; padding: 15px; border-radius: 8px; margin-bottom: 20px; text-align: center; }
    .filter-container { padding: 10px; margin-bottom: 20px; background-color: #e9ecef; border-radius: 8px; text-align: right; }
    #orgFilter { padding: 8px; border-radius: 5px; border: 1px solid #ccc; font-size: 1em; }
    .info-section { display: flex; justify-content: space-around; padding: 15px; background-color: #fff; border-radius: 8px; margin: 20px 0; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .info-item { text-align: center; }
    .info-item .label { font-size: 0.8em; font-weight: bold; color: #6c757d; text-transform: uppercase; letter-spacing: 0.5px; }
    .info-item .value { font-size: 1.2em; color: #343a40; margin-top: 5px; }
    .summary-section { margin-bottom: 20px; padding: 20px; background-color: white; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .summary-container { display: flex; flex-wrap: wrap; gap: 20px; }
    .summary-box { border: 1px solid #ccc; border-radius: 8px; padding: 15px; text-align: center; flex-grow: 1; min-width: 150px; }
    .summary-box .value { font-size: 2.5em; font-weight: bold; }
    .summary-box .label { font-size: 1em; color: #555; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 30px; }
    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
    th { background-color: #4A5568; color: white; }
    tr:nth-child(even) { background-color: #f2f2f2; }
    .status-Failed { background-color: #FDEDEC !important; }
    .status-Pending-Reboot { background-color: #FFF9C4 !important; }
    .status-Offline { background-color: #E0E0E0 !important; }
    a { color: #3498DB; text-decoration: none; }
    a:hover { text-decoration: underline; }
</style>")
        [void]$htmlBuilder.AppendLine('</head><body>')
        [void]$htmlBuilder.AppendLine("<div class='header-container'><h1>$ReportTitle</h1><a href='https://github.com/neoeny152/ninjapatchreport' target='_blank'>GitHub</a></div>")
        
        # --- Add Patch Tuesday Warning if applicable ---
        if ($IsCurrentMonthReport) {
            $firstDay = $ReportDate.AddDays(-($ReportDate.Day - 1))
            $firstTuesday = $firstDay
            while ($firstTuesday.DayOfWeek -ne 'Tuesday') { $firstTuesday = $firstTuesday.AddDays(1) }
            $secondTuesday = $firstTuesday.AddDays(7)

            if ((Get-Date) -lt $secondTuesday) {
                [void]$htmlBuilder.AppendLine("<div class='warning-banner'><b>Note:</b> This report is for the current month and is being run before the second Tuesday ('Patch Tuesday'). Compliance data may not be fully representative until after this month's patches are released and deployed.</div>")
            }
        }

        # --- Informational Header ---
        $generatedOnString = (Get-Date).ToString("dddd, MMMM dd, yyyy h:mm tt")
        $reportPeriodString = $ReportDate.ToString("MMMM yyyy")
        [void]$htmlBuilder.AppendLine("<div class='info-section'>")
        [void]$htmlBuilder.AppendLine("  <div class='info-item'><div class='label'>Generated On</div><div class='value'>$generatedOnString</div></div>")
        [void]$htmlBuilder.AppendLine("  <div class='info-item'><div class='label'>Report Period</div><div class='value'>$reportPeriodString</div></div>")
        [void]$htmlBuilder.AppendLine("  <div class='info-item'><div class='label'>Total Devices</div><div class='value'>$TotalDeviceCount</div></div>")
        [void]$htmlBuilder.AppendLine("</div>")

        # --- Organization Filter Dropdown ---
        $uniqueOrgs = $ReportData.Organization | Sort-Object -Unique
        if ($uniqueOrgs.Count -gt 1) {
            [void]$htmlBuilder.AppendLine("<div class='filter-container'>")
            [void]$htmlBuilder.AppendLine("  <label for='orgFilter'>Filter by Organization: </label>")
            [void]$htmlBuilder.AppendLine("  <select id='orgFilter' name='orgFilter'>")
            [void]$htmlBuilder.AppendLine("    <option value='all'>All Organizations</option>")
            foreach ($org in $uniqueOrgs) {
                $orgNameAttribute = $org -replace "'", "&apos;"
                [void]$htmlBuilder.AppendLine("    <option value='${orgNameAttribute}'>$($org)</option>")
            }
            [void]$htmlBuilder.AppendLine("  </select>")
            [void]$htmlBuilder.AppendLine("</div>")
        }

        [void]$htmlBuilder.AppendLine("<div class='summary-section'><h2>Workstation Summary</h2><div class='summary-container'><div class='summary-box' style='background-color:$wsComplianceColor'><div class='value'>$($WorkstationDeviceStats.Compliance)%</div><div class='label'>Compliance</div></div><div class='summary-box'><div class='value'>$($WorkstationDeviceStats.Total)</div><div class='label'>Total Machines</div></div><div class='summary-box'><div class='value'>$($WorkstationDeviceStats.Compliant)</div><div class='label'>Compliant</div></div><div class='summary-box'><div class='value'>$($WorkstationDeviceStats.NonCompliant)</div><div class='label'>Non-Compliant</div></div></div></div>")
        
        [void]$htmlBuilder.AppendLine("<div class='summary-section'><h2>Server Summary</h2><div class='summary-container'><div class='summary-box' style='background-color:$svrComplianceColor'><div class='value'>$($ServerDeviceStats.Compliance)%</div><div class='label'>Compliance</div></div><div class='summary-box'><div class='value'>$($ServerDeviceStats.Total)</div><div class='label'>Total Machines</div></div><div class='summary-box'><div class='value'>$($ServerDeviceStats.Compliant)</div><div class='label'>Compliant</div></div><div class='summary-box'><div class='value'>$($ServerDeviceStats.NonCompliant)</div><div class='label'>Non-Compliant</div></div></div></div>")
        
        # --- Split non-compliant devices into Servers and Workstations ---
        $nonCompliantDevices = $ReportData | Where-Object { $_.PatchStatus -ne 'Installed' }
        $nonCompliantServers = $nonCompliantDevices | Where-Object { $_.OS_Version -like "*Server*" }
        $nonCompliantWorkstations = $nonCompliantDevices | Where-Object { $_.OS_Version -like "*Workstation*" -or $_.OS_Version -like "*Windows 1*" }

        # Generate Server Table
        [void]$htmlBuilder.AppendLine("<div id='server-table-container'>")
        [void]$htmlBuilder.AppendLine("<h2>Non-Compliant Servers</h2>")
        Generate-DeviceTable -htmlBuilder $htmlBuilder -DeviceData $nonCompliantServers -ApiBaseUrl $ApiBaseUrl
        [void]$htmlBuilder.AppendLine("</div>")

        # Generate Workstation Table
        [void]$htmlBuilder.AppendLine("<div id='workstation-table-container'>")
        [void]$htmlBuilder.AppendLine("<h2>Non-Compliant Workstations</h2>")
        Generate-DeviceTable -htmlBuilder $htmlBuilder -DeviceData $nonCompliantWorkstations -ApiBaseUrl $ApiBaseUrl
        [void]$htmlBuilder.AppendLine("</div>")

        # --- JavaScript for Filtering ---
        [void]$htmlBuilder.AppendLine("<script>
    const orgFilter = document.getElementById('orgFilter');
    if (orgFilter) {
        orgFilter.addEventListener('change', function() {
            const selectedOrg = this.value;
            
            filterTable('server-table-container', selectedOrg);
            filterTable('workstation-table-container', selectedOrg);
        });
    }

    function filterTable(containerId, selectedOrg) {
        const container = document.getElementById(containerId);
        if (!container) return;

        const table = container.querySelector('table');
        if (!table) return;

        const rows = table.getElementsByTagName('tr');
        let visibleRows = 0;

        for (let i = 1; i < rows.length; i++) { // Start at 1 to skip header row
            const row = rows[i];
            const rowOrg = row.getAttribute('data-org');
            
            if (selectedOrg === 'all' || rowOrg === selectedOrg) {
                row.style.display = '';
                visibleRows++;
            } else {
                row.style.display = 'none';
            }
        }
        
        // Hide the entire container (header and table) if no rows are visible
        if (visibleRows === 0) {
            container.style.display = 'none';
        } else {
            container.style.display = '';
        }
    }
</script>")

        [void]$htmlBuilder.AppendLine('</body></html>')
        
        $htmlBuilder.ToString() | Out-File -FilePath $OutputFile -Force
        Write-Host "✔ HTML report generated successfully." -ForegroundColor Green
    }

    #================================================================
    # SCRIPT EXECUTION
    #================================================================

    # --- Determine Report Date Range ---
    $isCurrentMonthReport = $false
    if ($PSBoundParameters.ContainsKey('reportMonth') -and $PSBoundParameters.ContainsKey('reportYear')) {
        Write-Host "Historical report requested for $reportMonth/$reportYear." -ForegroundColor Yellow
        $reportStartDate = Get-Date -Year $reportYear -Month $reportMonth -Day 1 -Hour 0 -Minute 0 -Second 0
    } else {
        Write-Host "Defaulting to current month for report." -ForegroundColor Green
        $reportStartDate = Get-Date -Day 1 -Hour 0 -Minute 0 -Second 0
        $isCurrentMonthReport = $true
    }
    $reportEndDate = $reportStartDate.AddMonths(1)
    $startDateString = $reportStartDate.ToString('yyyy-MM-dd')
    $endDateString = $reportEndDate.ToString('yyyy-MM-dd')

    $accessToken = Get-AccessToken_Ninja -ApiBaseUrl $API_BASE_URL -ClientId $NINJA_CLIENT_ID -ClientSecret $NINJA_CLIENT_SECRET -Scope $NINJA_SCOPE
    
    $allOrgs = Get-AllNinjaApiData -AccessToken $accessToken -ApiBaseUrl $API_BASE_URL -Endpoint 'organizations'
    $allLocations = Get-AllNinjaApiData -AccessToken $accessToken -ApiBaseUrl $API_BASE_URL -Endpoint 'locations'
    $orgLookup = @{}; $allOrgs.ForEach({ $orgLookup[$_.id] = $_.name })
    $locLookup = @{}; $allLocations.ForEach({ $locLookup[$_.id] = $_.name })

    $allDevicesRaw = Get-AllNinjaApiData -AccessToken $accessToken -ApiBaseUrl $API_BASE_URL -Endpoint 'devices-detailed'
    
    $reportData = @{}
    foreach ($device in $allDevicesRaw) {
        if ($device.nodeClass -notlike "WINDOWS*") { continue }
        
        $daysOffline = 0
        $uptimeDays = "N/A"

        if ($device.offline) {
            $lastContactDate = [System.DateTimeOffset]::FromUnixTimeSeconds($device.lastContact).DateTime
            $daysOffline = [math]::Round((New-TimeSpan -Start $lastContactDate -End (Get-Date)).TotalDays, 0)
        } else { # Device is online
            if ($device.os -and $device.os.lastBootTime -ne 0) {
                $lastBoot = [System.DateTimeOffset]::FromUnixTimeSeconds($device.os.lastBootTime).DateTime
                $uptimeDays = [math]::Round((New-TimeSpan -Start $lastBoot -End (Get-Date)).TotalDays, 0)
            }
        }
        
        $reportData[$device.id] = [PSCustomObject]@{
            DeviceId     = $device.id
            DeviceName   = $device.systemName
            LastUser     = if ($device.lastLoggedInUser) { $device.lastLoggedInUser } else { "N/A" }
            OS_Version   = if ($device.os) { "$($device.os.name) $($device.os.releaseId)" } else { $device.nodeClass }
            Organization = if ($orgLookup.ContainsKey($device.organizationId)) { $orgLookup[$device.organizationId] } else { "N/A" }
            Location     = if ($locLookup.ContainsKey($device.locationId)) { $locLookup[$device.locationId] } else { "N/A" }
            PatchStatus  = if($device.offline) { "Offline" } else { "Not Patched" }
            InstalledKBs = [System.Collections.Generic.List[string]]::new()
            PendingKBs   = [System.Collections.Generic.List[string]]::new()
            LastContact  = Convert-FromUnixTime $device.lastContact
            UptimeDays   = $uptimeDays
            DaysOffline  = $daysOffline
            NeedsReboot  = if ($device.os) { $device.os.needsReboot } else { $false }
        }
    }
    
    Write-Host "`nQuerying patch data for the period starting $startDateString..."
    
    $patchNameFilter = { $_.name -like "*Cumulative Update*" -or $_.name -like "*.NET Framework*" }

    $installedPatches = Query-PatchData -AccessToken $accessToken -ApiBaseUrl $API_BASE_URL -Endpoint '/v2/queries/os-patch-installs' -Parameters @{status = 'INSTALLED'; installedAfter = $startDateString; installedBefore = $endDateString} | Where-Object $patchNameFilter
    $failedPatches = Query-PatchData -AccessToken $accessToken -ApiBaseUrl $API_BASE_URL -Endpoint '/v2/queries/os-patch-installs' -Parameters @{status = 'FAILED'; installedAfter = $startDateString; installedBefore = $endDateString} | Where-Object $patchNameFilter
    
    # Only query for pending/approved patches for the current month's report
    if ($isCurrentMonthReport) {
        $pendingPatches = Query-PatchData -AccessToken $accessToken -ApiBaseUrl $API_BASE_URL -Endpoint '/v2/queries/os-patches' -Parameters @{status = 'PENDING'} | Where-Object $patchNameFilter
        $approvedPatches = Query-PatchData -AccessToken $accessToken -ApiBaseUrl $API_BASE_URL -Endpoint '/v2/queries/os-patches' -Parameters @{status = 'APPROVED'} | Where-Object $patchNameFilter
        foreach ($patch in $pendingPatches) { if ($reportData.ContainsKey($patch.deviceId)) { $reportData[$patch.deviceId].PendingKBs.Add($patch.kbNumber) } }
        foreach ($patch in $approvedPatches) { if ($reportData.ContainsKey($patch.deviceId)) { $reportData[$patch.deviceId].PendingKBs.Add($patch.kbNumber) } }
    }

    foreach ($patch in $failedPatches) { if ($reportData.ContainsKey($patch.deviceId)) { $reportData[$patch.deviceId].PatchStatus = "Failed" } }
    foreach ($patch in $installedPatches) { if ($reportData.ContainsKey($patch.deviceId)) { $reportData[$patch.deviceId].InstalledKBs.Add($patch.kbNumber); if ($reportData[$patch.deviceId].PatchStatus -ne 'Failed') { $reportData[$patch.deviceId].PatchStatus = "Installed" } } }

    $allProcessedDevices = $reportData.Values | ForEach-Object {
        if ($_.PatchStatus -eq 'Installed' -and $_.NeedsReboot) {
            $_.PatchStatus = "Pending Reboot"
        }
        $_ 
    }

    $finalReportObjects = $allProcessedDevices | Where-Object { $_.DaysOffline -le $inactiveDeviceThresholdDays }
    Write-Host "`nFiltered for active devices. Kept $($finalReportObjects.Count) of $($allDevicesRaw.Count) total devices."

    $statusSortOrder = @{ 'Failed' = 1; 'Pending Reboot' = 2; 'Not Patched' = 3; 'Offline' = 4; 'Installed' = 5 }
    $finalReportObjects = $finalReportObjects | Sort-Object @{Expression = { $_.DaysOffline }}, @{Expression = { $statusSortOrder[$_.PatchStatus] }}, DeviceName
    
    # --- Generate the main, consolidated report ---
    $workstations = $finalReportObjects | Where-Object { $_.OS_Version -like "*Workstation*" -or $_.OS_Version -like "*Windows 1*" }
    $servers = $finalReportObjects | Where-Object { $_.OS_Version -like "*Server*" }
    
    $wsCompliantCount = ($workstations | Where-Object { $_.PatchStatus -eq 'Installed' }).Count
    $wsTotalCount = $workstations.Count
    $workstationDeviceStats = @{
        Total = $wsTotalCount
        Compliant = $wsCompliantCount
        NonCompliant = $wsTotalCount - $wsCompliantCount
        Compliance = if ($wsTotalCount -gt 0) { [math]::Round(($wsCompliantCount / $wsTotalCount) * 100) } else { 100 }
    }

    $svrCompliantCount = ($servers | Where-Object { $_.PatchStatus -eq 'Installed' }).Count
    $svrTotalCount = $servers.Count
    $serverDeviceStats = @{
        Total = $svrTotalCount
        Compliant = $svrCompliantCount
        NonCompliant = $svrTotalCount - $svrCompliantCount
        Compliance = if ($svrTotalCount -gt 0) { [math]::Round(($svrCompliantCount / $svrTotalCount) * 100) } else { 100 }
    }

    ConvertTo-HtmlReport -ReportData $finalReportObjects -OutputFile $outputHtmlFile -ApiBaseUrl $API_BASE_URL -WorkstationDeviceStats $workstationDeviceStats -ServerDeviceStats $serverDeviceStats -TotalDeviceCount $finalReportObjects.Count -ReportDate $reportStartDate -IsCurrentMonthReport $isCurrentMonthReport

    # --- (Optional) Generate a separate report for each organization ---
    if ($generateOrgReports -eq 1) {
        $uniqueOrgs = $finalReportObjects.Organization | Select-Object -Unique
        foreach ($orgName in $uniqueOrgs) {
            $orgDevices = $finalReportObjects | Where-Object { $_.Organization -eq $orgName }
            
            $orgWorkstations = $orgDevices | Where-Object { $_.OS_Version -like "*Workstation*" -or $_.OS_Version -like "*Windows 1*" }
            $orgServers = $orgDevices | Where-Object { $_.OS_Version -like "*Server*" }

            $orgWsCompliantCount = ($orgWorkstations | Where-Object { $_.PatchStatus -eq 'Installed' }).Count
            $orgWsTotalCount = $orgWorkstations.Count
            $orgWorkstationStats = @{
                Total = $orgWsTotalCount
                Compliant = $orgWsCompliantCount
                NonCompliant = $orgWsTotalCount - $orgWsCompliantCount
                Compliance = if ($orgWsTotalCount -gt 0) { [math]::Round(($orgWsCompliantCount / $orgWsTotalCount) * 100) } else { 100 }
            }

            $orgSvrCompliantCount = ($orgServers | Where-Object { $_.PatchStatus -eq 'Installed' }).Count
            $orgSvrTotalCount = $orgServers.Count
            $orgServerStats = @{
                Total = $orgSvrTotalCount
                Compliant = $orgSvrCompliantCount
                NonCompliant = $orgSvrTotalCount - $orgSvrCompliantCount
                Compliance = if ($orgSvrTotalCount -gt 0) { [math]::Round(($orgSvrCompliantCount / $orgSvrTotalCount) * 100) } else { 100 }
            }
            
            $safeOrgName = $orgName -replace '[^a-zA-Z0-9]', '-'
            $orgHtmlFile = "C:\admin\PatchReport_$(Get-Date -Format 'yyyy-MM-dd')_$($safeOrgName).html"

            ConvertTo-HtmlReport -ReportData $orgDevices -OutputFile $orgHtmlFile -ApiBaseUrl $API_BASE_URL -WorkstationDeviceStats $orgWorkstationStats -ServerDeviceStats $orgServerStats -ReportTitle "$orgName Patch Report" -TotalDeviceCount $orgDevices.Count -ReportDate $reportStartDate -IsCurrentMonthReport $isCurrentMonthReport
        }
    }

    # --- Export full data to CSV ---
    $csvReportObjects = $reportData.Values | ForEach-Object { # Use all data for CSV
        $clone = $_ | Select-Object *
        $clone.InstalledKBs = $clone.InstalledKBs -join ', '
        $clone.PendingKBs = $clone.PendingKBs -join ', '
        $clone | Select-Object DeviceId, DeviceName, LastUser, OS_Version, Organization, Location, PatchStatus, InstalledKBs, PendingKBs, LastContact, UptimeDays, DaysOffline
    }
    
    $csvReportObjects | Export-Csv -Path $outputCsvFile -NoTypeInformation -Force
    Write-Host "`n✔ CSV report has been saved to:" -ForegroundColor Yellow
    Write-Host $outputCsvFile

    # --- Copy files to network share if path is provided ---
    if ($fileSharePath -and (Test-Path -Path $fileSharePath)) {
        Write-Host "`nAttempting to copy reports to file share: $fileSharePath"
        try {
            $shareFolder = Join-Path -Path $fileSharePath -ChildPath $timestamp
            New-Item -ItemType Directory -Path $shareFolder -Force | Out-Null
            Copy-Item -Path $outputHtmlFile -Destination $shareFolder
            Copy-Item -Path $outputCsvFile -Destination $shareFolder
            
            if ($generateOrgReports -eq 1) {
                $orgShareFolder = Join-Path -Path $shareFolder -ChildPath "Organizations"
                New-Item -ItemType Directory -Path $orgShareFolder -Force | Out-Null
                Get-ChildItem -Path "C:\admin" -Filter "PatchReport_$(Get-Date -Format 'yyyy-MM-dd')_*.html" | ForEach-Object {
                    Copy-Item -Path $_.FullName -Destination $orgShareFolder
                }
            }

            Write-Host "✔ Reports successfully copied to share." -ForegroundColor Green
        } catch {
            Write-Warning "Failed to copy files to share: $($_.Exception.Message)"
        }
    }

    # --- Cleanup old files in C:\admin ---
    Write-Host "`nCleaning up old report files..."
    $fileTypes = @("PatchReport_*.html", "PatchReport_*.csv", "PatchReportLog_*.log")
    foreach ($type in $fileTypes) {
        $files = Get-ChildItem -Path "C:\admin" -Filter $type | Sort-Object CreationTime -Descending
        if ($files.Count -gt 5) {
            $files | Select-Object -Skip 5 | Remove-Item
            Write-Host "   - Cleaned up $($files.Count - 5) old '$type' files."
        }
    }
    Write-Host "✔ Cleanup complete."

} catch {
    Write-Error "A critical error occurred during script execution: $($_.Exception.Message)"
} finally {
    Write-Host "`nStopping transcript log."
    Stop-Transcript
}

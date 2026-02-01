param(
    [string]$ExcelPath = 'D:\Repositories\GitHub\MicrosoftDocs\dynamics365-guidance\tools\business-process-catalog-for-azure-devops\ADO template guideline.xlsx',
    [string]$OutputPath = 'D:\Repositories\GitHub\MicrosoftDocs\dynamics365-guidance\tools\business-process-catalog-for-azure-devops',
    [string]$AzureDevopsConfigPath = 'D:\Repositories\GitHub\MicrosoftDocs\dynamics365-guidance\tools\business-process-catalog-for-azure-devops\azure-devops-config.json',
    [switch]$CreateOnlyJsonFiles
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $AzureDevopsConfigPath)) {
    throw "Config file not found: $AzureDevopsConfigPath"
}
$config = Get-Content -Path $AzureDevopsConfigPath -Raw | ConvertFrom-Json
$patEnvVar = $config.'personal-access-token-env-var'
if (-not $patEnvVar) {
    throw 'Config missing "personal-access-token-env-var".'
}

$initialPat = [Environment]::GetEnvironmentVariable($patEnvVar, 'Process')
$patWasMissing = [string]::IsNullOrWhiteSpace($initialPat)

if ($patWasMissing) {
    $securePat = Read-Host 'Enter Azure DevOps PAT (needs Project and Team Read, write, & manage scope)' -AsSecureString
    $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePat)
    try {
        $plainPat = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
    }
    if ([string]::IsNullOrWhiteSpace($plainPat)) {
        throw 'No PAT provided.'
    }
    $encodedPat = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(":$($plainPat)"))
    [Environment]::SetEnvironmentVariable($patEnvVar, $encodedPat, 'Process')
}

try {
function Get-ExcelWorksheetData {
    param(
        [string]$Path,
        [string]$SheetName
    )

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    try {
        $wb = $excel.Workbooks.Open($Path, $false, $true)
        $ws = $wb.Worksheets.Item($SheetName)
        $used = $ws.UsedRange
        if (-not $used) { return @() }

        $data = $used.Value2
        $rowCount = $used.Rows.Count
        $colCount = $used.Columns.Count

        # Build header map
        $headers = @()
        for ($c = 1; $c -le $colCount; $c++) {
            $h = $data[1, $c]
            if (-not $h) { $h = "Column$($c)" }
            $headers += $h
        }

        $rows = @()
        for ($r = 2; $r -le $rowCount; $r++) {
            $rowObj = [ordered]@{}
            for ($c = 1; $c -le $colCount; $c++) {
                $rowObj[$headers[$c - 1]] = $data[$r, $c]
            }
            $rows += [pscustomobject]$rowObj
        }
        return $rows
    }
    finally {
        if ($wb) { $wb.Close($false) }
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

$sheetName = 'Teams'
if (-not (Test-Path $ExcelPath)) {
    throw "Excel file not found: $ExcelPath"
}

$rows = Get-ExcelWorksheetData -Path $ExcelPath -SheetName $sheetName

# Export teams to JSON file
$jsonOutputPath = Join-Path $OutputPath 'teams.json'
Write-Host "Exporting teams to JSON..."

# Extract teams from Column A (Team names in first column)
$teamsArray = @()

foreach ($row in $rows) {
    # Get the first column value (Teams column)
    $teamName = $row.Teams
    if ($teamName -and $teamName.ToString().Trim()) {
        $teamsArray += $teamName
    }
}

$teamsExport = @{
    'exportDate' = (Get-Date -Format 'O')
    'source' = $ExcelPath
    'sheetName' = $sheetName
    'teams' = $teamsArray
}

$teamsExport | ConvertTo-Json -Depth 10 | Set-Content -Path $jsonOutputPath -Encoding UTF8
Write-Host "Teams exported to: $jsonOutputPath"
Write-Host "Total teams exported: $($teamsArray.Count)"

# Skip Azure DevOps import if CreateOnlyJsonFiles flag is set
if ($CreateOnlyJsonFiles) {
    Write-Host "CreateOnlyJsonFiles flag set. Skipping Azure DevOps import."
}
else {
    # Import teams into Azure DevOps
    $organization = $config.organization
    $projectName = $config.project
    $apiVersion = $config.'api-version'

    if (-not $organization -or -not $projectName) {
        Write-Warning "Organization or project not configured. Skipping Azure DevOps import."
    }
    else {
        Write-Host "`nImporting teams to Azure DevOps..."
        
        # Get PAT from environment
        $encodedPat = [Environment]::GetEnvironmentVariable($patEnvVar, 'Process')
        $authHeader = @{ Authorization = "Basic $encodedPat" }
        
        try {
            # Get existing teams in the project
            $teamsUri = "https://dev.azure.com/$organization/_apis/projects/$projectName/teams?api-version=$apiVersion"
            $existingTeams = Invoke-RestMethod -Uri $teamsUri -Headers $authHeader -Method Get
            $existingTeamNames = @{}
            foreach ($team in $existingTeams.value) {
                $existingTeamNames[$team.name] = $team.id
            }
            Write-Host "Found $($existingTeams.value.Count) existing teams"
            
            # Load teams from JSON
            $teamsData = Get-Content -Path $jsonOutputPath -Raw | ConvertFrom-Json
            
            $successCount = 0
            $skipCount = 0
            $failCount = 0
            
            # Import each team
            foreach ($teamName in $teamsData.teams) {
                
                # Check if team already exists
                if ($existingTeamNames.ContainsKey($teamName)) {
                    Write-Host "[INFO] Team '$teamName' already exists (ID: $($existingTeamNames[$teamName])). Skipping." -ForegroundColor Yellow
                    $skipCount++
                    continue
                }
                
                try {
                    # Create team request body
                    $teamBody = @{
                        name = $teamName
                        description = "Imported from ADO template guideline"
                    } | ConvertTo-Json
                    
                    Write-Host "Creating team: $teamName..."
                    
                    # Create the team
                    $response = Invoke-RestMethod -Uri $teamsUri -Headers $authHeader -Method Post -Body $teamBody -ContentType 'application/json'
                    
                    Write-Host "[SUCCESS] Created team '$teamName' with ID: $($response.id)" -ForegroundColor Green
                    $successCount++
                }
                catch {
                    Write-Error "[ERROR] Failed to create team '$teamName': $($_.Exception.Message)" -ForegroundColor Red
                    $failCount++
                }
            }
            
            Write-Host "`nImport complete: $successCount succeeded, $skipCount skipped, $failCount failed"
        }
        catch {
            Write-Error "Failed to import teams: $($_.Exception.Message)"
        }
    }
}

}
finally {
    Remove-Item Env:\$patEnvVar -ErrorAction SilentlyContinue
}

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
    $securePat = Read-Host 'Enter Azure DevOps PAT (needs Work Items Read, write, & manage scope)' -AsSecureString
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

$sheetName = 'Area paths'
if (-not (Test-Path $ExcelPath)) {
    throw "Excel file not found: $ExcelPath"
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Open($ExcelPath, $false, $true)
$ws = $wb.Worksheets.Item($sheetName)
$used = $ws.UsedRange
$rowCount = $used.Rows.Count
$colCount = $used.Columns.Count

# Export area paths to JSON file
$jsonOutputPath = Join-Path $OutputPath 'areaPaths.json'
Write-Host "Exporting area paths to JSON..."

# Extract area paths from the worksheet
# Structure: Each row represents one hierarchy level
# When a column has a value, that's the hierarchy level for that row
# Column F contains the team association
$areaPathsArray = @()
$currentTeam = $null
$hierarchyPath = @($null, $null, $null, $null, $null)  # Track current hierarchy (5 levels max)

for ($r = 2; $r -le $rowCount; $r++) {
    $col1 = $ws.Cells.Item($r, 1).Value()  # Level 1
    $col2 = $ws.Cells.Item($r, 2).Value()  # Level 2
    $col3 = $ws.Cells.Item($r, 3).Value()  # Level 3
    $col4 = $ws.Cells.Item($r, 4).Value()  # Level 4
    $col5 = $ws.Cells.Item($r, 5).Value()  # Level 5
    $col6 = $ws.Cells.Item($r, 6).Value()  # Team

    # Update current team if provided
    if ($col6) {
        $currentTeam = $col6.ToString().Trim()
    }

    # Determine which hierarchy level has a value in this row
    if ($col1 -and $col1.ToString().Trim()) {
        $hierarchyPath[0] = $col1.ToString().Trim().Replace('Project name', $config.project)
        $hierarchyPath[1] = $null
        $hierarchyPath[2] = $null
        $hierarchyPath[3] = $null
        $hierarchyPath[4] = $null
    }
    elseif ($col2 -and $col2.ToString().Trim()) {
        $hierarchyPath[1] = $col2.ToString().Trim()
        $hierarchyPath[2] = $null
        $hierarchyPath[3] = $null
        $hierarchyPath[4] = $null
    }
    elseif ($col3 -and $col3.ToString().Trim()) {
        $hierarchyPath[2] = $col3.ToString().Trim()
        $hierarchyPath[3] = $null
        $hierarchyPath[4] = $null
    }
    elseif ($col4 -and $col4.ToString().Trim()) {
        $hierarchyPath[3] = $col4.ToString().Trim()
        $hierarchyPath[4] = $null
    }
    elseif ($col5 -and $col5.ToString().Trim()) {
        $hierarchyPath[4] = $col5.ToString().Trim()
    }
    else {
        # Empty row, skip
        continue
    }

    # Create structured area path object
    $areaPathObj = [ordered]@{
        'team' = if ($currentTeam) { $currentTeam } else { $null }
        'level1' = if ($hierarchyPath[0]) { $hierarchyPath[0] } else { $null }
        'level2' = if ($hierarchyPath[1]) { $hierarchyPath[1] } else { $null }
        'level3' = if ($hierarchyPath[2]) { $hierarchyPath[2] } else { $null }
        'level4' = if ($hierarchyPath[3]) { $hierarchyPath[3] } else { $null }
        'level5' = if ($hierarchyPath[4]) { $hierarchyPath[4] } else { $null }
    }
    
    $areaPathsArray += $areaPathObj
}

# Clean up Excel resources
$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

$areaPathsExport = @{
    'exportDate' = (Get-Date -Format 'O')
    'source' = $ExcelPath
    'sheetName' = $sheetName
    'areaPaths' = $areaPathsArray
}

$areaPathsExport | ConvertTo-Json -Depth 10 | Set-Content -Path $jsonOutputPath -Encoding UTF8
Write-Host "Area paths exported to: $jsonOutputPath"
Write-Host "Total area paths exported: $($areaPathsArray.Count)"

# Skip Azure DevOps import if CreateOnlyJsonFiles flag is set
if ($CreateOnlyJsonFiles) {
    Write-Host "CreateOnlyJsonFiles flag set. Skipping Azure DevOps import."
}
else {
    # Import area paths into Azure DevOps
    $organization = $config.organization
    $projectName = $config.project
    $apiVersion = $config.'api-version'

    if (-not $organization -or -not $projectName) {
        Write-Warning "Organization or project not configured. Skipping Azure DevOps import."
    }
    else {
        Write-Host "`nImporting area paths to Azure DevOps..."
        
        # Get PAT from environment
        $encodedPat = [Environment]::GetEnvironmentVariable($patEnvVar, 'Process')
        $authHeader = @{ Authorization = "Basic $encodedPat" }
        
        try {
            # Get existing area paths
            $classificationUri = "https://dev.azure.com/$organization/$projectName/_apis/wit/classificationnodes?`$depth=10&api-version=$apiVersion"
            $existingClassificationNodes = Invoke-RestMethod -Uri $classificationUri -Headers $authHeader -Method Get
            $existingAreas = $existingClassificationNodes.value | Where-Object { $_.structureType -eq 'area' }

            # Save existing area to json
            ##
            $existingAreasPath = Join-Path $PSScriptRoot "existingAreas.json"
            $existingAreas | ConvertTo-Json -Depth 10 | Set-Content -Path $existingAreasPath -Encoding UTF8
            Write-Host "Saved existing areas to: $existingAreasPath" -ForegroundColor Cyan
            ##
            
            # Build map of existing area paths
            $existingAreaPaths = @{}
            function Map-ExistingAreas {
                param(
                    [PSObject]$Node,
                    [string]$ParentPath = ''
                )
                
                if ($Node.name) {
                    $currentPath = if ($ParentPath) { "$ParentPath\$($Node.name)" } else { $Node.name }
                    $existingAreaPaths[$currentPath] = $Node
                    
                    if ($Node.children) {
                        foreach ($child in $Node.children) {
                            Map-ExistingAreas -Node $child -ParentPath $currentPath
                        }
                    }
                }
            }
            
            Map-ExistingAreas -Node $existingAreas
            Write-Host "Found $($existingAreaPaths.Count) existing area paths"
            
            # Load area paths from JSON
            $areaPathsData = Get-Content -Path $jsonOutputPath -Raw | ConvertFrom-Json
            
            $successCount = 0
            $skipCount = 0
            $failCount = 0

            $processedCount = 0
            
            # Import each area path
            foreach ($areaPath in $areaPathsData.areaPaths) {

                # if ($processedCount -gt 30) { break } # TEMP LIMIT FOR TESTING
                $processedCount++

                # Build the path from the level properties
                $pathParts = @()
                # if ($areaPath.level1) { $pathParts += $areaPath.level1 }
                if ($areaPath.level2) { $pathParts += $areaPath.level2 }
                if ($areaPath.level3) { $pathParts += $areaPath.level3 }
                if ($areaPath.level4) { $pathParts += $areaPath.level4 }
                if ($areaPath.level5) { $pathParts += $areaPath.level5 }
                
                # Skip if no path parts (empty row)
                if ($pathParts.Count -eq 0) {
                    continue
                }
                
                $pathValue = $pathParts -join '\'
                
                # Check if area path already exists
                $pathValueCheck = $config.project + '\' + $pathValue
                if ($existingAreaPaths.ContainsKey($pathValueCheck)) {
                    Write-Host "[INFO] Area path '$pathValueCheck' already exists. Skipping." -ForegroundColor Yellow
                    $skipCount++
                    continue
                }
                
                try {
                    # Split path into nodes and create missing ones
                    # Write-Host "Processing area path: $pathValue"
                    $pathNodes = $pathValue -split '\\'
                    $currentPath = ''
                    
                    foreach ($nodeName in $pathNodes) {
                        $currentPath = if ($currentPath) { "$currentPath\$nodeName" } else { $nodeName }
                        # Write-Host "  Checking node: $nodeName (Path so far: $currentPath)"
                        
                        # Check if this path already exists
                        $checkPath = $config.project + '\' + $currentPath
                        if ($existingAreaPaths.ContainsKey($checkPath)) {
                            # Write-Host "[INFO] Area path '$checkPath' already exists. Skipping." -ForegroundColor Yellow
                            continue
                        }
                        
                        # Create the area node
                        $parentPath = if ($currentPath -match '\\') {
                            $currentPath.Substring(0, $currentPath.LastIndexOf('\'))
                        }
                        else {
                            ''
                        }
                        
                        $createUri = "https://dev.azure.com/$organization/$projectName/_apis/wit/classificationnodes/Areas$(if ($parentPath) { '/' + ($parentPath -replace '\\', '/') })`?api-version=$apiVersion"
                        
                        $areaBody = @{
                            name = $nodeName
                        } | ConvertTo-Json
                        
                        # Write-Host "Creating area path: $currentPath..."
                        # Write-Host "URI: $createUri" -ForegroundColor DarkGray
                        # Write-Host "Body: $areaBody" -ForegroundColor DarkGray
                        
                        $response = Invoke-RestMethod -Uri $createUri -Headers $authHeader -Method Post -Body $areaBody -ContentType 'application/json'
                        $existingAreaPaths["$($config.project)\$currentPath"] = $response
                        
                        Write-Host "[SUCCESS] Created area path: $currentPath" -ForegroundColor Green

                        # Check if area path needs to be assigned to a team
                        if ($areaPath.team) {
                            # Get team field values
                            $teamFieldValuesUri = "https://dev.azure.com/$organization/$projectName/$($areaPath.team)/_apis/work/teamsettings/teamfieldvalues?api-version=$apiVersion"
                            $teamFieldValues = Invoke-RestMethod -Uri $teamFieldValuesUri -Headers $authHeader -Method Get

                            # Save existing team field values to json
                            <#
                            $teamFieldValuesPath = Join-Path $PSScriptRoot "existingTeamFieldValues_$($areaPath.team)_$($currentPath.replace('\', '_')).json"
                            $teamFieldValues | ConvertTo-Json -Depth 10 | Set-Content -Path $teamFieldValuesPath -Encoding UTF8
                            Write-Host "Saved team field values for team '$($areaPath.team)' to: $teamFieldValuesPath" -ForegroundColor Cyan
                            # if ($areaPath.team -eq "Accounting") { exit }
                            #>

                            $updateTeamFieldValues = @{}
                            $shouldUpdate = $false
                            # Set default value if empty
                            $updateTeamFieldValues.defaultValue = $teamFieldValues.defaultValue
                            if (-not $teamFieldValues.defaultValue -or [string]::IsNullOrWhiteSpace($teamFieldValues.defaultValue)) {
                                $updateTeamFieldValues.defaultValue = "$($config.project)\$currentPath"
                                $shouldUpdate = $true
                            }
                            # Add to values if not already present; value may be a partial path with includeChildren set to true
                            $pathWithProject = "$($config.project)\$currentPath"
                            $updateTeamFieldValues.values = $teamFieldValues.values
                            $includeCurrentPath = $true
                            # Write-Host "Checking existing team field values for team '$($areaPath.team)' and area path '$pathWithProject'..."
                            foreach ($value in $teamFieldValues.values) {
                                # Write-Host "  Existing team field value: $($value.value) (Include children: $($value.includeChildren))" -ForegroundColor DarkGray
                                if ($pathWithProject.Contains($value.value) -and $value.includeChildren) {
                                    $includeCurrentPath = $false
                                    break
                                }
                            }
                            if ($includeCurrentPath) {
                                $updateTeamFieldValues.values += @{
                                    value = $pathWithProject
                                    includeChildren = $true
                                }
                                $shouldUpdate = $true
                            }
                            if ($shouldUpdate) {
                                $updateBody = $updateTeamFieldValues | ConvertTo-Json -Depth 5
                                # Write-Host "Updating team field values for team '$($areaPath.team)' to include area path '$pathWithProject'..."
                                $response = Invoke-RestMethod -Uri $teamFieldValuesUri -Headers $authHeader -Method Patch -Body $updateBody -ContentType 'application/json'
                                Write-Host "[SUCCESS] Updated team field values for team '$($areaPath.team)' to include area path '$pathWithProject'" -ForegroundColor Green
                            }
                        }

                        $successCount++
                    }
                }
                catch {
                    Write-Error "[ERROR] Failed to create area path '$pathValue': $($_.Exception.Message)"
                    $failCount++
                }
            }
            
            Write-Host "`nImport complete: $successCount succeeded, $skipCount skipped, $failCount failed"
        }
        catch {
            Write-Error "Failed to import area paths: $($_.Exception.Message)"
        }
    }
}

}
finally {
    Remove-Item Env:\$patEnvVar -ErrorAction SilentlyContinue
}

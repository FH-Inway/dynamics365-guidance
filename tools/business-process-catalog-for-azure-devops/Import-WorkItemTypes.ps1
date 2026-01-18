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

$sheetName = 'Work item types'
if (-not (Test-Path $ExcelPath)) {
    throw "Excel file not found: $ExcelPath"
}

$rows = Get-ExcelWorksheetData -Path $ExcelPath -SheetName $sheetName

$items = @()
foreach ($row in $rows) {
    $name = $row.'Work item type'
    if (-not $name) { continue }

    $description = $row.'Help text'
    $color = $row.'Color'
    # Ensure color is a 6-character hex string
    if ($color) {
        # Remove decimal if present (e.g., "222222.0" -> "222222")
        $color = $color -replace '\.0$', ''
        # Ensure it's uppercase and exactly 6 characters
        $color = $color.ToUpper().Substring([Math]::Max(0, $color.Length - 6)).PadLeft(6, '0')
    }
    $icon = $row.'Icon'
    $disabled = $row.'Custom work item type'
    $inherits = $row.'Inherit from'

    $obj = [ordered]@{
        name         = $name
        description  = $description
        color        = $color
        icon         = $icon
        isDisabled   = if ($disabled -eq 'Disabled') { $true } else { $false }
        inheritsFrom = if ($inherits) { $inherits } else { $null }
    }

    $items += [pscustomobject]$obj
}

$items | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 (Join-Path $OutputPath 'workItemTypes.json')
Write-Host "Wrote $($items.Count) work item types to $(Join-Path $OutputPath 'workItemTypes.json')"

# Skip Azure DevOps import if CreateOnlyJsonFiles flag is set
if ($CreateOnlyJsonFiles) {
    Write-Host "CreateOnlyJsonFiles flag set. Skipping Azure DevOps import."
}
else {
    # Import work item types into Azure DevOps
    $organization = $config.organization
    $processName = $config.process
    $apiVersion = $config.'api-version'

    if (-not $organization -or -not $processName) {
        Write-Warning "Organization or process not configured. Skipping Azure DevOps import."
    }
    else {
        Write-Host "`nImporting work item types to Azure DevOps..."
        
        # Get PAT from environment
        $encodedPat = [Environment]::GetEnvironmentVariable($patEnvVar, 'Process')
        $authHeader = @{ Authorization = "Basic $encodedPat" }
        
        # Get process ID from process name
        $processListUri = "https://dev.azure.com/$organization/_apis/work/processes?api-version=$apiVersion"
        try {
            $processes = Invoke-RestMethod -Uri $processListUri -Headers $authHeader -Method Get
            $process = $processes.value | Where-Object { $_.name -eq $processName }
            
            if (-not $process) {
                throw "Process '$processName' not found in organization '$organization'."
            }
            
            $processId = $process.typeId
            Write-Host "Found process '$processName' with ID: $processId"
            
            # Get existing work item types
            $existingWitsUri = "https://dev.azure.com/$organization/_apis/work/processes/$processId/workitemtypes?api-version=$apiVersion"
            $existingWits = Invoke-RestMethod -Uri $existingWitsUri -Headers $authHeader -Method Get
            $existingWitMap = @{}
            foreach ($wit in $existingWits.value) {
                $existingWitMap[$wit.name] = $wit
            }
            
            # Create each work item type
            $successCount = 0
            $failCount = 0
            $skipCount = 0
            $updateCount = 0
            foreach ($item in $items) {
                # Check if work item type already exists
                if ($existingWitMap.ContainsKey($item.name)) {
                    $existingWit = $existingWitMap[$item.name]
                    
                    # Check if we need to disable it
                    if ($item.isDisabled -and -not $existingWit.isDisabled) {
                        # System work item types can only be updated when they were updated once in the gui first, see https://developercommunity.visualstudio.com/t/rest-api-work-item-types-update-errors/1044106.
                        Write-Host "  [INFO] Manually disable existing work item type: $($item.name) with reference name $($existingWit.referenceName)" -ForegroundColor Cyan
                        <#
                        $updateUri = "https://dev.azure.com/$organization/_apis/work/processes/$processId/workitemtypes/$($existingWit.referenceName)?api-version=$apiVersion"
                        $updateBody = @{ isDisabled = $true } | ConvertTo-Json -Depth 5
                        
                        try {
                            $result = Invoke-RestMethod -Uri $updateUri -Headers $authHeader -Method Patch -Body $updateBody -ContentType 'application/json'
                            Write-Host "  [UPDATE] Disabled: $($item.name)" -ForegroundColor Cyan
                            $updateCount++
                        }
                        catch {
                            Write-Host "  [ERROR] Failed to disable: $($item.name) - $($_.Exception.Message)" -ForegroundColor Red
                            $failCount++
                        }
                        #>
                    }
                    else {
                        Write-Host "  [SKIP] Skipped: $($item.name) (already exists)" -ForegroundColor Yellow
                        $skipCount++
                    }
                    continue
                }
                
                $createUri = "https://dev.azure.com/$organization/_apis/work/processes/$processId/workitemtypes?api-version=$apiVersion"
                
                $body = @{
                    name = $item.name
                    color = $item.color
                    icon = $item.icon
                }
                
                if ($item.description) {
                    $body.description = $item.description
                }
                if ($item.isDisabled) {
                    $body.isDisabled = $item.isDisabled
                }
                if ($item.inheritsFrom) {
                    $body.inherits = $item.inheritsFrom
                }
                
                $jsonBody = $body | ConvertTo-Json -Depth 5
                
                try {
                    $result = Invoke-RestMethod -Uri $createUri -Headers $authHeader -Method Post -Body $jsonBody -ContentType 'application/json'
                    Write-Host "  [OK] Created: $($item.name)" -ForegroundColor Green
                    $successCount++
                }
                catch {
                    Write-Host "  [ERROR] Failed: $($item.name) - $($_.Exception.Message)" -ForegroundColor Red
                    $failCount++
                }
            }
            
            Write-Host "`nImport complete: $successCount succeeded, $skipCount skipped, $updateCount updated, $failCount failed"
        }
        catch {
            Write-Error "Failed to import work item types: $($_.Exception.Message)"
        }
    }
}

}
finally {
    Remove-Item Env:\$patEnvVar -ErrorAction SilentlyContinue
}
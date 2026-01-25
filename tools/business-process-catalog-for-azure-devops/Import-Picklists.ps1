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

$sheetName = 'Picklists'
if (-not (Test-Path $ExcelPath)) {
    throw "Excel file not found: $ExcelPath"
}

$rows = Get-ExcelWorksheetData -Path $ExcelPath -SheetName $sheetName

# Export picklists to JSON file
$jsonOutputPath = Join-Path $OutputPath 'picklists.json'
Write-Host "Exporting picklists to JSON..."

# Transpose the data: each column becomes a picklist with an array of values
$picklistsArray = @()

if ($rows.Count -gt 0) {
    # Get all property names (column headers)
    $headers = $rows[0].PSObject.Properties | ForEach-Object { $_.Name }
    
    # For each column, collect all non-empty values
    foreach ($header in $headers) {
        $values = @()
        foreach ($row in $rows) {
            $value = $row.$header
            if ($value -and $value.ToString().Trim()) {
                $values += $value.ToString().Trim()
            }
        }
        
        # Only include the picklist if it has values
        if ($values.Count -gt 0) {
            $picklistsArray += @{
                'name' = $header
                'items' = $values
            }
        }
    }
}

$picklistsObject = @{
    'exportDate' = (Get-Date -Format 'O')
    'source' = $ExcelPath
    'sheetName' = $sheetName
    'picklists' = $picklistsArray
}

$picklistsObject | ConvertTo-Json -Depth 10 | Set-Content -Path $jsonOutputPath -Encoding UTF8
Write-Host "Picklists exported to: $jsonOutputPath"

# Skip Azure DevOps import if CreateOnlyJsonFiles flag is set
if ($CreateOnlyJsonFiles) {
    Write-Host "CreateOnlyJsonFiles flag set. Skipping Azure DevOps import."
}
else {
    # Import picklists into Azure DevOps
    $organization = $config.organization
    $processName = $config.process
    $apiVersion = $config.'api-version'

    if (-not $organization -or -not $processName) {
        Write-Warning "Organization or process not configured. Skipping Azure DevOps import."
    }
    else {
        Write-Host "`nImporting picklists to Azure DevOps..."
        
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
            
            # Get existing picklists
            $listsUri = "https://dev.azure.com/$organization/_apis/work/processes/lists?api-version=$apiVersion"
            $existingLists = Invoke-RestMethod -Uri $listsUri -Headers $authHeader -Method Get
            $existingPicklistNames = @{}
            foreach ($list in $existingLists.value) {
                $existingPicklistNames[$list.name] = $list.id
            }
            Write-Host "Found $($existingLists.value.Count) existing picklists"
            
            # Load picklists from JSON
            $picklistsData = Get-Content -Path $jsonOutputPath -Raw | ConvertFrom-Json
            
            $successCount = 0
            $skipCount = 0
            $failCount = 0
            
            # Import each picklist
            foreach ($picklist in $picklistsData.picklists) {
                $picklistName = $picklist.name
                
                # Check if picklist already exists
                if ($existingPicklistNames.ContainsKey($picklistName)) {
                    Write-Host "[INFO] Picklist '$picklistName' already exists (ID: $($existingPicklistNames[$picklistName])). Skipping." -ForegroundColor Yellow
                    $skipCount++
                    continue
                }
                
                try {
                    # Create picklist request body
                    $picklistBody = @{
                        id = $null
                        name = $picklistName
                        type = "String"  # Default to String type
                        items = $picklist.items
                        isSuggested = $false
                    } | ConvertTo-Json
                    
                    Write-Host "Creating picklist: $picklistName with $($picklist.items.Count) items..."
                    
                    # Create the picklist
                    $response = Invoke-RestMethod -Uri $listsUri -Headers $authHeader -Method Post -Body $picklistBody -ContentType 'application/json'
                    
                    Write-Host "[SUCCESS] Created picklist '$picklistName' with ID: $($response.id)" -ForegroundColor Green
                    $successCount++
                }
                catch {
                    Write-Host "[ERROR] Failed to create picklist '$picklistName': $($_.Exception.Message)" -ForegroundColor Red
                    $failCount++
                }
            }
            
            Write-Host "`nImport complete: $successCount succeeded, $skipCount skipped, $failCount failed"
        }
        catch {
            Write-Error "Failed to import picklists: $($_.Exception.Message)"
        }
    }
}

}
finally {
    Remove-Item Env:\$patEnvVar -ErrorAction SilentlyContinue
}
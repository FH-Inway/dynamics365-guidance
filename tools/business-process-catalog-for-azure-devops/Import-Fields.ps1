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

$sheetName = 'Fields'
if (-not (Test-Path $ExcelPath)) {
    throw "Excel file not found: $ExcelPath"
}

$rows = Get-ExcelWorksheetData -Path $ExcelPath -SheetName $sheetName

# Export fields to JSON file
$jsonOutputPath = Join-Path $OutputPath 'fields.json'
Write-Host "Exporting fields to JSON..."

# Convert rows to field objects array
$fieldsArray = @()

foreach ($row in $rows) {
    $fieldObj = @{}
    $row.PSObject.Properties | ForEach-Object {
        $fieldObj[$_.Name] = $_.Value
    }
    $fieldsArray += $fieldObj
}

$fieldsExport = @{
    'exportDate' = (Get-Date -Format 'O')
    'source' = $ExcelPath
    'sheetName' = $sheetName
    'fields' = $fieldsArray
}

$fieldsExport | ConvertTo-Json -Depth 10 | Set-Content -Path $jsonOutputPath -Encoding UTF8
Write-Host "Fields exported to: $jsonOutputPath"

# Skip Azure DevOps import if CreateOnlyJsonFiles flag is set
if ($CreateOnlyJsonFiles) {
    Write-Host "CreateOnlyJsonFiles flag set. Skipping Azure DevOps import."
}
else {
    # Import fields into Azure DevOps
    $organization = $config.organization
    $processName = $config.process
    $apiVersion = $config.'api-version'

    if (-not $organization -or -not $processName) {
        Write-Warning "Organization or process not configured. Skipping Azure DevOps import."
    }
    else {
        Write-Host "`nImporting fields to Azure DevOps..."
        
        # Get PAT from environment
        $encodedPat = [Environment]::GetEnvironmentVariable($patEnvVar, 'Process')
        $authHeader = @{ Authorization = "Basic $encodedPat" }
        
        # Get existing fields from Azure DevOps
        $fieldsUri = "https://dev.azure.com/$organization/_apis/wit/fields?api-version=$apiVersion"
        # Get picklists (used for picklist-backed fields)
        $picklistsUri = "https://dev.azure.com/$organization/_apis/work/processes/lists?api-version=$apiVersion"
        try {
            $existingFields = Invoke-RestMethod -Uri $fieldsUri -Headers $authHeader -Method Get
            $existingFieldNames = @{}
            foreach ($field in $existingFields.value) {
                $existingFieldNames[$field.referenceName] = $field.name
            }
            Write-Host "Found $($existingFields.value.Count) existing fields"

            $existingPicklists = Invoke-RestMethod -Uri $picklistsUri -Headers $authHeader -Method Get
            $picklistsByName = @{}
            foreach ($plist in $existingPicklists.value) {
                $picklistsByName[$plist.name] = $plist.id
            }
            Write-Host "Found $($existingPicklists.value.Count) existing picklists"
            
            # Load fields from JSON
            $fieldsData = Get-Content -Path $jsonOutputPath -Raw | ConvertFrom-Json
            
            # Filter to only custom fields (Custom field = "Yes") and limit to first 5 for testing
            $customFields = $fieldsData.fields | Where-Object { $_.'Custom field' -eq 'Yes' }
            Write-Host "Found $($customFields.Count) custom fields to import"
            
            $successCount = 0
            $skipCount = 0
            $warningCount = 0
            $failCount = 0
            
            # Import each custom field
            foreach ($field in $customFields) {
                $referenceName = $field.'Reference name'
                $fieldName = $field.'Field name'
                
                # Check for empty reference names
                if ([string]::IsNullOrWhiteSpace($referenceName)) {
                    Write-Host "[WARNING] Field '$fieldName' has empty reference name. Skipping." -ForegroundColor Yellow
                    $skipCount++
                    continue
                }
                
                # Skip system fields (those starting with System. or Microsoft.)
                if ($referenceName -match '^(System\.|Microsoft\.)') {
                    Write-Host "[WARNING] Field '$fieldName' (Reference: $referenceName) is a system field. Skipping." -ForegroundColor Yellow
                    $skipCount++
                    continue
                }
                
                # Check if field already exists
                if ($existingFieldNames.ContainsKey($referenceName)) {
                    Write-Host "[WARNING] Custom field '$fieldName' (Reference: $referenceName) already exists. Skipping." -ForegroundColor Yellow
                    $warningCount++
                    continue
                }
                
                try {
                    # Determine field type from Excel data
                    $fieldType = $field.'Field type'
                    
                    # Map field types to API field types
                    $apiFieldType = switch ($fieldType) {
                        'Boolean' { 'boolean' }
                        'HTML' { 'html' }
                        'String' { 'string' }
                        'Integer' { 'integer' }
                        'Decimal' { 'double' }
                        'DateTime' { 'dateTime' }
                        'Identity' { 'identity' }
                        'PicklistString' { 'string' }
                        'TreePath' { 'treePath' }
                        default { 'string' }
                    }

                    $isIdentity = $false
                    if ($fieldType -eq 'Identity') {
                        $isIdentity = $true
                    }

                    $isPicklist = $false
                    $picklistId = $null
                    if ($fieldType -eq 'PicklistString') {
                        $isPicklist = $true
                        if ($picklistsByName.ContainsKey($fieldName)) {
                            $picklistId = $picklistsByName[$fieldName]
                        }
                        else {
                            Write-Host "[WARNING] Picklist for field '$fieldName' not found. Expected picklist name: '$fieldName'. Skipping field." -ForegroundColor Yellow
                            $skipCount++
                            continue
                        }
                    }
                    
                    # Create field request body
                    $fieldBody = @{
                        name = $fieldName
                        referenceName = $referenceName
                        description = $field.Description -replace "`n", " "
                        type = $apiFieldType
                        usage = "workItem"
                        readOnly = $false
                        isQueryable = $true
                        canSortBy = $true
                        isIdentity = $isIdentity
                        isPicklist = $isPicklist
                        picklistId = $picklistId
                    } | ConvertTo-Json
                    
                    Write-Host "Creating custom field: $fieldName (Type: $apiFieldType)..."
                    
                    # Create the field
                    $response = Invoke-RestMethod -Uri $fieldsUri -Headers $authHeader -Method Post -Body $fieldBody -ContentType 'application/json'
                    
                    Write-Host "[SUCCESS] Created custom field '$fieldName' with reference: $($response.referenceName)" -ForegroundColor Green
                    $successCount++
                }
                catch {
                    Write-Host "[ERROR] Failed to create custom field '$fieldName': $($_.Exception.Message)" -ForegroundColor Red
                    $failCount++
                }
            }
            
            Write-Host "`nImport complete: $successCount succeeded, $warningCount skipped (already exist), $skipCount skipped (system/empty/missing picklist), $failCount failed"
        }
        catch {
            Write-Error "Failed to import fields: $($_.Exception.Message)"
        }
    }
}

}
finally {
    Remove-Item Env:\$patEnvVar -ErrorAction SilentlyContinue
}
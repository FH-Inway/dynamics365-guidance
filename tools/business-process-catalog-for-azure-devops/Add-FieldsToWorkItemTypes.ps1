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

# Build work item type → fields mapping from worksheet
if (-not $rows -or $rows.Count -eq 0) {
    throw "Worksheet '$sheetName' has no data."
}

$headerNames = @()
$firstRow = $rows[0]
foreach ($p in $firstRow.PSObject.Properties) {
    $headerNames += $p.Name
}

$baseHeaders = $headerNames | Select-Object -First 10
$fieldHeaders = $headerNames | Select-Object -Skip 10

$output = @()
foreach ($row in $rows) {
    $obj = [ordered]@{}
    foreach ($h in $baseHeaders) {
        $obj[$h] = $row.$h
    }

    $fields = @()
    foreach ($fh in $fieldHeaders) {
        $val = $row.$fh
        if ($null -ne $val -and ($val -eq 'X' -or $val -eq 'x')) {
            $fields += $fh
        }
    }
    $obj['fields'] = $fields

    $output += [pscustomobject]$obj
}

$witFieldsPath = Join-Path $OutputPath 'workItemTypesFields.json'
$output | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 $witFieldsPath
Write-Host "Wrote $($output.Count) work item type-field mappings to $witFieldsPath"

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
        Write-Host "`nAssociating fields with work item types in Azure DevOps..."
        
        # Get PAT from environment
        $encodedPat = [Environment]::GetEnvironmentVariable($patEnvVar, 'Process')
        $authHeader = @{ Authorization = "Basic $encodedPat" }
        
        # Get process ID from process name
        $processListUri = "https://dev.azure.com/$organization/_apis/work/processes?api-version=$apiVersion"
        # Get existing fields from Azure DevOps
        $fieldsUri = "https://dev.azure.com/$organization/_apis/wit/fields?api-version=$apiVersion"
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
            Write-Host "Found $($existingWits.value.Count) existing work item types"
            
            # Save existing work item types to JSON
            <#
            $existingWitsPath = Join-Path $PSScriptRoot "existingWorkItemTypes.json"
            $existingWits | ConvertTo-Json -Depth 10 | Set-Content -Path $existingWitsPath -Encoding UTF8
            Write-Host "Saved existing work item types to: $existingWitsPath" -ForegroundColor Cyan
            #>

            # Get existing fields
            $existingFields = Invoke-RestMethod -Uri $fieldsUri -Headers $authHeader -Method Get
            $existingFieldNames = @{}
            foreach ($field in $existingFields.value) {
                $existingFieldNames[$field.referenceName] = $field.name
            }
            Write-Host "Found $($existingFields.value.Count) existing fields"
            
            # Save existing fields to JSON
            <#
            $existingFieldsPath = Join-Path $PSScriptRoot "existingFields.json"
            $existingFields | ConvertTo-Json -Depth 10 | Set-Content -Path $existingFieldsPath -Encoding UTF8
            Write-Host "Saved existing fields to: $existingFieldsPath" -ForegroundColor Cyan
            #>
            
            # Process each work item type from the JSON
            $successCount = 0
            $skipCount = 0
            $failCount = 0
            $warningCount = 0
            
            # For testing: only process the first work item type
            $processedCount = 0
            
            foreach ($witMapping in $output) {
                # if ($processedCount -ge 2) { break }
                
                $witName = $witMapping.'Work item type'
                if (-not $witName) { continue }

                if ($witMapping.'Custom work item type' -ne 'Yes') {
                    Write-Host "  [SKIP] Work item type '$witName' is not custom. Skipping." -ForegroundColor Yellow
                    $skipCount++
                    continue
                }
                # if ($witName -ne 'Job') { continue } # For testing: only process Test case
                
                # Check if work item type exists in Azure DevOps
                if (-not $existingWitMap.ContainsKey($witName)) {
                    Write-Warning "  [SKIP] Work item type '$witName' not found in process"
                    $skipCount++
                    continue
                }
                
                $wit = $existingWitMap[$witName]
                $witRefName = $wit.referenceName
                
                # Get existing fields for this work item type
                $witFieldsUri = "https://dev.azure.com/$organization/_apis/work/processes/$processId/workItemTypes/$witRefName/fields?api-version=$apiVersion"
                $existingWitFields = Invoke-RestMethod -Uri $witFieldsUri -Headers $authHeader -Method Get
                
                # Save existing fields for this work item type to JSON
                <#
                $safeWitName = $witName -replace '[\\/:*?"<>|]', '_'
                $existingWitFieldsPath = Join-Path $PSScriptRoot "existingFields_$safeWitName.json"
                $existingWitFields | ConvertTo-Json -Depth 10 | Set-Content -Path $existingWitFieldsPath -Encoding UTF8
                Write-Host "  Saved fields for '$witName' to: $existingWitFieldsPath" -ForegroundColor Cyan
                #>
                
                $existingFieldRefNames = $existingWitFields.value | Select-Object -ExpandProperty referenceName
                
                # Get desired fields from JSON
                $desiredFields = $witMapping.fields
                if (-not $desiredFields) { $desiredFields = @() }
                
                # Check for fields in Azure DevOps that are not in JSON (excluding system fields)
                foreach ($existingFieldRef in $existingFieldRefNames) {
                    $existingField = $existingWitFields.value | Where-Object { $_.referenceName -eq $existingFieldRef }
                    # Skip system and inherited fields (customization = "system", customization = "inherited")
                    if ($existingField.customization -eq 'system') { continue }
                    if ($existingField.customization -eq 'inherited') { continue }
                    
                    if ($existingField.name -notin $desiredFields) {
                        Write-Warning "  [WARN] Field '$($existingField.name)' is associated with '$witName' but not defined in JSON"
                        $warningCount++
                    }
                }
                
                # Add missing fields
                $existingFieldNames = $existingWitFields.value | Select-Object -ExpandProperty name
                $fieldsToAdd = $desiredFields | Where-Object { $_ -notin $existingFieldNames }
                
                foreach ($fieldName in $fieldsToAdd) {
                    if ([string]::IsNullOrWhiteSpace($fieldName)) { continue }

                    $fieldToAdd = $existingFields.value | Where-Object { $_.name -eq $fieldName }
                    if (-not $fieldToAdd) {
                        Write-Warning "  [SKIP] Field '$fieldName' not found in Azure DevOps. Cannot add to '$witName'."
                        $skipCount++
                        continue
                    }
                    
                    $addFieldUri = "https://dev.azure.com/$organization/_apis/work/processes/$processId/workItemTypes/$witRefName/fields?api-version=$apiVersion"
                    $addFieldBody = @{
                        referenceName = $fieldToAdd.referenceName
                    }
                    if ($fieldToAdd.type -eq 'boolean') {
                        # For boolean fields, make them required and set default value to "0"
                        $addFieldBody['required'] = $true
                        $addFieldBody['defaultValue'] = '0'
                    }
                    $addFieldBody = $addFieldBody | ConvertTo-Json -Depth 5
                    
                    try {
                        $result = Invoke-RestMethod -Uri $addFieldUri -Headers $authHeader -Method Post -Body $addFieldBody -ContentType 'application/json'
                        Write-Host "  [OK] Added field '$fieldName' (ref: $($fieldToAdd.referenceName)) to '$witName'" -ForegroundColor Green
                        # Write-Host "       Result: $($result | ConvertTo-Json -Depth 5)" -ForegroundColor DarkGray
                        $successCount++
                    }
                    catch {
                        Write-Host "  [ERROR] Failed to add field '$fieldName' to '$witName': $($_.Exception.Message)" -ForegroundColor Red
                        $failCount++
                    }
                }
                
                # Report if no fields needed to be added
                if ($fieldsToAdd.Count -eq 0) {
                    Write-Host "  [SKIP] All fields already associated with '$witName'" -ForegroundColor Yellow
                }
                
                $processedCount++
            }
            
            Write-Host "`nAssociation complete: $successCount fields added, $skipCount work item types skipped, $warningCount warnings, $failCount failed"
        }
        catch {
            Write-Error "Failed to associate fields with work item types: $($_.Exception.Message)"
        }
    }
}

}
finally {
    Remove-Item Env:\$patEnvVar -ErrorAction SilentlyContinue
}
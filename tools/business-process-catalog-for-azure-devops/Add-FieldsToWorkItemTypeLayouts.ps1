param(
    [string]$JsonFilesPath = 'D:\Repositories\GitHub\MicrosoftDocs\dynamics365-guidance\tools\business-process-catalog-for-azure-devops',
    [string]$AzureDevopsConfigPath = 'D:\Repositories\GitHub\MicrosoftDocs\dynamics365-guidance\tools\business-process-catalog-for-azure-devops\azure-devops-config.json'
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

    # Read fields.json file with layout information
    $fieldsJsonPath = Join-Path $JsonFilesPath 'fields.json'
    if (-not (Test-Path $fieldsJsonPath)) {
        throw "fields.json file not found: $fieldsJsonPath"
    }
    $fields = Get-Content -Path $fieldsJsonPath -Raw | ConvertFrom-Json
    if (-not $fields) {
        throw "No data found in fields.json file: $fieldsJsonPath"
    }
    $fields = $fields.fields
    if (-not $fields) {
        throw "No 'fields' array found in fields.json file: $fieldsJsonPath"
    }

    # Read workItemTypesFields.json file with work item type -> fields mapping
    $witFieldsJsonPath = Join-Path $JsonFilesPath 'workItemTypesFields.json'
    if (-not (Test-Path $witFieldsJsonPath)) {
        throw "workItemTypesFields.json file not found: $witFieldsJsonPath"
    }
    $witFieldsMapping = Get-Content -Path $witFieldsJsonPath -Raw | ConvertFrom-Json
    if (-not $witFieldsMapping) {
        throw "No data found in workItemTypesFields.json file: $witFieldsJsonPath"
    }

    # Import fields into Azure DevOps
    $organization = $config.organization
    $processName = $config.process
    $apiVersion = $config.'api-version'

    if (-not $organization -or -not $processName) {
        Write-Warning "Organization or process not configured. Skipping Azure DevOps import."
    }
    else {
        Write-Host "`nUpdating layouts of work item types in Azure DevOps..."
        
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

            $processedCount = 0
            
            foreach ($witMapping in $witFieldsMapping) {
                # if ($processedCount -ge 2) { break } # For testing: only process a few work item types
                
                $witName = $witMapping.'Work item type'
                if (-not $witName) { continue }

                if ($witMapping.'Custom work item type' -ne 'Yes') {
                    Write-Host "  [SKIP] Work item type '$witName' is not custom. Skipping." -ForegroundColor Yellow
                    $skipCount++
                    continue
                }
                # if ($witName -ne 'Action item') { continue } # For testing: only process a specific work item type
                
                # Check if work item type exists in Azure DevOps
                if (-not $existingWitMap.ContainsKey($witName)) {
                    Write-Warning "  [SKIP] Work item type '$witName' not found in process"
                    $skipCount++
                    continue
                }
                
                $wit = $existingWitMap[$witName]
                $witRefName = $wit.referenceName
                
                # Save existing fields for this work item type to JSON
                <#
                $safeWitName = $witName -replace '[\\/:*?"<>|]', '_'
                $existingWitFieldsPath = Join-Path $PSScriptRoot "existingFields_$safeWitName.json"
                $existingWitFields | ConvertTo-Json -Depth 10 | Set-Content -Path $existingWitFieldsPath -Encoding UTF8
                Write-Host "  Saved fields for '$witName' to: $existingWitFieldsPath" -ForegroundColor Cyan
                #>
                
                # Get desired fields from JSON
                $desiredFields = $witMapping.fields
                if (-not $desiredFields) { $desiredFields = @() }
                
                # Get existing work item type layout
                $layoutUri = "https://dev.azure.com/$organization/_apis/work/processes/$processId/workItemTypes/$witRefName/layout?api-version=$apiVersion"
                $existingLayout = Invoke-RestMethod -Uri $layoutUri -Headers $authHeader -Method Get
                <#
                $existingLayoutJsonPath = Join-Path $PSScriptRoot "existingLayout_$($witName -replace '[\\/:*?"<>|]', '_').json"
                $existingLayout | ConvertTo-Json -Depth 10 | Set-Content -Path $existingLayoutJsonPath -Encoding UTF8
                Write-Host "  Saved existing layout for '$witName' to: $existingLayoutJsonPath" -ForegroundColor Cyan
                #>

                # Iterate over desired fields and add missing fields to layout
                foreach ($desiredField in $desiredFields) {
                    # Get layout information for the desired field
                    $fieldLayout = $fields | Where-Object { $_."Field name" -eq $desiredField }
                    $field = $existingFields.value | Where-Object { $_.name -eq $fieldLayout.'Field name' }
                    if (-not $field) {
                        Write-Warning "    [WARN] Field '$desiredField' not found in existing fields. Skipping."
                        $warningCount++
                        continue
                    }
                    $page = $existingLayout.pages | Where-Object { $_.label -eq $fieldLayout.'Page name' }
                    if (-not $page) {
                        # Create the page
                        $addPageUri = "https://dev.azure.com/$organization/_apis/work/processes/$processId/workItemTypes/$witRefName/layout/pages?api-version=$apiVersion"
                        $pagePayload = @{
                            "label" = $fieldLayout.'Page name'
                            "visible" = $true
                            "sections" = @()
                        }
                        $pageBody = $pagePayload | ConvertTo-Json -Depth 10
                        try {
                            $page = Invoke-RestMethod -Uri $addPageUri -Headers $authHeader -Method Post -Body $pageBody -ContentType 'application/json'
                            Write-Host "    [ADD] Created page '$($fieldLayout.'Page name')' in layout of '$witName'."
                            $existingLayout.pages += $page
                        }
                        catch {
                            Write-Error "    [FAIL] Failed to create page '$($fieldLayout.'Page name')' in layout of '$witName': $($_.Exception.Message)"
                            $failCount++
                            continue
                        }
                    }
                    $section = $page.sections | Where-Object { $_.id -eq $fieldLayout.'Group location' }
                    if (-not $section) {
                        Write-Warning "    [WARN] Section '$($fieldLayout.'Group location')' not found in page '$($fieldLayout.'Page name')' of layout of '$witName'. Skipping field '$desiredField'."
                        $warningCount++
                        continue
                    }
                    $group = $section.groups | Where-Object { $_.label -eq $fieldLayout.'Group name' }
                    if (-not $group) {
                        # Create the group
                        $isHTMLField = $false
                        $addGroupUri = "https://dev.azure.com/$organization/_apis/work/processes/$processId/workItemTypes/$witRefName/layout/pages/$($page.id)/sections/$($section.id)/groups?api-version=$apiVersion"
                        $groupPayload = @{
                            "label" = $fieldLayout.'Group name'
                            "order" = $fieldLayout.'Group sequence'
                            "visible" = $true
                            "controls" = @()
                        }
                        # HTML fields require their own group and group and control need to be created together
                        # https://github.com/artgarciams/copyWorkItemType/blob/38d761f7ffa50b4cfe949671c6860627fc7d0e01/ProjectAndGroup.psm1#L280
                        if ($fieldLayout.'Field type' -eq 'HTML') {
                            $groupPayload.controls += @{
                                "id" = $field.referenceName
                                "label" = $fieldLayout.'Field name'
                                "visible" = $true
                                "readOnly" = $false
                                "isContribution" = $false
                            }
                            $isHTMLField = $true
                        }
                        $groupBody = $groupPayload | ConvertTo-Json -Depth 10
                        try {
                            $group = Invoke-RestMethod -Uri $addGroupUri -Headers $authHeader -Method Post -Body $groupBody -ContentType 'application/json'
                            Write-Host "    [ADD] Created group '$($fieldLayout.'Group name')' in section '$($fieldLayout.'Group location')' of page '$($fieldLayout.'Page name')' in layout of '$witName'."
                            $section.groups += $group
                            if ($isHTMLField) {
                                Write-Host "    [ADD] Added HTML field '$($fieldLayout.'Field name')' to newly created group '$($fieldLayout.'Group name')' in layout of '$witName'."
                                $successCount++
                                continue
                            }
                        }
                        catch {
                            Write-Error "    [FAIL] Failed to create group '$($fieldLayout.'Group name')' in layout of '$witName': $($_.Exception.Message)"
                            $failCount++
                            continue
                        }
                    }
                    # Check if field is already in layout
                    $label = $fieldLayout.'Field name'
                    if ( [string]::IsNullOrWhiteSpace($fieldLayout.'Label') -eq $false) {
                        $label = $fieldLayout.'Label'
                    }
                    $existingField = $group.controls | Where-Object { $_.label -eq $label }
                    if ($existingField) {
                        Write-Host "    [SKIP] Field '$($fieldLayout.'Field name')', label '$label' already exists in layout of '$witName'."
                        continue
                    }
                    $existingField = $group.controls | Where-Object { $_.id -eq $field.referenceName }
                    if ($existingField) {
                        Write-Host "    [SKIP] Field '$($fieldLayout.'Field name')', reference name '$($fieldLayout.'Reference name')' already exists in layout of '$witName'."
                        continue
                    }
                    # Add field to group
                    $controlPayload = @{
                        "id" = $field.referenceName
                        "controlType" = "FieldControl"
                        "label" = $label
                        "order" = $fieldLayout.'Field sequence'
                        "visible" = $true
                        "readOnly" = $false
                        "isContribution" = $false
                    }
                    $addControlUri = "https://dev.azure.com/$organization/_apis/work/processes/$processId/workItemTypes/$witRefName/layout/groups/$($group.id)/controls?api-version=$apiVersion"
                    $controlBody = $controlPayload | ConvertTo-Json -Depth 10
                    try {
                        $control = Invoke-RestMethod -Uri $addControlUri -Headers $authHeader -Method Post -Body $controlBody -ContentType 'application/json'
                        Write-Host "    [ADD] Added field '$($fieldLayout.'Field name')', label '$label' to group '$($fieldLayout.'Group name')' in layout of '$witName'."
                        $group.controls += $control
                        $successCount++
                    }
                    catch {
                        Write-Error "    [FAIL] Failed to add field '$($fieldLayout.'Field name')' to layout of '$witName' in page '$($fieldLayout.'Page name')', section '$($fieldLayout.'Group location')', group '$($fieldLayout.'Group name')': $($_.Exception.Message)"
                        Write-Host "      Payload: $controlBody"                        
                        $failCount++
                        continue
                    }
                }
                $processedCount++
            }
            
            Write-Host "`nAssociation complete: $successCount fields added, $skipCount work item types skipped, $warningCount warnings, $failCount failed"
        }
        catch {
            Write-Error "Failed to add fields in work item type layouts: $($_.Exception.Message)"
        }
    }
}

finally {
    Remove-Item Env:\$patEnvVar -ErrorAction SilentlyContinue
}
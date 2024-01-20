# This script can be used to split the Microsoft Dynamics 365 Business Process Catalog CSV file
# into smaller files for import into Azure DevOps.
# See https://learn.microsoft.com/en-us/dynamics365/guidance/business-processes/about-import-catalog-devops

# Usage: .\Split-CatalogCSVForDevOpsImport.ps1 -CatalogFile "C:\Users\<username>\Downloads\Business Process Catalog ADO Upload DEC2023.csv"

param (
  [Parameter(Mandatory=$true)]
  [string]$CatalogFile
)

# Function to create split file
function CreateSplitFile {
  param (
    # parameter with csv content
    [Parameter(Mandatory=$true)]
    [array]$CsvContent,
    [Parameter(Mandatory=$true)]
    [string]$CatalogFileNameWithoutExtension,
    [Parameter(Mandatory=$true)]
    [int]$FileCounter,
    [Parameter(Mandatory=$true)]
    [string]$OriginalHeaderLine
  )

  # Create a new CSV file with the header
  $CsvContent | Export-Csv -Path "$CatalogFileNameWithoutExtension-Part$FileCounter.csv" -NoTypeInformation

  # Replace the first line in the file with the original header line
  $splitFile = Get-Content -Path "$CatalogFileNameWithoutExtension-Part$FileCounter.csv"
  $splitFile[0] = $OriginalHeaderLine
  $splitFile | Set-Content -Path "$CatalogFileNameWithoutExtension-Part$FileCounter.csv"
}

# The actual script starts here
$catalogFileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($CatalogFile)

# Import the CSV file
# Read the CSV file as text, because of multiple "Title" columns
$text = Get-Content -Path $CatalogFile

# Split the text into lines
$lines = $text -split "`n"

$originalHeaderLine = $lines[0]

# Split the header into columns
$header = $lines[0] -split ","

# Make the column names unique
for ($i = 0; $i -lt $header.Length; $i++) {
  for ($j = 0; $j -lt $i; $j++) {
    if ($header[$i] -eq $header[$j]) {
      $header[$i] += "_$i"
    }
  }
}

# Replace the header in the text
$lines[0] = $header -join ","

# Convert the text to CSV
$csv = $lines -join "`n" | ConvertFrom-Csv

# Initialize counters
$rowCounter = 0
$fileCounter = 1

# Split the header into columns ($csv[0] cannot be used, because it contains the headers along with the first line of data)
$headerColumns = $lines[0] -split ","

# Create a new object with properties corresponding to the headers
$header = New-Object PSObject
foreach ($column in $headerColumns) {
  $header | Add-Member -MemberType NoteProperty -Name $column -Value $null
}

# Create a new CSV file with the header
$header | Export-Csv -Path "$catalogFileNameWithoutExtension-Part$fileCounter.csv" -NoTypeInformation

# Initialize $newCsv as an empty array
$newCsv = @()

foreach ($i in 0..($csv.Count - 1)) {

  # If the current row is not an "Epic"
  if ($csv[$i].'Work Item Type' -ne 'Epic') {
    # Add the row to the current CSV file
    $newCsv += $csv[$i]
    $rowCounter++
  }

  # If the current row is an "Epic" 
  # and the rows collected so far plus the number of rows until the next "Epic" is greater than 1000
  # split the file

  # Determine the number of rows until the next "Epic"
  $nextEpicCount = 1 # Include the Epic line itself in the count
  for ($j = $i+1; $j -lt $csv.Count -and $csv[$j].'Work Item Type' -ne 'Epic'; $j++) {
    $nextEpicCount++
  }

  if ($rowCounter + $nextEpicCount -gt 1000) {
    $createSplitFileParameters = @{
      CsvContent = $newCsv
      CatalogFileNameWithoutExtension = $catalogFileNameWithoutExtension
      FileCounter = $fileCounter
      OriginalHeaderLine = $originalHeaderLine
    }
    CreateSplitFile @createSplitFileParameters

    # Prepare next split file
    $fileCounter++
    $header | Export-Csv -Path "$catalogFileNameWithoutExtension-Part$fileCounter.csv" -NoTypeInformation
    $newCsv = @()
    $rowCounter = 0
  }
  
}

# Export the last CSV file
if ($newCsv.Count -gt 0) {
  $createSplitFileParameters = @{
    CsvContent = $newCsv
    CatalogFileNameWithoutExtension = $catalogFileNameWithoutExtension
    FileCounter = $fileCounter
    OriginalHeaderLine = $originalHeaderLine
  }
  CreateSplitFile @createSplitFileParameters
}
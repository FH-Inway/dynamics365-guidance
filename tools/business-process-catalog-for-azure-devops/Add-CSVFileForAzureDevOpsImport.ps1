# Creates a csv file based on the following format:
# ID,Work Item Type,Title,Title,Title,Description,Assigned To,Business Owner,Business Process Lead,Area Path,Tags,State,Priority,Risk,Effort,Business Value,Time Criticality,Business Outcome Category,Value Area,Process Sequence ID,Discipline
$csvFileName = "business-process-catalog.csv"
New-Item -Path $csvFileName -ItemType File -Force
$csvFileContent = "ID,Work Item Type,Title,Title,Title,Description,Assigned To,Business Owner,Business Process Lead,Area Path,Tags,State,Priority,Risk,Effort,Business Value,Time Criticality,Business Outcome Category,Value Area,Process Sequence ID,Discipline"
$csvFileContent | Out-File -FilePath $csvFileName -Encoding utf8

# Iterate through the .md files of the guidance/business-processes folder where the file name does end with "-introduction.md"
# Store the beginning of the file name in a collection

$introductionFiles = Get-ChildItem -Path guidance\business-processes -Filter *-introduction.md -Recurse
$processNames = @()
foreach ($introductionFile in $introductionFiles) {
    $processName = $introductionFile.Name.Replace("-introduction.md", "")
    $processNames += $processName

    # Create a row in the csv file with the following values:
    # ID: leave empty
    # Work Item Type: Epic
    # Title: the process name, but the first letter is capitalized and the dash is replaced by a space
    # Title: leave empty
    # Title: leave empty
    # Description: the content of the first paragraph in the introduction file where the heading contains the word "overview"; the paragraph content should be converted to plain text
    # Assigned To: leave empty
    # Business Owner: leave empty
    # Business Process Lead: leave empty
}

# Iterate through the .md files of the guidance/business-processes folder where the file name does begin with one of the $processNames, but does not end with one of the following values:
# -introduction.md
# -overview.md
# -areas.md
# For each file, create a new row in the csv file with the following values:
# ID: leave empty
# Work Item Type: Epic
# Title: the file name of the current file, but with the process name and .md extension removed; the first letter is capitalized and the dash is replaced by a space
# Title: 
# Title: 
# Description: the content of the first paragraph in the current file where the heading contains the words "introduction to"; the paragraph content should be converted to plain text
# Assigned To: leave empty
# Business Owner: leave empty
# Business Process Lead: leave empty
# Area Path: add the process name to the string "DevOps Product Catalog Working Instance/"; capitalize the first letter of the process name and replace the dash by a space
# Tags: leave empty
# State: New
# Priority: 1
# Risk: leave empty
# Effort: leave empty
# Business Value: leave empty
# Time Criticality: leave empty
# Business Outcome Category: leave empty
# Value Area: Business
# Process Sequence ID: leave empty
# Discipline: leave empty

foreach ($processName in $processNames) {
    $processFiles = Get-ChildItem -Path guidance\business-processes -Filter $processName*.md -Recurse
    foreach ($processFile in $processFiles) {
        if ($processFile.Name -notmatch "(-introduction\.md|-overview\.md|-areas\.md)$") {
        # if ($processFile.Name -notlike "*-introduction.md" -and $processFile.Name -notlike "*-overview.md") {
            $processFileName = $processFile.Name.Replace($processName, "").Replace(".md", "")
            $processFileName = $processFileName.Substring(1, $processFileName.Length - 1)
            $processFileName = $processFileName.Substring(0, 1).ToUpper() + $processFileName.Substring(1, $processFileName.Length - 1)
            $processFileName = $processFileName.Replace("-", " ")
            $processFileName = $processFileName.Replace("  ", " ")
            


            # create a new row in the csv file with the following values:
            # ID: leave empty
            # Work Item Type: Epic
            # Title: the file name of the current file, but with the process name and .md extension removed; the first letter is capitalized and the dash is replaced by a space
            # Title: 
            # Title: 
            # Description: leave empty // TODO: the content of the first paragraph in the current file where the heading contains the words "introduction to"; the paragraph content should be converted to plain text
            # Assigned To: leave empty
            # Business Owner: leave empty
            # Business Process Lead: leave empty
            # Area Path: add the process name to the string "DevOps Product Catalog Working Instance/"; capitalize the first letter of the process name and replace the dash by a space
            # Tags: leave empty
            # State: New
            # Priority: 1
            # Risk: leave empty
            # Effort: leave empty
            # Business Value: leave empty
            # Time Criticality: leave empty
            # Business Outcome Category: leave empty
            # Value Area: Business
            # Process Sequence ID: leave empty
            # Discipline: leave empty
            Write-Host $processFile.Name
            $csvFileContent = "`n,Epic,$processFileName,,,,,,DevOps Product Catalog Working Instance/$processName,New,1,,,,Business,,"
            $csvFileContent | Out-File -FilePath $csvFileName -Encoding utf8 -Append
        }
    }
}

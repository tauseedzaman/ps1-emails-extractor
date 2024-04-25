#===========================================================================
#  ?ABOUT
#  Author         : Tauseed Zaman
#  Email          : tauseedzaman@pm.me    
#  CreatedOn      : 25/04/2024
#  Description    : This PowerShell script extracts email addresses from a text file and saves them to a CSV. It checks for the existence of the source file, reads its content, and extracts emails from it and puts them into a csv file 
#===========================================================================

# Define paths
$SOURCE_TXT_FILE_PATH = "./source.txt"
$OUTPUT_CSV_FILE_PATH = "./output.csv"

# Check if the source text file exists
if (-not (Test-Path "$SOURCE_TXT_FILE_PATH")) {
    Write-Host "The specified text file does not exist."
    exit
}

# Read the content of the text file
$fileContent = Get-Content -Path "$SOURCE_TXT_FILE_PATH" -Raw

# Initialize an empty array to store extracted email addresses
$emails = @()

# Construct regular expression pattern to match email addresses
$emailPattern = '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,3}\b'

# Use regular expression to find email addresses in the file content
$emailMatches = [regex]::Matches($fileContent, $emailPattern)

# Filter out the captured email addresses
foreach ($match in $emailMatches) {
    $emails += [PSCustomObject]@{
        Email = $match.Value
    }
}

# Export the extracted email addresses to a CSV file
$emails | Export-Csv -Path "$OUTPUT_CSV_FILE_PATH" -NoTypeInformation

Write-Host "Emails extracted from $SOURCE_TXT_FILE_PATH file and saved to $OUTPUT_CSV_FILE_PATH successfully."

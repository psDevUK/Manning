function Extract-ADInfo {
    param (
        [string]$FilePath,
        [string]$OutputPath
    )
    
    # Read the text file
    $content = Get-Content -Path $FilePath
    
    # Define the regex pattern
    $pattern = "(?<=GivenName\s+:\s)(\w+)|(?<=Surname\s+:\s)(\w+)|(?<=EmailAddress\s+:\s)([\w.-]+@[\w.-]+)|(?<=Department\s+:\s)(\w+)"
    
    # Find matches in the content
    $matches = $content -match $pattern
    
    # Create an array to hold the extracted data
    $data = @()
    
    # Iterate through matches and create a custom object
    for ($i = 0; $i -lt $matches.Count; $i += 4) {
        $givenName = ($matches[$i] -replace "GivenName\s+:","").Trim()
        $surname = ($matches[$i+1] -replace "SurName\s+:","").Trim()
        $emailAddress = ($matches[$i+2] -replace "emailAddress\s+:","").Trim()
        $department = ($matches[$i+3] -replace "department\s+:","").Trim()
        
        # Create a custom object
        $person = [PSCustomObject]@{
            GivenName = $givenName
            Surname = $surname
            EmailAddress = $emailAddress
            Department = $department
        }
        
        # Add the person object to the data array
        $data += $person
    }
    #Test if output exists
    if (Test-Path -Path $OutputPath)
    {
    Remove-Item -Path $OutputPath -Force
    }
    # Export the data to a CSV file
    $data | Export-Csv -Path $OutputPath -NoTypeInformation
}
function ConvertTo-HTMLReport {
    param (
        [string]$CSVFilePath,
        [string]$OutputFilePath,
        [string]$ReportTitle
    )
    
    # Import the CSV data
    $data = Import-Csv -Path $CSVFilePath | ? Department -eq "Marketing"
    
    # Convert the data to an HTML table with border and style
    $htmlTable = $data | ConvertTo-Html -As Table -PreContent "<table style='border-collapse: collapse; border: 2px solid black;'><colgroup><col style='background-color: #f2f2f2;'></colgroup>" -PostContent "</table>" -Property *
    
    # Create the HTML report content with style
    $htmlReport = @"
<!DOCTYPE html>
<html>
<head>
<title>$ReportTitle</title>
<style>
body {
    font-family: Arial, sans-serif;
}

h2 {
    color: #990000;
}

table, th, td {
  border: 1px solid black;
  border-collapse: collapse;
}
</style>
</head>
<body>
<h2>$ReportTitle</h2>
$htmlTable
</body>
</html>
"@
    
    # Save the HTML report to a file
    $htmlReport | Out-File -FilePath $OutputFilePath
}


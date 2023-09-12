# Import the CSV file
$csv = Import-Csv -Path .\DLs.csv

# Loop through each row in the CSV
foreach ($row in $csv) {
    # Get the distribution list from the current row
    $distributionList = $row.Name

    # Check if the distribution list exists
    $exists = Get-DistributionGroup -Identity $distributionList -ErrorAction SilentlyContinue

    if ($exists) {
        Write-Host "The distribution list $distributionList exists." -ForegroundColor Red
    } else {
        Write-Host "The distribution list $distributionList does not exist." -ForegroundColor Green
    }
}

Get-MgOrganization
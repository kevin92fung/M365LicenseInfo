<#  
With this script, you will generate two exported files to help you  
manage and analyze Microsoft 365 licenses and users in your tenant.  

The first file contains detailed info about all paid licenses,  
showing assigned and available license counts.  

The second file lists all users with key attributes such as display name,  
email, assigned licenses (matched to product names using the latest  
Microsoft SKU list), department, job title, office location, and more.  

The script downloads the latest SKU reference from Microsoft for accurate  
product naming. Both reports are first exported as CSV files, then converted  
to formatted Excel tables with autofilter enabled for easy sorting/filtering.  

You can apply filters on user attributes (e.g., office location, department,  
job title) within the script to focus your report on specific groups.  

**Requirements:**  
- You must run this script with a Global Administrator account for your tenant.  
- Microsoft Graph PowerShell module is required. Recommended to use PowerShell 7+  
  for best compatibility with Microsoft Graph.  

**Quick install guide for Microsoft Graph module:**  
Run in PowerShell 7+ console:  
Install-Module Microsoft.Graph -Scope CurrentUser -AllowClobber  

Use the resulting Excel tables to quickly filter and analyze licenses and users.  
#>



# Connect to MG-Graph
connect-mggraph -Scopes "User.Read.All","Directory.Read.All","AuditLog.Read.All"

# Download CSV with product names from Microsoft documentation
$pageUrl = "https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference"

# Download the webpage
$response = Invoke-WebRequest -Uri $pageUrl

# Extract the link to the CSV file containing license information
$csvLink = $response.Links |
    Where-Object { $_.href -like "https://download.microsoft.com/*" -and $_.href -like "*.csv" } |
    Select-Object -ExpandProperty href

# Verify that a CSV link was found
if (-not $csvLink) {
    Write-Error "No CSV link found."
    return
}

# Download the CSV file to a temporary location
$csvPath = "$env:TEMP\licensingplans.csv"
Invoke-WebRequest -Uri $csvLink -OutFile $csvPath

# Import license information from the CSV file
$results = Import-Csv -Path $csvPath

# List of SkuIds for licenses that are free and should be excluded
$excludedSkuIds = @(
    "f30db892-07e9-47e9-837c-80727f46fd3d",     # Microsoft Power Automate Free
    "5b631642-bd26-49fe-bd20-1daaa972ef80",     # Microsoft Power Apps for Developer
    "8f0c5670-4e56-4892-b06d-91c085d7004f",     # App Connect IW
    "6470687e-a428-4b7a-bef2-8a291ad947c9",     # Microsoft Store
    "093e8d14-a334-43d9-93e3-30589a8b47d0",     # Rights Management Service Basic Content Protection
    "e0dfc8b9-9531-4ec8-94b4-9fec23b05fc8",     # Microsoft Teams Exploratory Dept
    "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235",     # Microsoft Fabric (Free)
    "3f9f06f5-3c31-472c-985f-62d9c10ec167",     # Power Pages vTrial for Makers
    "8c4ce438-32a7-4ac5-91a6-e22ae08d9c8b"      # Rights Management Adhoc
)

# Retrieve all licenses in the tenant that are not free licenses
$GetAllLicenses = Get-MgSubscribedSku | Where-Object { $_.SkuId -notin $excludedSkuIds }

# Build a list of each license and its usage (assigned vs total)
$TenantLicenseCount = foreach ($License in $GetAllLicenses) {
    # Match SkuId with product name from CSV
    $matchedProducts = $results | Where-Object { $_.GUID -eq $License.SkuId }
    $UniqueNames = $matchedProducts.Product_Display_Name | Sort-Object -Unique
    $displayName = $UniqueNames -join ", "

    # Create an object with license information
    [PSCustomObject]@{
        Product  = $displayName
        Assigned = $License.ConsumedUnits
        Total    = $License.PrepaidUnits.Enabled
    }
}

# Retrieve all users and build a report with their licenses and information
# You can filter the users using Where-Object in different ways, for example:
# Where-Object { $_.JobTitle -eq "Support" }                    # Filter by exact job title
# Where-Object { $_.Department -eq "IT" }                       # Filter by department
# Where-Object { $_.UserPrincipalName -like "*@contoso.com" }   # Filter by domain
# Where-Object { $_.AccountEnabled -eq $true }                  # Only enabled accounts
# Where-Object { $_.City -eq "Stockholm" }                      # Filter by city
# Where-Object { $_.DisplayName -like "*Admin*" }               # Display name contains 'Admin'

$report = Get-MgUser -All | # Update Where-Object condition to use filter
Where-Object { $_.AssignedLicenses.City -ne "test" } |
Select-Object @{

    Name = 'Name'
    Expression = { $_.DisplayName }
}, @{

    Name = 'Licenses'
    Expression = {
        $licenses = Get-MgUserLicenseDetail -All -UserId $_.Id
        $names = foreach ($lic in $licenses) {
            # Skip licenses in the exclusion list
            if ($excludedSkuIds -contains $lic.SkuId) { continue }

            # Try to match the license with a product name
            $match = $results | Where-Object { $_.GUID -eq $lic.SkuId }
            if ($match) {
                $match.Product_Display_Name
            } else {
                "Unknown license ($($lic.SkuId))"
            }
        }
        # Return a newline-separated string with licenses
        ($names | Sort-Object -Unique) -join "`n"
    }
}, @{

    Name = 'Mail'
    Expression = { $_.Mail }
}, @{

    Name = 'Department'
    Expression = { $_.Department }
}, @{

    Name = 'Job Title'
    Expression = { $_.JobTitle }
}, @{

    Name = 'Office'
    Expression = { $_.OfficeLocation }
}, @{

    Name = 'City'
    Expression = { $_.City }
}

# Ask the user to provide a path to save the report files
$SaveFile = Read-Host "Enter path where the files should be saved"

# Retrieve the tenant name to use as part of the filename
$TenantName = (Get-MgOrganization).Displayname

# Create filenames for license and user report
$LicensesFileName = "$TenantName-Licenses.csv"
$UserFileName = "$TenantName-Users.csv"

# Combine path and filename in a cross-platform compatible way
$TenantFullPath = Join-Path -Path $SaveFile -ChildPath $LicensesFileName
$UserFullPath = Join-Path -Path $SaveFile -ChildPath $UserFileName

# Export the reports as CSV with UTF8 with BOM for correct åäö display in Excel
$TenantLicenseCount | Export-Csv -Path $TenantFullPath -Encoding utf8BOM -NoTypeInformation
$report | Export-Csv -Path $UserFullPath -Encoding utf8BOM -NoTypeInformation

# Create an Excel COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Select which files to convert
$csvFiles = @($TenantFullPath, $UserFullPath)

foreach ($csvFile in $csvFiles) {
    # Open CSV in Excel
    $workbook = $excel.Workbooks.Open($csvFile)
    $worksheet = $workbook.Sheets.Item(1)

    # Calculate last row and column to define table size
    $lastRow = $worksheet.UsedRange.Rows.Count
    $lastCol = $worksheet.UsedRange.Columns.Count
    $range = $worksheet.Range("A1").Resize($lastRow, $lastCol)

    # Add a table to the range
    $table = $worksheet.ListObjects.Add(1, $range, $null, 1)
    $table.Name = "DataTable"
    $table.TableStyle = "TableStyleMedium2"  # Change to another style if desired

    # Change file extension to .xlsx
    $xlsxPath = [System.IO.Path]::ChangeExtension($csvFile, ".xlsx")

    # Autofit columns and rows
    $worksheet.Columns.AutoFit()
    $worksheet.Rows.AutoFit()

    # Filter out empty cells
    $columnIndex = 2  # Change this to your target column (1 = A, 2 = B, etc.)

    # Enable autofilter if not already active
    if (-not $worksheet.AutoFilter) {
        $table.Range.AutoFilter()
    }

    # Set filter: show only rows where the column is NOT empty
    # xlFilterValues = 7
    # [TextFilterCriteria1] := "<>" means "not empty"
    $table.Range.AutoFilter($columnIndex, "<>")

    # Save as xlsx
    $workbook.SaveAs($xlsxPath, 51)  # 51 = xlOpenXMLWorkbook (xlsx)
    $workbook.Close($false)
}

# Quit Excel
$excel.Quit()

# Remove CSV files
foreach ($csv in $csvFiles){
    Remove-Item -Path $csv -Force
}

Disconnect-MgGraph

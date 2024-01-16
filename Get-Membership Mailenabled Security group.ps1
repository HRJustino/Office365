# Install and import the MSOnline module if not already installed
# Install-Module -Name MSOnline -Force -AllowClobber
# Import-Module MSOnline

# Connect to Exchange Online (you may need to install and import the ExchangeOnlineManagement module)
Connect-ExchangeOnline 

# Specify your Mail-Enabled Security Group's Display Name or Email Address
$groupName = 'Insert Group Name'

# Get group members
$groupMembers = Get-DistributionGroupMember -Identity $groupName

# Create an array to store the results
$report = @()

# Iterate through each group member
foreach ($member in $groupMembers) {
    # Retrieve information directly based on the member's email address
    $memberObject = Get-Recipient -Identity $member.PrimarySmtpAddress

    # If a valid member object is found, create a custom object with desired information
    if ($memberObject) {
        $reportObject = [PSCustomObject]@{
            GroupName   = $groupName
            UserName    = $memberObject.DisplayName
            Email       = $memberObject.PrimarySmtpAddress
        }

        # Add the object to the report array
        $report += $reportObject
    } else {
        # If not a valid member object, handle it accordingly (skip, log, etc.)
        Write-Host "Skipping member: $($member.PrimarySmtpAddress)"
    }
}

# Export the report to a CSV file
$report | Export-Csv -Path "C:\root\temp.csv" -NoTypeInformation

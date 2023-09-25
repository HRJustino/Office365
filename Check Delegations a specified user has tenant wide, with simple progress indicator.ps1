# Prompt the user for their choice of action
Clear-Host
Write-Host "Check Delegated access a specific user has access to:"
Write-Host "1. Get Mailbox Information"
Write-Host "2. Exit"
$actionChoice = Read-Host

# Function to get mailbox information
function Get-MailboxInformation {
    Set-ExecutionPolicy RemoteSigned
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline

    # Define the user you want to check
    $specifiedUser = Read-Host -Prompt "Enter the email address of the user to check"

    Write-Host "Fetching mailbox information. Please wait..."
    # Initialize an array to store the results
    $results = @()

    # Get all mailboxes (both user and shared) in your tenant
    $mailboxes = Get-Mailbox -ResultSize Unlimited 

    # Initialize a counter for the progress
    $progress = 0
    $totalMailboxes = $mailboxes.Count

    # Initialize a boolean flag to check if any delegated access was found
    $delegatedAccessFound = $false

    # Iterate through each mailbox
    foreach ($mailbox in $mailboxes) {
        $mailboxDisplayName = $mailbox.DisplayName

        try {
            $mailboxPermissions = Get-MailboxPermission -Identity $mailbox.DistinguishedName -ErrorAction Stop
        } catch {
            Write-Host "Error fetching permissions for mailbox: $mailboxDisplayName"
            Write-Host "Error details: $_"
            continue  # Skip this mailbox and move to the next
        }

        # Check if the specified user has been delegated access to the mailbox
        $delegatedAccess = $mailboxPermissions | Where-Object { $_.User -eq $specifiedUser }

        if ($delegatedAccess) {
            $result = [PSCustomObject]@{
                "Mailbox" = $mailboxDisplayName
                "DelegatedUser" = $specifiedUser
            }
            $results += $result

            # Set the flag to indicate that delegated access was found
            $delegatedAccessFound = $true
        }

        # Update the progress and display a percentage with 1 decimal place
        $progress++
        $percentage = [math]::Round(($progress / $totalMailboxes) * 100, 1)
        Write-Host "Fetching Mailboxes: $percentage%"
    }

    # If no delegated access was found, display a message
    if (-not $delegatedAccessFound) {
        Write-Host "No delegated access found for user: $specifiedUser"
    } else {
        # Display results in a GridView
        $results | Out-GridView
    }
}

# Based on the user's choice, either get mailbox information or manage mailbox permissions
if ($actionChoice -eq '1') {
    Get-MailboxInformation
} elseif ($actionChoice -eq '2') {
    exit
} else {
    Write-Host "Invalid choice. Please select a valid action (1 or 2)."
}

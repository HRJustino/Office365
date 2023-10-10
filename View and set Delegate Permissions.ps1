# Function to view delegation permissions on a specified user
function View-Permission {
    param (
        [string]$EmailAddress
    )

    # Purpose: View delegation permissions on a specified user's mailbox
    
    Write-Host "Viewing Permissions for $EmailAddress"
    
    # Get FullAccess permissions
    $fullAccessPermissions = Get-MailboxPermission -Identity $EmailAddress | Where-Object { $_.AccessRights -eq "FullAccess" }

    # Get Send As permissions
    $sendAsPermissions = Get-RecipientPermission -Identity $EmailAddress | Where-Object { $_.AccessRights -eq "SendAs" }

    if ($fullAccessPermissions.Count -gt 0) {
        Write-Host "Full Access Permissions:"
        $fullAccessPermissions | ForEach-Object {
            Write-Host "User/Group: $($_.User)"
        }
    } else {
        Write-Host "No Full Access Permissions found for this mailbox."
    }

    if ($sendAsPermissions.Count -gt 0) {
        Write-Host "Send As Permissions:"
        $sendAsPermissions | ForEach-Object {
            Write-Host "User/Group: $($_.Trustee)"
        }
    } else {
        Write-Host "No Send As Permissions found for this mailbox."
    }

    # Prompt the user to press Enter to return to the main menu
    Read-Host "Press Enter to return to the main menu..."

    # Clear the screen
    Clear-Host
}

# Function to add user/group permissions to a shared mailbox
function Add-Permission {
    param (
        [string]$EmailAddress
    )

    # Purpose: Add user/group permissions (Full Access and Send As) to a shared mailbox

    Write-Host "Adding Permissions to $EmailAddress"
    $UserOrGroup = Read-Host "Enter the User/Group to add with Full Access and Send As Permissions, Press CTRL + C to break"
    
    # Add FullAccess permission
    Add-MailboxPermission -Identity $EmailAddress -User $UserOrGroup -AccessRights FullAccess

    # Add Send As permission
    Add-RecipientPermission -Identity $EmailAddress -Trustee $UserOrGroup -AccessRights SendAs

    Write-Host "Permissions added successfully."
    Start-Sleep -Seconds 4
}

# Function to remove user/group permissions from a shared mailbox
function Remove-Permission {
    param (
        [string]$EmailAddress
    )

    # Purpose: Remove user/group permissions (Full Access and Send As) from a shared mailbox

    Write-Host "Removing Permissions from $EmailAddress"
    $UserOrGroup = Read-Host "Enter the User/Group to remove from Full Access and Send As Permissions, Press CTRL + C to break"
    
    # Remove FullAccess permission
    Remove-MailboxPermission -Identity $EmailAddress -User $UserOrGroup -AccessRights FullAccess -Confirm:$false

    # Remove Send As permission
    Remove-RecipientPermission -Identity $EmailAddress -Trustee $UserOrGroup -AccessRights SendAs -Confirm:$false

    Write-Host "Permissions removed successfully."
    Start-Sleep -Seconds 4
}

#
# Function to validate email address
function Validate-EmailAddress($email) {
    # Purpose: Validate if an email address is in a valid format
    $emailPattern = "^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$"
    return $email -match $emailPattern
}

# Main menu for shared mailbox permissions management
while ($true) {
    Clear-Host
    Write-Host "Shared Mailbox Permissions Management, this script is used to view and manage delegate settings"
    Write-Host "1. View Permissions"
    Write-Host "2. Add Permissions"
    Write-Host "3. Remove Permissions"
    Write-Host "4. Exit"
    $choice = Read-Host "Enter your choice"
   
    switch ($choice) {
        1 {
        Set-ExecutionPolicy RemoteSigned
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline
            $EmailAddress = Read-Host "Enter Email Address of the shared mailbox you want to view permissions for, Press CTRL + C to break"
            if (Validate-EmailAddress $EmailAddress) {
                View-Permission -EmailAddress $EmailAddress
            } else {
                Write-Host "Invalid email address. Please enter a valid email address."
                Start-Sleep -Seconds 2
            }
        }
        2 {
        Set-ExecutionPolicy RemoteSigned
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline
            $EmailAddress = Read-Host "Enter Email Address of the shared mailbox you want to manage permissions for, Press CTRL + C to break"
            if (Validate-EmailAddress $EmailAddress) {
                Add-Permission -EmailAddress $EmailAddress
            } else {
                Write-Host "Invalid email address. Please enter a valid email address."
                Start-Sleep -Seconds 2
            }
        }
        3 {
        Set-ExecutionPolicy RemoteSigned
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline
            $EmailAddress = Read-Host "Enter Email Address of the shared mailbox you want to manage permissions for, Press CTRL + C to break"
            if (Validate-EmailAddress $EmailAddress) {
                Remove-Permission -EmailAddress $EmailAddress
            } else {
                Write-Host "Invalid email address. Please enter a valid email address."
                Start-Sleep -Seconds 2
            }
        }
        4 {
            # Exit
            exit
        }
        default {
            Write-Host "Invalid choice. Please select a valid option."
            Start-Sleep -Seconds 2
        }
    }
}

#################################################
# Global Functions (Setup, Connect, and Prompts)
#################################################

function Setup-Environment {
    # There are 5 main steps; each step is divided into 2 sub-steps (reading/checking and applying/finalizing)
    $totalMainSteps = 5
    $subStepIncrement = (100 / $totalMainSteps) / 2

    # --- Step 1: Execution Policy ---
    $progress = 0
    Write-Progress -Activity "Setup Environment" -Status "Step 1/5: Reading current Execution Policy" -PercentComplete $progress
    try {
        $policy = Get-ExecutionPolicy
    } catch {
        Write-Progress -Activity "Setup Environment" -Status "Step 1/5: Failed to get Execution Policy" -PercentComplete $progress -Completed
        Write-Host "Error: Unable to read Execution Policy." -ForegroundColor Red
        return
    }
    $progress += $subStepIncrement
    Write-Progress -Activity "Setup Environment" -Status "Step 1/5: Comparing and applying Execution Policy fix" -PercentComplete $progress
    if ($policy -ne "RemoteSigned") {
        try {
            Set-ExecutionPolicy RemoteSigned -Force
            Write-Host "Execution Policy set to RemoteSigned."
        } catch {
            Write-Host "Error: Failed to set Execution Policy." -ForegroundColor Red
        }
    } else {
        Write-Host "Execution Policy already set to RemoteSigned."
    }
    $progress = 100 / $totalMainSteps  # End of step 1
    Write-Progress -Activity "Setup Environment" -Status "Step 1 completed" -PercentComplete $progress

    # --- Step 2: ExchangeOnlineManagement Module ---
    $step = 2
    $progress = ((($step - 1) * 100) / $totalMainSteps)
    Write-Progress -Activity "Setup Environment" -Status "Step 2/5: Searching for ExchangeOnlineManagement module" -PercentComplete $progress
    try {
        $moduleFound = Get-Module -ListAvailable -Name ExchangeOnlineManagement
    } catch {
        Write-Progress -Activity "Setup Environment" -Status "Step 2/5: Failed to search for module" -PercentComplete $progress -Completed
        Write-Host "Error: Unable to search for ExchangeOnlineManagement module." -ForegroundColor Red
        return
    }
    $progress += $subStepIncrement
    Write-Progress -Activity "Setup Environment" -Status "Step 2/5: Installing module if missing" -PercentComplete $progress
    if (-not $moduleFound) {
        try {
            Install-Module -Name ExchangeOnlineManagement -RequiredVersion 3.6.0 -Force
            Write-Host "ExchangeOnlineManagement module installed."
        } catch {
            Write-Host "Error: Failed to install ExchangeOnlineManagement module." -ForegroundColor Red
        }
    } else {
        Write-Host "ExchangeOnlineManagement module already installed."
    }
    $progress = (2 * 100 / $totalMainSteps)
    Write-Progress -Activity "Setup Environment" -Status "Step 2 completed" -PercentComplete $progress

    # --- Step 3: PowerShellGet Module ---
    $step = 3
    $progress = ((($step - 1) * 100) / $totalMainSteps)
    Write-Progress -Activity "Setup Environment" -Status "Step 3/5: Searching for PowerShellGet module" -PercentComplete $progress
    try {
        $psGetFound = Get-Module -ListAvailable -Name PowerShellGet
    } catch {
        Write-Progress -Activity "Setup Environment" -Status "Step 3/5: Failed to search for module" -PercentComplete $progress -Completed
        Write-Host "Error: Unable to search for PowerShellGet module." -ForegroundColor Red
        return
    }
    $progress += $subStepIncrement
    Write-Progress -Activity "Setup Environment" -Status "Step 3/5: Installing module if missing" -PercentComplete $progress
    if (-not $psGetFound) {
        try {
            Install-Module -Name PowerShellGet -Force -AllowClobber
            Write-Host "PowerShellGet module installed."
        } catch {
            Write-Host "Error: Failed to install PowerShellGet module." -ForegroundColor Red
        }
    } else {
        Write-Host "PowerShellGet module already installed."
    }
    $progress = (3 * 100 / $totalMainSteps)
    Write-Progress -Activity "Setup Environment" -Status "Step 3 completed" -PercentComplete $progress

    # --- Step 4: Office365 Connectivity Test ---
    $step = 4
    $progress = ((($step - 1) * 100) / $totalMainSteps)
    Write-Progress -Activity "Setup Environment" -Status "Step 4/5: Testing Office365 connectivity" -PercentComplete $progress
    try {
        Invoke-WebRequest -Uri "https://outlook.office365.com" -UseBasicParsing -ErrorAction Stop | Out-Null
        Write-Host "Office365 connection test passed."
    } catch {
        Write-Progress -Activity "Setup Environment" -Status "Step 4/5: Office365 connectivity test FAILED" -PercentComplete $progress -Completed
        Write-Host "Office365 connection test failed. Check your network." -ForegroundColor Red
        return
    }
    $progress += $subStepIncrement
    $progress = (4 * 100 / $totalMainSteps)
    Write-Progress -Activity "Setup Environment" -Status "Step 4 completed" -PercentComplete $progress

    # --- Step 5: Enable TLS12 ---
    $step = 5
    $progress = ((($step - 1) * 100) / $totalMainSteps)
    Write-Progress -Activity "Setup Environment" -Status "Step 5/5: Enabling TLS12" -PercentComplete $progress
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Write-Host "TLS12 enabled."
    } catch {
        Write-Progress -Activity "Setup Environment" -Status "Step 5/5: Failed to enable TLS12" -PercentComplete $progress -Completed
        Write-Host "Error: Unable to enable TLS12." -ForegroundColor Red
        return
    }
    $progress = 100
    Write-Progress -Activity "Setup Environment" -Status "Setup Environment completed" -PercentComplete $progress -Completed

    Pause
}

function Connect-Exchange {
    $progress = 0
    Write-Progress -Activity "Connect to Exchange Online" -Status "Initializing account selection..." -PercentComplete $progress
    Write-Host "Select the admin account to connect:" 
    Write-Host "1. domain_admin@domain.com"
    Write-Host "2. domain_admin@domain.com"
    Write-Host "3. domain_admin@domain.com"
    Write-Host "4. domain_admin@domain.com"
    Write-Host "5. domain_admin@domain.com"
    Write-Host "6. domain_admin@domain.com"
    $adminChoice = Read-Host "Enter your choice (1-6)"
    $progress += 30
    Write-Progress -Activity "Connect to Exchange Online" -Status "Connecting to chosen account..." -PercentComplete $progress
    switch ($adminChoice) {
        "1" { Connect-ExchangeOnline -UserPrincipalName domain_admin@domain.com }
        "2" { Connect-ExchangeOnline -UserPrincipalName domain_admin@domain.com }
        "3" { Connect-ExchangeOnline -UserPrincipalName domain_admin@domain.com }
        "4" { Connect-ExchangeOnline -UserPrincipalName domain_admin@domain.com }
        "5" { Connect-ExchangeOnline -UserPrincipalName domain_admin@domain.com }
        "6" { Connect-ExchangeOnline -UserPrincipalName domain_admin@domain.com }
        default { 
            Write-Host "Invalid selection, please try again." -ForegroundColor Yellow
            Connect-Exchange  # Restart if input is invalid
            return
        }
    }
    $progress = 100
    Write-Progress -Activity "Connect to Exchange Online" -Status "Connected" -PercentComplete $progress -Completed
    Write-Host "Connected to Exchange Online."
    Pause
}


function Connect-Graph {
    $progress = 0
    Write-Progress -Activity "Connect to Microsoft Graph" -Status "Initializing..." -PercentComplete $progress
    Write-Host "Connecting to Microsoft Graph..."
    try {
        Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All"
        $progress = 100
        Write-Progress -Activity "Connect to Microsoft Graph" -Status "Connected" -PercentComplete $progress -Completed
        Write-Host "Connected to Microsoft Graph."
    } catch {
        Write-Progress -Activity "Connect to Microsoft Graph" -Status "Failed" -PercentComplete $progress -Completed
        Write-Host "Failed to connect to Microsoft Graph." -ForegroundColor Red
    }
    Pause
}

# This function combines setup, Exchange, and Graph connection.
function Setup-And-Connect {
    Setup-Environment
    Connect-Exchange
    $graphPrompt = Read-Host "Would you like to connect to Microsoft Graph? (Y/N)"
    if ($graphPrompt -match "^[Yy]") {
        Connect-Graph
    } else {
        Write-Host "Skipping Microsoft Graph connection." -ForegroundColor Yellow
        Pause
    }
}

#################################################
# 365_View_All Section (View All 365 Details)
#################################################

# Helper function for outputting data
function Output-Data {
    param (
        [array]$Data,
        [string]$Title
    )
    $Data = $Data | Sort-Object DisplayName
    Write-Host "Select output method:"
    Write-Host "1. PowerShell Window"
    Write-Host "2. Grid View"
    Write-Host "3. Export to CSV (Downloads Folder)"
    $choice = Read-Host "Enter your choice"
    switch ($choice) {
        "1" { $Data | Format-Table -AutoSize }
        "2" { $Data | Out-GridView -Title $Title }
        "3" { 
            $path = "$env:USERPROFILE\Downloads\$($Title -replace ' ', '_').csv"
            $Data | Export-Csv -Path $path -NoTypeInformation
            Write-Host "Data exported to $path" -ForegroundColor Green
        }
        default { 
            Write-Host "Invalid choice. Showing in PowerShell window." 
            $Data | Format-Table -AutoSize 
        }
    }
    Write-Host "`nPress Enter to continue..." -ForegroundColor Cyan
    Read-Host
}

function Show-ViewAll365 {

    function View-AllMailboxes {
        # Fetch mailboxes and update progress per mailbox
        Write-Progress -Activity "View All Mailboxes" -Status "Fetching mailboxes..." -PercentComplete 0
        $mailboxes = Get-Mailbox | Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails
        $total = $mailboxes.Count
        $counter = 0
        foreach ($item in $mailboxes) {
            $counter++
            $percent = [math]::Round(($counter / $total * 100), 0)
            Write-Progress -Activity "View All Mailboxes" -Status "Processing mailbox $counter of $total" -PercentComplete $percent
        }
        Write-Progress -Activity "View All Mailboxes" -Status "Completed" -PercentComplete 100 -Completed
        Output-Data -Data $mailboxes -Title "All Mailboxes"
    }

    function View-UserMailboxes {
        Write-Progress -Activity "View User Mailboxes" -Status "Fetching user mailboxes..." -PercentComplete 0
        $mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox | Select-Object DisplayName, PrimarySmtpAddress
        $total = $mailboxes.Count
        $counter = 0
        foreach ($item in $mailboxes) {
            $counter++
            $percent = [math]::Round(($counter / $total * 100), 0)
            Write-Progress -Activity "View User Mailboxes" -Status "Processing mailbox $counter of $total" -PercentComplete $percent
        }
        Write-Progress -Activity "View User Mailboxes" -Status "Completed" -PercentComplete 100 -Completed
        Output-Data -Data $mailboxes -Title "User Mailboxes"
    }

    function View-SharedMailboxes {
        Write-Progress -Activity "View Shared Mailboxes" -Status "Fetching shared mailboxes..." -PercentComplete 0
        $mailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox | Select-Object DisplayName, PrimarySmtpAddress
        $total = $mailboxes.Count
        $counter = 0
        foreach ($item in $mailboxes) {
            $counter++
            $percent = [math]::Round(($counter / $total * 100), 0)
            Write-Progress -Activity "View Shared Mailboxes" -Status "Processing mailbox $counter of $total" -PercentComplete $percent
        }
        Write-Progress -Activity "View Shared Mailboxes" -Status "Completed" -PercentComplete 100 -Completed
        Output-Data -Data $mailboxes -Title "Shared Mailboxes"
    }

    function View-AllEntraUsers {
        Write-Progress -Activity "View All Entra Users" -Status "Fetching Entra users..." -PercentComplete 0
        $users = Get-MgUser -All | Select-Object DisplayName, UserPrincipalName, Mail
        $total = $users.Count
        $counter = 0
        foreach ($user in $users) {
            $counter++
            $percent = [math]::Round(($counter / $total * 100), 0)
            Write-Progress -Activity "View All Entra Users" -Status "Processing user $counter of $total" -PercentComplete $percent
        }
        Write-Progress -Activity "View All Entra Users" -Status "Completed" -PercentComplete 100 -Completed
        Output-Data -Data $users -Title "All Entra Users"
    }

    function View-LicensedEntraUsers {
        Write-Progress -Activity "View Licensed Entra Users" -Status "Fetching Entra users..." -PercentComplete 0
        $users = Get-MgUser -All | Select-Object DisplayName, UserPrincipalName, Mail
        $licensedUsers = @()
        $total = $users.Count
        $counter = 0
        foreach ($user in $users) {
            $licenses = Get-MgUserLicenseDetail -UserId $user.UserPrincipalName -ErrorAction SilentlyContinue
            if ($licenses) { $licensedUsers += $user }
            $counter++
            $percent = [math]::Round(($counter / $total * 100), 0)
            Write-Progress -Activity "View Licensed Entra Users" -Status "Filtering licensed users ($counter of $total)" -PercentComplete $percent
        }
        Write-Host "Total Licensed Users: $($licensedUsers.Count)" -ForegroundColor Green
        Write-Progress -Activity "View Licensed Entra Users" -Status "Completed" -PercentComplete 100 -Completed
        Output-Data -Data $licensedUsers -Title "Licensed Entra Users"
    }

    function View-UnlicensedEntraUsers {
        Write-Progress -Activity "View Unlicensed Entra Users" -Status "Fetching Entra users..." -PercentComplete 0
        $users = Get-MgUser -All | Select-Object DisplayName, UserPrincipalName, Mail
        $unlicensedUsers = @()
        $total = $users.Count
        $counter = 0
        foreach ($user in $users) {
            $licenses = Get-MgUserLicenseDetail -UserId $user.UserPrincipalName -ErrorAction SilentlyContinue
            if (-not $licenses) { $unlicensedUsers += $user }
            $counter++
            $percent = [math]::Round(($counter / $total * 100), 0)
            Write-Progress -Activity "View Unlicensed Entra Users" -Status "Filtering unlicensed users ($counter of $total)" -PercentComplete $percent
        }
        Write-Host "Total Unlicensed Users: $($unlicensedUsers.Count)" -ForegroundColor Green
        Write-Progress -Activity "View Unlicensed Entra Users" -Status "Completed" -PercentComplete 100 -Completed
        Output-Data -Data $unlicensedUsers -Title "Unlicensed Entra Users"
    }

    function View-ExternalEntraUsers {
        Write-Progress -Activity "View External Entra Users" -Status "Fetching external Entra users..." -PercentComplete 0
        $internalDomains = @("frese.dk", "frese.eu", "frese.co.uk", "fresevalves.com", "fresemetal.dk", "de-valves.dk")
        $users = Get-MgUser -All | Where-Object { $_.UserType -eq "Guest" -or ($internalDomains -notcontains ($_.UserPrincipalName -split "@")[1]) } | Select-Object DisplayName, UserPrincipalName, Mail
        $total = $users.Count
        $counter = 0
        foreach ($user in $users) {
            $counter++
            $percent = [math]::Round(($counter / $total * 100), 0)
            Write-Progress -Activity "View External Entra Users" -Status "Processing user $counter of $total" -PercentComplete $percent
        }
        Write-Host "Total Guest Users: $($users.Count)" -ForegroundColor Green
        Write-Progress -Activity "View External Entra Users" -Status "Completed" -PercentComplete 100 -Completed
        Output-Data -Data $users -Title "External Guest Users"
    }

    function View-InactiveUsers {
        Write-Progress -Activity "View Inactive Users" -Status "Initializing..." -PercentComplete 0
        $daysThreshold = Read-Host "Enter the number of days to check for inactive users"
        $cutoffDate = (Get-Date).AddDays(-[int]$daysThreshold)
        Write-Progress -Activity "View Inactive Users" -Status "Fetching users..." -PercentComplete 5
        $users = Get-MgUser -All -Property SignInActivity | Select-Object DisplayName, UserPrincipalName, Mail, SignInActivity
        $inactiveUsers = @()
        $total = $users.Count
        $counter = 0
        foreach ($user in $users) {
            if ($user.SignInActivity.LastSignInDateTime -lt $cutoffDate -or $user.SignInActivity.LastSignInDateTime -eq $null) {
                $inactiveUsers += $user
            }
            $counter++
            $percent = [math]::Round(($counter / $total * 100), 0)
            Write-Progress -Activity "View Inactive Users" -Status "Filtering inactive users ($counter of $total)" -PercentComplete $percent
        }
        Write-Host "Total Inactive Users Found: $($inactiveUsers.Count)" -ForegroundColor Green
        Write-Progress -Activity "View Inactive Users" -Status "Completed" -PercentComplete 100 -Completed
        Output-Data -Data $inactiveUsers -Title "Inactive Users (Last login older than $daysThreshold days)"
    }

    Clear-Host
    Write-Host "Microsoft 365 Management - View All 365 Details" -ForegroundColor Cyan
    Write-Host "1. View All Mailboxes"
    Write-Host "2. View All User Mailboxes"
    Write-Host "3. View All Shared Mailboxes"
    Write-Host "4. View All Entra Users"
    Write-Host "5. View Licensed Entra Users"
    Write-Host "6. View Unlicensed Entra Users"
    Write-Host "7. View External Guest Users"
    Write-Host "8. View Inactive Users"
    Write-Host "9. Return to Main Menu"
    $subChoice = Read-Host "Enter your choice"
    switch ($subChoice) {
        "1" { View-AllMailboxes }
        "2" { View-UserMailboxes }
        "3" { View-SharedMailboxes }
        "4" { View-AllEntraUsers }
        "5" { View-LicensedEntraUsers }
        "6" { View-UnlicensedEntraUsers }
        "7" { View-ExternalEntraUsers }
        "8" { View-InactiveUsers }
        "9" { return }
        default { Write-Host "Invalid choice. Returning to Main Menu." }
    }
    Pause
    return
}

#################################################
# 365_Exchange_Permissions Section (Manage Exchange Permissions)
#################################################

function Manage-ExchangePermissions {
    function Show-Menu-ExchangePermissions {
        Clear-Host
        Write-Host "============================="
        Write-Host " Office365 Exchange Permissions Menu"
        Write-Host "============================="
        Write-Host "1. Check Specific Mailbox Calendar Permissions"
        Write-Host "2. Set Specific Mailbox Calendar Permissions"
        Write-Host "3. Check Specific Mailbox Size Limits"
        Write-Host "4. Update Specific Mailbox Size Limits"
        Write-Host "5. Return to Main Menu"
        Write-Host "============================="
    }

    function Check-CalendarPermissions {
        Write-Progress -Activity "Check Calendar Permissions" -Status "Fetching calendar folder info..." -PercentComplete 0
        $mailbox = Read-Host "Enter the mailbox to check (e.g., user@domain.com)"
        $calendarFolder = (Get-MailboxFolderStatistics -Identity $mailbox | Where-Object { $_.FolderType -eq "Calendar" }).Identity
        if ($calendarFolder) {
            $folderParts = $calendarFolder -split "\\"
            Write-Progress -Activity "Check Calendar Permissions" -Status "Processing folder parts..." -PercentComplete 30
            if ($folderParts.Count -ge 2) {
                $folderName = $folderParts[1]
                $correctedIdentity = "$($mailbox):\$($folderName)"
                $permissions = Get-MailboxFolderPermission -Identity $correctedIdentity
                if ($permissions) {
                    $total = $permissions.Count
                    $counter = 0
                    foreach ($perm in $permissions) {
                        $counter++
                        $percent = [math]::Round(($counter / $total * 100), 0)
                        Write-Progress -Activity "Check Calendar Permissions" -Status "Processing permission $counter of $total" -PercentComplete $percent
                        Start-Sleep -Milliseconds 100
                    }
                }
                Write-Host "Checked calendar permissions for $mailbox."
            } else {
                Write-Host "Error: Unexpected calendar folder format for $mailbox." -ForegroundColor Red
            }
        } else {
            Write-Host "Error: Calendar folder not found for $mailbox." -ForegroundColor Red
        }
        Write-Progress -Activity "Check Calendar Permissions" -Status "Completed" -PercentComplete 100 -Completed
        Pause
    }

    function Set-CalendarPermissions {
        Write-Progress -Activity "Set Calendar Permissions" -Status "Fetching calendar folder info..." -PercentComplete 0
        $mailbox = Read-Host "Enter the user's email address (e.g., user@domain.com)"
        $calendarFolder = (Get-MailboxFolderStatistics -Identity $mailbox | Where-Object { $_.FolderType -eq "Calendar" }).Identity
        if ($calendarFolder) {
            $folderParts = $calendarFolder -split "\\"
            Write-Progress -Activity "Set Calendar Permissions" -Status "Processing folder parts..." -PercentComplete 30
            if ($folderParts.Count -ge 2) {
                $folderName = $folderParts[1]
                $correctedIdentity = "$($mailbox):\$($folderName)"
                # Simulate incremental progress during the update
                for ($i = 1; $i -le 5; $i++) {
                    $percent = $i * 20
                    Write-Progress -Activity "Set Calendar Permissions" -Status "Applying permissions... ($i/5)" -PercentComplete $percent
                    Start-Sleep -Milliseconds 150
                }
                Set-MailboxFolderPermission -Identity $correctedIdentity -User Default -AccessRights LimitedDetails
                Write-Host "Calendar permissions for $mailbox set to LimitedDetails."
            } else {
                Write-Host "Error: Unexpected calendar folder format for $mailbox." -ForegroundColor Red
            }
        } else {
            Write-Host "Error: Calendar folder not found for $mailbox."
        }
        Write-Progress -Activity "Set Calendar Permissions" -Status "Completed" -PercentComplete 100 -Completed
        Pause
    }

    function Check-MailboxSizeLimits {
        Write-Progress -Activity "Check Mailbox Size Limits" -Status "Fetching mailbox size limits..." -PercentComplete 0
        $mailbox = Read-Host "Enter the mailbox to check size limits for (e.g., user@domain.com)"
        $limits = Get-Mailbox -Identity $mailbox | Select-Object DisplayName, UserPrincipalName, MaxSendSize, MaxReceiveSize
        # If more than one property is returned, simulate processing per property
        $props = $limits | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
        $total = $props.Count
        $counter = 0
        foreach ($prop in $props) {
            $counter++
            $percent = [math]::Round(($counter / $total * 100), 0)
            Write-Progress -Activity "Check Mailbox Size Limits" -Status "Processing property $prop ($counter of $total)" -PercentComplete $percent
            Start-Sleep -Milliseconds 100
        }
        $limits | Format-Table -AutoSize
        Write-Progress -Activity "Check Mailbox Size Limits" -Status "Completed" -PercentComplete 100 -Completed
        Pause
    }

    function Update-MailboxSize {
        Write-Progress -Activity "Update Mailbox Size" -Status "Updating mailbox size limits..." -PercentComplete 0
        $mailbox = Read-Host "Enter the mailbox to update size limits for (e.g., user@domain.com)"
        $newSize = 50MB
        # Simulate a short progress loop
        for ($i = 1; $i -le 5; $i++) {
            $percent = $i * 20
            Write-Progress -Activity "Update Mailbox Size" -Status "Applying update... ($i/5)" -PercentComplete $percent
            Start-Sleep -Milliseconds 150
        }
        Set-Mailbox -Identity $mailbox -MaxReceiveSize $newSize -MaxSendSize $newSize
        Write-Host "Updated size limits for $mailbox to $newSize."
        Write-Progress -Activity "Update Mailbox Size" -Status "Completed" -PercentComplete 100 -Completed
        Pause
    }

    Show-Menu-ExchangePermissions
    $expChoice = Read-Host "Enter your choice (1-5)"
    switch ($expChoice) {
        "1" { Check-CalendarPermissions }
        "2" { Set-CalendarPermissions }
        "3" { Check-MailboxSizeLimits }
        "4" { Update-MailboxSize }
        "5" { return }
        default { Write-Host "Invalid selection. Returning to Main Menu." }
    }
    return
}

#################################################
# 365_Check_Delegated_Mailbox_Access Section
#################################################

function Check-DelegatedMailboxAccess {
    function Check-DelegatedMailboxAccess-Core {
        Write-Progress -Activity "Delegated Mailbox Access Check" -Status "Initializing..." -PercentComplete 0
        $specifiedUser = Read-Host "Enter the user to check delegated mailbox access (e.g., user@domain.com)"
        $results = @()
        $mailboxes = Get-Mailbox -ResultSize Unlimited
        $totalMailboxes = $mailboxes.Count
        $processedMailboxes = 0
        foreach ($mailbox in $mailboxes) {
            $mailboxPermissions = Get-MailboxPermission -Identity $mailbox.DistinguishedName
            $delegatedAccess = $mailboxPermissions | Where-Object { $_.User -eq $specifiedUser }
            if ($delegatedAccess) {
                $result = [PSCustomObject]@{
                    "Mailbox"       = $mailbox.DisplayName
                    "DelegatedUser" = $specifiedUser
                }
                $results += $result
            }
            $processedMailboxes++
            $progressPercent = [math]::Round(($processedMailboxes / $totalMailboxes * 100), 0)
            Write-Progress -Activity "Delegated Mailbox Access Check" -Status "Checking mailbox permissions ($processedMailboxes of $totalMailboxes)" -PercentComplete $progressPercent
        }
        Write-Progress -Activity "Delegated Mailbox Access Check" -Status "Completed" -PercentComplete 100 -Completed
        Write-Host "Select output option:"
        Write-Host "1: View results in PowerShell console"
        Write-Host "2: Display results in GridView"
        Write-Host "3: Export results to CSV"
        $outputChoice = Read-Host "Enter your choice (1-3)"
        switch ($outputChoice) {
            "1" { $results | Format-Table -AutoSize }
            "2" { $results | Out-GridView }
            "3" { 
                $csvPath = Read-Host "Enter the full file path for CSV export (e.g., C:\Path\DelegatedAccessResults.csv)"
                $results | Export-Csv -Path $csvPath -NoTypeInformation
                Write-Host "Results exported to $csvPath."
            }
            default { 
                Write-Host "Invalid selection. Defaulting to console output."
                $results | Format-Table -AutoSize
            }
        }
        Pause
    }

    Clear-Host
    Write-Host "============================="
    Write-Host " Delegated Mailbox Access Check"
    Write-Host "============================="
    Write-Host "1. Check Delegated Mailbox Access"
    Write-Host "2. Return to Main Menu"
    Write-Host "============================="
    $delegatedChoice = Read-Host "Enter your choice"
    switch ($delegatedChoice) {
        "1" { Check-DelegatedMailboxAccess-Core }
        "2" { return }
        default { Write-Host "Invalid selection. Returning to Main Menu." }
    }
    return
}

#################################################
# 365_Sharedmailbox_Sent_Folder_Fix Section
#################################################

function Fix-SharedMailboxSentFolder {
    function Fix-SentFolderBehavior {
        Write-Progress -Activity "Fix Shared Mailbox Sent Folder" -Status "Fetching mailbox info..." -PercentComplete 0
        $mailbox = Read-Host "Enter the shared mailbox to fix (e.g., user@domain.com)"
        Get-Mailbox -Identity $mailbox | 
            Select-Object DisplayName, MessageCopyForSentAsEnabled, MessageCopyForSendOnBehalfEnabled | 
            Format-Table -AutoSize
        # Simulate progress during fix application
        for ($i = 1; $i -le 5; $i++) {
            $percent = $i * 20
            Write-Progress -Activity "Fix Shared Mailbox Sent Folder" -Status "Applying fix... ($i/5)" -PercentComplete $percent
            Start-Sleep -Milliseconds 150
        }
        Set-Mailbox -Identity $mailbox -MessageCopyForSentAsEnabled $true -MessageCopyForSendOnBehalfEnabled $true
        Write-Host "Fixed sent folder behavior for $mailbox. Properties have been set to True."
        Write-Progress -Activity "Fix Shared Mailbox Sent Folder" -Status "Completed" -PercentComplete 100 -Completed
        Pause
    }

    function Show-Menu-Sent {
        Clear-Host
        Write-Host "============================="
        Write-Host "   Shared Mailbox Sent Folder Fix"
        Write-Host "============================="
        Write-Host "1. Fix Sent Folder Behavior for Shared Mailbox"
        Write-Host "2. Return to Main Menu"
        Write-Host "============================="
    }

    Show-Menu-Sent
    $sentChoice = Read-Host "Enter your choice (1-2)"
    switch ($sentChoice) {
        "1" { Fix-SentFolderBehavior }
        "2" { return }
        default { Write-Host "Invalid selection. Returning to Main Menu." }
    }
    return
}

#################################################
# Main Menu
#################################################

do {
    Clear-Host
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "         365 Administration Menu" -ForegroundColor Cyan
    Write-Host "========================================="
    Write-Host "1. Setup Environment & Connect (Exchange & Graph)"
    Write-Host "2. View All 365 Details"
    Write-Host "3. Manage 365 Exchange Permissions"
    Write-Host "4. Check Delegated Mailbox Access"
    Write-Host "5. Fix Shared Mailbox Sent Folder Issues"
    Write-Host "6. Exit"
    Write-Host "========================================="
    $mainChoice = Read-Host "Enter your choice (1-6)"
    switch ($mainChoice) {
        "1" { 
            Write-Host "Running Setup, Exchange & Graph Connection..." -ForegroundColor Green
            Setup-And-Connect
        }
        "2" { 
            Write-Host "Running 365_View_All..." -ForegroundColor Green
            Show-ViewAll365
        }
        "3" { 
            Write-Host "Running 365_Exchange_Permissions..." -ForegroundColor Green
            Manage-ExchangePermissions
        }
        "4" { 
            Write-Host "Running 365_Check_Delegated_Mailbox_Access..." -ForegroundColor Green
            Check-DelegatedMailboxAccess
        }
        "5" { 
            Write-Host "Running 365_Sharedmailbox_Sent_Folder_Fix..." -ForegroundColor Green
            Fix-SharedMailboxSentFolder
        }
        "6" { 
            Write-Host "Exiting..." -ForegroundColor Red
            break
        }
        default { 
            Write-Host "Invalid selection. Please choose an option from 1 to 6." -ForegroundColor Yellow
        }
    }
    Pause
} while ($true)

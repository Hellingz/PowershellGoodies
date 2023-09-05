## Office 365 ##
## Author: Tobias Helling ##
## Date: 2023-09-05 ##

# Get all user mailbox memberships #

# Checks if module is installed

if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
    # Is not installed, installing it
    Write-Host 'ExchangeOnlineManagement module is not installed. Installing it...'
    Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser
}

# Importing module
Import-Module ExchangeOnlineManagement

# Connecting to Exchange Online Server
Connect-ExchangeOnline

# Ask whos permissions we are going to look for
$specificMemberEmail = Read-Host 'Enter user UPN'

# Store data in array
$results = @()

# Get list of all mailboxers
$mailboxes = Get-Mailbox -ResultSize Unlimited

# Loop through all shared mailboxes
foreach ($mailbox in $mailboxes) {
    # Get permission for mailbox
    $permissions = Get-MailboxPermission -Identity $mailbox.DistinguishedName

    # Loop through all permissions to find specific member
    foreach ($permission in $permissions) {
        if ($permission.User -eq $specificMemberEmail) {
            $result = New-Object PSObject -Property @{
                Identity = $mailbox.DistinguishedName
                User = $specificMemberEmail
                AccessRights = $permission.AccessRights
            }

            # Adds a row to our results array
            $results += $result

            Write-Host "Found shared mailbox: $($mailbox.DisplayName)" -ForegroundColor DarkGreen
        }
    }
}

# Get list of all distribution lists
$distributionGroups = Get-DistributionGroup -ResultSize Unlimited

# Loop through all distribution lists
foreach ($group in $distributionGroups) {
    # Get members from distribution list
    $members = Get-DistributionGroupMember -Identity $group.DistinguishedName

    # Loop through members to find specific member
    foreach ($member in $members) {
        if ($member.PrimarySmtpAddress -eq $specificMemberEmail) {
            $result = New-Object PSObject -Property @{
                Identity = $group.DistinguishedName
                User = $specificMemberEmail
                MemberType = "DistributionGroup"
            }

            # Adds a row to our results array
            $results += $result
            
            # Skriv ut en notering om att distributionsgruppen har kontrollerats
            Write-Host "Found distribution group: $($group.DisplayName)" -ForegroundColor DarkGreen
        }
    }
}

# Exportera resultaten till en CSV-fil
$results | Export-Csv -Path $home\Downloads\MailboxMembership.csv -NoTypeInformation
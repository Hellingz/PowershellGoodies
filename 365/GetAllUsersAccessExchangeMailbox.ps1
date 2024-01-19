if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
    # Is not installed, installing it
    Write-Host 'ExchangeOnlineManagement module is not installed. Installing it...'
    Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser
}

# Importing module
Import-Module ExchangeOnlineManagement

# Connecting to Exchange Online Server
Connect-ExchangeOnline

# Store data in array
$results = @()

# Get list of all mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited

# Loop through all mailboxes
foreach ($mailbox in $mailboxes) {
    # Get permissions for the mailbox
    $permissions = Get-MailboxPermission -Identity $mailbox.DistinguishedName

    # Loop through all permissions for the mailbox
    foreach ($permission in $permissions) {
        # Exclude NT AUTHORITY\SELF
        if ($permission.User -ne "NT AUTHORITY\SELF") {
            # Get UPN for the Office 365 user
            $userUPN = if ($permission.User -like "*\*") { 
                $permission.User | Split-Path -Leaf
            } else {
                (Get-User -Identity $permission.User).UserPrincipalName
            }

            $result = [PSCustomObject]@{
                Identity     = $mailbox.UserPrincipalName
                User         = $userUPN
                AccessRights = $permission.AccessRights
            }

            # Adding result row
            $results += $result

            Write-Host "Hittade delad brevlåda: $($mailbox.DisplayName)" -ForegroundColor DarkGreen
        }
    }
}

# Export results to CSV file
$results | Export-Csv -Path $home\Downloads\MinExport.csv -NoTypeInformation

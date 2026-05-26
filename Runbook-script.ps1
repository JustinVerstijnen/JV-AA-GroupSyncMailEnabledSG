param(
    [Parameter(Mandatory = $false)]
    [string]$Organization = "justinverstijnen.onmicrosoft.com",

    [Parameter(Mandatory = $false)]
    [string]$TargetGroupIdentity = "SG-SendAs@justinverstijnen.nl",

    [Parameter(Mandatory = $false)]
    [string[]]$AllowedDomains = @(
        "justinverstijnen.nl"
    ),

    [Parameter(Mandatory = $false)]
    [string[]]$ExcludedPrimarySmtpAddresses = @(
    ),

    [Parameter(Mandatory = $false)]
    [bool]$WhatIfMode = $true
)

Connect-ExchangeOnline `
    -ManagedIdentity `
    -Organization $Organization `
    -ShowBanner:$false

$ErrorActionPreference = "Stop"

try {
    Write-Output "Starting synchronization."
    Write-Output "Tenant: $Organization"
    Write-Output "Target group: $TargetGroupIdentity"
    Write-Output "WhatIfMode: $WhatIfMode"

    $AllowedDomainLookup = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    foreach ($Domain in $AllowedDomains) {
        if (-not [string]::IsNullOrWhiteSpace($Domain)) {
            [void]$AllowedDomainLookup.Add($Domain.Trim().ToLower())
        }
    }

    $ExcludedUserLookup = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    foreach ($ExcludedAddress in $ExcludedPrimarySmtpAddresses) {
        if (-not [string]::IsNullOrWhiteSpace($ExcludedAddress)) {
            [void]$ExcludedUserLookup.Add($ExcludedAddress.Trim().ToLower())
        }
    }

    Write-Output "Retrieving existing group members..."

    $ExistingMembers = Get-DistributionGroupMember `
        -Identity $TargetGroupIdentity `
        -ResultSize Unlimited

    $ExistingMemberLookup = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    foreach ($Member in $ExistingMembers) {
        if ($Member.PrimarySmtpAddress) {
            [void]$ExistingMemberLookup.Add($Member.PrimarySmtpAddress.ToString().ToLower())
        }

        if ($Member.WindowsEmailAddress) {
            [void]$ExistingMemberLookup.Add($Member.WindowsEmailAddress.ToString().ToLower())
        }

        if ($Member.ExternalDirectoryObjectId) {
            [void]$ExistingMemberLookup.Add($Member.ExternalDirectoryObjectId.ToString().ToLower())
        }

        if ($Member.Name) {
            [void]$ExistingMemberLookup.Add($Member.Name.ToString().ToLower())
        }
    }

    Write-Output "Number of existing group members: $($ExistingMembers.Count)"

    Write-Output "Retrieving user mailboxes..."

    $UserMailboxes = Get-EXOMailbox `
        -RecipientTypeDetails UserMailbox `
        -ResultSize Unlimited `
        -Properties PrimarySmtpAddress,HiddenFromAddressListsEnabled,DisplayName,UserPrincipalName

    $EligibleUsers = foreach ($Mailbox in $UserMailboxes) {
        if (-not $Mailbox.PrimarySmtpAddress) {
            continue
        }

        $PrimarySmtpAddress = $Mailbox.PrimarySmtpAddress.ToString().ToLower()
        $DomainPart = ($PrimarySmtpAddress -split "@")[-1]

        if (-not $AllowedDomainLookup.Contains($DomainPart)) {
            continue
        }

        if ($ExcludedUserLookup.Contains($PrimarySmtpAddress)) {
            continue
        }

        if ($Mailbox.HiddenFromAddressListsEnabled -eq $true) {
            continue
        }

        $Mailbox
    }

    Write-Output "Number of eligible users: $($EligibleUsers.Count)"

    $AddedCount = 0
    $AlreadyMemberCount = 0
    $FailedCount = 0

    foreach ($User in $EligibleUsers) {
        $PrimarySmtpAddress = $User.PrimarySmtpAddress.ToString().ToLower()

        if ($ExistingMemberLookup.Contains($PrimarySmtpAddress)) {
            $AlreadyMemberCount++
            Write-Output "Already a member, skipping: $PrimarySmtpAddress"
            continue
        }

        try {
            if ($WhatIfMode) {
                Write-Output "WHATIF: would add to group: $PrimarySmtpAddress"
            }
            else {
                Add-DistributionGroupMember `
                    -Identity $TargetGroupIdentity `
                    -BypassSecurityGroupManagerCheck `
                    -Member $PrimarySmtpAddress `
                    -Confirm:$false

                [void]$ExistingMemberLookup.Add($PrimarySmtpAddress)
                Write-Output "Added to group: $PrimarySmtpAddress"
            }

            $AddedCount++
        }
        catch {
            $FailedCount++
            Write-Warning "Failed to add $PrimarySmtpAddress : $($_.Exception.Message)"
        }
    }

    Write-Output "Synchronization completed."
    Write-Output "Added or WhatIf: $AddedCount"
    Write-Output "Already a member: $AlreadyMemberCount"
    Write-Output "Failed: $FailedCount"
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Write-Output "Exchange Online session closed."
}

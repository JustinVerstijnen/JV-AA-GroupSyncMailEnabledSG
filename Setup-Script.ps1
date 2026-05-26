<#
.SYNOPSIS
    One-time setup script for an Azure Automation Managed Identity to manage
    Exchange Online mail-enabled security group membership.

.DESCRIPTION
    This script performs the following actions:

    1. Grants the Exchange.ManageAsApp application permission to the Managed Identity.
    2. Assigns a Microsoft Entra directory role to the Managed Identity.
    3. Creates the Exchange Online Service Principal pointer if it does not exist.
    4. Assigns the Exchange RBAC role "Security Group Creation and Membership"
       to the Managed Identity service principal.
    5. Validates the resulting assignments.

.REQUIREMENTS
    - Run this script interactively as an admin.
    - Your account must have sufficient rights in Microsoft Graph and Exchange Online.
    - Microsoft.Graph PowerShell SDK.
    - ExchangeOnlineManagement PowerShell module.

.NOTES
    Use the Object ID / Principal ID and Client ID from:
    Azure Automation Account -> Identity -> System assigned

    Do NOT use values from "App registrations" unless you are working with a normal app registration.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$Organization,

    [Parameter(Mandatory = $true)]
    [string]$ManagedIdentityObjectId,

    [Parameter(Mandatory = $true)]
    [string]$ManagedIdentityClientId,

    [Parameter(Mandatory = $false)]
    [string]$ManagedIdentityDisplayName = "Azure Automation Managed Identity - SendAs Group Sync",

    [Parameter(Mandatory = $false)]
    [string]$EntraDirectoryRoleName = "Exchange Recipient Administrator",

    [Parameter(Mandatory = $false)]
    [string]$ExchangeManagementRoleName = "Security Group Creation and Membership",

    [Parameter(Mandatory = $false)]
    [string]$ExchangeRoleAssignmentName = "MI-SendAs-SecurityGroupMembership",

    [Parameter(Mandatory = $false)]
    [string]$ExchangeOnlineResourceAppId = "00000002-0000-0ff1-ce00-000000000000",

    [Parameter(Mandatory = $false)]
    [string]$ExchangeManageAsAppRoleId = "dc50a0fb-09a3-484d-be87-e023b12c6440",

    [Parameter(Mandatory = $false)]
    [bool]$AssignEntraDirectoryRole = $true,

    [Parameter(Mandatory = $false)]
    [bool]$AssignExchangeApplicationRbacRole = $true
)

$ErrorActionPreference = "Stop"

function Write-Step {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host $Message -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
}

function Ensure-Module {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModuleName
    )

    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "Module '$ModuleName' is not installed. Installing for CurrentUser..."
        Install-Module $ModuleName -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module $ModuleName -Force
}

Write-Step "Input validation"

if ($Organization -notlike "*.onmicrosoft.com") {
    Write-Warning "The Organization value is usually the tenant's primary .onmicrosoft.com domain, for example contoso.onmicrosoft.com."
}

Write-Host "Organization: $Organization"
Write-Host "Managed Identity Object ID: $ManagedIdentityObjectId"
Write-Host "Managed Identity Client ID: $ManagedIdentityClientId"
Write-Host "Managed Identity Display Name: $ManagedIdentityDisplayName"
Write-Host "Entra Directory Role: $EntraDirectoryRoleName"
Write-Host "Exchange Management Role: $ExchangeManagementRoleName"
Write-Host "Exchange Role Assignment Name: $ExchangeRoleAssignmentName"

Write-Step "Loading required PowerShell modules"

Ensure-Module -ModuleName "Microsoft.Graph"
Ensure-Module -ModuleName "ExchangeOnlineManagement"

Write-Step "Connecting to Microsoft Graph"

$GraphScopes = @(
    "AppRoleAssignment.ReadWrite.All",
    "Application.Read.All",
    "RoleManagement.ReadWrite.Directory"
)

Connect-MgGraph -Scopes $GraphScopes -NoWelcome

try {
    Write-Step "Validating Managed Identity service principal in Microsoft Entra ID"

    $ManagedIdentityServicePrincipal = Get-MgServicePrincipal -ServicePrincipalId $ManagedIdentityObjectId -ErrorAction Stop

    Write-Host "Found Managed Identity service principal:"
    $ManagedIdentityServicePrincipal |
        Select-Object DisplayName, Id, AppId |
        Format-List |
        Out-String |
        Write-Host

    if ($ManagedIdentityServicePrincipal.AppId -ne $ManagedIdentityClientId) {
        Write-Warning "The supplied Client ID does not match the AppId found for the supplied Object ID."
        Write-Warning "Supplied Client ID: $ManagedIdentityClientId"
        Write-Warning "Found AppId: $($ManagedIdentityServicePrincipal.AppId)"
        throw "Managed Identity Object ID and Client ID do not appear to belong to the same service principal."
    }

    Write-Step "Granting Exchange.ManageAsApp to the Managed Identity"

    $ExchangeOnlineResource = Get-MgServicePrincipal `
        -Filter "AppId eq '$ExchangeOnlineResourceAppId'" `
        -ErrorAction Stop

    if (-not $ExchangeOnlineResource) {
        throw "Office 365 Exchange Online resource service principal was not found in this tenant."
    }

    Write-Host "Found Office 365 Exchange Online resource:"
    $ExchangeOnlineResource |
        Select-Object DisplayName, Id, AppId |
        Format-List |
        Out-String |
        Write-Host

    $ExchangeManageAsAppRole = $ExchangeOnlineResource.AppRoles | Where-Object {
        $_.Id -eq $ExchangeManageAsAppRoleId -or $_.Value -eq "Exchange.ManageAsApp"
    }

    if (-not $ExchangeManageAsAppRole) {
        throw "Exchange.ManageAsApp app role was not found on the Office 365 Exchange Online resource."
    }

    $ExistingAppRoleAssignment = Get-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $ManagedIdentityObjectId `
        -All | Where-Object {
            $_.ResourceId -eq $ExchangeOnlineResource.Id -and
            $_.AppRoleId -eq $ExchangeManageAsAppRole.Id
        }

    if ($ExistingAppRoleAssignment) {
        Write-Host "Exchange.ManageAsApp is already assigned to the Managed Identity."
    }
    else {
        New-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $ManagedIdentityObjectId `
            -PrincipalId $ManagedIdentityObjectId `
            -AppRoleId $ExchangeManageAsAppRole.Id `
            -ResourceId $ExchangeOnlineResource.Id | Out-Null

        Write-Host "Exchange.ManageAsApp has been assigned to the Managed Identity."
    }

    if ($AssignEntraDirectoryRole -eq $true) {
        Write-Step "Assigning Microsoft Entra directory role to the Managed Identity"

        $RoleDefinition = Get-MgRoleManagementDirectoryRoleDefinition `
            -Filter "DisplayName eq '$EntraDirectoryRoleName'" `
            -ErrorAction Stop

        if (-not $RoleDefinition) {
            throw "Microsoft Entra role '$EntraDirectoryRoleName' was not found."
        }

        $ExistingDirectoryRoleAssignment = Get-MgRoleManagementDirectoryRoleAssignment `
            -Filter "principalId eq '$ManagedIdentityObjectId'" `
            -All | Where-Object {
                $_.RoleDefinitionId -eq $RoleDefinition.Id -and
                $_.DirectoryScopeId -eq "/"
            }

        if ($ExistingDirectoryRoleAssignment) {
            Write-Host "Microsoft Entra role '$EntraDirectoryRoleName' is already assigned to the Managed Identity."
        }
        else {
            New-MgRoleManagementDirectoryRoleAssignment `
                -PrincipalId $ManagedIdentityObjectId `
                -RoleDefinitionId $RoleDefinition.Id `
                -DirectoryScopeId "/" | Out-Null

            Write-Host "Microsoft Entra role '$EntraDirectoryRoleName' has been assigned to the Managed Identity."
        }
    }
    else {
        Write-Host "Skipping Microsoft Entra directory role assignment because AssignEntraDirectoryRole is false."
    }
}
finally {
    Disconnect-MgGraph | Out-Null
}

Write-Step "Connecting to Exchange Online as admin"

Connect-ExchangeOnline `
    -Organization $Organization `
    -ShowBanner:$false

try {
    Write-Step "Creating or validating Exchange Online Service Principal pointer"

    $ExistingExchangeServicePrincipal = Get-ServicePrincipal -ErrorAction SilentlyContinue | Where-Object {
        $_.AppId -eq $ManagedIdentityClientId -or
        $_.ObjectId -eq $ManagedIdentityObjectId -or
        $_.DisplayName -eq $ManagedIdentityDisplayName
    }

    if ($ExistingExchangeServicePrincipal) {
        Write-Host "Exchange Online Service Principal pointer already exists:"
        $ExistingExchangeServicePrincipal |
            Select-Object DisplayName, ObjectId, AppId |
            Format-List |
            Out-String |
            Write-Host
    }
    else {
        Write-Host "Creating Exchange Online Service Principal pointer..."

        New-ServicePrincipal `
            -AppId $ManagedIdentityClientId `
            -ObjectId $ManagedIdentityObjectId `
            -DisplayName $ManagedIdentityDisplayName | Out-Null

        Write-Host "Exchange Online Service Principal pointer has been created."

        $ExistingExchangeServicePrincipal = Get-ServicePrincipal -ErrorAction Stop | Where-Object {
            $_.AppId -eq $ManagedIdentityClientId -or
            $_.ObjectId -eq $ManagedIdentityObjectId
        }
    }

    if (-not $ExistingExchangeServicePrincipal) {
        throw "Exchange Online Service Principal pointer could not be found after creation."
    }

    $ExchangeServicePrincipalObjectId = $ExistingExchangeServicePrincipal.ObjectId

    Write-Host "Exchange Service Principal Object ID used for RBAC assignment: $ExchangeServicePrincipalObjectId"

    if ($AssignExchangeApplicationRbacRole -eq $true) {
        Write-Step "Assigning Exchange RBAC role to the Managed Identity Service Principal"

        $ManagementRole = Get-ManagementRole -Identity $ExchangeManagementRoleName -ErrorAction Stop

        if (-not $ManagementRole) {
            throw "Exchange management role '$ExchangeManagementRoleName' was not found."
        }

        $ExistingExchangeRoleAssignment = Get-ManagementRoleAssignment `
            -Role $ExchangeManagementRoleName `
            -ErrorAction SilentlyContinue | Where-Object {
                $_.RoleAssigneeType -eq "ServicePrincipal" -and
                (
                    $_.RoleAssigneeName -eq $ExchangeServicePrincipalObjectId -or
                    $_.RoleAssigneeName -eq $ManagedIdentityObjectId -or
                    $_.Name -eq $ExchangeRoleAssignmentName
                )
            }

        if ($ExistingExchangeRoleAssignment) {
            Write-Host "Exchange RBAC role assignment already exists:"
            $ExistingExchangeRoleAssignment |
                Select-Object Name, Role, RoleAssigneeName, RoleAssigneeType |
                Format-List |
                Out-String |
                Write-Host
        }
        else {
            New-ManagementRoleAssignment `
                -Name $ExchangeRoleAssignmentName `
                -Role $ExchangeManagementRoleName `
                -App $ExchangeServicePrincipalObjectId | Out-Null

            Write-Host "Exchange RBAC role '$ExchangeManagementRoleName' has been assigned to the Managed Identity Service Principal."
        }
    }
    else {
        Write-Host "Skipping Exchange Application RBAC role assignment because AssignExchangeApplicationRbacRole is false."
    }

    Write-Step "Validation"

    Write-Host "Exchange Online Service Principal:"
    Get-ServicePrincipal -Identity $ExchangeServicePrincipalObjectId |
        Select-Object DisplayName, ObjectId, AppId |
        Format-List |
        Out-String |
        Write-Host

    Write-Host "Exchange RBAC assignments for this Service Principal:"
    Get-ManagementRoleAssignment -ErrorAction SilentlyContinue | Where-Object {
        $_.RoleAssigneeType -eq "ServicePrincipal" -and
        (
            $_.RoleAssigneeName -eq $ExchangeServicePrincipalObjectId -or
            $_.RoleAssigneeName -eq $ManagedIdentityObjectId
        )
    } |
        Select-Object Name, Role, RoleAssigneeName, RoleAssigneeType |
        Format-Table -AutoSize |
        Out-String |
        Write-Host

    Write-Step "Completed"

    Write-Host "Setup completed successfully."
    Write-Host "Wait a few minutes for permissions to propagate, then test the Azure Automation runbook again."
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}

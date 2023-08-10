<#
Copyright (C) 2023 Kyle Gary Smith

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Affero General Public License as published
by the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU Affero General Public License for more details.

You should have received a copy of the GNU Affero General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
#>

<#
.SYNOPSIS
    Script that sets permission levels for SharePoint Online sites using PnP PowerShell.

.DESCRIPTION
    The script sets permission levels for SharePoint Online sites using PnP PowerShell. The script prompts the user to confirm the action before making any changes.

.PARAMETER SPOTenant
    SharePoint tenant URL.

.PARAMETER SPOSite
    SharePoint Online site URL.

.PARAMETER IdentityType
    Identity type: User or Group.

.PARAMETER Identity
    User or group identity.

.PARAMETER PermissionLevel
    Permission level to set.

.PARAMETER UpdateLibraries
    Update all document libraries in the site and subsites.

.EXAMPLE
    Set-SPOGobalPerms -SPOTenant 'contoso' -SPOSite 'finance' -IdentityType 'User' -Identity 'jane@contoso.com' -PermissionLevel 'Read' -UpdateLibraries $true
#>

param(
    # SharePoint Tenant
    [Parameter(mandatory = $true )]
    [string]
    $SPOTenant,

    # SPO Site
    [Parameter(mandatory = $false)]
    [string]
    $SPOSite,

    # Identity Type
    [Parameter(mandatory = $true)]
    [ValidateSet(
        "User",
        "Group"
    )]
    [string]
    $IdentityType,

    # User/Group Identity
    [Parameter(mandatory = $true)]
    [string]
    $Identity,

    # Permission Level
    [Parameter(Mandatory = $true)]
    [ValidateSet(
        "Full Control",
        "Design",
        "Edit",
        "Contribute",
        "Review",
        "Read",
        "Restricted View",
        "Approve",
        "Manage Hierarchy",
        "Restricted Read",
        "Restricted Interfaces for Translation",
        "Contribute without delete",
        "Moderate",
        "Create new subsites",
        "View Only"
    )]
    [string]
    $PermissionLevel,

    # Update for all libraries?
    [Parameter(mandatory = $false)]
    [boolean]
    $UpdateLibraries
)

#region Functions
function Confirm-Action {
    Write-host  "Your current selected site is $SPOSiteURL (this will include all subsites), you are applying the '$PermissionLevel' permission to the '$Identity' User/Group." -ForegroundColor Yellow -BackgroundColor Black
    Write-host  "Would you like to continue? Y/N:" -ForegroundColor Yellow -BackgroundColor Black
    
    $Confirmation = Read-Host

    If ($Confirmation -eq "N") {
        Write-host  "'N' selected, exiting" -ForegroundColor Yellow
        Start-Sleep -Seconds 10

        Break
    }
    elseif ($Confirmation -eq "Y") {
        Write-host  "'Y' selected, continuing..." -ForegroundColor Green

    }
    else {
        Write-host  "Invalid option selected, exiting" -ForegroundColor Red
        Start-Sleep -Seconds 10

        Break
    }
}

function Set-SharePointOnlinePermissions {
    param (
        # Site
        [Parameter(mandatory = $true)]
        [Microsoft.SharePoint.Client.Web]
        $Site,
 
        # PnP Identity
        [Parameter(mandatory = $true)]
        $PnPIdentity,
 
        # Permission Level
        [Parameter(mandatory = $true)]
        [string]
        $PermissionLevel,

        # Identity Type
        [Parameter(mandatory = $true)]
        [string]
        $IdentityType
    )

    $relativePath = ($Site.ServerRelativeUrl).Replace("/sites/$SPOSite", "")

    #Switch if site is root, or subsite
    switch ($Site.ServerRelativeUrl) {
        $SPOSiteURLChild {
            Write-Host "Site: '$($Site.Title)' is the root site"

            #Add with perms
            try {
                Write-host  "Changing permissions on site: '$($Site.Title)'"
                If ($IdentityType -eq "User") {                    
                    PnP.PowerShell\Set-PnPWebPermission -User $PnPLoginName -AddRole $PermissionLevel
                }
                elseif ($IdentityType -eq "Group") {
                    PnP.PowerShell\Set-PnPWebPermission -Group $PnPIdentity -AddRole $PermissionLevel
                }
                Write-host  "Completed for site: '$($Site.Title)'"
            }
            catch {
                Write-host  -ForegroundColor Red -BackgroundColor Black "Unable to set permission: '$PermissionLevel' for Identity: '$Identity' on site $($Site.Title): $($Error[0].Exception)"
                $Site.Url += $FailedSites
            }
        }
        Default { 
            Write-Host "Site: '$($Site.Title)' is a subsite"

            #Add with perms
            try {
                Write-host  "Changing permissions on site: '$($Site.Title)'"
                If ($IdentityType -eq "User") {
                    PnP.PowerShell\Set-PnPWebPermission -User $PnPLoginName -AddRole $PermissionLevel -Identity $relativePath
                }
                elseif ($IdentityType -eq "Group") {                    
                    PnP.PowerShell\Set-PnPWebPermission -Group $PnPIdentity -AddRole $PermissionLevel -Identity $relativePath                  
                }
                Write-host  "Completed for site: '$($Site.Title)'"
            }
            catch {
                Write-host -ForegroundColor Red -BackgroundColor Black "Unable to set permission: '$PermissionLevel' for Identity: '$Identity' on site $($Site.Title): $($Error[0].Exception)"
                $Site.Url += $FailedSites
            }
        }
    }
}

function Get-SPOPermissionCheck() {
    param (
        # Site
        [Parameter(mandatory = $true)]
        [Microsoft.SharePoint.Client.Web]
        $Site,

        # PnP Identity
        [Parameter(mandatory = $true)]
        $PnPLoginName,

        # Permission Level
        [Parameter(mandatory = $true)]
        [string]
        $PermissionLevel
    )

    #Reset result var
    $result = $null
    
    Foreach ($RoleAssignment in $Site.RoleAssignments) {

        #Get extended properties
        PnP.PowerShell\Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member -ErrorAction SilentlyContinue

        #Expand properties to own vars
        $member = $RoleAssignment.Member
        $loginName = $member.LoginName
        $roleBindings = $RoleAssignment.RoleDefinitionBindings
        $roleBindingsName = $roleBindings.name

        #If both user/group and permission match for one of the site role assignments, success!
        if (($loginName -eq $PnPLoginName) -and ($roleBindingsName -eq $PermissionLevel)) {
            $result = $true
        }
    }

    #If result is empty (meaning no matches, fail)
    if ([string]::IsNullOrWhiteSpace($result)) {
        $result = $false
    }

    #Keep this explicitly typed (it hates it otherwise)
    return [bool]$result
}

#endregion

#region Var
$Modules = @(
    "Microsoft.Online.SharePoint.PowerShell"
    "PnP.PowerShell"
)
#Reset Sites to null
$Sites = @()

#Reset failed sites
$FailedSites = @()

#Set SPO Site to default if one is not specified
if ([string]::IsNullOrWhiteSpace($SPOSite)) {
    Write-host  "SPO Site not specified, using default site"
    $SPOSiteURL = "https://$SPOTenant.sharepoint.com/"
}
else {
    $SPOSiteURL = "https://$SPOTenant.sharepoint.com/sites/$SPOSite"
}

$SPOSiteURLChild = $SPOSiteURL.Replace("https://$SPOTenant.sharepoint.com", "")

#endregion

#region do later for nice to haves

#region Check, get and update modules
foreach ($Module in $Modules) {
    if (!(Get-Module -Name $Module -ListAvailable | Select-Object Name, Version)) {
        Write-host  "Module $Module is not installed, installing now..."
        try {
            Install-Module -Name $Module
            Write-host  "Module $Module installed"
        }
        catch {
            Write-host -ForegroundColor Red -BackgroundColor Black  "Unable to install module $Module, installaion failed with error $($Error[0].Exception)"
        }
    }
    else {
        Write-host  "Module $Module is installed, continuing"
    }

    #Update Module
    <#
    try {
        Write-Host "Attempting to update module $Module"
        Update-Module -Name $Module -Force
        Write-host  "Module $Module updated"
    }
    catch {
        Write-host -ForegroundColor Red -BackgroundColor Black   "Unable to update module $Module, update failed with error: $($Error[0].Exception)" 
    }
   #>
}
#endregion

############################
##Update for all libraries##
############################

#endregion

#Connect to tenant
try {
    PnP.PowerShell\Connect-PnPOnline -Url $SPOSiteURL -UseWebLogin
    Write-host  "SPO site $SPOTenant connected"
}
catch {
    Write-host -ForegroundColor Red -BackgroundColor Black  "Unable to connect to site $SPOTenant, connection failed with error: $($Error[0].Exception)"
}

#Get all sites
try {
    #Add subsites
    $Sites += PnP.PowerShell\Get-PnPSubWeb -Recurse -Includes "HasUniqueRoleAssignments", "RoleAssignments"

    #Add root site (the object gets angry if you do this first)
    $Sites += PnP.PowerShell\Get-PnPWeb -Includes "HasUniqueRoleAssignments", "RoleAssignments"

    Write-host  "$($Sites.Count) sites were found"
}
catch {
    Write-host -ForegroundColor Red -BackgroundColor Black  "Unable to get sites, command failed with error: $($Error[0].Exception) "
}

#Get User/Group
try {
    If ($IdentityType -eq "User") {
        $PnPIdentity = PnP.PowerShell\Get-PnPUser  | Where-Object Email -Match $Identity
        $PnPLoginName = $PnPIdentity.LoginName
        Write-host  "User PnP Identity is $($PnPIdentity.Email)"
    }
            
    elseif ($IdentityType -eq "Group") {
        $PnPIdentity = PnP.PowerShell\Get-PnPGroup | Where-Object Title -Match $Identity
        $PnPLoginName = $PnPIdentity.LoginName
    }

    #Fail if SPO doesn't know what you mean
    If ($null -eq $PnPLoginName) {
        throw "No Identity found for identity: $($Identity)"
    }
}
catch {
    { write-host -ForegroundColor Red -BackgroundColor Black "Unable to get PnPIdentity, command failed with error: $($Error[0].Exception) " }
}


#Confirm
Confirm-Action

#Change Permissions on site if applicable
foreach ($Site in $Sites) {
    Write-Host ""
    Write-Host " -------------------------------------------------------------------------------------------- "

    Write-Host "Checking site: '$($Site.Title)'" 

    #Change if the site does not inherit permissions
    If ($Site.HasUniqueRoleAssignments) {

        Write-Host "Site: '$($Site.Title)' has unique permissions" -ForegroundColor Yellow -BackgroundColor Black

        #If the permission check returns false, change the permissions
        if ( -not (Get-SPOPermissionCheck -Site $Site -PnPLoginName $PnPLoginName -permissionlevel $PermissionLevel)) {
            Write-Host "Need to change Site: '$($Site.Title)'"
            Set-SharePointOnlinePermissions -Site $Site -PnPIdentity $PnPLoginName -permissionlevel $PermissionLevel -IdentityType $IdentityType
        }
        else {
            Write-Host "Permissions on site: '$($Site.Title)' were correct" -ForegroundColor Green -BackgroundColor Black
        }
    }
    else {
        Write-Host "Site '$($Site.title)' has inherited permissions, ignoring..." -ForegroundColor Green -BackgroundColor Black
    }  
}

#Return the failed sites
if (-not ([string]::IsNullOrWhiteSpace($FailedSites))) {
    Write-host  "Completed with errors" -ForegroundColor Red
    Write-Host "$($FailedSites.Count) sites failed:"
    foreach ($FailedSite in $FailedSites) {
        Write-Host "Site $($FailedSite.ServerRelativeUrl) failed"
    }
}

#No sites failed, congrats!
else {
    Write-host  "Completed without errors" -ForegroundColor Green
}

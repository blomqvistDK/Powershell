<#
 
.SYNOPSIS
    Runs through all GPOs and checks if a given target has permissions and adds result to a text file. 
 
.DESCRIPTION
    This script will produce a text file containing all GPO details of the GPOs where the target either have or doesnt have permissions to. 
 
.PARAMETER TargetName
    Name of the target 
 
.PARAMETER TargetType
    Type of target (Group, User or Computer)

.PARAMETER hasPermissions
    Boolean to specify if we want to find all the GPO's that the target has permission to (True) or does not have permission to (False)
   
.EXAMPLE
    Get-GPOPermissionReport -TargetName "Authenticated Users" -TargetType Group -hasPermission $False
    Lists all the GPO's where Authenticated Users doesnt have access

.EXAMPLE
    Get-GPOPermissionReport -TargetName "Authenticated Users" -TargetType Group -hasPermission $True
    Lists all the GPO's where Authenticated Users have access
  
.NOTES
    Cmdlet name:      Get-GPOPermissionReport
    Author:           Mikael Blomqvist
    DateCreated:      2017-07-06
 
#>


[CmdletBinding(DefaultParameterSetName="String")] 
Param(
 [Parameter(Mandatory=$True, Position=0, HelpMessage="TargetName")] 
            [string]$TargetName,
 [Parameter(Mandatory=$True, Position=1, HelpMessage="TargetType(User/Computer/Group)")] 
            [string]$TargetType,
 [Parameter(Mandatory=$True, Position=2, HelpMessage="Find all GPO's where target has permission(True/False)")] 
            [boolean]$hasPermission
)


#Load GPO module
 Import-Module GroupPolicy
 
 $date = Get-Date -Format yyyyMMddhhmmss

 $file = "$PSScriptRoot\GPOPermissions_$date.csv"

 #Prepare the text file with the correct headings
 Add-Content $file "Name,LinksPath,WMI Filter,CreatedTime,ModifiedTime,CompVerDir,CompVerSys,UserVerDir,UserVerSys,CompEnabled,UserEnabled,SecurityFilter,GPO Enabled,Enforced"

#Get all GPOs in current domain
 $GPOs = Get-GPO -All

#Check we have GPOs
 if ($GPOs) 
 {
    foreach ($GPO in $GPOs) 
    {
        $GPOName = $GPO.DisplayName
        $GPOGuid = $GPO.id
        
        #Retrieve the permission for the specified target
        $GPPermissions = Get-GPPermissions -Guid $GPOGuid -TargetName $TargetName -TargetType $TargetType -ErrorAction SilentlyContinue

        #Check if target have permissions or not and set permission to either True(has permission) or False(no permission)
        If($GPPermissions -eq $null)
        { 
            $permission = $False
        }
        Else
        {
            $permission = $True
        }

        # if permissions matches what we look for either true or false
        if ($permission -eq $hasPermission) 
        { 
            [xml]$gpocontent =  Get-GPOReport -Guid $GPOGuid -ReportType xml
            $LinksPaths = $gpocontent.GPO.LinksTo
            $Wmi = Get-GPO -Guid $GPOGuid | Select-Object WmiFilter
 
            $CreatedTime = $gpocontent.GPO.CreatedTime
            $ModifiedTime = $gpocontent.GPO.ModifiedTime
            $CompVerDir = $gpocontent.GPO.Computer.VersionDirectory
            $CompVerSys = $gpocontent.GPO.Computer.VersionSysvol
            $CompEnabled = $gpocontent.GPO.Computer.Enabled
            $UserVerDir = $gpocontent.GPO.User.VersionDirectory
            $UserVerSys = $gpocontent.GPO.User.VersionSysvol
            $UserEnabled = $gpocontent.GPO.User.Enabled
            $SecurityFilter = ((Get-GPPermissions -Guid $GPOGuid -All | ?{$_.Permission -eq "GpoApply"}).Trustee | ?{$_.SidType -ne "Unknown"}).name -Join ','
        
            if($LinksPaths -ne $null)
            {
                foreach ($LinksPath in $LinksPaths)
                {
                    Add-Content $file "$GPOName,$($LinksPath.SOMPath),$(($wmi.WmiFilter).Name),$CreatedTime,$ModifiedTime,$CompVerDir,$CompVerSys,$UserVerDir,$UserVerSys,$CompEnabled,$UserEnabled,""$($SecurityFilter)"",$($LinksPath.Enabled),$($LinksPath.NoOverride)"
                }   
            }
            else
            {
                Add-Content $file "$GPOName,$($LinksPath.SOMPath),$(($wmi.WmiFilter).Name),$CreatedTime,$ModifiedTime,$CompVerDir,$CompVerSys,$UserVerDir,$UserVerSys,$CompEnabled,$UserEnabled,""$($SecurityFilter)"",$($LinksPath.Enabled),$($LinksPath.NoOverride)"
            } 
        }
    }
}
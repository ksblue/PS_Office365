####################################################################
#
# This is a script to disable a selected service for all users that
# are licensed with a selected SKU.  It will preserve any services
# that you may have previously disabled on each license.
#
# Change AccountSkuId to the value of the SKU/license you want to update.
# Change DisableNewServiceName to be the name of the service you want
# to disable.
# Change the number value on Get-MsolUser -MaxResults to a small number 
# if you want to try this out on a few accounts first.
#
# This script requires the legacy MSOnline V1 PowerShell module for Azure 
# Active Directory.  More information on where to find the downloads  
# here:
# https://technet.microsoft.com/en-us/library/dn975125.aspx
#
# If you are running this script multiple times, it is not necessary to
# do the Connect-MsolService cmdlet every time, only the first time.
# The other times, you can comment out that line.  (Put a # in front of it.)
#
####################################################################

$AccountSkuId = "yourtenant:STANDARDWOFFPACK_FACULTY"
$DisableNewServiceName = "TEAMS1"

#Requires -Module MSOnline

Connect-MsolService -Credential (Get-Credential)

# Get all licensed users.
Write-Host "Querying licensed users."
$users = Get-MsolUser -MaxResults 100000 | 
   Where-Object {$_.isLicensed -eq $true} |
   Sort-Object -Property userprincipalname

# Loop thru all licensed users.
Write-Host "Processing licensed users."
foreach ($user in $users)
{
   $upn = $user.userprincipalname

   # Loop thru all of a user's licenses.
   foreach ($license in $user.Licenses) 
   {
      if ($license.AccountSkuId -eq $AccountSkuId) 
      {
         # Once we find the license we want to update, make a list of all the currently disabled services on that license.
         $disabledServices = @($license.servicestatus | 
            Where-Object {$_.provisioningstatus -eq "Disabled"} | 
            Select-Object -ExpandProperty serviceplan |
            Select-Object -ExpandProperty servicename |
            Sort-Object)

         # If the new service is not already disabled, add it to the list of disabled services to update the license.
         if ($DisableNewServiceName -notin $disabledServices)
         {
            Write-Host "Disabling $DisableNewServiceName on $AccountSkuId for $upn"
            $disabledServices += $DisableNewServiceName
            $licenseOption = New-MsolLicenseOptions -AccountSkuId $AccountSkuId -DisabledPlans $disabledServices
            Set-MsolUserLicense -UserPrincipalName $upn -LicenseOptions $licenseOption 
         }
      }     
   }
} 

Write-Host "All done!"

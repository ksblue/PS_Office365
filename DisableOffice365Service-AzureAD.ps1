########################################################################
#
# This is a script to disable a selected service for all users that
# are licensed with a selected SKU.
#
# Change SkuName to the name of the SKU/license you want to update.
# Change ServiceToDisable to the name of the service you want to disable.
# Change the number value on Get-AzureADUser -Top to a small number 
# if you want to try this out on a few accounts first.
#
# This script requires the Azure Active Directory V2 PowerShell module.
# More information on where to find the downloads here:
# https://technet.microsoft.com/en-us/library/dn975125.aspx
#
# If you are running this script multiple times, it is not necessary to
# do the Connect-AzureAD cmdlet every time, only the first time.
# The other times, you can comment out that line.  (Put a # in front of it.)
#
########################################################################

$SkuName = "STANDARDWOFFPACK_FACULTY"
$ServiceToDisable = "STREAM_O365_E3"

#Requires -Module AzureAD

Connect-AzureAD

# Find the SKU by name.
$sku = Get-AzureADSubscribedSku | 
   Where-Object {$_.skupartnumber -eq $SkuName}

if ($sku -eq $null)
{
   Write-Host "$SkuName - SKU not found."
}
elseif ($sku.ServicePlans.ServicePlanName -notcontains $ServiceToDisable)
{
   Write-Host "$ServiceToDisable - Service not found."
}
else
{
   # Get all users with this license.
   Write-Host "Querying licensed users."
   $users = Get-AzureADUser -Top 100 | 
      Where-Object {$_.AssignedLicenses -ne $null -and $_.AssignedLicenses.skuid -contains $sku.SkuId} |
      Sort-Object -Property UserPrincipalName

   # Loop thru all licensed users.
   Write-Host "Processing licensed users."
   foreach ($user in $users)
   {
      $upn = $user.UserPrincipalName

      $userLicense = Get-AzureADUserLicenseDetail -ObjectId $upn | 
         Where-Object {$_.SkuId -eq $sku.SkuId}
   
      # Make a list of all the currently disabled services on that user's license.
      $userDisabledServices = @($userLicense.ServicePlans | 
         Where-Object {$_.ProvisioningStatus -eq "Disabled"} | 
         Select-Object -ExpandProperty ServicePlanName |
         Sort-Object)

      # If the new service is not disabled, add it to the list of already disabled services and update the user's license. 
      if ($ServiceToDisable -notin $userDisabledServices)
      {
         Write-Host "Disabling $ServiceToDisable from $SkuName for $upn"
         $license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
         $license.SkuId = $sku.SkuId
         
         $license.DisabledPlans = ($userLicense.ServicePlans | Where-Object {$_.ProvisioningStatus -eq "Disabled"}).ServicePlanId
         $license.DisabledPlans += ($sku.ServicePlans | Where-Object {$_.ServicePlanName -eq $ServiceToDisable}).ServicePlanId
        
         $licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
         $licenses.AddLicenses = $license 
         Set-AzureADUserLicense -ObjectId $user.ObjectId -AssignedLicenses $licenses
      }
   } 

   Write-Host "All done!"
}

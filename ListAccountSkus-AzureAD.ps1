####################################################################
#
# This script will produce an easy to read list of all of the SKUs 
# and their services assigned to your tenant.
#
# This script requires the Azure Active Directory V2 PowerShell module.
# More information on where to find the downloads here:
# https://technet.microsoft.com/en-us/library/dn975125.aspx
#
# If you are running this script multiple times, it is not necessary to
# do the Connect-MsolService cmdlet every time, only the first time.
# The other times, you can comment out that line.  (Put a # in front of it.)
#
####################################################################

#Requires -Module AzureAD

Connect-AzureAD

$asku = Get-AzureADSubscribedSku

$sku = foreach ($a in $asku) 
{
   foreach ($s in $a.serviceplans)
   {
      New-Object psobject -Property @{
         SkuId              = $a.skupartnumber
         ServiceName        = $s.serviceplanname
         ProvisioningStatus = $s.ProvisioningStatus
      }   
   }
}

$sku | 
Sort-Object -Property Skuid,ServiceName | 
Format-Table -Property ServiceName,ProvisioningStatus -GroupBy SkuId

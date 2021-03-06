####################################################################
#
# This script will produce an easy to read list of all of the SKUs 
# and their services assigned to your tenant.
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

#Requires -Module MSOnline

Connect-MsolService -Credential (Get-Credential)

$sku = foreach ($m in (Get-MsolAccountSku)) 
{
   foreach ($s in $m.ServiceStatus)
   {
      New-Object psobject -Property @{
         AccountSkuId       = $m.AccountSkuId
         ServiceName        = $s.ServicePlan.ServiceName
         ProvisioningStatus = $s.ProvisioningStatus.ToString()
      }   
   }
}

$sku | 
Sort-Object -Property AccountSkuid,ServiceName | 
Format-Table -Property ServiceName,ProvisioningStatus -GroupBy AccountSkuId

<#

.SYNOPSIS
   Script to send alert if new Office 365 SKUs or services are found.

.DESCRIPTION
   This script compares the current list of Office 365 SKUs and 
   services with a list of SKUs and services saved the previous time
   this script ran.  If there are any new SKUs or services, an 
   e-mail alert is sent. 

.PARAMETER LogFile
   Path and filename for script log file.

.PARAMETER PreviousSkuFile
   Path and filename for the file containing the SKU ids and 
   service names that existed the last time the script ran.  
   This is used to compare with the SKU ids and service names 
   that currently exist.

.PARAMETER EmailFrom
   The from e-mail address to use when sending an e-mail.

.PARAMETER EmailTo
   The to e-mail address(es) to use when sending an e-mail.

.PARAMETER EmailSmtp
   The SMTP server to use when sending an e-mail.

.PARAMETER O365Cred
   Path and filename for xml file with Office 365 credentials.

.PARAMETER RetriesCount
   The number of times to attempt to connect to Office 365 MSOL
   before quitting.

.PARAMETER RetriesSleep
   The number of seconds to sleep in between attempts to retry
   connecting to Office 365 MSOL.

.NOTES
   NAME:    Office365NewSkuAlert.ps1
   AUTHOR:  Shaun Blue
   DATE:    August 21, 2015

#>
[cmdletbinding()]
param (
   [string]$LogFile,
   [string]$PreviousSkuFile,
   [string]$EmailFrom,
   [string[]]$EmailTo,
   [string]$EmailSmtp,
   [string]$O365Cred,
   [int]$RetriesCount,
   [int]$RetriesSleep 
)

#Requires -Modules CommonFunctions, Office365Functions, MSOnline
#Requires -Version 3.0

Set-StrictMode -Version latest

Write-LogMessage -Path $LogFile -Message "$($MyInvocation.mycommand.name) started. ---"

try
{
   Write-LogMessage -Path $LogFile -Message "Connecting to Office 365 MSOL."
   if (Open-Office365SessionScript -Credential (Import-Credential -Path $O365Cred) -Msol -LogFile $LogFile -RetriesCount $RetriesCount -RetriesSleep $RetriesSleep)
   {
      $previousSkus = Import-Csv -Path $PreviousSkuFile -ErrorAction Stop | 
      Select-Object -Property AccountSkuId,@{Name="ServiceNames";Expression={$_.servicenames -split ","}}

      $currentSkus = Get-MsolAccountSku -ErrorAction stop | 
      Select-Object -Property AccountSkuId,@{Name="ServiceNames";Expression={$_.servicestatus.serviceplan.servicename | Sort-Object}}

      if (@($previousSkus).Count -gt 0)
      {
         $parms = @{
            ReferenceObject  = $previousSkus | Select-Object -ExpandProperty AccountSkuId
            DifferenceObject = $currentSkus | Select-Object -ExpandProperty AccountSkuId
            IncludeEqual     = $true
            ErrorAction      = "Stop"
         }
         $comp = Compare-Object @parms 

         $newSkus = $comp |
         Where-Object {$_.sideindicator -eq "=>"} | 
         Select-Object -ExpandProperty inputobject

         $sameSkus = $comp |
         Where-Object {$_.sideindicator -eq "=="} | 
         Select-Object -ExpandProperty inputobject

         $body = ""
         $CRLF = "`r`n"

         if ($newSkus -ne $null)
         {
            foreach ($n in $newSkus)
            {
               $body += "New SKU:  $n$CRLF"
               Write-LogMessage -Path $LogFile -Message "New SKU - $n"
               foreach ($s in ($currentSkus | Where-Object {$_.accountskuid -eq $n}).servicenames)
               {
                  $body += "   $s$CRLF"
                  Write-LogMessage -Path $LogFile -Message "   $s"
               }
            }
            $body += "$CRLF"
         }

         if ($sameSkus -ne $null)
         {
            foreach ($s in $sameSkus)
            {
               $parms = @{
                  ReferenceObject  = ($previousSkus | Where-Object {$_.accountskuid -eq $s}).servicenames
                  DifferenceObject = ($currentSkus | Where-Object {$_.accountskuid -eq $s}).servicenames
                  ErrorAction      = "Stop"
               }
               $newServices = Compare-Object @parms |
               Where-Object {$_.sideindicator -eq "=>"}

               if ($newServices -ne $null)
               {
                  $body += "New Service(s) for Existing SKU:  $s$CRLF"
                  Write-LogMessage -Path $LogFile -Message "New Service(s) for Existing SKU - $s"
                  foreach ($x in $newServices)
                  {
                     $body += "   $($x.inputobject)$CRLF"
                     Write-LogMessage -Path $LogFile -Message "   $($x.InputObject)"
                  }
                  $body += "$CRLF"
               }
            }
         }

         if ($body -eq "")
         {
            Write-LogMessage -Path $LogFile -Message "No changes."
         }
         else
         {
            $parms = @{
               From        = $EmailFrom
               To          = $EmailTo
               Subject     = "New Office 365 SKUs or Services"
               SmtpServer  = $EmailSmtp
               Body        = $body
               ErrorAction = "Stop" 
            }
            Send-MailMessage @parms 
            Write-LogMessage -Path $LogFile -Message "E-mail alert sent to $EmailTo"
         }
      }

      $currentSkus | 
      Select-Object -Property AccountSkuId,@{Name="ServiceNames";Expression={($_.servicenames | Sort-Object) -join ","}} | 
      Export-Csv -Path $PreviousSkuFile -NoTypeInformation -ErrorAction Stop
   }
   else
   {
      Write-LogMessage -Path $LogFile -Message "*** Could not connect to Office 365 MSOL. ***"
   }
}
catch
{
   Write-LogMessage -Path $LogFile -Message "EXCEPTION:"
   Write-LogException -Path $LogFile -Exception $_
}

Write-LogMessage -Path $LogFile -Message "$($MyInvocation.mycommand.name) finished."

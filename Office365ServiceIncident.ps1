<#

.SYNOPSIS
   Script to e-mail the latest Office 365 service incidents.

.DESCRIPTION
   This script queries the Office 365 service health dashboard, 
   checking for any new incidents posted since the last time the 
   script ran.  If any new incidents are found, they are formatted 
   into an e-mail and sent to the e-mail administrators.

.PARAMETER ServiceUrl
   The Office 365 Service URL.

.PARAMETER LastRunFile
   Path for file that stores the date/time this script was run last.

.PARAMETER LogFile
   Path for output file of script log results.

.PARAMETER EmailFrom
   From address for e-mail.

.PARAMETER EmailTo
   To addresses for e-mail.

.PARAMETER EmailErrorTo
   To addresses for error message e-mail.

.PARAMETER EmailSmtp
   SMTP server for sending e-mail.

.PARAMETER O365Cred
   Path and filename for xml file with Office 365 credentials.

.NOTES
   NAME:    Office365ServiceIncident.ps1
   AUTHOR:  Shaun Blue
   DATE:    October 17, 2014

   03/27/2017 Use new LSU modules.  Also change to only report the lastest updates
   instead of reporting all updates for incidents with a recent update.
   05/02/2017 Modified format of message details.

#>
[cmdletbinding()]
param (
   [string]$ServiceUrl,
   [string]$LastRunFile,
   [string]$LogFile,
   [string]$EmailFrom,
   [string[]]$EmailTo,
   [string[]]$EmailErrorTo,
   [string]$EmailSmtp,
   [string]$O365Cred 
)

#Requires -Modules CommonFunctions
#Requires -Version 3.0

Set-StrictMode -Version latest

$scriptname = $MyInvocation.mycommand.name
Write-LogMessage -Path $LogFile -Message "$scriptname started. ---"

try
{
   # Read from LastRunFile the last time this script was run.
   # If not found, use the current date.
   try
   {
      $lastrundatetime = [datetime]::SpecifyKind((Get-Content -Path $LastRunFile -ErrorAction stop ),[datetimekind]::local)
   }
   catch
   {
      $lastrundatetime = Get-Date 
   }
   $lastrundatetime = $lastrundatetime.ToUniversalTime()

   $emailbody = "<html><head><style>" +
   “BODY{font-family: Arial; font-size: 10pt;}” +
   “TABLE{border: 1px solid black; border-collapse: collapse;}” +
   “TH{border: 1px solid black; background: #dddddd; padding: 5px; text-align: left }” +
   “TD{border: 1px solid black; padding: 5px; }” +
   “</style></head><body>"

   # Use credentials stored in file.
   $cred = Import-Credential -Path $O365Cred -ErrorAction Stop
   # Obtain cookie for authentication.
   $jsonPayload = (@{userName=$cred.username;password=$cred.GetNetworkCredential().password;} | ConvertTo-Json).tostring()
   $cookie = (Invoke-RestMethod -ContentType "application/json" -Method Post -Uri "$ServiceUrl/Register" -Body $jsonPayload).RegistrationCookie
   # Get events.
   $jsonPayload = (@{lastCookie=$cookie;locale="en-US";preferredEventTypes=@(0)} | ConvertTo-Json).tostring()
   $newevents = (Invoke-RestMethod -ContentType "application/json" -Method Post -Uri "$ServiceUrl/GetEvents" -Body $jsonPayload).events |
   Where-Object {$_.lastupdatedtime -gt $lastrundatetime -and ($_.messages.publishedtime | Where-Object {$_ -gt $lastrundatetime})}

   # If any new alerts are found, format an e-mail and send it.
   if (@($newevents).count -gt 0)
   {
      foreach ($n in $newevents) 
      {
         $emailbody += "<br /><table>" +
         "<tr><th>Service/Feature</th><td>$($n.AffectedServiceHealthStatus[0].ServiceName) / $($n.AffectedServiceHealthStatus[0].ServiceFeatureStatus[0].FeatureName)</td></tr>" +
         "<tr><th>Incident</th><td>$($n.title)</td></tr>" +
         "<tr><th>Current Status</th><td>$($n.status)</td></tr>" + 
         "<tr><th>Start Time</th><td>$($n.starttime.ToLocalTime().tostring("F"))</td></tr>"
         if ($n.EndTime -eq $null)
         {
            $emailbody += "<tr><th>End Time</th><td></td></tr>"
         }
         else
         {
            $emailbody += "<tr><th>End Time</th><td>$($n.endtime.ToLocalTime().tostring("F"))</td></tr>"
         }

         # Select only the latest updates to an event.
         $msgnbr = 0
         foreach ($m in ($n.Messages | Sort-Object -Property publishedtime))
         {
            $msgnbr++
            if ($m.publishedtime -gt $lastrundatetime)
            {
               $messagetext = $m.MessageText -replace "`n","<br />"            
               $emailbody += "<tr><th>Update Time</th><td>$($m.publishedtime.ToLocalTime().tostring("F"))</td></tr>" +
               "<tr><th>Update Number</th><td>$msgnbr</td></tr>" +
               "<tr><th>Details</th><td>$($messagetext)</td></tr>"
            }
         }

         $emailbody += "</table><br />"
      }
      $emailbody += "</body></html>"
      $parms = @{
         From        = $EmailFrom
         To          = $EmailTo
         Subject     = "Office 365 Service Health update(s)"
         SmtpServer  = $EmailSmtp
         Body        = $emailbody
         BodyAsHtml  = $true 
         ErrorAction = "Stop"
      }
      Send-MailMessage @parms 
   }

   # Update the LastRunFile with the date/time of this run.
   Set-Content -Path $LastrunFile -Value (Get-Date)
}
catch
{
   # If something goes wrong, send an e-mail to whoever needs to fix this.
   Write-LogMessage -Path $LogFile -Message "EXCEPTION:"
   Write-LogException -Path $LogFile -Exception $_ 
   $parms = @{
      From        = $EmailFrom
      To          = $EmailErrorTo
      Subject     = "Office 365 Service Health update(s) - ERROR"
      SmtpServer  = $EmailSmtp
      Body        = "An error has occurred in $scriptname."
      Attachments = $LogFile
   }
   Send-MailMessage @parms  
}

$cred = $null 
$cookie = $null
$jsonPayload = $null 

Write-LogMessage -Path $LogFile -Message "New updates = $(@($newevents).count)"
Write-LogMessage -Path $LogFile -Message "$scriptname finished."

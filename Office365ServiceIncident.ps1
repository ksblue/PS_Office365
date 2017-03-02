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
   So far, it has been this:
   https://api.admin.microsoftonline.com/shdtenantcommunications.svc

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

   $emailbody = "<html><head>" +
   "<style>" +
   “BODY{font-family: Arial; font-size: 10pt;}” +
   “TABLE{border: 1px solid black; border-collapse: collapse;}” +
   “TH{border: 1px solid black; background: #dddddd; padding: 5px; text-align: left }” +
   “TD{border: 1px solid black; padding: 5px; }” +
   “</style>” +
   "</head><body>Service Health updates posted to the Office 365 dashboard:<br /><br />"

   # Use credentials stored in file.
   $cred = Import-Credential -Path $O365Cred -ErrorAction Stop
   # Obtain cookie for authentication.
   $jsonPayload = (@{userName=$cred.username;password=$cred.GetNetworkCredential().password;} | convertto-json).tostring()
   $cookie = (Invoke-RestMethod -ContentType "application/json" -Method Post -Uri "$ServiceUrl/Register" -Body $jsonPayload).RegistrationCookie
   # Get events.
   $jsonPayload = (@{lastCookie=$cookie;locale="en-US";preferredEventTypes=@(0)} | convertto-json).tostring()
   $events = (Invoke-RestMethod -ContentType "application/json" -Method Post -Uri "$ServiceUrl/GetEvents" -Body $jsonPayload)
   $newevents = $events.events | Where-Object {$_.lastupdatedtime -gt $lastrundatetime}

   # If any new alerts are found, format an e-mail and send it.
   if (@($newevents).count -gt 0)
   {
      foreach ($n in $newevents) 
      {
         $emailbody += "<br /><table>"
         $emailbody += "<tr><th>SERVICE / FEATURE</th><th>INCIDENT</th><th>CURRENT STATUS</th><th>DATE AND TIME</th><th>DETAILS</th></tr>"
         $messages = $n.Messages | Sort-Object -Property publishedtime -Descending
         $msgcount = 0
         $totalmsgs = @($messages).Count
         foreach ($m in $messages)
         {
            $msgcount++
            $messagetext = $m.MessageText -replace "`n","<br />"
            if ($msgcount -eq 1)
            {
               $emailbody += "<tr><td rowspan=`"$totalmsgs`" style=`"vertical-align:top`">$($n.AffectedServiceHealthStatus[0].ServiceName) / <br />$($n.AffectedServiceHealthStatus[0].ServiceFeatureStatus[0].FeatureName)</td>" + 
               "<td rowspan=`"$totalmsgs`" style=`"vertical-align:top`">$($n.title)</td>" + 
               "<td rowspan=`"$totalmsgs`" style=`"vertical-align:top`">$($n.status)</td>" + 
               "<td>$($m.publishedtime.ToLocalTime().tostring("F"))</td>" + 
               "<td>$($messagetext)</td></tr>"
            }
            else
            {
               $emailbody += "<tr><td>$($m.publishedtime.ToLocalTime().tostring("F"))</td><td>$($messagetext)</td></tr>"
            }
         }
         $emailbody += "</table><br />"
      }
      $emailbody += "</body></html>"
      $parms = @{
         From        = $EmailFrom
         To          = $EmailTo
         Subject     = "Office 365 service health update"
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
   # In case something goes wrong, try to send an e-mail to whoever needs to fix this.
   Write-LogMessage -Path $LogFile -Message "EXCEPTION:"
   Write-LogException -Path $LogFile -Exception $_ 
   $parms = @{
      From        = $EmailFrom
      To          = $EmailErrorTo
      Subject     = "Office 365 service health update - ERROR"
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

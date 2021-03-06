Function Export-Credential
{
   <#

   .SYNOPSIS
      Function to export a credential object to XML format.

   .DESCRIPTION
      This function accepts a credential object and exports it to a file
      in XML format.
      
   .EXAMPLE
      Export-Credential -Credential (Get-Credential) -Path c:\SaveCredential.xml

   .PARAMETER Credential
      Credential object containing the id and password to be saved and
      exported to a file.

   .PARAMETER Path
      Path and file name for XML credential file.

   #>
   [cmdletbinding()]
   Param (
      [parameter(Mandatory=$true)]
      $Credential,
      [parameter(Mandatory=$true)]
      [string]$Path
   )

   $Credential = $Credential | Select-Object *
   $Credential.password = $Credential.Password | ConvertFrom-SecureString
   $Credential | Export-Clixml -Path $Path

}


Function Import-Credential
{
   <#

   .SYNOPSIS
      Function to import credentials to credential object from
      XML file.

   .DESCRIPTION
      This function reads a credential file saved in XML format by 
      Export-Credential and creates a system credential object.
      
   .EXAMPLE
      Import-Credential -Path c:\temp\CredFile.xml

   .PARAMETER Path
      Path and file name of XML credential file.

   #>
   [cmdletbinding()]
   Param (
      [parameter(Mandatory=$true)]
      [string]$Path
   )

   $cred = Import-Clixml -Path $Path
   $cred.password = $cred.Password | ConvertTo-SecureString
   New-Object System.Management.Automation.PSCredential($cred.username, $cred.password)

}


Function Write-HostAndLogMessage
{
   <#

   .SYNOPSIS
      Function to write message to output and log file.

   .DESCRIPTION
      This function writes a message to output and to the log file
      specified in the function path.
      
   .EXAMPLE
      Write-HostAndLogMessage -Path $LogFile -Message "This is the message."

   .PARAMETER Path
      Path and file name for log file.

   .PARAMETER Message
      Message text to write to output and log file.

   #>
   [cmdletbinding()]
   Param (
      [parameter(Mandatory=$true)]
      [string]$Path,
      [string]$Message
   )

   $timeNow = get-date -Format s
   Write-Output "$timenow - $Message"
   Add-Content -Path $Path -Value "$timenow - $Message"

}


Function Write-HostMessage
{
   <#
   .SYNOPSIS
      Function to write message to output.

   .DESCRIPTION
      This function writes a message to output.

   .EXAMPLE
      Write-HostMessage -Message "This is the message."

   .PARAMETER Message
      Message text to write to output.

   #>
   [cmdletbinding()]
   Param (
      [string]$Message
   )

   Write-Output "$(Get-Date -Format s) - $Message"

}


Function Write-LogException
{
   <#

   .SYNOPSIS
      Function to write exception information to log file.

   .DESCRIPTION
      This function writes exception information to a log file
      specified in the function call.

   .EXAMPLE
      Write-LogException -Path $LogFile -Exception $_

   .PARAMETER Path
      Path and file name for log file.

   .PARAMETER Exception

   #>
   [cmdletbinding()]
   Param (
      [parameter(Mandatory=$true)]
      [string]$Path,
      $Exception
   )

   $timeNow = get-date -Format s
   Add-Content -Path $Path -Value "$timenow - categoryinfo:  $($Exception.categoryinfo)"
   Add-Content -Path $Path -Value "$timenow - exception:  $($Exception.exception)"
   Add-Content -Path $Path -Value "$timenow - fullyqualifiederrorid:  $($Exception.fullyqualifiederrorid)"
   Add-Content -Path $Path -Value "$timenow - invocationinfo:  $($Exception.invocationinfo)"
   Add-Content -Path $Path -Value "$timenow - targetobject:  $($Exception.targetobject)"

}


Function Write-LogMessage
{
   <#

   .SYNOPSIS
      Function to write message to log file.

   .DESCRIPTION
      This function writes a message to the log file specified in
      the function call.

   .EXAMPLE
      Write-LogMessage -Path $Logfile -Message "This is the message."

   .PARAMETER Path
      Path and file name for log file.

   .PARAMETER Message
      Message text to write to log file.

   #>
   [cmdletbinding()]
   Param (
      [parameter(Mandatory=$true)]
      [string]$Path,
      [string]$Message
   )

   Add-Content -Path $Path -Value "$(Get-Date -Format s) - $Message"

}

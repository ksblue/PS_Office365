Function Close-Office365Session
{
   <#

   .SYNOPSIS
      Function to close Office 365 remote sessions.

   .DESCRIPTION
      This function will remove any remote PowerShell sessions
      created by the Open-Office365Session function that 
      are open. Use this to cleanup open sessions at the end of
      a script.

   .PARAMETER Compliance
      Switch indicating to close a Compliance session.

   .PARAMETER Exchange
      Switch indicating to close an Exchange Online session.

   .PARAMETER SkypeForBusiness
      Switch indicating to close a Skype for Business session.

   .PARAMETER All
      Swtch indicating to close all possible Office 365 sessions.

   .EXAMPLE
      Close-Office365Session -Exchange

   .NOTES
      AUTHOR:  Shaun Blue
      DATE:    February 2017

   #>
   [cmdletbinding()]
   Param (
      [switch]$Compliance,
      [switch]$Exchange,
      [switch]$SkypeForBusiness,
      [switch]$All
   )

   if ($Compliance -or $All)
   {
      if ($global:SessionCompliance -ne $null)
      {
         Remove-PSSession -Session $global:SessionCompliance
      }
   }

   if ($Exchange -or $All)
   {
      if ($global:SessionExchange -ne $null)
      {
         Remove-PSSession -Session $global:SessionExchange
      }
   }

   if ($SkypeForBusiness -or $All)
   {
      if ($global:SessionSkypeForBusiness -ne $null)
      {
         Remove-PSSession -Session $global:SessionSkypeForBusiness
      }
   }

}


Function Open-Office365Session
{
   <#

   .SYNOPSIS
      Function to open Office 365 session(s).

   .DESCRIPTION
      This function opens Office 365 sessions for the services
      requested using the credential in the object provided.
      Inidividual services or all can be requested.

   .PARAMETER Credential
      Object containing the credential id and password to be used
      to connect to Office 365 service(s).

   .PARAMETER Compliance
      Switch indicating to establish a Compliance session.

   .PARAMETER Exchange
      Switch indicating to establish an Exchange Online session.

   .PARAMATER Msol
      Switch indicating to establish MS Online Services session.

   .PARAMETER Sharepoint
      Switch indicating to establish a Sharepoint Online session.

   .PARAMETER SkypeForBusiness
      Switch indicating to establish a Skype for Business session.

   .PARAMETER All
      Swtch indicating to open all possible Office 365 sessions.

   .EXAMPLE
      Open-Office365Session -All -Credential $cred

   .EXAMPLE
      Open-Office365Session -Credential (Get-Credential) -Exchange -Msol

   .NOTES
      AUTHOR:  Shaun Blue
      DATE:    February 2017

   #>
   [cmdletbinding()]
   Param (
      [System.Management.Automation.PSCredential]$Credential,
      [switch]$Compliance,
      [switch]$Exchange,
      [switch]$Msol,
      [switch]$Sharepoint,
      [switch]$SkypeForBusiness,
      [switch]$All
   )

   if ($Credential -eq $null)
   {
      $Credential = Get-Credential
   }

   if ($Compliance -or $All)
   {
      $global:SessionCompliance = $null
      try
      {
         $parms = @{
            ConfigurationName = "Microsoft.Exchange"
            ConnectionUri     = "https://ps.compliance.protection.outlook.com/powershell-liveid"
            Credential        = $Credential
            Authentication    = "Basic"
            AllowRedirection  = $true
            WarningAction     = "SilentlyContinue"
            ErrorAction       = "Stop"
         }
         $global:SessionCompliance = New-PSSession @parms

         $parms = @{
            ModuleInfo          = (Import-PSSession -Session $global:SessionCompliance -AllowClobber -DisableNameChecking)
            Global              = $true
            DisableNameChecking = $true
            ErrorAction         = "Stop"
         }
         Import-Module @parms
      }
      catch
      {
         throw $_
      }
   }

   if ($Exchange -or $All)
   {
      $global:SessionExchange = $null
      try
      {
         $parms = @{
            ConfigurationName = "Microsoft.Exchange"
            ConnectionUri     = "https://outlook.office365.com/powershell-liveid"
            Credential        = $Credential
            Authentication    = "Basic"
            AllowRedirection  = $true
            WarningAction     = "SilentlyContinue"
            ErrorAction       = "Stop"
         }
         $global:SessionExchange = New-PSSession @parms

         $parms = @{
            ModuleInfo          = (Import-PSSession -Session $global:SessionExchange -AllowClobber -DisableNameChecking)
            Global              = $true
            DisableNameChecking = $true
            ErrorAction         = "Stop"
         }
         Import-Module @parms
      }
      catch
      {
          throw $_
      }
   }

   if ($Msol -or $All)
   {
      $modulename = "MSOnline"
      try
      {
         $parms = @{
            Name        = $modulename
            Global      = $true
            ErrorAction = "Stop"
         }
         Import-Module @parms 

         $parms = @{
            Credential  = $Credential
            ErrorAction = "Stop"
         }
         Connect-MsolService @parms
      }
      catch
      {
         if ($_.fullyqualifiederrorid -like "*ModuleNotFound*")
         {
            Write-Error "$modulename module needed to make MSOL connection.  Please install."
         }
         else
         {
            throw $_
         }
      }
   }
      
   if ($Sharepoint -or $All)
   {
      $modulename = "Microsoft.Online.SharePoint.PowerShell"
      try
      {
         $parms = @{
            Name                = $modulename
            Global              = $true
            DisableNameChecking = $true
            ErrorAction         = "Stop"
         }
         Import-Module @parms 

         $parms = @{
            Url         = "https://yourtenant-admin.sharepoint.com"
            Credential  = $Credential
            ErrorAction = "Stop"
         }
         Connect-SPOService @parms
      }
      catch
      {
         if ($_.fullyqualifiederrorid -like "*ModuleNotFound*")
         {
            Write-Error "$modulename module needed to make SharePoint Online connection.  Please install."
         }
         else
         {
            throw $_
         }
      }
   }
      
   if ($SkypeForBusiness -or $All)
   {
      $global:SessionSkypeForBusiness = $null
      $modulename = "LyncOnlineConnector"
      try
      {
         $parms = @{
            Name        = $modulename
            Force       = $true
            ErrorAction = "Stop"
         }
         Import-Module @parms 

         $parms = @{
            Credential  = $Credential
            ErrorAction = "Stop"
         }
         $global:SessionSkypeForBusiness = New-CsOnlineSession @parms 

         $parms = @{
            ModuleInfo          = (Import-PSSession -Session $global:SessionSkypeForBusiness -AllowClobber -DisableNameChecking)
            Global              = $true
            DisableNameChecking = $true
            ErrorAction         = "Stop"
         }
         Import-Module @parms 
      }
      catch
      {
         if ($_.fullyqualifiederrorid -like "*ModuleNotFound*")
         {
            Write-Error "$modulename module needed to make Skype for Business connection.  Please install."
         }
         else
         {
            throw $_
         }
      }
   }
      
}


Function Open-Office365SessionScript
{
   <#

   .SYNOPSIS
      Function to open Office 365 session(s) in a script.

   .DESCRIPTION
      This function calls the open Office 365 sessions function 
      for a script.  It allows for a script log file and retries in 
      case the inital connection attempt fails.

   .PARAMETER Credential
      Object containing the credential id and password to be used
      to connect to Office 365 service(s).

   .PARAMETER Compliance
      Switch indicating to establish a Compliance session.

   .PARAMETER Exchange
      Switch indicating to establish an Exchange Online session.

   .PARAMATER Msol
      Switch indicating to establish MS Online Services session.

   .PARAMETER Sharepoint
      Switch indicating to establish a Sharepoint Online session.

   .PARAMETER SkypeForBusiness
      Switch indicating to establish a Skype for Business session.

   .PARAMETER All
      Swtch indicating to open all possible Office 365 sessions.

   .PARAMETER LogFile
      Path for script log file.

   .PARAMETER RetriesCount
      The number of times to attempt to connect to an Office 365 
      service before quitting.

   .PARAMETER RetriesSleep
      The number of seconds to sleep in between attempts to retry
      connecting to an Office 365 service.

   .EXAMPLE
      Open-Office365Session -Credential $cred -Exchange -LogFile $LogFile
      -RetriesCount 5 -RetriesSleep 15

   .NOTES
      AUTHOR:  Shaun Blue
      DATE:    February 2017

   #>
   [cmdletbinding()]
   Param (
      [System.Management.Automation.PSCredential]$Credential,
      [switch]$Compliance,
      [switch]$Exchange,
      [switch]$Msol,
      [switch]$Sharepoint,
      [switch]$SkypeForBusiness,
      [switch]$All,
      [string]$LogFile,
      [int]$RetriesCount,
      [int]$RetriesSleep
   )

   $sessionConnected = $false 
   $connectionCount = 0
   do
   {
      $connectionCount++
      Write-LogMessage -Path $LogFile -Message "Attempting connection $connectionCount to Office 365."
      try
      {
         $parms = @{
            Compliance       = $Compliance
            Exchange         = $Exchange
            Msol             = $Msol
            Sharepoint       = $Sharepoint
            SkypeForBusiness = $SkypeForBusiness
            All              = $All
            Credential       = $Credential
            ErrorAction      = "Stop"
         }
         Open-Office365Session @parms
         $sessionConnected = $true 
         Write-LogMessage -Path $LogFile -Message "Connected to Office 365."
      }
      catch
      {
         Write-LogException -Path $LogFile -Exception $_ 
         if ($connectionCount -lt $RetriesCount)
         {
            Write-LogMessage -Path $LogFile -Message "Could not connect to Office 365.  Try again."
            Start-Sleep $RetriesSleep
         }
         else
         {
            Write-LogMessage -Path $LogFile -Message "Could not connect to Office 365.  Stopping."
         }
      }
   } until ($sessionConnected -or $connectionCount -ge $RetriesCount)

   $sessionConnected

}

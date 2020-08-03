<#

This is not working now.  It looks as if something broke.  
I am coming back to this after a few months and think that it could be cleaner.
Suggest splitting into two scripts, one that uses command line only one uses out-grid as a choice.

#>

function Uninstall-Software  
{
  <#
      .SYNOPSIS
      Uninstall unneeded or unwanted software
  #>

  [CmdletBinding(SupportsShouldProcess)]
  param
  (
    [Parameter(Mandatory,HelpMessage = 'Software DisplayName', Position = 0)]
    [String]$SoftwareName
  )

  $HKLMPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
  
  $SoftwareList = (Get-ChildItem -Path $HKLMPath   |
    Get-ItemProperty  | 
    Where-Object -FilterScript {
      $_.DisplayName -match $SoftwareName
    } | Select-Object -Property DisplayName, UninstallString)
    
    $SoftwareList = $SoftwareList | Out-GridView -PassThru -Title 'Verify Software to Remove'
    

  #$SoftwareList

  $MSIExecCount = 0
  $EXECount = 0
  foreach ($app in $SoftwareList) 
  {
  if(($app.UninstallString) -match 'MSIEXEC'){
  Write-Verbose -Message 'MSIEXEC'
  $MSIExecCount = $MSIExecCount + 1
  $MSIExecCount
  }
  elseif(($app.UninstallString) -match 'EXE'){
  Write-Verbose -Message ('EXE {0}' -f $app.UninstallString)
  $EXECount = $EXECount + 1
  $EXECount
  }
  }

    Write-Verbose -Message ('App - {0}' -f $app)
    #$App.UninstallString
    If ($app.UninstallString) 
    {
      Write-Verbose -Message ('App Uninstall - {0}' -f $app.UninstallString)

      $uninst = ($app.UninstallString)
      Write-Verbose -Message ('UninstallString - {0}'  -f  $uninst)

      $GUID = ($uninst.split('{')).trim('}')[1]
      Write-Verbose -Message ('App GUID - {0}' -f  $GUID)

      $app = Get-WmiObject -Class Win32_Product -ComputerName $env:COMPUTERNAME| Where-Object{$_ -match $GUID}
      #$app.Uninstall()
      #Start-Process -FilePath 'msiexec.exe' -ArgumentList "/X $GUID /passive" -Wait

      Write-Verbose -Message $uninst
    }
  }

Uninstall-Software -SoftwareName 'Java 8 Update 161' -Verbose
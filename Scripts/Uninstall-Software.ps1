#requires -Version 3.0 -Modules PackageManagement

function Uninstall-Software  
{
  <#
      .SYNOPSIS
      Uninstall unneeded or unwanted software
  #>
  
  [CmdletBinding(SupportsShouldProcess)]
  param
  (
    [Parameter(Mandatory = $false,HelpMessage = 'Software DisplayName', Position = 0)]
    [String[]]$RemoveSoftware
  )
  
  if(-not $RemoveSoftware)
  {
    $InstalledSoftware = Get-Package
    $RemovalSoftware = ($InstalledSoftware |
    Select-Object -Property Name |
    Out-GridView -PassThru -Title 'Verify Software to Remove')
  }
  Else
  {
    foreach($EachSoftware in $RemovalSoftware)
    {
      $RemovalSoftware = Get-Package -Name $EachSoftware
    }
  }

  
  
  foreach($EachSoftware in $RemovalSoftware)
  {
    Get-Package -Name $($EachSoftware.Name) |  Uninstall-Package -WhatIf
  }
}

Uninstall-Software -RemoveSoftware 'Cisco eReader', 'Raspberry Pi Imager' 
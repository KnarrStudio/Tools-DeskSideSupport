#requires -Version 3.0 -Modules PackageManagement

function Uninstall-Software  
{
  <#
      .SYNOPSIS
      Uninstalls software based on Parameter, File or Pick list.

      .DESCRIPTION
      This function was designed to go into another script, but will work on it's own too.
      Call the function and pass either a Name, list of names or us the file feature if there is a need to make it more dynamic.  
      If you don't know the software name then use the "Pick" parameter, but at that point you might as well use "Add-Remove Programs"

      .PARAMETER SoftwareList
      Use this to add to software right to the function call.  This can be one or many

      .PARAMETER File
      Use this if you want to pull from a .txt file that you created.

      .PARAMETER PickMenu
      Use this to select the software to remove on the fly, but at that point you might as well use "Add-Remove Programs"

      .PARAMETER Add
      This is used to help you build or edit the "File"

      .EXAMPLE
      Uninstall-Software -SoftwareList Value
      Uninstalls the software that is listed in "Value"

      .EXAMPLE
      Uninstall-Software -File Value -Add New
      Use this to create the software removal file

      Uninstall-Software -File Value -Add Update
      Use this to update a software removal file that you are already using

      .EXAMPLE
      Uninstall-Software -PickMenu
      This is used when you don't know the software or file that you want to remove by name and allows you to pick it from a list.  This helps to ensure that the script will function as expected.

      .NOTES
      Place additional notes here.

      .LINK
      https://github.com/KnarrStudio/Tools-DeskSideSupport/blob/master/Scripts/Uninstall-Software.ps1

      .INPUTS
      1. Direct input at the function call
      2. From a txt file

      .OUTPUTS
      1. The removal of software
      2. A txt file that can be used in future
  #>


  [CmdletBinding(SupportsShouldProcess,DefaultParameterSetName = 'Default')]
  param
  (
    [Parameter(Mandatory,HelpMessage = 'Add help message for user', Position = 0,ParameterSetName = 'String')]
    [String[]]$SoftwareList,
    [Parameter(Mandatory,HelpMessage = 'Add help message for user', Position = 0,ParameterSetName = 'File')]
    [ValidateScript({
          If($_ -match '.txt')
          {
            $true
          }
          Else
          {
            Throw 'Input file needs to be plan text'
          }
    })][String]$File,
    [Parameter(Mandatory,HelpMessage = 'Add help message for user', Position = 0,ParameterSetName = 'PickMenu')]
    [Switch]$PickMenu,
    [Parameter(Position = 1,ParameterSetName = 'File')]
    [ValidateSet('New','Updates')]
    [String]$Add = $null
  )
  
  if($PSBoundParameters.Values.Count -eq 0) 
  {
    Get-Help -Name Uninstall-Software
    return
  }
  
  $VerboseMessage = 'Verbose Message'
  Write-Verbose -Message ('{0}' -f $VerboseMessage)
  
  $InstalledSoftware = Get-Package
          
  switch($PSBoundParameters.Keys)
  {
    'File'
    {
      Write-Verbose -Message ('switch: {0}' -f $($PSBoundParameters.Keys))
      if(-not $Add)
      {
        $SoftwareList = Get-Content -Path $File
      }
      elseif($Add -eq 'New')
      {
        $SoftwareList = ($InstalledSoftware |
          Select-Object -Property Name |
        Out-GridView -PassThru -Title 'Software Pick List').name 
        $SoftwareList | Out-File -FilePath $File -Force
      }
      elseif($Add -eq 'Updates')
      {
        $SoftwareList = ($InstalledSoftware |
          Select-Object -Property Name |
        Out-GridView -PassThru -Title 'Software Pick List').name 
        $SoftwareList | Out-File -FilePath $File -Append
      }
    }
    'PickMenu'
    { 
      Write-Verbose -Message ('switch: {0}' -f $($PSBoundParameters.Keys))
      $SoftwareList = ($InstalledSoftware |
        Select-Object -Property Name |
      Out-GridView -PassThru -Title 'Software Pick List').name
    }
    'Default'
    {
      Write-Verbose -Message ('switch: {0}' -f $($PSBoundParameters.Keys))
    }
  }

  if(-not $Add)
  {
    Write-Verbose -Message ('foreach Software - Uninstall-Package')
    foreach($EachSoftware in $SoftwareList)
    {
      $EachSoftware = $EachSoftware.Trim()
      Write-Verbose -Message ('foreach Software: {0}' -f $EachSoftware)
      try
      {
        Get-Package -Name $EachSoftware |  Uninstall-Package -WhatIf -ErrorAction Stop
      }
      catch
      {
        Write-Warning -Message ('foreach Software: {0}' -f $EachSoftware)
      }
    }
  }
}

# Examples:
# Uninstall-Software -SoftwareList 'Microsoft Edge Dev', 'Raspberry Pi Imager', 'ITPS.OMCS.Tools' -Verbose

# Uninstall-Software -File 'D:\GitHub\KnarrStudio\Tools-StandAloneSystems\Configfiles\SoftwareList.txt' -Add Updates

# Uninstall-Software -PickMenu -Verbose

Get-Help -Online Uninstall-Software
  

# SIG # Begin signature block
# MIIFvQYJKoZIhvcNAQcCoIIFrjCCBaoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUJNOTtdYrneFN2M7E32u8EVNz
# q1KgggNLMIIDRzCCAi+gAwIBAgIQdr8M9RYBq65KZ4TmHgXXMDANBgkqhkiG9w0B
# AQsFADAnMSUwIwYDVQQDDBxBUk5FU0VOLkVSSUsuSk9ITi4xMDIwODk2MTA0MB4X
# DTIyMDIxOTEzMzgzN1oXDTIzMDIxOTEzNTgzN1owJzElMCMGA1UEAwwcQVJORVNF
# Ti5FUklLLkpPSE4uMTAyMDg5NjEwNDCCASIwDQYJKoZIhvcNAQEBBQADggEPADCC
# AQoCggEBALlVtaiHeQFsBgEhZ5ZG9plThVyZDU9vBLUJDhhkkralrdADvFnEIPDK
# 1qehJAKdIPB3uBBWw5rXRdaNWJsiHso6PL9CZIOXskT1ft6pSWOgHoGV/RzBhnk8
# HkwKneYDxx9tp3MBLT6AGiCZTm2Qu+E1J8jFFFusW1JDCx0TNQ1bP/DtS/TfrnyS
# EqxuVHvFjaI3XfZG0JGec36Jult8wOHr1Iv/YQRKXTyU63pi0gOddfD6zyrkuqL2
# kCP/+mgLtRx56zrfloc+tI9MMIRW0KH7e6I+yWFi+SPfaXvT3hmHksdumwVoADzl
# 2jFieaCY8Ah8ZPh5nLuevUmXYhi+JR0CAwEAAaNvMG0wDgYDVR0PAQH/BAQDAgeA
# MBMGA1UdJQQMMAoGCCsGAQUFBwMDMCcGA1UdEQQgMB6CHEFSTkVTRU4uRVJJSy5K
# T0hOLjEwMjA4OTYxMDQwHQYDVR0OBBYEFAzIoPmuXXn/OOLkRKM9x7lv8udKMA0G
# CSqGSIb3DQEBCwUAA4IBAQArNEJElgTjiFKHLfmHwhcJkj8TAHiYURP+y0hFMehT
# ey7Z5Gst2UY3/PftdQ0TyRQpmMpdEjspC+kJDLtlzq6ynqFhhnZRcD85xh89eVJT
# C8ArY925wr3rJiuA2ktIXbF5oBSZw/MKSa/GOSlcsj0C7iWIUJ8zjI9r422LButq
# +OA5Kzc2yPGCD19bu0TGqffLm6eB3j+phPm3Aya/DkpA2Xv0Kex+dxmr75Cg67Go
# SArwX6PBluYglQgwwNziD9GdQJ0zNAW0ynd5y+/FIKU+BZCeEv6EbIdCzub4nB7N
# cdhCgeQFqmQ8pRXIO0ke0biEctoRV5cRKvZIRCrEHiZpMYIB3DCCAdgCAQEwOzAn
# MSUwIwYDVQQDDBxBUk5FU0VOLkVSSUsuSk9ITi4xMDIwODk2MTA0AhB2vwz1FgGr
# rkpnhOYeBdcwMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAA
# MBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgor
# BgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBSAemsKk7FmfkkRJKGudUemqQOv+jAN
# BgkqhkiG9w0BAQEFAASCAQCLhJvUSFzsCcV8+Kfz6DGaUOEe+aIY10xRRgoDYQ2N
# MjIKXHl1I2QI/abF/OgBPjqPMgGRm0NRUWBYjgQUIdSz7fHYz3vu3/BBvmzYIlpm
# ad0Lr9HcRZqPkObhqAUk6sxUD3xnww7M9+VlsuOQDETKthx7A1Rmu2mdnIwbPT1G
# lCe1pPZQxNcwy7DhZp/5VOXUy2cvjv6dkpfn6tnrW6FxOQLOVnVqgefaHwXD2EZ7
# GALFH/BFujNrKNXrdfWiKKmePE1xXEtQD+M7kqbjY85nKORRUQ48o4rAsWHuIJ3/
# HJm7eT9fbZIgnWgTyYAWFZW3LdYxsY3rUzAuh5s2jXXu
# SIG # End signature block

#requires -Version 3.0  -RunAsAdministrator
<#PSScriptInfo

    .VERSION 1.0

    .GUID ebdf766e-f61c-49a4-a764-1102ed0ac4dc

    .AUTHOR Erik@home

    .COMPANYNAME KnarrStudio

    .COPYRIGHT 2021 KnarrStudio

    .RELEASENOTES
    Quick script to automate a manual process

#>
<#
    .SYNOPSIS
    Automates the steps used to repair Windows Updates. 

    .DESCRIPTION
    Automates the steps used to repair Windows Updates. 
    The steps can be found in the Advanced section of the "Troubleshoot problems updating Windows 10" page. See link
    
    PowerShells the following steps:
    net.exe stop wuauserv 
    net.exe stop cryptSvc 
    net.exe stop bits 
    net.exe stop msiserver 
    ren C:\Windows\SoftwareDistribution -NewName SoftwareDistribution.old 
    ren C:\Windows\System32\catroot2 -NewName Catroot2.old 
    net.exe start wuauserv 
    net.exe start cryptSvc 
    net.exe start bits 
    net.exe start msiserver 


    .EXAMPLE
    As an admin, run: Repair-WindowsUpdate

    .NOTES
    

    .LINK
    https://support.microsoft.com/help/4089834?ocid=20SMC10164Windows10

#>

$UpdateServices = 'wuauserv', 'cryptSvc', 'bits', 'msiserver'
$RenameFiles = "$env:windir\SoftwareDistribution", "$env:windir\System32\catroot2"
function Set-ServiceState 
{
  <#
      .SYNOPSIS
      Start or stop Services based on "Stop / Start" switch
  #>
  param(
    [Parameter(Mandatory,HelpMessage = 'list of services that to stop or start')][string[]]$services,
    [Switch]$Stop,
    [Switch]$Start
  )
  if ($Stop)
  {
    ForEach ($service in $services)
    {
      try
      {
        Stop-Service -InputObject $service -PassThru
      }
      catch
      {
        Stop-Service -InputObject $service -Force
      }
    }
  }
  if ($Start)
  {
    ForEach ($service in $services)
    {
      Start-Service -InputObject $service
    }
  }
}
function Rename-Files
{
  <#
      .SYNOPSIS
      Renames files to ".old"
  #>
  param(
    [Parameter(Mandatory,HelpMessage = 'list of files to be renamed with ".old"')][string[]]$Files
  )
  ForEach($File in $Files)
  {
    Rename-Item -Path $File -NewName ('{0}.old' -f $File) -Force
  }
}
Set-ServiceState -services $UpdateServices -Stop
Rename-Files -Files $RenameFiles
Set-ServiceState -services $UpdateServices -Start

# SIG # Begin signature block
# MIIFvQYJKoZIhvcNAQcCoIIFrjCCBaoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUJyM9N9FyzDNwBdgTK0PgdFxI
# stugggNLMIIDRzCCAi+gAwIBAgIQdr8M9RYBq65KZ4TmHgXXMDANBgkqhkiG9w0B
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
# BgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBRG2vxxzfiEsU2AeEDO4aeB4cw5izAN
# BgkqhkiG9w0BAQEFAASCAQCN0iXKuEuqa7WBnHqLX3valjCPJczsh/2YshIPRhdS
# j9JopGbmPbQD+TGCGeHhpB1aaVlvWk8p4mxWHgrQphW0bLjbWDwO57RkGNKQsSnV
# LjujEHeVU23GvaH7JxgAXpf1VUSFIgbuXksVIzvUU0UcsfHgONx5fwI6ezwAb1jh
# QbyUtPHR3kKM3cHbFDzm5Cx/zpGt8t7M2altFWJPlIkb3FeMQXCSo72AaONPlFzp
# weG2Fkk7WtESwmYz+8mh7tFXJpyTCdEnyBwaVpKYcL1Hmha9Ia+lEq49zPhjWZUy
# IyndsmSWajmiteUfg5gK6YYtpVfN7ted2eQFwqjtHjL0
# SIG # End signature block

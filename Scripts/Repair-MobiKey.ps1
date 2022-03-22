#requires -Version 3.0
#requires -runasadministrator

Write-Host 'Edit the $MobikeyCertsPath variable and remove this line.' ; Return
$MobikeyCertsPath = '\\Networkshare\mobikey\certs'

function Repair-MobiKey 
{
  <#
      .SYNOPSIS
      Repairs the Mobkey Installation
   
      .EXAMPLE
      Repair-MobiKey -MobikeyCertsPath '\\Networkshare\mobikey\certs'
      
      .EXAMPLE
      Repair-MobiKey -MobikeyCertsPath '\\Networkshare\mobikey\certs' -verbose
         
      .NOTES
      You will need to make the following changes

      .LINK
      https://github.com/KnarrStudio/Tools-DeskSideSupport
      The first link is opened by Get-Help -Online Repair-MobiKey
  #>

  [CmdletBinding(HelpUri = 'https://github.com/KnarrStudio/Tools-DeskSideSupport',
  ConfirmImpact = 'Medium')]
  [OutputType([String])]
  Param
  (
    # Param1 help description
    [Parameter(Mandatory,HelpMessage = 'UNC or other path to where the certs are stored', Position = 0)]
    [string]$MobikeyCertsPath 
  )
  
  Begin
  {
    $ServiceName = 'Route1*'
    $MobikeyLocalPath = "$env:ProgramData\Mobikey\certs"
  }
  Process
  {
    # First Stop the Mobikey Service 
    if((Get-Service -Name $ServiceName).Status -ne 'Stopped')
    {
      Write-Verbose -Message ('Mobikey Service is Running')
      Stop-Service -Name $ServiceName
      (Get-Service -Name $ServiceName).WaitForStatus('Stopped')
      Write-Verbose -Message ('Mobikey Service has been stopped')
    }
    
    # Delete the old files out of the ProgramData Dir
    if (Test-Path -Path $MobikeyLocalPath)
    {
      Write-Verbose -Message ('Delete old Certificate files')
      Get-ChildItem -Path $MobikeyLocalPath -Recurse | Remove-Item -Force #-WhatIf
    }
  
    # Copy files from the File share to the local drive
    if (Test-Path -Path $MobikeyCertsPath)
    {
      Write-Verbose -Message ('Copy Files from Network to Localhost')
      Copy-Item -Path $MobikeyCertsPath -Include *.* -Destination $MobikeyLocalPath -Force
    }
    # Start the Mobikey service 
    Write-Verbose -Message ('Start Mobikey Service')
    Start-Service -Name $ServiceName
  }
  End
  {
    $ServiceStatus = (Get-Service -Name $ServiceName).Status
    Write-Verbose -Message ('Mobikey Service is {0}' -f $ServiceStatus)
  }
}

Repair-MobiKey -MobikeyCertsPath $MobikeyCertsPath -Verbose





# SIG # Begin signature block
# MIIFvQYJKoZIhvcNAQcCoIIFrjCCBaoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUNMlJzlIP+v6apRiayoJ4oC3P
# l1+gggNLMIIDRzCCAi+gAwIBAgIQdr8M9RYBq65KZ4TmHgXXMDANBgkqhkiG9w0B
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
# BgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBSjmp+5lRZtmJdjeecWtCgcfp2WNjAN
# BgkqhkiG9w0BAQEFAASCAQCfMl81RYq7s+lCF2GaNnex17h8ezDuDYPlEOWw9OtP
# Kcp12xFl4cQE8MeOZ3wj3C0ZoZqLispNkgnCQZ3VwkJj/8a32eatECN8sy1m4TKo
# +GCWjkbNe1xFRVUwdxqEFbh8BYNeyoNTnPXlf/bjy1I8pfcNU3au0ToTXoqe2AGz
# DyN8ZywBP0GdzOjZBVfBpzwL+W9Sgx+phuwq7kRQLd948CzQFyHOAWPgLsbOW1TT
# jTkP8/kKK4ZhVWpMCqHuxv5JZ6zA8uQpLmCHIF+3QQv8mVp6RiokYgcLXRAYaqCN
# AlGaX54AxF+k2S1L2XOqnKHhD9ds1xW2rEMvqJJ6EviN
# SIG # End signature block

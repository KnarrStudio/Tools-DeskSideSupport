#requires -Version 3.0 -Modules NetTCPIP
#
function Start-RemoteUpdates
{

  param
  (
    [Parameter(Mandatory,  Position = 0)]
    [string]
    $ComputerList,
    
    [Parameter(Mandatory , Position = 1)]
    [string]
    $RoboSource,
    
    [Parameter(Mandatory, Position = 2)]
    [string]
    $computer = 'generikstorage',
    
    [Parameter(Mandatory, Position = 3)]
    $UpdateList
  )
  
  Begin{
    # region Variables 
    $testport = '5985' #,'5986'
    $computernames = Get-Content -Path $ComputerList
    # endregion Variables
  
    # region Scriptblock
    $RoboCopy = {
      param(
        [Parameter(Mandatory)]$RoboSource,
        [Parameter(Mandatory)]$ComputerName #= 'generikstorage'
      ) 
      ('{0}\system32\robocopy.exe {1} \\{2}\c$ *.* /r:1 /w:1' -f $env:windir, $RoboSource, $ComputerName)
    }
    # endregion Scriptblock
  }
  Process
  {FOREACH($computer in $computernames) 
    {
      Write-Verbose -Message ('Working on {0}' -f $computer)
    
      if(Test-Path -Path ('\\{0}\c$' -f $computer))
      {
        Invoke-Command -ScriptBlock $RoboCopy -ComputerName $computer -Session -RoboSource $RoboSource
      }
      If (Test-Connection -ComputerName $computer -Count 1) 
      {
        Write-Verbose -Message 'Checking winrm service'
        Get-Service -Name winrm -ComputerName $computer | Start-Service
      
        If (Test-NetConnection -ComputerName yahoo.com  -Port $testport -InformationLevel Quiet) 
        {
          Write-Verbose -Message 'Starting remote session'
          $PsSession = New-PSSession -ComputerName $computer
          Invoke-Command -Session $PsSession -ScriptBlock {
            for ($i = 0; $i -lt $UpdateList.Count; $i++)
            {
              msiexec.exe /update $($UpdateList[$i]) /quiet /norestart
            } 
          }
          Write-Verbose -Message 'Removing PSSession'
          Remove-PSSession -Session $PsSession
        }
      }
    }
  }
  End
  {}
}

    
$SplatUpdates = @{
  ComputerList = 'c:\Temp\templist.txt'
  RoboSource   = 'C:\Temp\Patching'
  UpdateList   = @('C:\KB3203392_conv-x-none.msp', 'C:\KB3203393_word-x-none.msp', 'C:\KB3213555_mso-x-none.msp', 'C:\KB4011045_word-x-none.msp')
}
  
Start-RemoteUpdates @SplatUpdates





# SIG # Begin signature block
# MIIFvQYJKoZIhvcNAQcCoIIFrjCCBaoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUw3eN9ZDyOklnyN4G5uC4jDyl
# GvygggNLMIIDRzCCAi+gAwIBAgIQdr8M9RYBq65KZ4TmHgXXMDANBgkqhkiG9w0B
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
# BgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBSyw+fpAPTX4BnF1/qlCg50PYdH8DAN
# BgkqhkiG9w0BAQEFAASCAQCccvYfFccx4hIGBqd1F8BS9xg4W1lI68kqL92Zx021
# OXjKSNtCo7HvP3oupeVm8IwDzuxCtsufe6BzMp8EigEMzeTC4NLRolosIgh9BgrQ
# CWA6D0qE43BqHYNowFN+wkDEhSku8ufrypLF1lsWeeUFAOglJWbqcszL1c4m1qnW
# Av+smxjrRzU5tZo7AT5ecUrWj4wiYqV03qZS6H989jFHdWAFCaX3knMc/fUhM7Cg
# nSI0lBOHjT0ILa6HXrVwMDDetu0HWfGSh2l6XinbYHeasYSauDQsWNZkBLib2mrz
# oCNYQlH5Wu349fEJCbZfb2/2wvxrW42YkxvP76c1uGNZ
# SIG # End signature block

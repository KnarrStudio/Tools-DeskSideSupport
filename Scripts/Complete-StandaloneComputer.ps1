#requires -Version 3.0
#Requires -RunAsAdministrator
#Requires -Modules Microsoft.PowerShell.LocalAccounts, Microsoft.PowerShell.Utility, PackageManagement

<#
    .SYNOPSIS
    This script is to help out with the building of the stand alone systems.
    It should be able to do the following by the time it is completed.

    .DESCRIPTION
    1. Add new user[s] 
    2. Add users to specific groups
    3. Uninstall unneeded software
    4. Build a standard folder structure that will be used for the care and feeding
    5. Make registry changes as required by the function of the device *Later versions

    .EXAMPLE
    Complete-StandaloneComputer.ps1

#>

Begin{
  $NewUsers = @{
    RFVuser    = @{
      FullName            = 'RFV User'
      Description         = 'Standard local Account'
      AccountGroup        = 'Users'
      AccountNeverExpires = $true
      Password            = '1qaz@WSX3edc$RFV'
    }
    RFVAdmin   = @{
      FullName            = 'RFV Admin'
      Description         = 'Local Admin Account for IT'
      AccountGroup        = 'Administrators'
      AccountNeverExpires = $true
      Password            = 'RFV@dm!nP@$$!!'
    }
    CRTech     = @{
      FullName            = 'TV Tech User'
      Description         = 'Standard TV Account'
      AccountGroup        = 'Users'
      AccountNeverExpires = $true
      Password            = '1qaz@WSX3edc$RFV'
    }
    CRAdmin    = @{
      FullName            = 'TV Administrator'
      Description         = 'TV Admin Account'
      AccountGroup        = 'Administrators'
      AccountNeverExpires = $true
      Password            = '1qaz@WSX3edc$RFV'
    }
    NineOneOne = @{
      FullName            = '911'
      Description         = 'Emergancy Access PW in "KeePass"'
      AccountGroup        = 'Administrators'
      AccountNeverExpires = $true
      Password            = '1qaz@WSX3edc$RFV'
    }
  }
  $NewFolderInfo = [ordered]@{
    CyberUpdates = @{
      Path       = 'D:\CyberUpdates'
      ACLGroup   = 'Administrators'
      ACLControl = 'Full Control'
      ReadMeText = 'This is the working folder for the monthly updates and scanning.'
      ReadMeFile = 'README.TXT'
    }
    ScanReports  = @{
      Path       = 'D:\CyberUpdates\ScanReports'
      ACLGroup   = 'Administrators'
      ACLControl = 'Full Control'
      ReadMeText = 'This is where the "IA" scans engines and reports will be kept.'
      ReadMeFile = 'README.TXT'
    }
  }
  
  $NewNicConfigSplat = @{
    InterfaceAlias = 'Ethernet' 
    IpAddress = '192.168.86.92' 
    PrefixLength = 24 
    DefaultGateway = '192.168.86.1'
  }
   $SetNicConfigSplat = @{
    InterfaceAlias = 'Ethernet' 
    IpAddress = '192.168.86.92' 
    PrefixLength = 24 
    }
  
  
  #>
  # Variables
  $ConfigFilesFolder = 'D:\GitHub\KnarrStudio\Tools-StandAloneSystems\Configfiles'

  #$NewFolderInfo = Import-PowerShellDataFile  -Path ('{0}\NewFolderInfo.psd1' -f $ConfigFilesFolder)
  #$NewUsers = Import-PowerShellDataFile  -Path ('{0}\NewUsers.psd1' -f $ConfigFilesFolder)
  $NewGroups = Import-PowerShellDataFile  -Path ('{0}\NewGroups.psd1' -f $ConfigFilesFolder)
  
  #$NewGroups = @('RFV_Users', 'RFV_Admins', 'TestGroup', 'Guests')
  #$Password911 = Read-Host "Enter a 911 Password" -AsSecureString
  #$PasswordUser = Read-Host -Prompt 'Enter a User Password' -AsSecureString
  #$CurrentUsers = Get-LocalUser
  #$CurrentGroups = Get-LocalGroup
  
  # House keeping
  function New-Folder {
    <#
        .SYNOPSIS
        Short Description
    #>
    param
    (
      [Parameter(Mandatory, Position = 0)]
      [Object]$NewFolderInfo
    )
    foreach($ItemKey in $NewFolderInfo.keys)
    {
      $NewFolderPath = $NewFolderInfo.$ItemKey.Path
      $NewFile = $NewFolderInfo.$ItemKey.ReadMeFile
      $FileText = $NewFolderInfo.$ItemKey.ReadMeText
      If(-not (Test-Path -Path $NewFolderPath))
      {
        New-Item -Path $NewFolderPath -ItemType Directory -Force #-WhatIf
        $FileText | Out-File -FilePath ('{0}\{1}' -f $NewFolderPath, $NewFile) #-WhatIf
      }
    }
  }
  function Add-LocalRFVGroups    {
    <#
        .SYNOPSIS
        Short Description
    #>
    param
    (
      [Parameter(Mandatory, Position = 0)]
      [Object]$GroupList
    )
    foreach($NewGroup in $GroupList.Keys)
    {
      $LocalUserGroups = (Get-LocalGroup).name

      if($($GroupList[$NewGroup].Name) -notin $LocalUserGroups)
      {
        Write-Verbose -Message ('Creating {0} Account' -f $NewGroup)
        New-LocalGroup -Name $($GroupList[$NewGroup].Name) -Description $($GroupList[$NewGroup].Description) -WhatIf
      }
    }  
  }
  function Add-RFVLocalUsers    {
    <#
        .SYNOPSIS
        Add new local users to the computer

    #>
    param
    (
      [Parameter(Mandatory, Position = 0)]
      [Object]$UserName,
      [Parameter(Mandatory, Position = 1)]
      [Object]$UserInfo
    )
    $LocalUserNames = (Get-LocalUser).name
    $SecurePassword = ConvertTo-SecureString -String ($UserInfo.Password) -AsPlainText -Force
    $UserDescription = ($UserInfo.Description)
    $UserFullName = ($UserInfo.FullName)
    If ($UserName -notin $LocalUserNames)
    {
      Write-Verbose -Message ('Creating {0} Account' -f $UserFullName)
      New-LocalUser -Name $UserName -Description $UserDescription -FullName $UserFullName  -Password $SecurePassword -WhatIf -ErrorAction SilentlyContinue
    }
  }
  function Add-RFVUsersToGroups  {
    <#
        .SYNOPSIS
        Short Description
    #>
    param
    (
      [Parameter(Mandatory, Position = 0)]
      [Object]$UserName,
      [Parameter(Mandatory, Position = 1)]
      [Object]$UserInfo
    )
    $UserPrimaryGroup = ($UserInfo.AccountGroup) 
    $GroupMembership = (Get-LocalGroupMember -Group $UserPrimaryGroup | Where-Object -Property ObjectClass -EQ -Value User).Name
    if($GroupMembership -match $UserName)
    {
      Write-Verbose -Message ('Adding "{0}" Account to {1} group' -f $($UserInfo.FullName), $UserPrimaryGroup) -Verbose
      Add-LocalGroupMember -Group $UserPrimaryGroup -Member $UserName -ErrorAction Stop
    }
  }
  function Uninstall-Software    {
    <#
        .SYNOPSIS
        Uninstalls software based on Parameter, File or Pick list.

        .LINK
        https://github.com/KnarrStudio/Tools-StandAloneSystems/blob/master/Scripts/Uninstall-Software.ps1
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
  function Set-WallPaper    {
    <#
        .SYNOPSIS
        Change Desktop picture/background
    #>
    param
    (
      [Parameter(Position = 0)]
      #[string]$BackgroundSource = "$env:HOMEDRIVE\Windows\Web\Wallpaper\Windows\img0.jpg",
      #[string]$BackupgroundDest = "$env:PUBLIC\Pictures\BG.jpg"
      [string]$BackgroundSource = "$env:HOMEDRIVE\Windows\Web\Wallpaper\Windows\img0.jpg",
      [string]$BackupgroundDest = "$env:HOMEDRIVE\Windows\Web\Wallpaper\Windows\img0.jpg"
    )
    If ((Test-Path -Path $BackgroundSource) -eq $false)
    {
      Copy-Item -Path $BackgroundSource -Destination $BackupgroundDest -Force -WhatIf
    }
    Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop\' -Name Wallpaper -Value $BackupgroundDest 
    Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop\' -Name TileWallpaper -Value '0'
    Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop\' -Name WallpaperStyle -Value '10' -Force
  }
  function Set-CdLetterToX    {
    <#
        .SYNOPSIS
        Test for a CD and change the drive Letter to X:
    #>
    param
    (
      [Parameter(Position = 0)]
      [Object]$CdDrive = (Get-WmiObject -Class Win32_volume -Filter 'DriveType=5'|   Select-Object -First 1)
    )
    if($CdDrive)
    {
      if(-not (Test-Path -Path X:\))
      {
        Write-Verbose -Message ('Changing{0} drive letter to X:' -f ([string]$CdDrive.DriveLetter))
        $CdDrive | Set-WmiInstance -Arguments @{
          DriveLetter = 'X:'
        }
      }
    }
  }
  function Set-NicConfiguration {
    <#
        .SYNOPSIS
        Test for a CD and change the drive Letter to X:
    #>
    param
    (
      [Parameter(Position = 0)]
      [hashtable]$NewNetConfigSplat,
      [Parameter(Position = 1)]
      [hashtable]$SetNetConfigSplat
    )
    
    
    $NIC = Get-NetIPInterface -AddressFamily IPv4 | sort -Property InterfaceMetric | select -f 1 
    if($NIC.Dhcp -eq 'Enabled'){
      Set-NetIPInterface  -InterfaceAlias $NIC.InterfaceAlias -AddressFamily IPv4 -Dhcp Disabled
    }
    Rename-NetAdapter -Name $NIC.InterfaceAlias -NewName 'Ethernet'
    
    
    #New-NetIPAddress @NewNetConfigSplat #-WhatIf
    Set-NetIPAddress  #@SetNetConfigSplat #-WhatIf
    
  }
}
Process{
  
  # Creates new folder structure 
  New-Folder -NewFolderInfo $NewFolderInfo
  
  # Changes the CD/DVD drive letter to "X" for standardization.
  Set-CdLetterToX
  
  # Sets the Wallpaper to a specific flavor for all users.
  Set-WallPaper
  
  # Adds new groups 
  Add-LocalRFVGroups -GroupList $NewGroups 
   
  # Adds new users based on the "NewUsers.psd1" file
  ForEach ($UserName in $NewUsers.Keys) 
  {
    $UserInfo = $NewUsers[$UserName]
    Add-RFVLocalUsers -UserName $UserName -userinfo $UserInfo
    Add-RFVUsersToGroups -UserName $UserName -UserInfo $UserInfo
  }
  
  Get-Help -Online -Name Uninstall-Software
  #Uninstall-Software -File 'C:\Temp\SoftwareList.txt' -Add New

  # Sets the network configuration on the NIC
  #Set-NicConfiguration $NicConfigSplat
}
End{
}

# SIG # Begin signature block
# MIIFvQYJKoZIhvcNAQcCoIIFrjCCBaoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUFY7nQePoBtIbjxWO2fiAqiWR
# HB2gggNLMIIDRzCCAi+gAwIBAgIQdr8M9RYBq65KZ4TmHgXXMDANBgkqhkiG9w0B
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
# BgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBR1kr2OD7yNjYEY0ogQDL1FT/HwjTAN
# BgkqhkiG9w0BAQEFAASCAQB/s5PciEzLz7he4yK+Xe1MWqdu8WEIK17pLldlfakY
# AWQPORty+dbGNWh4wqN2KM7nHbf5D9wlwS/UqGh6UxsxeBhfm7/ugiAKW2/agUQY
# Xjog0xtT4UH7p26wXKPSc29J86gGY2o1VvCx0Yb8mD6DgaA0uJ7AgVbSHBpGBI+K
# /OXL+RhC71Csw6kwLrbe2MMnChKHxv9oUKhirpgZwYVUhPQ06DuOZPt360xZMhFc
# lHvqhgKdxdxwxqpWRXf5gUpL3tPv+DhJU5QY76+I+NqqHVJJdJv3za+dMqnj9ITM
# vLK/AEpbjOePkzTQeXKqZ6nIyBhYbLOPrqZp955N82Yx
# SIG # End signature block

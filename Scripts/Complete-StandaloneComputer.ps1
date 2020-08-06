#requires -Version 3.0
#Requires -RunAsAdministrator
#Requires -Modules Microsoft.PowerShell.LocalAccounts
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
# Add new users
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
      Path       = 'C:\CyberUpdates'
      ACLGroup   = 'Administrators'
      ACLControl = 'Full Control'
      ReadMeText = 'This is the working folder for the monthly updates and scanning.'
      ReadMeFile = 'README.TXT'
    }
    ScanReports  = @{
      Path       = 'C:\CyberUpdates\ScanReports'
      ACLGroup   = 'Administrators'
      ACLControl = 'Full Control'
      ReadMeText = 'This is where the "IA" scans engines and reports will be kept.'
      ReadMeFile = 'README.TXT'
    }
  }
  # Variables
  $NewGroups = @('RFV_Users', 'RFV_Admins', 'TestGroup', 'Guests')
  # $Password911 = Read-Host "Enter a 911 Password" -AsSecureString
  #$PasswordUser = Read-Host -Prompt 'Enter a User Password' -AsSecureString
  #$CurrentUsers = Get-LocalUser
  #$CurrentGroups = Get-LocalGroup
  # House keeping
  function New-Folder  
  {
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
        New-Item -Path $NewFolderPath -ItemType Directory -Force -WhatIf
        $FileText | Out-File -FilePath $NewFolderPath"\"$NewFile -WhatIf
      }
    }
  }
  function Add-LocalRFVGroups  
  {
    <#
        .SYNOPSIS
        Short Description
    #>
    param
    (
      [Parameter(Mandatory, Position = 0)]
      [String]$NewGroups
    )
    $LocalUserGroups = (Get-LocalGroup).name
    ForEach($NewGroup in $NewGroups)
    {
      if($NewGroup -notin $LocalUserGroups)
      {
        Write-Verbose -Message ('Creating {0} Account' -f $NewGroup)
        New-LocalGroup -Name $NewGroup -Description $NewGroup -WhatIf
      }
    }
  }
  function Add-RFVLocalUsers  
  {
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
  function Add-RFVUsersToGroups
  {
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
  function Uninstall-Software  
  {
    <#
        .SYNOPSIS
        Uninstall unneeded or unwanted software
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param
    (
      [Parameter(Mandatory, Position = 0)]
      [String]$SoftwareName
    )
    function Get-SoftwareList
    {
      param
      (
        [Parameter(Mandatory, ValueFromPipeline, HelpMessage = 'Data to filter')]
        [Object]$InputObject
      )
      process
      {
        if ($InputObject.DisplayName -match $SoftwareName)
        {
          $InputObject
        }
      }
    }
    $SoftwareList = $null
    $app = (Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall  |
      Get-ItemProperty | 
      Get-SoftwareList |
    Select-Object -Property DisplayName, UninstallString)
    #$SoftwareList
    ForEach ($app in $SoftwareList) 
    {
      #$App.UninstallString
      If ($app.UninstallString) 
      {
        $uninst = ($app.UninstallString)
        $GUID = ($uninst.split('{')[1]).trim('}')
        Start-Process -FilePath 'msiexec.exe' -ArgumentList "/X $GUID /passive" -Wait
        #Write-Host $uninst
      }
    }
  }
  function Set-WallPaper  
  {
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
  function Set-CdLetterToX  
  {
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
}
Process{
  New-Folder -NewFolderInfo $NewFolderInfo
  ForEach ($UserName in $NewUsers.Keys) 
  {
    $UserInfo = $NewUsers[$UserName]
    Add-LocalRFVGroups -NewGroups ($UserInfo.AccountGroup)
    Add-RFVLocalUsers -UserName $UserName -userinfo $UserInfo
    Add-RFVUsersToGroups -UserName $UserName -UserInfo $UserInfo
  }
  #Add-LocalRFVGroups -NewGroups $NewGroups
  #Add-RFVLocalUsers -NewUsers $NewUsers
  #Set-CdLetterToX
  #Set-WallPaper
  #Uninstall-Software
}
End{}

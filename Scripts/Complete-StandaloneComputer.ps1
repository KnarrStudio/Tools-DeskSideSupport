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
  
  # Variables
  #$NewFolderInfo = Import-PowerShellDataFile  -Path "D:\GitHub\KnarrStudio\Tools-StandAloneSystems\Configfiles\NewFolderInfo.psd1"
  #$NewUsers = Import-PowerShellDataFile  -Path "D:\GitHub\KnarrStudio\Tools-StandAloneSystems\Configfiles\NewUsers.psd1"
  #$NewGroup = Import-PowerShellDataFile  -Path "D:\GitHub\KnarrStudio\Tools-StandAloneSystems\Configfiles\NewGroups.psd1"
  
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
        New-Item -Path $NewFolderPath -ItemType Directory -Force
        $FileText | Out-File -FilePath ('{0}\{1}' -f $NewFolderPath, $NewFile)
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
  
  <#  ForEach ($UserName in $NewUsers.Keys) 
      {
      $UserInfo = $NewUsers[$UserName]
      Add-LocalRFVGroups -NewGroups ($UserInfo.AccountGroup)
      Add-RFVLocalUsers -UserName $UserName -userinfo $UserInfo
      Add-RFVUsersToGroups -UserName $UserName -UserInfo $UserInfo
  }#>

  #Set-CdLetterToX
  
  #Set-WallPaper
  
  Get-Help -Online -Name Uninstall-Software
  #Uninstall-Software -File 'C:\Temp\SoftwareList.txt' -Add New

}
End{}

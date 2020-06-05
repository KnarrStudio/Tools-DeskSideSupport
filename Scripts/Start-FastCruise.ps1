#requires -Version 3.0
function Start-FastCruise
{
  <#
      .SYNOPSIS
      Short Description
      .DESCRIPTION
      Detailed Description
      .EXAMPLE
      Start-FastCruise
      explains how to use the command
      can be multiple lines
      .EXAMPLE
      Start-FastCruise
      another example
      can have as many examples as you like
  #>
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory = $false, Position = 0)]
    [Object]
    $FastCruiseReport = 'C:\temp\Reports\FastCruise_2020-May.csv' #("$env:HOMEDRIVE\Temp\Reports\FastCruise_{0}.csv" -f $(Get-Date -Format yyyy-MMMM))
  )
   
  Begin
  {
    
    $YearMonth = Get-Date -Format yyyy-MMMM
    # $FastCruiseReport = ("$env:HOMEDRIVE\Temp\Reports\FastCruise_{0}.csv" -f $YearMonth)
    
    Write-Verbose -Message ('Testing the Report Path: {0}' -f $FastCruiseReport)
    if(-not (Test-Path -Path $FastCruiseReport))
    {
      Write-Verbose -Message 'Test Failed.  Creating the Report now.'
      $null = New-Item -Path $FastCruiseReport -ItemType File
    } 
    
    $Ans = 'z'
    $Orange = 'Blue'
    $CompImport = Import-Csv -Path $FastCruiseReport
    
    # Select last status of system.
    Write-Verbose -Message "Getting last status of workstation: $env:COMPUTERNAME"
    $LatestStatus = $CompImport |
    Where-Object -FilterScript {
      $PSItem.Color -match $Orange
    } |
    Select-Object -Last 1 
    
    function Script:Get-Location
    {
      <#
          .SYNOPSIS
          Short Description
          .DESCRIPTION
          Detailed Description
      #>
      [CmdletBinding()]
      param
      (
        [Parameter(Mandatory = $false, Position = 0)]
        $Desk       = @('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I')
      )
      
      $Location = [ordered]@{
        Department = [ordered]@{
          MCDO = [ordered]@{
            Building = [ordered]@{
              AV29  = [ordered]@{
                Room = @(
                  8, 
                  9, 
                  10, 
                  11, 
                  12, 
                  13, 
                  14, 
                  15, 
                  16, 
                  17, 
                  18, 
                  19, 
                  20
                )
              }
              AV34  = [ordered]@{
                Room = @(
                  1, 
                  6, 
                  7, 
                  8, 
                  201, 
                  202, 
                  203, 
                  204, 
                  205
                )
              }
              ELC3  = [ordered]@{
                Room = @(
                  100, 
                  101, 
                  102, 
                  103, 
                  104, 
                  105, 
                  106, 
                  107
                )
              }
              ELC31 = [ordered]@{
                Room = @(
                  1
                )
              }
              ELC32 = [ordered]@{
                Room = @(
                  1
                )
              }
              ELC33 = [ordered]@{
                Room = @(
                  1
                )
              }
              ELC34 = [ordered]@{
                Room = @(
                  1
                )
              }
              ELC35 = [ordered]@{
                Room = @(
                  1
                )
              }
              ELC36 = [ordered]@{
                Room = @(
                  1
                )
              }
            }
          }
          CA   = [ordered]@{
            Building = [ordered]@{
              AV29 = [ordered]@{
                Room = @(
                  1, 
                  2, 
                  3, 
                  4, 
                  5, 
                  6, 
                  7, 
                  23, 
                  24, 
                  25, 
                  26, 
                  27, 
                  28, 
                  29, 
                  30
                )
              }
              AV34 = [ordered]@{
                Room = @(
                  1, 
                  2, 
                  3, 
                  200, 
                  214
                )
              }
              AV44 = [ordered]@{
                Room = @(
                  1
                )
              }
              AV45 = [ordered]@{
                Room = @(
                  1
                )
              }
              AV46 = [ordered]@{
                Room = @(
                  1
                )
              }
              AV47 = [ordered]@{
                Room = @(
                  1, 
                  2
                )
              }
              AV48 = [ordered]@{
                Room = @(
                  1, 
                  2
                )
              }
            }
          }
          PRO  = [ordered]@{
            Building = [ordered]@{
              AV34 = [ordered]@{
                Room = @(
                  210, 
                  211, 
                  212, 
                  213
                )
              }
              ELC4 = [ordered]@{
                Room = @(
                  1, 
                  100, 
                  101, 
                  102, 
                  103, 
                  104, 
                  105, 
                  106, 
                  107
                )
              }
            }
          }
          TJ   = [ordered]@{
            Building = [ordered]@{
              AV34 = [ordered]@{
                Room = @(
                  2, 
                  3, 
                  13, 
                  11
                )
              }
              ELC2 = [ordered]@{
                Room = @(
                  1
                )
              }
            }
          }
        }
      }
      
      [string]$Script:LclDept = $Location.Department.Keys | Out-GridView -Title 'Department' -PassThru
      [string]$Script:LclBuild = $Location.Department[$LclDept].Building.Keys | Out-GridView -Title 'Building' -PassThru
      [string]$Script:LclRm = $Location.Department[$LclDept].Building[$LclBuild].Room | Out-GridView -Title 'Room' -PassThru
      [string]$Script:LclDesk = $Desk | Out-GridView -Title 'Desk' -PassThru
    } # End Function
    
    function Open-Form
    {
      [CmdletBinding(HelpUri = 'https://github.com/KnarrStudio/Tools-DeskSideSupport',
      ConfirmImpact = 'Medium')]
      [OutputType([String])]
      Param
      (
        # Param1 help description
        [Parameter(Mandatory,HelpMessage = 'The data label', Position = 0)]
        [string]  $FormLabel
      )
      Add-Type -AssemblyName System.Windows.Forms
      Add-Type -AssemblyName System.Drawing
      
      $form = New-Object -TypeName System.Windows.Forms.Form
      $form.Text = 'Computer Description'
      $form.Size = New-Object -TypeName System.Drawing.Size -ArgumentList (300, 200)
      $form.StartPosition = 'CenterScreen'
      
      $okButton = New-Object -TypeName System.Windows.Forms.Button
      $okButton.Location = New-Object -TypeName System.Drawing.Point -ArgumentList (75, 120)
      $okButton.Size = New-Object -TypeName System.Drawing.Size -ArgumentList (75, 23)
      $okButton.Text = 'OK'
      $okButton.DialogResult = [Windows.Forms.DialogResult]::OK
      $form.AcceptButton = $okButton
      $form.Controls.Add($okButton)
      
      $cancelButton = New-Object -TypeName System.Windows.Forms.Button
      $cancelButton.Location = New-Object -TypeName System.Drawing.Point -ArgumentList (150, 120)
      $cancelButton.Size = New-Object -TypeName System.Drawing.Size -ArgumentList (75, 23)
      $cancelButton.Text = 'Cancel'
      $cancelButton.DialogResult = [Windows.Forms.DialogResult]::Cancel
      $form.CancelButton = $cancelButton
      $form.Controls.Add($cancelButton)
      
      $label = New-Object -TypeName System.Windows.Forms.Label
      $label.Location = New-Object -TypeName System.Drawing.Point -ArgumentList (10, 20)
      $label.Size = New-Object -TypeName System.Drawing.Size -ArgumentList (280, 20)
      $label.Text = $FormLabel
      $form.Controls.Add($label)
      
      $textBox = New-Object -TypeName System.Windows.Forms.TextBox
      $textBox.Location = New-Object -TypeName System.Drawing.Point -ArgumentList (10, 40)
      $textBox.Size = New-Object -TypeName System.Drawing.Size -ArgumentList (260, 20)
      $form.Controls.Add($textBox)
      
      $form.Topmost = $true
      
      $form.Add_Shown({
          $textBox.Select()
      })
      $result = $form.ShowDialog()
      
      if ($result -eq [Windows.Forms.DialogResult]::OK)
      {
        $x = $textBox.Text
        $x
      }
    }

    function Get-McAfeeVersion { 
      [CmdletBinding()]
      param ([Object]$Computer) 
      $ProductVer = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SOFTWARE\McAfee\DesktopProtection').GetValue('szProductVer') 
      $EngineVer = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SOFTWARE\McAfee\AVEngine').GetValue('EngineVersionMajor') 
      $DatVer = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SOFTWARE\McAfee\AVEngine').GetValue('AVDatVersion') 
 
      $ComputerStat['McAfee Product version'] = $ProductVer
      $ComputerStat['McAfee Engine version'] = $EngineVer
      $ComputerStat['McAfee Dat version'] = $DatVer
    }
    
    Function Get-InstalledSoftware
    {
      [cmdletbinding(DefaultParameterSetName = 'SortList',SupportsPaging = $true)]
      Param(
        
        [Parameter(Mandatory = $true,HelpMessage = 'At least part of the software name to test', Position = 0,ParameterSetName = 'SoftwareName')]
        [String[]]$SoftwareName,
        [Parameter(ParameterSetName = 'SortList')]
        [Parameter(ParameterSetName = 'SoftwareName')]
        [ValidateSet('InstallDate', 'DisplayName','DisplayVersion')] 
        [String]$SortList = 'InstallDate'
        
      )
      
      Begin { 
        $SoftwareOutput = @()
        $InstalledSoftware = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*)
      }
      
      Process {
        Try 
        {
          if($SoftwareName -eq $null) 
          {
            $SoftwareOutput = $InstalledSoftware |
            #Sort-Object -Descending -Property $SortList |
            Select-Object -Property @{
              Name = 'Date Installed'
              Exp  = {
                $_.Installdate
              }
            }, @{
              Name = 'Version'
              Exp  = {
                $_.DisplayVersion
              }
            }, DisplayName #, UninstallString 
          }
          Else 
          {
            foreach($Item in $SoftwareName)
            {
              $SoftwareOutput += $InstalledSoftware |
              Where-Object -Property DisplayName -Match -Value $Item |
              Select-Object -ExpandProperty @{
                Name = 'Version'
                Exp  = {
                  $_.DisplayVersion
                }
              }#, DisplayName # , UninstallString 
            }
          }
        }
        Catch 
        {
          # get error record
          [Management.Automation.ErrorRecord]$e = $_
          
          # retrieve information about runtime error
          $info = New-Object -TypeName PSObject -Property @{
            Exception = $e.Exception.Message
            Reason    = $e.CategoryInfo.Reason
          }
          
          # output information. Post-process collected info, and log info (optional)
          $info
        }
      }
      
      End{ 
        Switch ($SortList){
          'DisplayName' 
          {
            $SoftwareOutput |
            Sort-Object -Property displayname
          }
          'DisplayVersion' 
          {
            $SoftwareOutput |
            Sort-Object -Property 'Version'
          }
          'UninstallString'
          {

          }
          default  
          {
            $SoftwareOutput |
            Sort-Object -Property 'Date Installed'
          } # 'InstallDate'
          
        }
      }
    }
    
    $McAfee = Get-InstalledSoftware -SoftwareName 'McAfee'
    $MozillaVersion = Get-InstalledSoftware -SoftwareName 'Mozilla Firefox'
    $AdobeVersion = Get-InstalledSoftware -SoftwareName Adobe 
    $LatestWSUSupdate = (New-Object -ComObject 'Microsoft.Update.AutoUpdate'). Results 
    
    Write-Verbose -Message ('Latest Status: {0}' -f $LatestStatus)
    #$LatestStatus | Format-List -Property ComputerName, Building, Room, @{Label='Desk A to Z from left';Expression={$_.Desk}},Notes
    $LatestStatus | Format-List -Property ComputerName, Building, Room, Desk, Notes
    
    Write-Verbose -Message 'Setting up the ComputerStat hash'
    $ComputerStat = [ordered]@{
      'ComputerName'  = "$env:COMPUTERNAME"
      'UserName'      = "$env:USERNAME"
      'Date'          = "$(Get-Date)"
      'Firefox Version' = $MozillaVersion
      'Adobe Version' = $AdobeVersion
      'McAfee Product version' = ''
      'McAfee Engine version' = ''
      'McAfee Dat version' = ''
      'WSUS Search Success' = ''
      'WSUS Install Success' = ''
      'Department'    = ''
      'Building'      = ''
      'Room'          = ''
      'Desk'          = ''
      'Phone'         = ''
      'Notes'         = ''
    }
    
    
  }
  
  Process
  {
    $LatestStatus
    Do
    {
      $Ans = Read-Host -Prompt 'Is this correct? Y/N'
    }
    While(($Ans -ne 'N') -and ($Ans -ne 'Y')) 
    
    # $PowerPointResult
    # $AdobeResult
    
    
    if($Ans -eq 'N')
    {
      Get-Location
      Write-Verbose -Message ('OSD-OMC-{0}-{1}-{2}{3}' -f $LclDept, $LclBuild, $LclRm, $LclDesk)
      
      $ComputerStat['Department'] = $LclDept     
      $ComputerStat['Building'] = $LclBuild
      $ComputerStat['Room'] = $LclRm
      $ComputerStat['Desk'] = $LclDesk
    }
    else
    {
      $ComputerStat['Building'] = $($LatestStatus.Building)
      $ComputerStat['Room'] = $($LatestStatus.Room)
      $ComputerStat['Desk'] = $($LatestStatus.Desk)
      $ComputerStat['Color'] = $($LatestStatus.Color)
      $ComputerStat['MS Office Test'] = $PowerPointResult = 7
      $ComputerStat['Adobe Test'] = $AdobeResult = 8
    }
  }
  
  END
  {
    
    if($Mcafee){Get-McAfeeVersion -Computer $env:COMPUTERNAME}
    else{
      $ComputerStat['McAfee Product version'] = 'Not Found'
      $ComputerStat['McAfee Engine version'] = 'Not Found'
      $ComputerStat['McAfee Dat version'] = 'Not Found'
    }

    
    $ComputerStat['WSUS Search Success'] = $LatestWSUSupdate.LastSearchSuccessDate
    $ComputerStat['WSUS Install Success'] = $LatestWSUSupdate.LastInstallationSuccessDate

    
    $Phone = Open-Form -FormLabel 'Nearest Phone Number (or last 4)'
    $ComputerStat['Phone'] = $Phone
    
    [string]$Notes = Read-Host -Prompt 'Notes'
    $ComputerStat['Notes'] = $Notes
    
    $ComputerStat  |
    ForEach-Object -Process {
      [pscustomobject]$_
    } |
    Export-Csv -Path $FastCruiseReport -NoTypeInformation -Append
    
    
    Write-Verbose -Message 'Show Last Cruisers'
    Get-Content -Path $FastCruiseReport |
    Select-Object -Last 5 |
    Format-Table
  }
}

Start-FastCruise -Verbose
$r = Import-Csv -Path 'C:\temp\Reports\FastCruise_2020-May.csv'
$r |
Sort-Object -Property Department, Building |
ForEach-Object -Process {
('{0} = OSD-OMC-{1}-{2}-{3}{4}' -f $_.ComputerName, $_.Department, $_.Building, $_.Room, $_.Desk)
}



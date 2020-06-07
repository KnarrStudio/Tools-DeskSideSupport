#requires -Version 3.0

#Edit the splats to customize the script
$FastCruiseSplat = @{
  FastCruiseReportPath = 'C:\temp\Reports'
  FastCruiseFile       = 'FastCruise.csv' 
  Verbose              = $true
}
$PDFApplicationTestSplat = @{
  TestFile    = 'O:\OMC-S\IT\Scripts\FastCruise\FastCruiseTestFile.pdf'
  TestProgram = "${env:ProgramFiles(x86)}\Adobe\Acrobat 2015\Acrobat\Acrobat.exe"
  ProcessName = 'Acrobat'
}
$PowerPointApplicationTestSplat = @{
  TestFile    = 'O:\OMC-S\IT\Scripts\FastCruise\FastCruiseTestFile.pptx'
  TestProgram = "${env:ProgramFiles(x86)}\Microsoft Office\Office16\POWERPNT.EXE"
  ProcessName = 'POWERPNT'
}
<#$WordpadApplicationTestSplat = @{
    TestFile    = "$env:windir\DtcInstall.log"
    TestProgram = "$env:ProgramFiles\Windows NT\Accessories\wordpad.exe"
    ProcessName = 'wordpad'
    }
#>

function Start-FastCruise
{
  param
  (
    [Parameter(Mandatory, Position = 0)]
    [String]$FastCruiseReportPath,
    [Parameter(Mandatory, Position = 0)]
    [ValidateScript({
          If($_ -match '.csv')
          {
            $true
          }
          Else
          {
            Throw 'Input file needs to be CSV'
          }
    })][String]$FastCruiseFile

  )
   
  Begin
  {
    Write-Verbose -Message 'Setup Variables'
    $Ans = 'z'
        
    Write-Verbose -Message 'Setup Report' 
    $YearMonth = Get-Date -Format yyyy-MMMM
    $FastCruiseFile = [String]$($FastCruiseFile.Replace('.',"_$YearMonth."))
    $FastCruiseReport = ('{0}\{1}' -f $FastCruiseReportPath, $FastCruiseFile)
    Write-Verbose -Message ('{0}' -f $FastCruiseReport) 
    
    Write-Verbose -Message ('Testing the Report Path: {0}' -f $FastCruiseReport)
    if(-not (Test-Path -Path $FastCruiseReport))
    {
      Write-Verbose -Message 'Test Failed.  Creating the Report now.'
      $null = New-Item -Path $FastCruiseReport -ItemType File
    } 
    function Start-ApplicationTest
    {
      param
      (
        [Parameter(Mandatory, Position = 0)]
        [string]$FunctionTest,
        [Parameter(Mandatory, Position = 1)]
        [string]$TestFile,
        [Parameter(Mandatory, Position = 2)]
        [string]$TestProgram,
        [Parameter(Mandatory, Position = 3)]
        [string]$ProcessName
      )
      $DescriptionLists = [Ordered]@{
        FunctionResult = 'Good', 'Failed'
      }

      if($FunctionTest -eq 'Y')
      {
        try
        {
          Write-Verbose -Message ('Attempting to open {0} with {1}' -f $TestFile, $ProcessName)
          Start-Process -FilePath $TestProgram -ArgumentList $TestFile

          Write-Host -Object ('The Fast Cruise Script will continue after {0} has been closed.' -f $ProcessName) -BackgroundColor Red -ForegroundColor Yellow
          Write-Verbose -Message ('Wait-Process: {0}' -f $ProcessName)
          Wait-Process -Name $ProcessName
        
          $TestResult = $DescriptionLists.FunctionResult | Out-GridView -Title $ProcessName -OutputMode Single
        }
        Catch
        {
          Write-Verbose -Message 'TestResult: Failed'
          $TestResult = $DescriptionLists.FunctionResult[1]
        }
      }
      else
      {
        Write-Verbose -Message 'TestResult: Bypassed'
        $TestResult = 'Bypassed'
      }
      Return $TestResult
    }
     
    function Get-LastComputerStatus
    {
      <#
          .SYNOPSIS
          Return the last status of system based on what was in the current Fast Cruise Report
      #>
      param
      (
        [Parameter(Mandatory, Position = 0)]
        [String]$FastCruiseReport
      )
  
      Write-Verbose -Message 'Importing the Fast Cruise Report'
      $CompImport = Import-Csv -Path $FastCruiseReport
  
      # Select last status of system.
      Write-Verbose -Message "Getting last status of workstation: $env:COMPUTERNAME"
      
      try
      {
        $LatestStatus = $CompImport |
        Where-Object -FilterScript {
          $PSItem.ComputerName -eq $env:COMPUTERNAME
        } |
        Select-Object -Last 1 
        if($LatestStatus -eq $null)
        {
          Write-Output -InputObject 'Unable to find an existing record for this system.'
          $Ans = 'NoHistory'
        }
      }
      Catch
      {
        # get error record
        [Management.Automation.ErrorRecord]$e = $_

        # retrieve information about runtime error
        $info = New-Object -TypeName PSObject -Property @{
          Exception = $e.Exception.Message
        }
      
        # output information. Post-process collected info, and log info (optional)
        $info
      }
      Return $LatestStatus
    }
    
    function Script:Get-Location
    {
      <#
          .SYNOPSIS
          Get-Location of workstation
      #>
      [CmdletBinding()]
      
      [Object[]]$Desk       = @('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I')
      
      $Location = [Ordered]@{
        Department = [Ordered]@{
          MCDO = [Ordered]@{
            Building = [Ordered]@{
              AV29  = [Ordered]@{
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
              AV34  = [Ordered]@{
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
              ELC3  = [Ordered]@{
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
              ELC31 = [Ordered]@{
                Room = @(
                  1
                )
              }
              ELC32 = [Ordered]@{
                Room = @(
                  1
                )
              }
              ELC33 = [Ordered]@{
                Room = @(
                  1
                )
              }
              ELC34 = [Ordered]@{
                Room = @(
                  1
                )
              }
              ELC35 = [Ordered]@{
                Room = @(
                  1
                )
              }
              ELC36 = [Ordered]@{
                Room = @(
                  1
                )
              }
            }
          }
          CA   = [Ordered]@{
            Building = [Ordered]@{
              AV29 = [Ordered]@{
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
              AV34 = [Ordered]@{
                Room = @(
                  1, 
                  2, 
                  3, 
                  200, 
                  214
                )
              }
              AV44 = [Ordered]@{
                Room = @(
                  1
                )
              }
              AV45 = [Ordered]@{
                Room = @(
                  1
                )
              }
              AV46 = [Ordered]@{
                Room = @(
                  1
                )
              }
              AV47 = [Ordered]@{
                Room = @(
                  1, 
                  2
                )
              }
              AV48 = [Ordered]@{
                Room = @(
                  1, 
                  2
                )
              }
            }
          }
          PRO  = [Ordered]@{
            Building = [Ordered]@{
              AV34 = [Ordered]@{
                Room = @(
                  210, 
                  211, 
                  212, 
                  213
                )
              }
              ELC4 = [Ordered]@{
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
          TJ   = [Ordered]@{
            Building = [Ordered]@{
              AV34 = [Ordered]@{
                Room = @(
                  2, 
                  3, 
                  13, 
                  11
                )
              }
              ELC2 = [Ordered]@{
                Room = @(
                  1
                )
              }
            }
          }
        }
      }
      
      [string]$Script:LclDept = $Location.Department.Keys | Out-GridView -Title 'Department' -OutputMode Single
      [string]$Script:LclBuild = $Location.Department[$LclDept].Building.Keys | Out-GridView -Title 'Building' -OutputMode Single
      [string]$Script:LclRm = $Location.Department[$LclDept].Building[$LclBuild].Room | Out-GridView -Title 'Room' -OutputMode Single
      [string]$Script:LclDesk = $Desk | Out-GridView -Title 'Desk' -OutputMode Single
    } # End Location-Function
    
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

    function Get-McAfeeVersion 
    { 
      param ([Parameter(Mandatory)][Object]$Computer) 
      $ProductVer = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SOFTWARE\McAfee\DesktopProtection').GetValue('szProductVer') 
      $EngineVer = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SOFTWARE\McAfee\AVEngine').GetValue('EngineVersionMajor') 
      $DatVer = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SOFTWARE\McAfee\AVEngine').GetValue('AVDatVersion') 
 
      $ComputerStat['McAfee Product version'] = $ProductVer
      $ComputerStat['McAfee Engine version'] = $EngineVer
      $ComputerStat['McAfee Dat version'] = $DatVer
    }
    
    Function Get-InstalledSoftware
    {
      [cmdletbinding(DefaultParameterSetName = 'SortList',SupportsPaging)]
      Param(
        
        [Parameter(Mandatory,HelpMessage = 'At least part of the software name to test', Position = 0,ParameterSetName = 'SoftwareName')]
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
    

    $AdobeVersion = Get-InstalledSoftware -SoftwareName Adobe 
    <#bookmark Software Versions #>
    $MozillaVersion = (Get-InstalledSoftware -SoftwareName 'Mozilla Firefox').version
    $McAfeeVersion  = (Get-InstalledSoftware -SoftwareName 'McAfee Agent').version

    <#bookmark Windows Updates #>    
    $LatestWSUSupdate = (New-Object -ComObject 'Microsoft.Update.AutoUpdate'). Results 
    
    
    Write-Verbose -Message 'Setting up the ComputerStat hash'
    $ComputerStat = [ordered]@{
      'ComputerName'         = "$env:COMPUTERNAME"
      'UserName'             = "$env:USERNAME"
      'Date'                 = "$(Get-Date)"
      'Firefox Version'      = $MozillaVersion
      'Adobe Version'        = $AdobeVersion
      'McAfee Product version' = ''
      'McAfee Engine version' = ''
      'McAfee Dat version'   = ''
      'WSUS Search Success'  = $LatestWSUSupdate.LastSearchSuccessDate
      'WSUS Install Success' = $LatestWSUSupdate.LastInstallationSuccessDate
      'Department'           = ''
      'Building'             = ''
      'Room'                 = ''
      'Desk'                 = ''
    }
  } #End BEGIN region
  
  Process
  {
   
    Write-Verbose -Message 'Getting Last Status recorded'
    $LatestStatus = (Get-LastComputerStatus -FastCruiseReport $FastCruiseReport) 
    Write-Output -InputObject 'Latest Status'
    $LatestStatus | Select-Object -Property Computername, Department, Building, Room, Desk

    # Location Varification
    Do
    {
      $Ans = Read-Host -Prompt 'Is this information correct? Y/N'
    }
    While(($Ans -ne 'N') -and ($Ans -ne 'Y')) 
    
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
    }
    
    <#bookmark Application Test #> 
    Do
    {
      $FunctionTest = Read-Host -Prompt 'Perform Function Tests (MS Office and Adobe) Y/N'
    }
    While(($FunctionTest -ne 'N') -and ($FunctionTest -ne 'Y')) 
    # Start-ApplicationTest -FunctionTest $FunctionTest @WordpadApplicationTestSplat -Verbose
    
    $AdobeResult = Start-ApplicationTest -FunctionTest $FunctionTest @PDFApplicationTestSplat
    $PowerPointResult = Start-ApplicationTest -FunctionTest $FunctionTest @PowerPointApplicationTestSplat
    
    $ComputerStat['MS Office Test'] = $PowerPointResult
    $ComputerStat['Adobe Test'] = $AdobeResult
    
    <#bookmark Windows Update Status #> 
    $ComputerStat['WSUS Search Success'] = $LatestWSUSupdate.LastSearchSuccessDate
    $ComputerStat['WSUS Install Success'] = $LatestWSUSupdate.LastInstallationSuccessDate

    <#bookmark Local phone number #> 
    $Phone = Open-Form -FormLabel 'Nearest Phone Number (or last 4)'
    $ComputerStat['Phone'] = $Phone
    
    <#bookmark Fast cruise notes #>
    [string]$Notes = Open-Form -FormLabel 'Notes'
    $ComputerStat['Notes'] = $Notes
  } #End PROCESS region
  
  END
  {
    
    if($McAfeeVersion)
    {
      Get-McAfeeVersion -Computer $env:COMPUTERNAME
    }
    else
    {
      $ComputerStat['McAfee Product version'] = 'Not Found'
      $ComputerStat['McAfee Engine version'] = 'Not Found'
      $ComputerStat['McAfee Dat version'] = 'Not Found'
    }

    $ComputerStat  |
    ForEach-Object -Process {
      [pscustomobject]$_
    } |
    Export-Csv -Path $FastCruiseReport -NoTypeInformation -Append
    
    
    Write-Output -InputObject 'The information recorded'
    $ComputerStat | Format-Table

    <#bookmark Fast cruising shipmates #>
    Write-Output -InputObject 'Fast Cruise shipmates'
    Import-Csv -Path $FastCruiseReport |
    Select-Object -Last 4 -Property Date, Username, Building, Room, Phone |
    Format-Table
  } #End END region
}

Clear-Host #Clears the console.  This shouldn't be needed once the script can be run directly from PS
Start-FastCruise @FastCruiseSplat # Make sure you have updated and completed the "Splats" at the top of the script



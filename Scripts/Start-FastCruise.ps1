#requires -Version 3.0

#Edit the splats to customize the script
$FastCruiseSplat = @{
  FastCruiseReportPath = 'C:\temp\Report'
  FastCruiseFile       = 'FastCruise.csv' 
  Verbose              = $true
}
$PDFApplicationTestSplat = @{
  TestFile    = '\\FastCruise\FastCruiseTestFile.pdf'
  TestProgram = "${env:ProgramFiles(x86)}\Adobe\Acrobat 2015\Acrobat\Acrobat.exe"
  ProcessName = 'Acrobat'
}
$PowerPointApplicationTestSplat = @{
  TestFile    = '\\FastCruise\FastCruiseTestFile.pptx'
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
    #$LocationVerification = $null
        
    Write-Verbose -Message 'Setup Report' 
    $YearMonth = Get-Date -Format yyyy-MMMM
    $FastCruiseFile = [String]$($FastCruiseFile.Replace('.',('_{0}.' -f $YearMonth)))
    $FastCruiseReport = ('{0}\{1}' -f $FastCruiseReportPath, $FastCruiseFile)
    #$FastCruiseReport = "C:\temp\Reports\FastCruise_Test.csv"
    Write-Verbose -Message ('{0}' -f $FastCruiseReport) 
    
    Write-Verbose -Message ('Testing the Report Path: {0}' -f $FastCruiseReportPath)
    if(-not (Test-Path -Path $FastCruiseReportPath))
    {
      Write-Verbose -Message 'Test Failed.  Creating the Directory now.'
      $null = New-Item -Path $FastCruiseReportPath -ItemType Directory -Force
    } 
    Write-Verbose -Message ('Testing the Report Path: {0}' -f $FastCruiseReport)
    if(-not (Test-Path -Path $FastCruiseReport))
    {
      Write-Verbose -Message 'Test Failed.  Creating the File now.'
      $null = New-Item -Path $FastCruiseReport -ItemType File -Force
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
      Write-Verbose -Message ('Enter Function: {0}' -f $PSCmdlet.MyInvocation.MyCommand.Name)
      if($FunctionTest -eq 'Yes')
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
    } # End ApplicationTest-Function
     
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
  
      Write-Verbose -Message ('Enter Function: {0}' -f $PSCmdlet.MyInvocation.MyCommand.Name)
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
    } # End ComputerStatus-Function
    
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
    
    function Show-VbForm
    {
      [cmdletbinding(DefaultParameterSetName = 'Message')]
      param(
        [Parameter(Mandatory=$false,Position = 0,ParameterSetName = 'Message')]
        [Switch]$YesNoBox,
        [Parameter(Mandatory=$False,Position = 0,ParameterSetName = 'Input')]
        [Switch]$InputBox,
        [Parameter(Mandatory=$true,Position = 1)]
        [string]$Message,
        [Parameter(Mandatory=$false,Position = 2)]
        [string]$TitleBar = 'Fast Cruise'
    
      )
      Write-Verbose -Message ('Enter Function: {0}' -f $PSCmdlet.MyInvocation.MyCommand.Name)
      
      Add-Type -AssemblyName Microsoft.VisualBasic
      
      if($InputBox){
        $Response = [Microsoft.VisualBasic.Interaction]::InputBox($Message, $TitleBar)
      }
      if($YesNoBox){  
        $Response = [Microsoft.VisualBasic.Interaction]::MsgBox($Message, 'YesNo,SystemModal,MsgBoxSetForeground', $TitleBar)
      }
      Return $Response
    } # End VbForm-Function
    
    Function Get-InstalledSoftware
    {
      [cmdletbinding(SupportsPaging)]
      Param(
        
        [Parameter(Mandatory=$false,HelpMessage = 'At least part of the software name to test', Position = 0)]
        [String[]]$SoftwareName,
        [ValidateSet('DisplayName','DisplayVersion')] 
        [String]$SelectParameter
      )
      
      Begin { 
        Write-Verbose -Message ('Enter Function: {0}' -f $PSCmdlet.MyInvocation.MyCommand.Name)
      
        $SoftwareOutput = @()
        $InstalledSoftware = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*)
      }
      
      Process {
        Try 
        {
          if($SoftwareName -eq $null) 
          {
            $SoftwareOutput = $InstalledSoftware |
            Select-Object -Property Installdate, DisplayVersion, DisplayName #, UninstallString 
          }
          Else 
          {
            foreach($Item in $SoftwareName)
            {
              $SoftwareOutput += $InstalledSoftware |
              Where-Object -Property DisplayName -Match -Value $Item |
              Select-Object -Property Installdate, DisplayVersion, DisplayName #, UninstallString 
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
          }
          
          # output information. Post-process collected info, and log info (optional)
          $info
        }
      }
      
      End{  
        Switch ($SelectParameter){
          'DisplayName' 
          {
            $SoftwareOutput.displayname
          }
          'DisplayVersion' 
          {
            $SoftwareOutput.DisplayVersion
          }
          default  
          {
            $SoftwareOutput
          
          }
        }
      }
    } # End InstalledSoftware-Function
    
    <#bookmark Software Versions #>
    $AdobeVersion = Get-InstalledSoftware -SoftwareName Adobe -SelectParameter DisplayVersion
    $MozillaVersion = Get-InstalledSoftware -SoftwareName 'Mozilla Firefox' -SelectParameter DisplayVersion
    $McAfeeVersion  = Get-InstalledSoftware -SoftwareName 'McAfee Agent' -SelectParameter DisplayVersion
    #$TestSoftware  = Get-InstalledSoftware -SoftwareName 'Vmware' -SelectParameter DisplayVersion
    

    <#bookmark Windows Updates #>    
    $LatestWSUSupdate = (New-Object -ComObject 'Microsoft.Update.AutoUpdate'). Results 
    
    
    Write-Verbose -Message 'Setting up the ComputerStat hash'
    $ComputerStat = [ordered]@{
      'ComputerName'         = "$env:COMPUTERNAME"
      'UserName'             = "$env:USERNAME"
      'Date'                 = "$(Get-Date)"
      'Firefox Version'      = $MozillaVersion
      'Adobe Version'        = $AdobeVersion
      'McAfee Version'       = $McAfeeVersion
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
    #Write-Output -InputObject 'Latest Status'
    #$LatestStatus | Select-Object -Property Computername, Department, Building, Room, Desk

    <#bookmark Location Verification #>
    $ComputerLocation = (@'

ComputerName: (Assest Tag)
- {0}

Department:
- {1}

Building:
- {2}

Room:
- {3}

Desk:
- {4}
          
'@ -f $LatestStatus.ComputerName, $LatestStatus.Department, $LatestStatus.Building, $LatestStatus.Room, $LatestStatus.Desk)

    $LocationVerification = Show-VbForm -YesNoBox -Message $ComputerLocation
    
    if($LocationVerification -eq 'No')
    {
      Get-Location
      Write-Verbose -Message ('Computer Description: ABC-DEF-{0}-{1}-{2}{3}' -f $LclDept, $LclBuild, $LclRm, $LclDesk)
      
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
    $FunctionTest = Show-VbForm -YesNoBox -Message 'Perform Applicaion Tests (MS Office and Adobe)?'     
    
    $AdobeResult = Start-ApplicationTest -FunctionTest $FunctionTest @PDFApplicationTestSplat
    $PowerPointResult = Start-ApplicationTest -FunctionTest $FunctionTest @PowerPointApplicationTestSplat
    
    $ComputerStat['MS Office Test'] = $PowerPointResult
    $ComputerStat['Adobe Test'] = $AdobeResult
    
    <#bookmark Windows Update Status #> 
    $ComputerStat['WSUS Search Success'] = $LatestWSUSupdate.LastSearchSuccessDate
    $ComputerStat['WSUS Install Success'] = $LatestWSUSupdate.LastInstallationSuccessDate

    <#bookmark Local phone number #> 
    $Phone = Show-VbForm -InputBox -Message 'Nearest Phone Number (or last 4):'
    $ComputerStat['Phone'] = $Phone
    
    <#bookmark Fast cruise notes #>
    [string]$Notes = Show-VbForm -InputBox -Message 'Notes about this cruise:'
    $ComputerStat['Notes'] = $Notes
  } #End PROCESS region
  
  END
  {
    
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



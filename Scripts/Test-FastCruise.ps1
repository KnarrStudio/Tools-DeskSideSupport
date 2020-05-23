#requires -Version 4.0


Begin
{

  $YearMonth = Get-Date -Format yyyy-MMMM
  $FastCruiseReport = ("$env:HOMEDRIVE\Temp\Reports\FastCruise_{0}.csv" -f $YearMonth)
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

  function script:Get-ComputerLocation 
  {
    $Department = @('CA', 'MCDO', 'OCA', 'PRO', 'TJ')

    $DescriptionCA = [ordered]@{
      Building = @('AV34', 'AV29', 'ELC1', 'ELC2', 'ELC5', 'ELC6', 'ELC7', 'ELC47', 'ELC48', 'ELC53')
      Room     = @(1, 2, 3, 4, 5, 6, 7, 23, 24, 25, 26, 27, 28, 29, 30, 200, 214)
      Desk     = @('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I')
    }

    $DescriptionMCDO = [ordered]@{
      Building = @('AV34', 'AV29', 'ELC3', 'ELC7', 'ELC31', 'ELC32', 'ELC33', 'ELC34', 'ELC35', 'ELC36')
      Room     = @(1, 6, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 100, 101, 102, 103, 104, 105, 106, 107, 201, 202, 203, 204, 205)
      Desk     = @('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I')
    }
  
    $DescriptionPRO = [ordered]@{
      Building = @('AV34', 'ELC4')
      Room     = @(1, 100, 101, 102, 103, 104, 105, 106, 107, 210, 211, 212, 213)
      Desk     = @('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I')
    }
  
    <#    $ComputerDescription  = [ordered]@{
        Building = ''
        Room     = ''
        Desk     = ''
    }#>

      $DescriptionSwitch = $Department | Out-GridView -PassThru
    $ComputerStat.Department = $DescriptionSwitch 
    
    Switch($DescriptionSwitch){
      'MCDO' 
      {
        $DescriptionMCDO.Keys |ForEach-Object -Process {
          $mylocation = $DescriptionMCDO.$_ | Out-GridView -Title $_ -PassThru
          $ComputerStat.$_ = $mylocation
          Write-Verbose -Message $ComputerStat
        }
      } # End MCDO

      'PRO' 
      {
        $DescriptionPRO.Keys |ForEach-Object -Process {
          $mylocation = $DescriptionPRO.$_ | Out-GridView -Title $_ -PassThru
          $ComputerStat.$_ = $mylocation
          Write-Verbose -Message $ComputerStat
        }
      } # End PRO

      Default 
      {
        $DescriptionCA.Keys |ForEach-Object -Process {
          $mylocation = $DescriptionCA.$_ | Out-GridView -Title $_ -PassThru
          $ComputerStat.$_ = $mylocation
          Write-Verbose -Message $ComputerStat
        }
      } # End CA
    } #End Switch $r
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
        } #'InstallDate'
      
      }
    }
  }

  $MozillaVersion = Get-InstalledSoftware -SoftwareName 'Mozilla Firefox'
  $AdobeVersion = Get-InstalledSoftware -SoftwareName Adobe 

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
  #$AdobeResult


  if($Ans -eq 'N')
  {
    
    Get-ComputerLocation 

    <#
        #[string]$Building = Read-Host -Prompt 'Building'

        $Building = $DescriptionLists.Building | Out-GridView -Title 'Building' -OutputMode Single
        $ComputerStat['Building'] = $Building

        #[string]$Room = Read-Host -Prompt 'Room'

        $Room = Open-Form  # $DescriptionLists.Room | Out-GridView -Title 'Room' -OutputMode Single
        $ComputerStat['Room'] = $Room

        #[string]$Desk = Read-Host -Prompt 'Desk'

        $Desk = $DescriptionLists.Desk | Out-GridView -Title 'Desk' -OutputMode Single
        $ComputerStat['Desk'] = $Desk

        [string]$Color = Read-Host -Prompt 'Color'
        $ComputerStat['Color'] = $Color
    #>
    
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
  Get-Content -Path $FastCruiseReport | Select-Object -Last 5
}




#requires -Version 4.0

Begin
{
  $DescriptionLists = @{
    Building = 'AV34', 'AV29', 'ELC1', 'ELC2', 'ELC3', 'ELC4', 'ELC5', 'ELC6', 'ELC7', 'ELC31', 'ELC32', 'ELC33', 'ELC34', 'ELC35', 'ELC36'
    Desk     = 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'
  }

  $YearMonth = Get-Date -Format yyyy-MMMM
  $FastCruiseReport = "C:\Temp\Reports\FastCruise_$YearMonth.csv"

  $Ans = '0'
  $Orange = 'Blue'
  $CompImport = Import-Csv -Path $FastCruiseReport
  
  # Select last status of system.
  Write-Verbose -Message "Getting last status of workstation: $env:COMPUTERNAME"
  $LatestStatus = $CompImport |
  Where-Object -FilterScript {
    $PSItem.Color -match $Orange
  } |
  Select-Object -Last 1 

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
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)
  
    $cancelButton = New-Object -TypeName System.Windows.Forms.Button
    $cancelButton.Location = New-Object -TypeName System.Drawing.Point -ArgumentList (150, 120)
    $cancelButton.Size = New-Object -TypeName System.Drawing.Size -ArgumentList (75, 23)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
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
  
    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
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
            Select-Object -Property @{
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
          Target    = $e.CategoryInfo.TargetName
          Script    = $e.InvocationInfo.ScriptName
          Line      = $e.InvocationInfo.ScriptLineNumber
          Column    = $e.InvocationInfo.OffsetInLine
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

  $MozillaVersion =  Get-InstalledSoftware -SoftwareName 'Mozilla Firefox'
  $AdobeVersion = Get-InstalledSoftware -SoftwareName Adobe 

  Write-Verbose -Message ('Latest Status: {0}' -f $LatestStatus)
  #$LatestStatus | Format-List -Property ComputerName, Building, Room, @{Label='Desk A to Z from left';Expression={$_.Desk}},Notes
  $LatestStatus | Format-List -Property ComputerName, Building, Room, Desk, Notes
  
  Write-Verbose -Message 'Setting up the ComputerStat hash'
  $ComputerStat = [ordered]@{
    'ComputerName' = "$env:COMPUTERNAME"
    'UserName'   = "$env:USERNAME"
    'Date'       = "$(Get-Date)"
    'Firefox Version' = $MozillaVersion
    'Adobe Version' = $AdobeVersion
    'Building'   = ''
    'Room'       = ''
    'Desk'       = ''
    'Color'      = ''
    'Notes'      = ''
  }
  
  Write-Verbose -Message "Testing the Report Path: $FastCruiseReport"
  if(-not (Test-Path $FastCruiseReport))
  {
    Write-Verbose -Message 'Test Failed.  Creating the Report now.'
    $Null = New-Item -Path $FastCruiseReport -ItemType File
  } 
}
Process
{
  Do
  {
    $Ans = Read-Host -Prompt 'Is this correct? Y/N'
  }
  While(($Ans -ne 'N') -and ($Ans -ne 'Y')) 

   # $PowerPointResult
    #$AdobeResult


  if($Ans -eq 'N')
  {
    #[string]$Building = Read-Host -Prompt 'Building'

    $Building = $DescriptionLists.Building | Out-GridView -Title 'Building' -OutputMode Single
    $ComputerStat['Building'] = $Building

    #[string]$Room = Read-Host -Prompt 'Room'

    $Room = Get-Form  # $DescriptionLists.Room | Out-GridView -Title 'Room' -OutputMode Single
    $ComputerStat['Room'] = $Room

    #[string]$Desk = Read-Host -Prompt 'Desk'

    $Desk = $DescriptionLists.Desk | Out-GridView -Title 'Desk' -OutputMode Single
    $ComputerStat['Desk'] = $Desk

    [string]$Color = Read-Host -Prompt 'Color'
    $ComputerStat['Color'] = $Color
  }
  else
  {
    $ComputerStat['Building'] = $($LatestStatus.Building)
    $ComputerStat['Room'] = $($LatestStatus.Room)
    $ComputerStat['Desk'] = $($LatestStatus.Desk)
    $ComputerStat['Color'] = $($LatestStatus.Color)
    $ComputerStat['MS Office Test'] = $PowerPointResult
    $ComputerStat['Adobe Test'] = $AdobeResult
  }
}
END
{
  [string]$Notes = Read-Host -Prompt 'Notes'
  $ComputerStat['Notes'] = $Notes

  $ComputerStat  |
  ForEach-Object -Process {
    [pscustomobject]$_
  } |
  Export-Csv -Path $FastCruiseReport -NoTypeInformation -Append
  
  
  Write-Verbose -Message 'Show Last 10 Cruisers'
  Get-Content -Path $FastCruiseReport | Select-Object -Last 5
}


<#{<#$Credential = Get-Credential
$Credential | Export-CliXml -Path .\Jaap.Cred
$Credential = Import-CliXml -Path .\Jaap.Cred
#>

$MenuObject = 'System.Management.Automation.Host.ChoiceDescription'
$red1 = New-Object -TypeName $MenuObject -ArgumentList '&Red1', 'Favorite color: Red1'
$blue1 = New-Object -TypeName $MenuObject -ArgumentList '&Blue1', 'Favorite color: Blue1'
$yellow1 = New-Object -TypeName $MenuObject -ArgumentList '&Yellow1', 'Favorite color: Yellow1'
$red2 = New-Object -TypeName $MenuObject -ArgumentList '&Red2', 'Favorite color: Red2'
$blue2 = New-Object -TypeName $MenuObject -ArgumentList '&Blue2', 'Favorite color: Blue2'
$yellow2 = New-Object -TypeName $MenuObject -ArgumentList '&Yellow2', 'Favorite color: Yellow2'
$red3 = New-Object -TypeName $MenuObject -ArgumentList '&Red3', 'Favorite color: Red3'
$blue3 = New-Object -TypeName $MenuObject -ArgumentList '&Blue3', 'Favorite color: Blue3'
$yellow3 = New-Object -TypeName $MenuObject -ArgumentList '&Yellow3', 'Favorite color: Yellow3'
$red4 = New-Object -TypeName $MenuObject -ArgumentList '&Red4', 'Favorite color: Red4'
$blue4 = New-Object -TypeName $MenuObject -ArgumentList '&Blue4', 'Favorite color: Blue4'
$yellow4 = New-Object -TypeName $MenuObject -ArgumentList '&Yellow4', 'Favorite color: Yellow4'
 
$options = [System.Management.Automation.Host.ChoiceDescription[]]($red1, $blue1, $yellow1,$red2, $blue2, $yellow2,$red3, $blue3, $yellow3,$red4, $blue4, $yellow4)


$title = 'Favorite color'
$message = 'What is your favorite color?'
$result = $host.ui.PromptForChoice($title, $message, $options, 0)

Write-Output $result}#>

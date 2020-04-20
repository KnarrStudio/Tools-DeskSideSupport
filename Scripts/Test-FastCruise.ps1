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

  Write-Verbose -Message ('Latest Status: {0}' -f $LatestStatus)
  #$LatestStatus | Format-List -Property ComputerName, Building, Room, @{Label='Desk A to Z from left';Expression={$_.Desk}},Notes
  $LatestStatus | Format-List -Property ComputerName, Building, Room, Desk, Notes
  
  Write-Verbose -Message 'Setting up the ComputerStat hash'
  $ComputerStat = [ordered]@{
    'ComputerName' = "$env:COMPUTERNAME"
    'UserName'   = "$env:USERNAME"
    'Date'       = "$(Get-Date)"
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
    $Ans = Read-Host -Prompt 'Is this correct? N/Y'
  }
  While(($Ans -ne 'N') -and ($Ans -ne 'Y')) 

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




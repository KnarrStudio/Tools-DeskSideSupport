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





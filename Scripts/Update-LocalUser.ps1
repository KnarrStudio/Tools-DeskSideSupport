$AdminUser = @{
  Name                = 'sysmaint'
  FullName            = '911'
  Description         = 'Emergency Access PW in "SCIT"'
  AccountNeverExpires = $true
  Password            = '1qaz@WSX3edc$RFV'
}
  
function Update-LocalUser 
{
  <#
      .SYNOPSIS
      Updates a user based on the hash settings


  #>
  param
  (
    [Parameter(Mandatory, Position = 0)]
    [Object]$UserToUpdate,
    [Parameter(Mandatory, Position = 1)]
    [Object]$UpdateInfo
  )
  $UserSplat = @{
    Name                = $UpdateInfo.Name
    FullName            = $UpdateInfo.FullName
    Description         = $UpdateInfo.Description
    AccountNeverExpires = $UpdateInfo.AccountNeverExpires
  }
  
  $LocalUserNames = (Get-LocalUser).name
  $LocalUserNames

  $UserName = [String]$UpdateInfo.Name
  Write-Verbose -Message $UserName

  $Password = ConvertTo-SecureString -String ($UpdateInfo.Password) -AsPlainText -Force
  Write-Verbose -Message $Password

  $UserDescription = ($UpdateInfo.Description)
  Write-Verbose -Message $UserDescription

  $UserFullName = ($UpdateInfo.FullName)
  Write-Verbose -Message $UserFullName

  Write-Verbose -Message $UpdateInfo

  if ($UserToUpdate -in $LocalUserNames)
  {
    Write-Verbose -Message ('Renaming {0} Account' -f $UserFullName)
    Rename-LocalUser -Name $UserToUpdate -NewName $UserName

    Write-Verbose -Message ('Setting {0} Account' -f $UserFullName)
    Set-LocalUser @UserSplat -Password $Password -ErrorAction SilentlyContinue
  }
  if ((Get-LocalUser -Name $UserName).Enabled -eq $false) 
  {
    Enable-LocalUser -Name $UserName
  }
}

Update-LocalUser -UserToUpdate sysmaint -UpdateInfo $AdminUser -Verbose
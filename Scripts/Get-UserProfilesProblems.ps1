$UserRegistry = @{
  Path = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*'
}
Get-ItemProperty @UserRegistry | Where-Object {
  $_.ProfileImagePath -match '.bak'
}
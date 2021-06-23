$ProfileString = '.bak'
$UserRegistry = @{
  Path = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*'
}
Get-ItemProperty @UserRegistry | Where-Object {
  $_.ProfileImagePath -match $ProfileString
}

<#
If there is an output, then open Regedit as an Admin:
The Registry key with the .bak extension contains the user's actual profile while the one without the .bak contains the Temp profile. 
Delete the Registry Key WITHOUT the .bak extension and rename the one with it to xxxxx1234 (without the .bak). 
Notice the fields on the right, there should be a value named RefCount, change the value to 0.
#>


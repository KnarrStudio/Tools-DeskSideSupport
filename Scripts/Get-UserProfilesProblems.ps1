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


# SIG # Begin signature block
# MIIFvQYJKoZIhvcNAQcCoIIFrjCCBaoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUrPJbEI/GRC7mY1JKzk69pW+p
# FcOgggNLMIIDRzCCAi+gAwIBAgIQdr8M9RYBq65KZ4TmHgXXMDANBgkqhkiG9w0B
# AQsFADAnMSUwIwYDVQQDDBxBUk5FU0VOLkVSSUsuSk9ITi4xMDIwODk2MTA0MB4X
# DTIyMDIxOTEzMzgzN1oXDTIzMDIxOTEzNTgzN1owJzElMCMGA1UEAwwcQVJORVNF
# Ti5FUklLLkpPSE4uMTAyMDg5NjEwNDCCASIwDQYJKoZIhvcNAQEBBQADggEPADCC
# AQoCggEBALlVtaiHeQFsBgEhZ5ZG9plThVyZDU9vBLUJDhhkkralrdADvFnEIPDK
# 1qehJAKdIPB3uBBWw5rXRdaNWJsiHso6PL9CZIOXskT1ft6pSWOgHoGV/RzBhnk8
# HkwKneYDxx9tp3MBLT6AGiCZTm2Qu+E1J8jFFFusW1JDCx0TNQ1bP/DtS/TfrnyS
# EqxuVHvFjaI3XfZG0JGec36Jult8wOHr1Iv/YQRKXTyU63pi0gOddfD6zyrkuqL2
# kCP/+mgLtRx56zrfloc+tI9MMIRW0KH7e6I+yWFi+SPfaXvT3hmHksdumwVoADzl
# 2jFieaCY8Ah8ZPh5nLuevUmXYhi+JR0CAwEAAaNvMG0wDgYDVR0PAQH/BAQDAgeA
# MBMGA1UdJQQMMAoGCCsGAQUFBwMDMCcGA1UdEQQgMB6CHEFSTkVTRU4uRVJJSy5K
# T0hOLjEwMjA4OTYxMDQwHQYDVR0OBBYEFAzIoPmuXXn/OOLkRKM9x7lv8udKMA0G
# CSqGSIb3DQEBCwUAA4IBAQArNEJElgTjiFKHLfmHwhcJkj8TAHiYURP+y0hFMehT
# ey7Z5Gst2UY3/PftdQ0TyRQpmMpdEjspC+kJDLtlzq6ynqFhhnZRcD85xh89eVJT
# C8ArY925wr3rJiuA2ktIXbF5oBSZw/MKSa/GOSlcsj0C7iWIUJ8zjI9r422LButq
# +OA5Kzc2yPGCD19bu0TGqffLm6eB3j+phPm3Aya/DkpA2Xv0Kex+dxmr75Cg67Go
# SArwX6PBluYglQgwwNziD9GdQJ0zNAW0ynd5y+/FIKU+BZCeEv6EbIdCzub4nB7N
# cdhCgeQFqmQ8pRXIO0ke0biEctoRV5cRKvZIRCrEHiZpMYIB3DCCAdgCAQEwOzAn
# MSUwIwYDVQQDDBxBUk5FU0VOLkVSSUsuSk9ITi4xMDIwODk2MTA0AhB2vwz1FgGr
# rkpnhOYeBdcwMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAA
# MBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgor
# BgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBT0dEN84RqrTWaBYHyai6q/fQp1hzAN
# BgkqhkiG9w0BAQEFAASCAQBqRn5t8waHNZM7P2f+1Scd6sJ4DP32O/krweUttIdZ
# Wz2BDRw8VRE3SjdwIJx3HJSUSWQWr+yNqEvqeJ2B9UMo5fwMe5LOXmAMvCPEp0Ql
# Alh9ZKs9TIo71Dww4g2TrpS0O9kzd+EaHwkIJNivQ1TuxV9+vDxJh0ZalO4sTkPf
# SuNLUS9dzMLAOl98js0M7PwFspqcGVERYm77Z5tGrMKMLNcYz3vm4oshlo4mNYj5
# blot4JgKTVw/3uq416hb652heHRmm8jjAUSYMaicHcSGD5lRCym9yxagVeMX7sE9
# 7RPoOvB3xGE5+92F/qBjDh4BKgj23awRh9a18qgkbKL7
# SIG # End signature block

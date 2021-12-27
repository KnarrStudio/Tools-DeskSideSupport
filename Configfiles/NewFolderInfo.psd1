  $NewFolderInfo = [ordered]@{
    CyberUpdates = @{
      Path       = 'C:\temp\ConfigFilesHashTablePsd\CyberUpdates'
      ACLGroup   = 'Administrators'
      ACLControl = 'Full Control'
      ReadMeText = 'This is the working folder for the monthly updates and scanning.'
      ReadMeFile = 'README.TXT'
    }
    ScanReports  = @{
      Path       = 'C:\temp\ConfigFilesHashTablePsd\CyberUpdates\ScanReports'
      ACLGroup   = 'Administrators'
      ACLControl = 'Full Control'
      ReadMeText = 'This is where the "IA" scans engines and reports will be kept.'
      ReadMeFile = 'README.TXT'
    }
  }
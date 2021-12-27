$NewUsers = @{
    RFVuser    = @{
      FullName            = 'RFV User'
      Description         = 'Standard local Account'
      AccountGroup        = 'Users'
      AccountNeverExpires = $true
      Password            = '1qaz@WSX3edc$RFV'
    }
    RFVAdmin   = @{
      FullName            = 'RFV Admin'
      Description         = 'Local Admin Account for IT'
      AccountGroup        = 'Administrators'
      AccountNeverExpires = $true
      Password            = 'RFV@dm!nP@$$!!'
    }
    CRTech     = @{
      FullName            = 'TV Tech User'
      Description         = 'Standard TV Account'
      AccountGroup        = 'Users'
      AccountNeverExpires = $true
      Password            = '1qaz@WSX3edc$RFV'
    }
    CRAdmin    = @{
      FullName            = 'TV Administrator'
      Description         = 'TV Admin Account'
      AccountGroup        = 'Administrators'
      AccountNeverExpires = $true
      Password            = '1qaz@WSX3edc$RFV'
    }
    NineOneOne = @{
      FullName            = '911'
      Description         = 'Emergancy Access PW in "KeePass"'
      AccountGroup        = 'Administrators'
      AccountNeverExpires = $true
      Password            = '1qaz@WSX3edc$RFV'
    }
  }
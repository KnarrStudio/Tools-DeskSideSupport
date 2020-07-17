#requires -Version 3.0 -Modules NetAdapter

#Edit the splats to customize the script
$FastCruiseSplat = @{
	FastCruiseReportPath = 'C:\temp\Report'
	FastCruiseFile = 'FastCruise.csv' 
	Verbose = $true
}
$PDFApplicationTestSplat = @{
	TestFile = '\\FastCruise\FastCruiseTestFile.pdf'
	TestProgram = "${env:ProgramFiles(x86)}\Adobe\Acrobat 2015\Acrobat\Acrobat.exe"
	ProcessName = 'Acrobat'
}
$PowerPointApplicationTestSplat = @{
	TestFile = '\\FastCruise\FastCruiseTestFile.pptx'
	TestProgram = "${env:ProgramFiles(x86)}\Microsoft Office\Office16\POWERPNT.EXE"
	ProcessName = 'POWERPNT'
}
<#$WordpadApplicationTestSplat = @{
    TestFile    = "$env:windir\DtcInstall.log"
    TestProgram = "$env:ProgramFiles\Windows NT\Accessories\wordpad.exe"
    ProcessName = 'wordpad'
    }
#>

# Edit the Variables
$SoftwareChecks = @(@('Adobe', 'Version'), @( 'Mozilla Firefox', 'Version'), @('McAfee Agent', 'Version')) #,@('VMware','Version'))

$jsonFilePath = 'C:\Users\Erik.Arnesen\Documents\GitHub\KnarrStudio\Tools-DeskSideSupport\Scripts\computerlocation.json' 


function Start-FastCruise
{
 param
(
	[Parameter(Mandatory, Position = 0)]
	[String]$FastCruiseReportPath,
	[Parameter(Mandatory, Position = 0)]
	[ValidateScript({
			If($_ -match '.csv')
			{
				$true
			}
			Else
			{
				Throw 'Input file needs to be CSV'
			}
		})][String]$FastCruiseFile
)

Begin
{
	Write-Verbose -Message 'Setup Variables'
	#$LocationVerification = $null
	#$ComputerName = $env:COMPUTERNAME

	Write-Verbose -Message 'Setup Report' 
	$YearMonth = Get-Date -Format yyyy-MMMM
	$FastCruiseFile = [String]$($FastCruiseFile.Replace('.',('_{0}.' -f $YearMonth)))
	$FastCruiseReport = ('{0}\{1}' -f $FastCruiseReportPath, $FastCruiseFile)
	#$FastCruiseReport = "C:\temp\Reports\FastCruise_Test.csv"
	Write-Verbose -Message ('{0}' -f $FastCruiseReport) 

	Write-Verbose -Message ('Testing the Report Path: {0}' -f $FastCruiseReportPath)
	if(-not (Test-Path -Path $FastCruiseReportPath))
	{
		Write-Verbose -Message 'Test Failed.  Creating the Directory now.'
		$null = New-Item -Path $FastCruiseReportPath -ItemType Directory -Force
	} 
	Write-Verbose -Message ('Testing the Report Path: {0}' -f $FastCruiseReport)
	if(-not (Test-Path -Path $FastCruiseReport))
	{
		Write-Verbose -Message 'Test Failed.  Creating the File now.'
		$null = New-Item -Path $FastCruiseReport -ItemType File -Force
	} 

	# Variables
	$Phone = $null

	Write-Verbose -Message 'Get-Content of Json File'
	try
	{
		$Script:PhysicalLocations = Get-Content -Path $jsonFile -ErrorAction Stop | ConvertFrom-Json 
		Write-Verbose -Message 'Physical Locations'
		$PhysicalLocations
	}
	catch
	{
		$PhysicalLocations = $null
	}

	function Start-ApplicationTest
 {
		param
		(
			[Parameter(Mandatory, Position = 0)]
			[string]$FunctionTest,
			[Parameter(Mandatory, Position = 1)]
			[string]$TestFile,
			[Parameter(Mandatory, Position = 2)]
			[string]$TestProgram,
			[Parameter(Mandatory, Position = 3)]
			[string]$ProcessName
		)
		$DescriptionLists = [Ordered]@{
			FunctionResult = 'Good', 'Failed'
		}
		Write-Verbose -Message ('Enter Function: {0}' -f $PSCmdlet.MyInvocation.MyCommand.Name)
		if($FunctionTest -eq 'Yes')
		{
			try
			{
				Write-Verbose -Message ('Attempting to open {0} with {1}' -f $TestFile, $ProcessName)
				#Start-Process -FilePath $TestProgram -ArgumentList $TestFile
				Start-Process -FilePath $TestFile

				Write-Host -Object ('The Fast Cruise Script will continue after {0} has been closed.' -f $ProcessName) -BackgroundColor Red -ForegroundColor Yellow
				Write-Verbose -Message ('Wait-Process: {0}' -f $ProcessName)
				Wait-Process -Name $ProcessName

				$TestResult = $DescriptionLists.FunctionResult | Out-GridView -Title $ProcessName -OutputMode Single
			}
			Catch
			{
				Write-Verbose -Message 'TestResult: Failed'
				$TestResult = $DescriptionLists.FunctionResult[1]
			}
		}
		else
		{
			Write-Verbose -Message 'TestResult: Bypassed'
			$TestResult = 'Bypassed'
		}
		Return $TestResult
	} # End ApplicationTest-Function

	function Get-LastComputerStatus
 {
		<#
          .SYNOPSIS
          Return the last status of system based on what was in the current Fast Cruise Report
      #>
		param
		(
			[Parameter(Mandatory, Position = 0)]
			[String]$FastCruiseReport
		)

		Write-Verbose -Message ('Enter Function: {0}' -f $PSCmdlet.MyInvocation.MyCommand.Name)
		Write-Verbose -Message 'Importing the Fast Cruise Report'
		$CompImport = Import-Csv -Path $FastCruiseReport

		# Select last status of system.
		Write-Verbose -Message "Getting last status of workstation: $env:COMPUTERNAME"

		try
		{
			$LatestStatus = $CompImport |
			Where-Object -FilterScript {
				$PSItem.ComputerName -eq $env:COMPUTERNAME
			} |
			Select-Object -Last 1 
			if($LatestStatus -eq $null)
			{
				Write-Output -InputObject 'Unable to find an existing record for this system.'
				$Script:Ans = 'NoHistory'
			}
		}
		Catch
		{
			# get error record
			[Management.Automation.ErrorRecord]$e = $_

			# retrieve information about runtime error
			$info = New-Object -TypeName PSObject -Property @{
				Exception = $e.Exception.Message
			}

			# output information. Post-process collected info, and log info (optional)
			$info
		}
		Return $LatestStatus
	} # End ComputerStatus-Function

	function Get-ComputerLocation 
	{
		<#
          .SYNOPSIS
          Get-ComputerLocation of workstation
      #>

		param
		(
			[Parameter(Mandatory = $false, Position = 0)]
			[Object]$jsonFilePath
		)

		function Convert-JSONToHash
 {
			param(
				$root
			)
			$hash = @{}

			$keys = $root |
			Get-Member -MemberType NoteProperty |
			Select-Object -ExpandProperty Name

			$keys | ForEach-Object -Process {
				$key = $_
				$obj = $root.$($_)
				if($obj -match '@{')
				{
					$nesthash = Convert-JSONToHash -root $obj
					$hash.add($key,$nesthash)
				}
				else
				{
					$hash.add($key,$obj)
				}
			}
			return $hash
		}

		[Object[]]$Desk = @('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q')

		if(Test-Path $jsonFilePath -ErrorAction SilentlyContinue)
		{
			$location = Convert-JSONToHash -root $(Get-Content -Path $jsonFilePath -ErrorAction SilentlyContinue | ConvertFrom-Json)
			[string]$Script:LclDept = $location.Department.keys | Out-GridView -Title 'Department' -OutputMode Single
			[string]$Script:LclBuild = $location.Department[$LclDept].Building.Keys | Out-GridView -Title 'Building' -OutputMode Single
			[string]$Script:LclRm = $location.Department[$LclDept].Building[$LclBuild].Room | Out-GridView -Title 'Room' -OutputMode Single
			[string]$Script:LclDesk = $Desk | Out-GridView -Title 'Desk' -OutputMode Single
		}
		else
		{
			[string]$Script:LclDept = Show-VbForm -InputBox -Message 'Department: MCDO, PRO, CA, Other' -TitleBar 'Department'
			[string]$Script:LclBuild = Show-VbForm -InputBox -Message 'Building: ELC44, AV34' -TitleBar 'Building'
			[string]$Script:LclRm = Show-VbForm -InputBox -Message 'Room Number:' -TitleBar 'Room'
			[string]$Script:LclDesk = $Desk | Out-GridView -Title 'Desk' -OutputMode Single
		}

		if($location -eq 'rainbow')
		{
			$location = [Ordered]@{
				Department = [Ordered]@{
					InternalHash = @{
						Building = @{
							None = @{
								Room = @(
									0
								)
							}
						}
					}
					MCDO = [Ordered]@{
						Building = [Ordered]@{
							AV29 = [Ordered]@{
								Room = @(
									8, 
									9, 
									10, 
									11, 
									12, 
									13, 
									14, 
									15, 
									16, 
									17, 
									18, 
									19, 
									20
								)
							}
							AV34 = [Ordered]@{
								Room = @(
									1, 
									6, 
									7, 
									8, 
									201, 
									202, 
									203, 
									204, 
									205
								)
							}
							ELC3 = [Ordered]@{
								Room = @(
									100, 
									101, 
									102, 
									103, 
									104, 
									105, 
									106, 
									107
								)
							}
							ELC31 = [Ordered]@{
								Room = @(
									1
								)
							}
							ELC32 = [Ordered]@{
								Room = @(
									1
								)
							}
							ELC33 = [Ordered]@{
								Room = @(
									1
								)
							}
							ELC34 = [Ordered]@{
								Room = @(
									1
								)
							}
							ELC35 = [Ordered]@{
								Room = @(
									1
								)
							}
							ELC36 = [Ordered]@{
								Room = @(
									1
								)
							}
						}
					}
					CA = [Ordered]@{
						Building = [Ordered]@{
							AV29 = [Ordered]@{
								Room = @(
									1, 
									2, 
									3, 
									4, 
									5, 
									6, 
									7, 
									23, 
									24, 
									25, 
									26, 
									27, 
									28, 
									29, 
									30
								)
							}
							AV34 = [Ordered]@{
								Room = @(
									1, 
									2, 
									3, 
									200, 
									214
								)
							}
							AV44 = [Ordered]@{
								Room = @(
									1
								)
							}
							AV45 = [Ordered]@{
								Room = @(
									1
								)
							}
							AV46 = [Ordered]@{
								Room = @(
									1
								)
							}
							AV47 = [Ordered]@{
								Room = @(
									1, 
									2
								)
							}
							AV48 = [Ordered]@{
								Room = @(
									1, 
									2
								)
							}
						}
					}
					PRO = [Ordered]@{
						Building = [Ordered]@{
							AV34 = [Ordered]@{
								Room = @(
									210, 
									211, 
									212, 
									213
								)
							}
							ELC4 = [Ordered]@{
								Room = @(
									1, 
									100, 
									101, 
									102, 
									103, 
									104, 
									105, 
									106, 
									107
								)
							}
						}
					}
					TJ = [Ordered]@{
						Building = [Ordered]@{
							AV34 = [Ordered]@{
								Room = @(
									2, 
									3, 
									13, 
									11
								)
							}
							ELC2 = [Ordered]@{
								Room = @(
									1
								)
							}
						}
					}
				}
			}
		}
		#>      
	} # End Location-Function

	function Show-VbForm
 {
		[cmdletbinding(DefaultParameterSetName = 'Message')]
		param(
			[Parameter(Position = 0,ParameterSetName = 'Message')]
			[Switch]$YesNoBox,
			[Parameter(Position = 0,ParameterSetName = 'Input')]
			[Switch]$InputBox,
			[Parameter(Mandatory,Position = 1)]
			[string]$Message,
			[Parameter(Position = 2)]
			[string]$TitleBar = 'Fast Cruise'
		)
        
		Write-Verbose -Message ('Enter Function: {0}' -f $PSCmdlet.MyInvocation.MyCommand.Name)

		Add-Type -AssemblyName Microsoft.VisualBasic

		if($InputBox)
		{
			$Response = [Microsoft.VisualBasic.Interaction]::InputBox($Message, $TitleBar)
		}
		if($YesNoBox)
		{
			$Response = [Microsoft.VisualBasic.Interaction]::MsgBox($Message, 'YesNo,SystemModal,MsgBoxSetForeground', $TitleBar)
		}
		Return $Response
	} # End VbForm-Function

	Function Get-InstalledSoftware
 {
		[cmdletbinding(SupportsPaging)]
		Param(

			[Parameter(HelpMessage = 'At least part of the software name to test',ValueFromPipeline, Position = 0)]
			[String[]]$SoftwareName,
			[ValidateSet('DisplayName','DisplayVersion')] 
			[String]$SelectParameter
		)

		Begin { 
			Write-Verbose -Message ('Enter Function: {0}' -f $PSCmdlet.MyInvocation.MyCommand.Name)

			$SoftwareOutput = @()
			$InstalledSoftware = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*)
		}

		Process {
			Try 
			{
				if($SoftwareName -eq $null) 
				{
					$SoftwareOutput = $InstalledSoftware |
					Select-Object -Property Installdate, DisplayVersion, DisplayName #, UninstallString 
				}
				Else 
				{
					foreach($Item in $SoftwareName)
					{
						$SoftwareOutput += $InstalledSoftware |
						Where-Object -Property DisplayName -Match -Value $Item |
						Select-Object -Property Installdate, DisplayVersion, DisplayName #, UninstallString 
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
				}

				# output information. Post-process collected info, and log info (optional)
				$info
			}
		}

		End{ 
			Switch ($SelectParameter){
				'DisplayName' 
				{
					$SoftwareOutput.displayname
				}
				'DisplayVersion' 
				{
					$SoftwareOutput.DisplayVersion
				}
				default 
				{
					$SoftwareOutput
				}
			}
		}
	} # End InstalledSoftware-Function

	function Get-MacAddress 
	{
		param(
			[Parameter(Position = 0)]
			[Switch]$LastFour
		)
		$MacAddress = (Get-NetAdapter -Physical | Where-Object -Property status -EQ -Value 'Up').macaddress
		if($LastFour)
		{
			$MacInfo = (($MacAddress.Split('-',5))[4]).replace('-',':')
		}
		else
		{
			$MacInfo = $MacAddress
		}
		$MacInfo
	}

	<#bookmark Windows Updates #> 
	$LatestWSUSupdate = (New-Object -ComObject 'Microsoft.Update.AutoUpdate'). Results 

	Write-Verbose -Message 'Setting up the ComputerStat hash'
	<#bookmark ComputerStat Hashtable #>
	$ComputerStat = [ordered]@{
		'ComputerName' = "$env:COMPUTERNAME"
        'SerialNumber' = 'N/A'
		'MacAddress' = 'N/A'
		'UserName' = "$env:USERNAME"
		'Date' = "$(Get-Date)"
		'WSUS Search Success' = 'N/A'
		'WSUS Install Success' = 'N/A'
		'Department' = 'N/A'
		'Building' = 'N/A'
		'Room' = 'N/A'
		'Desk' = 'N/A'
	}

} #End BEGIN region

Process
{
	<#bookmark Get-MacAddress #>
	Write-Verbose -Message 'Getting Mac Address'
	$ComputerStat['MacAddress'] = Get-MacAddress
    
    Write-Verbose -Message 'Getting Serial Number'
    $ComputerStat['SerialNumber'] = (Get-WmiObject win32_SystemEnclosure).serialnumber
    

	<#bookmark Software Versions #>
	#$ComputerStat['VmWare Version']  = Get-InstalledSoftware -SoftwareName 'Vmware' -SelectParameter DisplayVersion

	#$SoftwareChecks = @(@('Adobe', 'Version'), @( 'Mozilla Firefox', 'Version'), @('McAfee Agent', 'Version')) #,@('VMware','Version'))
	foreach($SoftwareItem in $SoftwareChecks)
	{
		$ComputerStat["$SoftwareItem"] = Get-InstalledSoftware -SoftwareName $SoftwareItem[0] -SelectParameter DisplayVersion
	}

	Write-Verbose -Message 'Getting Last Status recorded'
	$LatestStatus = (Get-LastComputerStatus -FastCruiseReport $FastCruiseReport) 
	#Write-Output -InputObject 'Latest Status'
	#$LatestStatus | Select-Object -Property Computername, Department, Building, Room, Desk

	<#bookmark Location Verification #>
	$ComputerLocation = (@'

ComputerName: (Assest Tag)
- {0}

Serial Number:
- {6}

Department:
- {1}

Building:
- {2}

Room:
- {3}

Desk:
- {4}

Phone
- {5}
          
'@ -f $LatestStatus.ComputerName, $LatestStatus.Department, $LatestStatus.Building, $LatestStatus.Room, $LatestStatus.Desk, $LatestStatus.Phone,$LatestStatus.SerialNumber)

	<#bookmark Application Test #> 
	$FunctionTest = Show-VbForm -YesNoBox -Message 'Perform Applicaion Tests (MS Office and Adobe)?' 

	$AdobeResult = Start-ApplicationTest -FunctionTest $FunctionTest @PDFApplicationTestSplat
	$PowerPointResult = Start-ApplicationTest -FunctionTest $FunctionTest @PowerPointApplicationTestSplat

	$ComputerStat['MS Office Test'] = $PowerPointResult
	$ComputerStat['Adobe Test'] = $AdobeResult

	$LocationVerification = Show-VbForm -YesNoBox -Message $ComputerLocation

	if($LocationVerification -eq 'No')
	{
		Get-ComputerLocation -jsonFilePath $jsonFilePath
		Write-Verbose -Message ('Computer Description: ABC-DEF-{0}-{1}-{2}{3}' -f $LclDept, $LclBuild, $LclRm, $LclDesk)

		$ComputerStat['Department'] = $LclDept 
		$ComputerStat['Building'] = $LclBuild
		$ComputerStat['Room'] = $LclRm
		$ComputerStat['Desk'] = $LclDesk
	}
	else
	{
		$ComputerStat['Building'] = $($LatestStatus.Building)
		$ComputerStat['Room'] = $($LatestStatus.Room)
		$ComputerStat['Desk'] = $($LatestStatus.Desk)
		$ComputerStat['Phone'] = $($LatestStatus.Phone)
	}

	if($LocationVerification -eq 'No')
	{
		<#bookmark Local phone number #> 
		$RegexPhone = '^\d{3}-\d{3}-\d{4}'
		While($Phone -notmatch $RegexPhone)
		{
			$Phone = Show-VbForm -InputBox -Message 'Nearest Phone Number (757-555-1234):'
		}
		$ComputerStat['Phone'] = $Phone
	}

	<#bookmark Windows Update Status #> 
	$ComputerStat['WSUS Search Success'] = $LatestWSUSupdate.LastSearchSuccessDate
	$ComputerStat['WSUS Install Success'] = $LatestWSUSupdate.LastInstallationSuccessDate

	<#bookmark Fast cruise notes #>
	[string]$Notes = Show-VbForm -InputBox -Message 'Notes about this cruise:'
	$ComputerStat['Notes'] = $Notes
} #End PROCESS region

END
{
	$ComputerStat |
	ForEach-Object -Process {
		[pscustomobject]$_
	} |
	Export-Csv -Path $FastCruiseReport -NoTypeInformation -Append


	Write-Output -InputObject 'The information recorded'
	$ComputerStat | Format-Table

	<#bookmark Fast cruising shipmates #>
	Write-Output -InputObject 'Fast Cruise shipmates'
	Import-Csv -Path $FastCruiseReport |
	Select-Object -Last 4 -Property Date, Username, Building, Room, Phone |
	Format-Table
} #End END region
}

Clear-Host #Clears the console.  This shouldn't be needed once the script can be run directly from PS
Start-FastCruise @FastCruiseSplat # Make sure you have updated and completed the "Splats" at the top of the script


function Send-eMail {
  [CmdletBinding(DefaultParameterSetName = 'Default')]
  param
  (
    [Parameter(Mandatory,HelpMessage = 'To email address(es)', Position = 0)]
    [String[]]$MailTo,
    [Parameter(Mandatory,HelpMessage = 'From email address', Position = 1)]
    [Object]$MailFrom,
    [Parameter(Mandatory,HelpMessage = 'Email subject', Position = 2)]
    [Object]$msgsubj,
    [Parameter(Mandatory,HelpMessage = 'SMTP Server(s)', Position = 3)]
    [String[]]$SmtpServers,
    [Parameter(Position = 4)]
    [AllowNull()]
    $MessageBody,
    [Parameter(Position = 5)]
    [AllowNull()]
    [Object]$AttachedFile,
    [Parameter(Position = 6)]
    [AllowEmptyString()]
    [string]$ErrorFile = ''
  )

  $DateTime = Get-Date -Format s

  if([string]::IsNullOrEmpty($MessageBody))
  {
    $MessageBody = ('{1} - Email generated from {0}' -f $env:computername, $DateTime)
    Write-Warning -Message 'Setting Message Body to default message'
  }
  elseif(($MessageBody -match '.txt') -or ($MessageBody -match '.htm'))
  {
    if(Test-Path $MessageBody)
    {
      [String]$MessageBody = Get-Content -Path $MessageBody
    }
  }
  elseif(-not ($MessageBody -is [String]))
  {
    $MessageBody = ('{0} - Original message was not sent as a String.' -f $DateTime)
  }
  else
  {
    $MessageBody = ("{0}`n{1}" -f $MessageBody, $DateTime)
  }
    
  if([string]::IsNullOrEmpty($ErrorFile))
  {
    $ErrorFile = New-TemporaryFile
    Write-Warning  -Message ('Setting Error File to: {0}' -f $ErrorFile)
  }
  $SplatSendMessage = @{
    From        = $MailFrom
    To          = $MailTo
    Subject     = $msgsubj
    Body        = $MessageBody
    Priority    = 'High'
    ErrorAction = 'Stop'
  }
  
  if($AttachedFile)
  {
    Write-Verbose -Message 'Inserting file attachment'
    $SplatSendMessage.Attachments = $AttachedFile
  }
  if($MessageBody.Contains('html'))
  {
    Write-Verbose -Message 'Setting Message Body to HTML'
    $SplatSendMessage.BodyAsHtml  = $true
  }
  
  foreach($SMTPServer in $SmtpServers)
  {
    try
    {
      Write-Verbose -Message ('Try to send mail thru {0}' -f $SMTPServer)
      Send-MailMessage -SmtpServer $SMTPServer  @SplatSendMessage
      # Write-Output $SMTPServer  @SplatSendMessage
      Write-Verbose -Message ('successful from {0}' -f $SMTPServer)
      Write-Host -Object ("`nsuccessful from {0}" -f $SMTPServer) -ForegroundColor green
      Break 
    } 
    catch 
    {
      $ErrorMessage  = $_.exception.message
      Write-Verbose -Message ("Error Message: `n{0}" -f $ErrorMessage)
      ('Unable to send message thru {0} server' -f $SMTPServer) | Out-File -FilePath $ErrorFile -Append
      ('- {0}' -f $ErrorMessage) | Out-File -FilePath $ErrorFile -Append
      Write-Verbose -Message ('Errors written to: {0}' -f $ErrorFile)
    }
  }
}

function Show-AsciiMenu {
  <#
      .SYNOPSIS
      Describe purpose of "Show-AsciiMenu" in 1-2 sentences.

      .DESCRIPTION
      Add a more complete description of what the function does.

      .PARAMETER Title
      Describe parameter -Title.

      .PARAMETER MenuItems
      Describe parameter -MenuItems.

      .PARAMETER TitleColor
      Describe parameter -TitleColor.

      .PARAMETER LineColor
      Describe parameter -LineColor.

      .PARAMETER MenuItemColor
      Describe parameter -MenuItemColor.

      .EXAMPLE
      Show-AsciiMenu -Title Value -MenuItems Value -TitleColor Value -LineColor Value -MenuItemColor Value
      Describe what this call does

      .NOTES
      Place additional notes here.

      .LINK
      URLs to related sites
      The first link is opened by Get-Help -Online Show-AsciiMenu

      .INPUTS
      List of input types that are accepted by this function.

      .OUTPUTS
      List of output types produced by this function.
  #>


  [CmdletBinding()]
  param
  (
    [string]$Title = 'Title',

    [String[]]$MenuItems = 'None',

    [string]$TitleColor = 'Red',

    [string]$LineColor = 'Yellow',

    [string]$MenuItemColor = 'Cyan'
  )
  Begin{
    # Set Variables
    $i = 1
    $Tab = "`t"
    $VertLine = '║'
  
    function Write-HorizontalLine
    {
      param
      (
        [Parameter(Position = 0)]
        [string]
        $DrawLine = 'Top'
      )
      Switch ($DrawLine) {
        Top 
        {
          Write-Host ('╔{0}╗' -f $HorizontalLine) -ForegroundColor $LineColor
        }
        Middle 
        {
          Write-Host ('╠{0}╣' -f $HorizontalLine) -ForegroundColor $LineColor
        }
        Bottom 
        {
          Write-Host ('╚{0}╝' -f $HorizontalLine) -ForegroundColor $LineColor
        }
      }
    }
    function Get-Padding
    {
      param
      (
        [Parameter(Mandatory, Position = 0)]
        [int]$Multiplier 
      )
      "`0"*$Multiplier
    }
    function Write-MenuTitle
    {
      Write-Host ('{0}{1}' -f $VertLine, $TextPadding) -NoNewline -ForegroundColor $LineColor
      Write-Host ($Title) -NoNewline -ForegroundColor $TitleColor
      if($TotalTitlePadding % 2 -eq 1)
      {
        $TextPadding = Get-Padding -Multiplier ($TitlePaddingCount + 1)
      }
      Write-Host ('{0}{1}' -f $TextPadding, $VertLine) -ForegroundColor $LineColor
    }
    function Write-MenuItems
    {
      foreach($menuItem in $MenuItems)
      {
        $number = $i++
        $ItemPaddingCount = $TotalLineWidth - $menuItem.Length - 6 #This number is needed to offset the Tab, space and 'dot'
        $ItemPadding = Get-Padding -Multiplier $ItemPaddingCount
        Write-Host $VertLine  -NoNewline -ForegroundColor $LineColor
        Write-Host ('{0}{1}. {2}{3}' -f $Tab, $number, $menuItem, $ItemPadding) -NoNewline -ForegroundColor $LineColor
        Write-Host $VertLine -ForegroundColor $LineColor
      }
    }
  }

  Process
  {
    $TitleCount = $Title.Length
    $LongestMenuItemCount = ($MenuItems | Measure-Object -Maximum -Property Length).Maximum
    Write-Debug -Message ('LongestMenuItemCount = {0}' -f $LongestMenuItemCount)

    if  ($TitleCount -gt $LongestMenuItemCount)
    {
      $ItemWidthCount = $TitleCount
    }
    else
    {
      $ItemWidthCount = $LongestMenuItemCount
    }

    if($ItemWidthCount % 2 -eq 1)
    {
      $ItemWidth = $ItemWidthCount + 1
    }
    else
    {
      $ItemWidth = $ItemWidthCount
    }
    Write-Debug -Message ('Item Width = {0}' -f $ItemWidth)
   
    $TotalLineWidth = $ItemWidth + 10
    Write-Debug -Message ('Total Line Width = {0}' -f $TotalLineWidth)
  
    $TotalTitlePadding = $TotalLineWidth - $TitleCount
    Write-Debug -Message ('Total Title Padding  = {0}' -f $TotalTitlePadding)
  
    $TitlePaddingCount = [math]::Floor($TotalTitlePadding / 2)
    Write-Debug -Message ('Title Padding Count = {0}' -f $TitlePaddingCount)

    $HorizontalLine = '═'*$TotalLineWidth
    $TextPadding = Get-Padding -Multiplier $TitlePaddingCount
    Write-Debug -Message ('Text Padding Count = {0}' -f $TextPadding.Length)


    Write-HorizontalLine -DrawLine Top
    Write-MenuTitle
    Write-HorizontalLine -DrawLine Middle
    Write-MenuItems
    Write-HorizontalLine -DrawLine Bottom
  }

  End
  {}
}

do{
$Raspberry = $null

#Show-AsciiMenu -Title 'THIS IS THE TITLE' -MenuItems 'Exchange Server', 'Active Directory', 'Sytem Center Configuration Manager', 'Lync Server' -TitleColor Red  -MenuItemColor green
Show-AsciiMenu -Title 'EXIT STRATAGY' -MenuItems 'Shutdown and Restart system', 'Rerun Fast Cruise','Exit' #-Debug
$Raspberry = Read-Host 'Select Number'

switch($Raspberry)
{
1 {Restart-Computer}
2 {Start-FastCruise @FastCruiseSplat}
3 {Exit}

<#
$EmailMessage =  Read-Host "Message to Send"
Send-eMail @SplatSendEmailHelpDesk -MessageBody $EmailMessage
#>

}

}Until ($Raspberry)


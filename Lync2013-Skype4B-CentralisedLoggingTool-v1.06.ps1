########################################################################
# Name: Centralised Logging Tool
# Version: 1.06  ( 25/01/2019 )
# Created On: 11/4/2013
# Created By: James Cussen
# Site: www.myskypelab.com
# Based on: Randy Wintle's original Centralized Logging Tool - http://blog.ucmadeeasy.com/
# Description: This script offers a GUI for the Lync Centralised Logging Tool.
# Notes: Execute script on Lync or Skype server.
# 
# 1.01 Enhancements
#		Fixed the format of the Start/Stop export date format (set when the script loads) to the region setting of the machine. 
#		This was previously hardcoded to Australian format in v1.00. Lync expects the date format in the format of the region 
#		of the computer (as set in Control Panel) for the Search command. Thanks to John Crouch for reporting this not working 
#		in the US in v1.00.
#
# 1.02 Enhancements
#		- Added SBA's to the Pool list. They also have the centralised logging service. - Nice pickup by John Crouch.
#		- Collocated Persistent Chat servers now don't get added to the pool list.
#
# 1.03 Enhancements
#		- Pools listbox now is in alphabetical order
#		- Scenarios listbox is now in alphabetical order
#		- Components listbox is now in alphabetical order
#		- Annoying counting in PowerShell window when the script loads has now been removed
#		- Added prerequisite check for Snooper and debugging tools when script loads
#		- Script is now signed
#		- Using UP and DOWN arrow keys on pools listbox now updates GUI correctly
#		- If pool does not have a status then an error ("ERROR: Pool status not found. Check that CLS is running correctly on pool.") will be displayed in the GUI
#		- A status message will now display to indicate to select a Pool when no pool is selected in the Pools listbox
#		- Fixed error that is displayed when Cancel button on browse dialog is clicked
#		- Fixed issue with Analysing files with snooper that have spaces in them
#
# 1.04 Enhancements
#		- Now supports Skype for Business!
#
# 1.05 Enhancements
#		- Now checks for Skype for Business Snooper location.
#		- Changed the default trace saving location to C:\Tracing\
#
# 1.06 Enhancements
#		- Added support for Skype for Business 2019
#
########################################################################

$theVersion = $PSVersionTable.PSVersion
$MajorVersion = $theVersion.Major

Write-Host ""
Write-Host "--------------------------------------------------------------"
Write-Host "PowerShell Version Check..." -foreground "yellow"
if($MajorVersion -eq  "1")
{
	Write-Host "This machine only has Version 1 PowerShell installed.  This version of Powershell is not supported." -foreground "red"
}
elseif($MajorVersion -eq  "2")
{
	Write-Host "This machine has Version 2 PowerShell installed. This version of Powershell is not supported." -foreground "red"
}
elseif($MajorVersion -eq  "3")
{
	Write-Host "This machine has version 3 PowerShell installed. CHECK PASSED!" -foreground "green"
}
elseif($MajorVersion -eq  "4")
{
	Write-Host "This machine has version 4 PowerShell installed. CHECK PASSED!" -foreground "green"
}
elseif($MajorVersion -eq  "5")
{
	Write-Host "This machine has version 5 PowerShell installed. CHECK PASSED!" -foreground "green"
}
else
{
	Write-Host "This machine has version $MajorVersion PowerShell installed. Unknown level of support for this version." -foreground "yellow"
}
Write-Host "--------------------------------------------------------------"
Write-Host ""

Function Get-MyModule 
{ 
Param([string]$name) 
	
	if(-not(Get-Module -name $name)) 
	{ 
		if(Get-Module -ListAvailable | Where-Object { $_.name -eq $name }) 
		{ 
			Import-Module -Name $name 
			return $true 
		} #end if module available then import 
		else 
		{ 
			return $false 
		} #module not available 
	} # end if not module 
	else 
	{ 
		return $true 
	} #module already loaded 
} #end function get-MyModule 


$Script:LyncModuleAvailable = $false
$Script:SkypeModuleAvailable = $false

Write-Host "--------------------------------------------------------------"
#Import Lync Module
if(Get-MyModule "Lync")
{
	Invoke-Expression "Import-Module Lync"
	Write-Host "Imported Lync Module..." -foreground "green"
	$Script:LyncModuleAvailable = $true
}
else
{
	Write-Host "Unable to import Lync Module..." -foreground "yellow"
}
#Import SkypeforBusiness Module
if(Get-MyModule "SkypeforBusiness")
{
	Invoke-Expression "Import-Module SkypeforBusiness"
	Write-Host "Imported SkypeforBusiness Module..." -foreground "green"
	$Script:SkypeModuleAvailable = $true
}
else
{
	Write-Host "Unable to import SkypeforBusiness Module... (Expected on a Lync 2013 system)" -foreground "yellow"
}


####################################################################################################
##LOAD Variables that may be needed
$filtertype= "NoFilter"

#Global Variables - Edit these to whatever you want the default to be
$script:snooperLocation = "C:\Program Files\Microsoft Lync Server 2013\Debugging Tools\Snooper.exe"
$script:pathDialog = "C:\Tracing\"

#Prereq Check
Write-Host "INFO: SNOOPER not found. Checking other drives..." -foreground "yellow"		
$AllDrives = Get-PSDrive -PSProvider 'FileSystem'
$FoundDrive = $false
foreach($Drive in $AllDrives) #Find Lync Snooper
{
	[string]$DriveName = $Drive.Root
	$TestDrive = "${DriveName}Program Files\Microsoft Lync Server 2013\Debugging Tools\Snooper.exe"
	Write-Host "INFO: Checking: $TestDrive" -foreground "yellow"
	
	if (Test-Path $TestDrive)
	{
		Write-Host "Found Lync SNOOPER Drive. Using: $TestDrive" -foreground "green"
		Write-Host
		$script:snooperLocation = $TestDrive
		$FoundDrive = $true
		break
	}
	
}
foreach($Drive in $AllDrives) #Override with newer Skype4B Snooper if available
{
	
	[string]$DriveName = $Drive.Root
	$TestDrive = "${DriveName}Program Files\Skype for Business Server 2015\Debugging Tools\Snooper.exe"
	Write-Host "INFO: Checking: $TestDrive" -foreground "yellow"
	
	if (Test-Path $TestDrive)
	{
		Write-Host "Found Skype for Business 2015 SNOOPER Drive. Using: $TestDrive" -foreground "green"
		Write-Host
		$script:snooperLocation = $TestDrive
		$FoundDrive = $true
		break
	}
	
	$TestDrive = "${DriveName}Program Files\Skype for Business Server 2019\ClsAgent\Snooper.exe"
	Write-Host "INFO: Checking: $TestDrive" -foreground "yellow"
	
	if (Test-Path $TestDrive)
	{
		Write-Host "Found Skype for Business 2019 SNOOPER Drive. Using: $TestDrive" -foreground "green"
		Write-Host
		$script:snooperLocation = $TestDrive
		$FoundDrive = $true
		break
	}
}
if(!$FoundDrive)
{
	$script:snooperLocation = ""
	Write-Host "ERROR: Could not find a drive with SNOOPER installed on it. Please install Lync 2013 Debugging Tools to access Snooper." -foreground "red" 
	Write-Host
}


#Creating Global Variables
$script:StatusOfPools = @()
$script:pools = @()
$script:scenariotype = ""
$script:pooltotrace= ""
$script:duration = ""
$script:AgentStatus = ""


##DEFINE FUNCTIONS TO BE USED LATER##
####################################################################################################

##START TRACING
function StartTracing
{
	$statustext.text = "Starting Log, Please Wait..."			

	
	$loopNo = 0
	foreach($StatusOfPool in $StatusOfPools)
	{
		$thisPoolName = $StatusOfPool.PoolName
		$thisPoolStatus = $StatusOfPool.PoolStatus
		$thisPoolScenario = $StatusOfPool.PoolScenario
		$thisPoolComponents = $StatusOfPool.Components
		if($pooltotrace -eq $thisPoolName -and $thisPoolStatus -eq "Stopped")
		{
		
			Write-Host "EXECUTING START"
		
			if($durationCheckBox.checked)
			{	
				
				##Build Cmd to start the tracing with duration
				[string]$clscmd = "Start-CsClsLogging -Scenario $scenariotype -pools $pooltotrace -duration $duration"
				
				##Execute CLS Start tracing cmd
				Write-Host "------------------------------------------------------"
				Write-Host "RUNNING COMMAND: $clscmd" -foreground "Green"
				
				$pinfo = New-Object System.Diagnostics.ProcessStartInfo
				$pinfo.FileName = "powershell.exe"
				$pinfo.RedirectStandardError = $true
				$pinfo.RedirectStandardOutput = $false
				$pinfo.UseShellExecute = $false
				$pinfo.Arguments = "-command `"$clscmd`""
				$p = New-Object System.Diagnostics.Process
				$p.StartInfo = $pinfo
				$p.Start() | Out-Null
				$p.WaitForExit()
				$stderr = $p.StandardError.ReadToEnd()

				
				if($Script:SkypeModuleAvailable) #Skype for Business Processing
				{
					Get-AgentStatus
				}
				else #Lync 2013 Processing
				{
					$script:StatusOfPools[$loopNo] = @(@{"PoolName" = "$pooltotrace";"PoolStatus" = "Started";"PoolScenario" = "$scenariotype";"Components" = "$scenariotype";"AlwaysOn" = $StatusOfPools[$loopNo].AlwaysOn})
					
					if($stderr.Contains("Failed") -or $stderr.Contains("failed"))
					{
						$statustext.text = "COMMAND FAILURE: Check the powershell window for details - Retrieving Pool status..."
						Write-Host "COMMAND FAILURE: $stderr" -ForegroundColor red
						Get-AgentStatus
						$statustext.text = "COMMAND FAILURE: Check the powershell window for details."
					}
					else
					{
						Write-Host "$stderr"
					}
				}
				Write-Host "------------------------------------------------------"

			}
			else
			{	
				##Build Cmd to start the tracing without duration
				[string]$clscmd = "Start-CsClsLogging -Scenario $scenariotype -pools $pooltotrace"
				
				##Execute CLS Start tracing cmd
				Write-Host "------------------------------------------------------"
				Write-Host "RUNNING COMMAND: $clscmd" -foreground "Green"
				
				$pinfo = New-Object System.Diagnostics.ProcessStartInfo
				$pinfo.FileName = "powershell.exe"
				$pinfo.RedirectStandardError = $true
				$pinfo.RedirectStandardOutput = $false
				$pinfo.UseShellExecute = $false
				$pinfo.Arguments = "-command `"$clscmd`""
				$p = New-Object System.Diagnostics.Process
				$p.StartInfo = $pinfo
				$p.Start() | Out-Null
				$p.WaitForExit()
				$stderr = $p.StandardError.ReadToEnd()

								
				if($Script:SkypeModuleAvailable) #Skype for Business Processing
				{
					Get-AgentStatus
				}
				else #Lync 2013 Processing
				{
					$script:StatusOfPools[$loopNo] = @(@{"PoolName" = "$pooltotrace";"PoolStatus" = "Started";"PoolScenario" = "$scenariotype";"Components" = "$scenariotype";"AlwaysOn" = $StatusOfPools[$loopNo].AlwaysOn})
					
					if($stderr.Contains("Failed") -or $stderr.Contains("failed"))
					{
						$statustext.text = "COMMAND FAILURE: Check the PowerShell window for details - Retrieving Pool status..."
						Write-Host "COMMAND FAILURE: $stderr" -ForegroundColor red
						Get-AgentStatus
						$statustext.text = "COMMAND FAILURE: Check the PowerShell window for details."
					}
					else
					{
						Write-Host "$stderr"
					}
				}
				Write-Host "------------------------------------------------------"
				
			}

			break
		}
		$loopNo++
	}
	

	$statustext.text = ""

	##Notify Host that the tracing is started
	Write-Host "Tracing Started" -Foreground "Green"
	Write-Host
	
	Update-Buttons
	
}
##END START TRACING
####################################################################################################


##UPDATE TRACING
function UpdateTracing
{		
	$statustext.text = "Updating Trace, Please wait..."	
	
	##Build Cmd to start the tracing
	if($durationCheckBox.Checked)
	{
		[string]$clscmd = "Update-CsClsLogging -pools $pooltotrace -duration $duration"
		
		##Execute CLS Start tracing cmd
		Write-Host "------------------------------------------------------"
		Write-Host "RUNNING COMMAND: $clscmd" -foreground "Green"
		Invoke-Expression $clscmd
		Write-Host "------------------------------------------------------"

		$statustext.text= "Trace Updating..."

		##Notify Host that the tracing is started
		Write-Host "Tracing Updated" -Foreground "Green" 
	}
	else
	{
		Write-Host "ERROR: No Duration has been specified." -ForegroundColor red
	}
	Update-Buttons
	$statustext.text = ""
}
##END UPDATE TRACING
####################################################################################################


##STOP TRACING
Function StopTracing
{
	$statustext.text = "Stopping Logging, Please Wait..."
	
	#Change the status of the pool to Stopped
	$loopNo = 0
	foreach($StatusOfPool in $StatusOfPools)
	{
		$thisPoolName = $StatusOfPool.PoolName
		$thisPoolStatus = $StatusOfPool.PoolStatus
		$thisPoolScenario = $StatusOfPool.PoolScenario
		$thisPoolComponents = $StatusOfPool.Components
		if($pooltotrace -eq $thisPoolName -and $thisPoolStatus -eq "Started")
		{
		
			Write-Host "EXECUTING STOP"
		
			##Build CLS Stop Cmd
			[string]$clscmd = "Stop-CsClsLogging -pools $thisPoolName -scenario $thisPoolScenario"
			Write-Host "------------------------------------------------------"
			Write-Host "RUNNING COMMAND: $clscmd" -foreground "green"
			
			$pinfo = New-Object System.Diagnostics.ProcessStartInfo
			$pinfo.FileName = "powershell.exe"
			$pinfo.RedirectStandardError = $true
			$pinfo.RedirectStandardOutput = $false
			$pinfo.UseShellExecute = $false
			$pinfo.Arguments = "-command `"$clscmd`""
			$p = New-Object System.Diagnostics.Process
			$p.StartInfo = $pinfo
			$p.Start() | Out-Null
			$p.WaitForExit()
			$stderr = $p.StandardError.ReadToEnd()
						
			
			if($Script:SkypeModuleAvailable) #Skype for Business Processing
			{
				Get-AgentStatus
				#Update-Buttons
			}
			else #Lync 2013 Processing
			{
				$script:StatusOfPools[$loopNo] = @(@{"PoolName" = "$pooltotrace";"PoolStatus" = "Stopped";"PoolScenario" = "";"Components" = "$thisPoolComponents"; "AlwaysOn" = $StatusOfPools[$loopNo].AlwaysOn})
				
				if($stderr.Contains("Failed") -or $stderr.Contains("failed"))
				{
					$statustext.text = "COMMAND FAILURE: Check the PowerShell window for details - Retrieving Pool status..."
					Write-Host "COMMAND FAILURE: $stderr" -ForegroundColor red
					Get-AgentStatus
					$statustext.text = "COMMAND FAILURE: Check the PowerShell window for details."
				}
				else
				{
					Write-Host "$stderr"
				}
				
			}
			Write-Host "------------------------------------------------------"
						
			Update-Buttons
			$statustext.text = ""
			break
		}
		$loopNo++
	}
	
	##Notify host tracing has stopped
	Write-Host "Tracing Ended" -Foreground "Green" 
	Write-Host
}
##END STOP TRACING
####################################################################################################


##START SEARCH EXPORT
function ExportLogs
{
	$compfilter = $complistbox.selecteditems
	$filtertype = $filterlist.selecteditem
	$filepath = $pathbox.text
	$loglevel = $levellistbox.selecteditem
	$poolListItemSelected = $poollistbox.SelectedItem
	$phoneFilter = $filterPhonetext.text
	$uriFilter = $filterURItext.text
	$callIDFilter = $filterCallIDtext.text
	$ipFilter = $filterIPtext.text
	$sipContentsFilter = $sipContentstext.text
	
	
	#Disable all GUI components
	$StartButton.enabled = $false
	$StopButton.enabled = $false
	$ExportButton.enabled = $false
	$UpdateButton.enabled = $true
	$ExportButton.enabled = $false
	#Filter components
	$levellistbox.enabled = $false
	$filterPhonetext.enabled = $false
	$filterURItext.enabled = $false
	$filterCallIDtext.enabled = $false
	$filterIPtext.enabled = $false
	$matchTypeCheckBox.enabled = $false
	$timeCheckBox.enabled = $false
	$complistbox.enabled = $false
	$sipContentstext.enabled = $false
					

	#create file name
	$FileName = $poolListItemSelected + "_TRACE_"
	
	#Create tracing Directory if it doesn't exist
	if (!(Test-Path $filepath)) {md $filepath}

	##Define full path path + name
	[string]$filepathfull = $filepath + $filename 

	$statustext.text = "Exporting... Path: " + $filepath 
	##Start building CLS Search Cmds
		
	#Default commands	
	[string]$clscmd = "Search-CsClsLogging -loglevel $loglevel -pools $poolListItemSelected"
	
	$filepathfull += "_LEVEL_$loglevel"

	#Start / End Time Filter
	if($timeCheckBox.checked)
	{
		$startTime = $startTimeTextBox.Text
		$endTime = $stopTimeTextBox.Text
		$clscmd += " -StartTime `"$startTime`" -EndTime `"$endTime`"" 
		
		$filepathStart = $startTime.Replace(":",".").Replace("/","-").Replace(" ","_")
		$filepathStop = $endTime.Replace(":",".").Replace("/","-").Replace(" ","_")
		$filepathfull += "_START-${filepathStart}_STOP-${filepathStop}"
	}
	# Can't be used in filenames: / ? < > \ : * | "
	if($phoneFilter -ne "None" -and $phoneFilter -ne "")
	{
		$clscmd += " -Phone `"$phoneFilter`"" 
		
		$filt = $phoneFilter.Replace("/","-").Replace("?","-").Replace("<","-").Replace(">","-").Replace("\","-").Replace(":","-").Replace("*","-").Replace("|","-")
		$filepathfull += "_PHONE-$filt"
	}
	if($uriFilter -ne "None" -and $uriFilter -ne "")
	{
		$clscmd += " -Uri `"$uriFilter`""
		
		$filt = $uriFilter.Replace("/","-").Replace("?","-").Replace("<","-").Replace(">","-").Replace("\","-").Replace(":","-").Replace("*","-").Replace("|","-")
		$filepathfull += "_URI-$filt"
	}
	if($callIDFilter -ne "None" -and $callIDFilter -ne "")
	{
		$clscmd += " -Callid `"$callIDFilter`""
		
		$filt = $callIDFilter.Replace("/","-").Replace("?","-").Replace("<","-").Replace(">","-").Replace("\","-").Replace(":","-").Replace("*","-").Replace("|","-")
		$filepathfull += "_CALLID-$filt"
	}
	if($ipFilter -ne "None" -and $ipFilter -ne "")
	{
		$clscmd += " -IP `"$ipFilter`""
		
		$filt = $ipFilter.Replace("/","-").Replace("?","-").Replace("<","-").Replace(">","-").Replace("\","-").Replace(":","-").Replace("*","-").Replace("|","-")
		$filepathfull += "_IP-$filt"
	}
		
	#Match All or Match Any
	if($matchTypeCheckBox.checked)
	{
		$clscmd += " -MatchAny"
		$filepathfull += "_MATCHANY"		
	}
	else
	{
		$clscmd += " -MatchAll"
		$filepathfull += "_MATCHALL"
	}
	#SIP Contents Filter
	if($sipContentsFilter -ne "None" -and $sipContentsFilter -ne "")
	{
		$clscmd += " -SipContents `"$sipContentsFilter`""
		
		$filt = $sipContentsFilter.Replace("/","-").Replace("?","-").Replace("<","-").Replace(">","-").Replace("\","-").Replace(":","-").Replace("*","-").Replace("|","-")
		$filepathfull += "_SIP-$filt"
	}
	#Component Filter - This is left till last as it's the least important information that need to be in the filename.
	if($compfilter -ne $null)
	{
		$clscmd += " -components "
		$filepathfull += "_COMP-"
		
		$firstCheck = $true
		foreach ($objItem in $compfilter)
		{
			if($firstCheck)
			{
				$firstCheck = $false
				$clscmd += "`"$objItem`""
				
				$filt = $objItem.Replace("/","-").Replace("?","-").Replace("<","-").Replace(">","-").Replace("\","-").Replace(":","-").Replace("*","-").Replace("|","-")
				$filepathfull += "+$objItem"
			}
			elseif($objItem -ne "None")
			{
				$clscmd += ",`"$objItem`""
				
				$filt = $objItem.Replace("/","-").Replace("?","-").Replace("<","-").Replace(">","-").Replace("\","-").Replace(":","-").Replace("*","-").Replace("|","-")
				$filepathfull += "+$objItem"
			}
		}
	}
	
	$filepathfull += "_" + (Get-Date -Format s).Replace(":","-")
	
	$filenameLength = $filepathfull.length
	
	#check length of file name, must be under 260 chars
	if($filenameLength -gt 255)
	{
		Write-Host "Filename is too long. Truncating Name." -ForegroundColor red 
		$filepathfull = $filepathfull.substring(0,255) + ".txt"
	}
	else
	{
		$filepathfull += ".txt"
	}
	
	$clscmd += " -OutputFilePath `"$filepathfull`""
	
	Write-Host "------------------------------------------------------"
	Write-Host "RUNNING COMMAND: $clscmd" -Foreground "Green"
	Write-Host 
	Invoke-Expression $clscmd
	Write-Host "------------------------------------------------------"
	
		
	##Notify the host that the files have been exported
	Write-Host "Files have been exported to:" -Foreground "Green"
	Write-Host "$filepathfull" -Foreground "Green"
	Write-Host

	Update-Buttons
	$statustext.text = ""

}
##END SEARCH EXPORT
###################################################################################################


##START FUNCTION TO QUERY COMPONENTS
function PopulateComponents
{
	$statustext.text = "Populating Available Components..."
	$poolListItemSelected = $poollistbox.SelectedItem
	$loopNo = 0
	$complistbox.items.Clear()
	
	foreach($StatusOfPool in $StatusOfPools)
	{
		$thisPoolName = $StatusOfPool.PoolName
		$thisPoolStatus = $StatusOfPool.PoolStatus
		$thisPoolScenario = $StatusOfPool.PoolScenario
		$thisPoolComponents = $StatusOfPool.Components
		if($poolListItemSelected -eq $thisPoolName -and !$compCheckbox.Checked)
		{
			if($thisPoolScenario -ne "")
			{
				#build string to add Global/ in to scenario name
				[string]$compscenariostring = "Global/" + $thisPoolScenario
			}
			elseif($thisPoolComponents -ne "")
			{
				#build string to add Global/ in to scenario name
				[string]$compscenariostring = "Global/" + $thisPoolComponents
			}
			else
			{
				#build string to add Global/ in to scenario name
				[string]$compscenariostring = "Global/AlwaysOn"
			}
			
			#grab the scenario being used
			$scenario = Get-CSCLSScenario -Identity $compscenariostring
			##run through the scenario to pull component names

			foreach ($sc in $scenario.provider) {[void] $complistbox.Items.Add($sc.name)}	

			$statustext.text = ""
			break
		}
		elseif($poolListItemSelected -eq $thisPoolName)
		{
			$scenario = Invoke-Expression "Get-CSCLSScenario -Identity Global/AlwaysOn"
			foreach ($sc in $scenario.provider) {[void] $complistbox.Items.Add($sc.name)}
			break			
		}
		$loopNo++
	}
	$statustext.text = ""
	
}
##END FUNCTION TO QUERY COMPONENTS
###################################################################################################


function Get-AgentStatus
{
	#SKYPE FOR BUSINESS CHECK
	if($Script:SkypeModuleAvailable)
	{
		$script:StatusOfPools = @()
		
		$Status = Show-CsClsLogging
		$Status
		
		#Write-Host "---------UNIQUE POOL CHECK----------" ##DEBUGGING
		$a = $Status | select-object PoolFQDN 
		$b = $Status | select-object PoolFQDN | select -uniq
		$UniquePools = Compare-object –referenceobject $b –differenceobject $a

		foreach($Pool in $UniquePools.InputObject)
		{
			$thePool = $Pool.PoolFQDN
			$UniqueScenrios = $Status | where-object {$_.PoolFQDN -eq $thePool} | select-object ScenarioName | Group-Object ScenarioName 
			
			#$UniqueScenrios  ##DEBUGGING
			#Write-Host "No of Scenarios: " $UniqueScenrios.length  ##DEBUGGING
			
			if($UniqueScenrios.length -gt 1)
			{
				Write-Host ""
				Write-Host "WARNING: Pool $thePool currently has servers running different scenarios. This may cause issues with the operation of this tool." -foreground "Yellow"
				Write-Host ""
				break
			}
		}
		#Write-Host "---------UNIQUE POOL END----------" ##DEBUGGING

		
		foreach($theStatus in $Status)
		{		
			$StatusMachineFQDN = $theStatus.MachineFqdn
			$StatusPool = $theStatus.PoolFqdn
			$StatusScenarioName = $theStatus.ScenarioName
			$StatusAlwaysOn = $theStatus.AlwaysOn
			$StatusResponseMessage = $theStatus.ResponseMessage
			
			#Write-Host "POOL: $StatusPool MACHINE: $StatusMachineFQDN SCENARIO: $StatusScenarioName ALWAYSON: $StatusAlwaysOn RESPONSE: $StatusResponseMessage " ##DEBUGGING
			
			if($StatusAlwaysOn -eq $false)
			{
				if($StatusScenarioName -ne "" -and $StatusScenarioName -ne $null)
				{
					$script:StatusOfPools += @(@{"PoolName" = "$StatusPool";"PoolStatus" = "Started";"PoolScenario" = "$StatusScenarioName";"Components" = "$StatusScenarioName"; "AlwaysOn" = "No"})
				}
				else
				{
					$script:StatusOfPools += @(@{"PoolName" = "$StatusPool";"PoolStatus" = "Stopped";"PoolScenario" = "";"Components" = ""; "AlwaysOn" = "No"})
				}
			}
			elseif($StatusAlwaysOn -eq $true)
			{
				if($StatusScenarioName -ne "" -and $StatusScenarioName -ne $null)
				{	
					$script:StatusOfPools += @(@{"PoolName" = "$StatusPool";"PoolStatus" = "Started";"PoolScenario" = "$StatusScenarioName";"Components" = "$StatusScenarioName"; "AlwaysOn" = "Yes"})
				}
				else
				{
					$script:StatusOfPools += @(@{"PoolName" = "$StatusPool";"PoolStatus" = "Stopped";"PoolScenario" = "";"Components" = ""; "AlwaysOn" = "Yes"})
				}
			}
			elseif($StatusAlwaysOn -eq $null)
			{
				Write-Host ""
				Write-Host "INFO: ${StatusPool} does not have an Always On state. Assuming the pool is unavailable." -foreground "Yellow"
			}
		}
		Write-Host ""
	}
	else # Or do it the Lync 2013 way...
	{
		$pinfo = New-Object System.Diagnostics.ProcessStartInfo
		$pinfo.FileName = "powershell.exe"
		$pinfo.RedirectStandardError = $true
		$pinfo.RedirectStandardOutput = $true
		$pinfo.UseShellExecute = $false
		$pinfo.Arguments = "-command Show-CsClsLogging"
		$p = New-Object System.Diagnostics.Process
		$p.StartInfo = $pinfo
		$p.Start() | Out-Null
		$p.WaitForExit()
		$stdout = $p.StandardOutput.ReadToEnd()
		$stderr = $p.StandardError.ReadToEnd()
		Write-Host "`$stderr $stderr"
		Write-Host "`$stdout $stdout"
		
		$script:AgentStatus =  $stderr
		Parse-Status
	}
}

function Parse-Status
{
	#Parsing Text like shown below:
	#Success Code - 0, Successful on 1 agents
	#Tracing Status:
	#2013STDFE001.domain.com (2013STDFE001 v5.0.8308.0) (AlwaysOn=No,Scenario=IncomingAndOutgoingCall,Started=3/30/2013 6:37:27 AM,By=2013STDPC\Administrator,Duration=0.04:00)
	#2013STDFE001.domain.com (2013STDFE001 v5.0.8308.0) (Same as pool)
	#exit code:  + 0
	
	$script:StatusOfPools = @()
	$AgentStrings = $AgentStatus.split("`r`n")
	
	foreach($AgentString in $AgentStrings)
	{
		
		foreach($pool in $pools)
		{
			
			if($AgentString.Contains($pool) -and $AgentString.Contains("AlwaysOn=Yes") -and !$AgentString.Contains("Same as pool"))
			{
				$script:StatusOfPools += @(@{"PoolName" = "$pool";"PoolStatus" = "";"PoolScenario" = "";"Components" = ""; "AlwaysOn" = "Yes"})
			}
			elseif($AgentString.Contains($pool) -and !$AgentString.Contains("Same as pool"))
			{
				$script:StatusOfPools += @(@{"PoolName" = "$pool";"PoolStatus" = "";"PoolScenario" = "";"Components" = ""; "AlwaysOn" = "No"})
			}
			
			if($AgentString.Contains($pool) -and $AgentString.Contains("Scenario=") -and !$AgentString.Contains("Same as pool"))
			{
				$pos = $AgentString.IndexOf("Scenario=")
				$rightPart = $AgentString.Substring($pos+9)
				$scenarioArray = $rightPart.Split(",")
				$scenarioName = $scenarioArray[0]
				
				$loopNo = 0
				foreach($StatusOfPool in $StatusOfPools)
				{
					if($StatusOfPool.PoolName -eq $pool)
					{
						$script:StatusOfPools[$loopNo] = @(@{"PoolName" = "$pool";"PoolStatus" = "Started";"PoolScenario" = "$scenarioName";"Components" = "$scenarioName";"AlwaysOn" = $StatusOfPool.AlwaysOn})
					}
					$loopNo++
				}
			}
			elseif($AgentString.Contains($pool) -and !$AgentString.Contains("Same as pool"))
			{
				$loopNo = 0
				foreach($StatusOfPool in $StatusOfPools)
				{
					if($StatusOfPool.PoolName -eq $pool)
					{
						$script:StatusOfPools[$loopNo] = @(@{"PoolName" = "$pool";"PoolStatus" = "Stopped";"PoolScenario" = "";"Components" = "";"AlwaysOn" = $StatusOfPool.AlwaysOn})
					}
					$loopNo++
				}
			}
		}
	}
}

function AlwaysOn-Change
{
	$poolListItemSelected = $poollistbox.SelectedItem
	$alwaysOnLabel.text = ""
	
	$loopNo = 0
	foreach($StatusOfPool in $StatusOfPools)
	{
		$thisPoolName = $StatusOfPool.PoolName
		$thisPoolAlwaysOn = $StatusOfPool.AlwaysOn
		if($poolListItemSelected -eq $thisPoolName -and $thisPoolAlwaysOn -eq "Yes")
		{	
			$statustext.text = "Disabling AlwaysOn for $thisPoolName..."
			$script:StatusOfPools[$loopNo] = @(@{"PoolName" = $StatusOfPools[$loopNo].PoolName;"PoolStatus" = $StatusOfPools[$loopNo].PoolStatus;"PoolScenario" = $StatusOfPools[$loopNo].PoolScenario;"Components" = $StatusOfPools[$loopNo].Components;"AlwaysOn" = "No"})

			##Build Cmd to start the tracing without duration
			[string]$clscmd = "Stop-CsClsLogging -Scenario AlwaysOn -pools $thisPoolName"
			
			Write-Host "------------------------------------------------------"
			Write-Host "RUNNING COMMAND: $clscmd" -foreground "Green"
			
			$pinfo = New-Object System.Diagnostics.ProcessStartInfo
			$pinfo.FileName = "powershell.exe"
			$pinfo.RedirectStandardError = $true
			$pinfo.RedirectStandardOutput = $false
			$pinfo.UseShellExecute = $false
			$pinfo.Arguments = "-command `"$clscmd`""
			$p = New-Object System.Diagnostics.Process
			$p.StartInfo = $pinfo
			$p.Start() | Out-Null
			$p.WaitForExit()
			$stderr = $p.StandardError.ReadToEnd()
			

			$statustext.text = ""
			
			if($stderr.Contains("Failed") -or $stderr.Contains("failed"))
			{
				$statustext.text = "COMMAND FAILURE: Check the PowerShell window for details - Retrieving Pool status..."
				Write-Host "COMMAND FAILURE: $stderr" -ForegroundColor red
				Get-AgentStatus
				$statustext.text = "COMMAND FAILURE: Check the PowerShell window for details."
			}
			else
			{
				Write-Host "$stderr"
			}
			Write-Host "------------------------------------------------------"
						
			
			break
		}
		elseif($poolListItemSelected -eq $thisPoolName -and $thisPoolAlwaysOn -eq "No")
		{
			$statustext.text = "Enabling AlwaysOn for $thisPoolName..."
			
			$script:StatusOfPools[$loopNo] = @(@{"PoolName" = $StatusOfPools[$loopNo].PoolName;"PoolStatus" = $StatusOfPools[$loopNo].PoolStatus;"PoolScenario" = $StatusOfPools[$loopNo].PoolScenario;"Components" = $StatusOfPools[$loopNo].Components;"AlwaysOn" = "Yes"})
			##Build Cmd to start the tracing without duration
			[string]$clscmd = "Start-CsClsLogging -Scenario AlwaysOn -pools $thisPoolName"
			Write-Host "------------------------------------------------------"
			Write-Host "RUNNING COMMAND: $clscmd" -foreground "green"

			$pinfo = New-Object System.Diagnostics.ProcessStartInfo
			$pinfo.FileName = "powershell.exe"
			$pinfo.RedirectStandardError = $true
			$pinfo.RedirectStandardOutput = $false
			$pinfo.UseShellExecute = $false
			$pinfo.Arguments = "-command `"$clscmd`""
			$p = New-Object System.Diagnostics.Process
			$p.StartInfo = $pinfo
			$p.Start() | Out-Null
			$p.WaitForExit()
			$stderr = $p.StandardError.ReadToEnd()

			$statustext.text = ""
			
			if($stderr.Contains("Failed") -or $stderr.Contains("failed"))
			{
				$statustext.text = "COMMAND FAILURE: Check the PowerShell window for details - Retrieving Pool status..."
				Write-Host "COMMAND FAILURE: $stderr" -ForegroundColor red
				Get-AgentStatus
				$statustext.text = "COMMAND FAILURE: Check the PowerShell window for details."
			}
			else
			{
				Write-Host "$stderr"
			}
			Write-Host "------------------------------------------------------"
			
			break
		}
		$loopNo++
	}

	Update-Buttons
}

function Update-Buttons
{
	$poolListItemSelected = $poollistbox.SelectedItem
	$scenarioListBoxItemSelected = $scenarioListBox.SelectedItem
	$statustext.text = ""
	
	if($poolListItemSelected -ne $null)
	{
		if($poolListItemSelected.contains("-----FRONT END SERVERS-----") -or $poolListItemSelected.contains("-----MEDIATION SERVERS-----") -or $poolListItemSelected.contains("-----PERSISTENT CHAT SERVERS-----") -or $poolListItemSelected.contains("-----EDGE SERVERS-----") -or $scenarioListBoxItemSelected -eq $null -or $poolListItemSelected -eq $null)
		{
			$StartButton.enabled = $false
			$StopButton.enabled = $false
			$ExportButton.enabled = $false
			$UpdateButton.enabled = $false
			$stopLabel.Text = ""
			$alwaysOnLabel.Text = ""
			$alwaysOnButton.enabled = $false
			#Filter components
			$levellistbox.enabled = $false
			$filterPhonetext.enabled = $false
			$filterURItext.enabled = $false
			$filterCallIDtext.enabled = $false
			$filterIPtext.enabled = $false
			$matchTypeCheckBox.enabled = $false
			$timeCheckBox.enabled = $false
			$complistbox.enabled = $false
			$levellistbox.enabled = $false
			$sipContentstext.enabled = $false
			$compCheckbox.enabled = $false
			$durationTextBox.Enabled = $false
			$durationCheckBox.Enabled = $false
			$startTimeTextBox.Enabled = $false
			$stopTimeTextBox.Enabled = $false
			$statustext.text = "INFO: Please select a pool..."
			
		}
		else
		{
			$durationCheckBox.Enabled = $true
			if($durationCheckBox.checked)
			{
				$durationTextBox.Enabled = $true
			}
			
			$timeCheckBox.enabled = $true
			if($timeCheckBox.checked)
			{
			$startTimeTextBox.Enabled = $true
			$stopTimeTextBox.Enabled = $true
			}
			
			$foundPool = $false
			$poolCount = $StatusOfPools.length
			$poolLoopNo = 0
			foreach($StatusOfPool in $StatusOfPools)
			{
				$thisPoolName = $StatusOfPool.PoolName
				$thisPoolStatus = $StatusOfPool.PoolStatus
				$thisPoolScenario = $StatusOfPool.PoolScenario
				$thisPoolComponents = $StatusOfPool.Components
				$thisPoolAlwaysOn = $StatusOfPool.AlwaysOn
				
				
				if($poolListItemSelected -eq $thisPoolName -and $thisPoolAlwaysOn -eq "Yes")
				{
					$alwaysOnButton.enabled = $true
					$alwaysOnLabel.Text = "AlwaysOn Enabled"
					$alwaysOnButton.forecolor = "green"
					$alwaysOnLabel.forecolor = "green"
				}
				else
				{
					$alwaysOnButton.enabled = $true
					$alwaysOnLabel.Text = "AlwaysOn Disabled"
					$alwaysOnButton.forecolor = "red"
					$alwaysOnLabel.forecolor = "red"
				}
				
				if($poolListItemSelected -eq $thisPoolName -and $thisPoolStatus -eq "Started")
				{
					$StartButton.enabled = $false
					$StopButton.enabled = $true
					$ExportButton.enabled = $false
					$UpdateButton.enabled = $true
					$ExportButton.enabled = $true

					
					#Filter components
					$levellistbox.enabled = $true
					$filterPhonetext.enabled = $true
					$filterURItext.enabled = $true
					$filterCallIDtext.enabled = $true
					$filterIPtext.enabled = $true
					$matchTypeCheckBox.enabled = $true
					$timeCheckBox.enabled = $true
					$complistbox.enabled = $true
					$sipContentstext.enabled = $true
					$compCheckbox.enabled = $true
					
					PopulateComponents
					$stopLabel.forecolor = "green"
					$stopLabel.Text = "Currently Logging: $thisPoolScenario"
				}
				elseif($poolListItemSelected -eq $thisPoolName -and $thisPoolStatus -eq "Stopped")
				{
					$StartButton.enabled = $true
					$StopButton.enabled = $false
					$ExportButton.enabled = $true
					$UpdateButton.enabled = $false
					
					#Filter components
					$levellistbox.enabled = $true
					$filterPhonetext.enabled = $true
					$filterURItext.enabled = $true
					$filterCallIDtext.enabled = $true
					$filterIPtext.enabled = $true
					$matchTypeCheckBox.enabled = $true
					$timeCheckBox.enabled = $true
					$complistbox.enabled = $true
					$sipContentstext.enabled = $true
					$compCheckbox.enabled = $true
					
					PopulateComponents
					if($thisPoolComponents -ne "")
					{
						$stopLabel.forecolor = "orange"
						$stopLabel.Text = "Export Available: $thisPoolComponents"
						$ExportButton.enabled = $true
					}
					else
					{
						$stopLabel.Text = ""
						$ExportButton.enabled = $true
					}
				}
				
				$poolLoopNo++
				if($poolListItemSelected -eq $thisPoolName)
				{
					break
				}
				else
				{
					if($poolCount -eq $poolLoopNo)
					{
						$StartButton.enabled = $false
						$StopButton.enabled = $false
						$ExportButton.enabled = $false
						$UpdateButton.enabled = $false
						$ExportButton.enabled = $false
						$AlwaysOnButton.enabled = $false

						#Filter components
						$levellistbox.enabled = $false
						$filterPhonetext.enabled = $false
						$filterURItext.enabled = $false
						$filterCallIDtext.enabled = $false
						$filterIPtext.enabled = $false
						$matchTypeCheckBox.enabled = $false
						$timeCheckBox.enabled = $false
						$complistbox.enabled = $false
						$sipContentstext.enabled = $false
						$compCheckbox.enabled = $false
						$stopLabel.Text = ""
						$alwaysOnLabel.Text = ""
						$statustext.text = "ERROR: Pool status not found. Check that CLS is running correctly on pool."
					}
				}
				
			}
		}
	}
	else
	{
		$StartButton.enabled = $false
		$StopButton.enabled = $false
		$ExportButton.enabled = $false
		$UpdateButton.enabled = $false
	}
		
}

##START DRAWING UI
###################################################################################################

#Load Params
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

##Draw Main UI
$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Centralised Logging Tool v1.06"
$objForm.Size = New-Object System.Drawing.Size(600,630) 
$objForm.StartPosition = "CenterScreen"
[byte[]]$WindowIcon = @(137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 32, 0, 0, 0, 32, 8, 6, 0, 0, 0, 115, 122, 122, 244, 0, 0, 0, 6, 98, 75, 71, 68, 0, 255, 0, 255, 0, 255, 160, 189, 167, 147, 0, 0, 0, 9, 112, 72, 89, 115, 0, 0, 11, 19, 0, 0, 11, 19, 1, 0, 154, 156, 24, 0, 0, 0, 7, 116, 73, 77, 69, 7, 225, 7, 26, 1, 36, 51, 211, 178, 227, 235, 0, 0, 5, 235, 73, 68, 65, 84, 88, 195, 197, 151, 91, 108, 92, 213, 21, 134, 191, 189, 207, 57, 115, 159, 216, 78, 176, 27, 98, 72, 226, 88, 110, 66, 66, 34, 185, 161, 168, 193, 73, 21, 17, 2, 2, 139, 75, 164, 182, 106, 145, 170, 190, 84, 74, 104, 65, 16, 144, 218, 138, 138, 244, 173, 69, 106, 101, 42, 129, 42, 149, 170, 162, 15, 168, 168, 151, 7, 4, 22, 180, 1, 41, 92, 172, 52, 196, 68, 105, 130, 19, 138, 98, 76, 154, 27, 174, 227, 248, 58, 247, 57, 103, 175, 62, 236, 241, 177, 199, 246, 140, 67, 26, 169, 251, 237, 236, 61, 179, 215, 191, 214, 191, 214, 191, 214, 86, 188, 62, 37, 252, 31, 151, 174, 123, 42, 224, 42, 72, 56, 138, 152, 99, 191, 175, 247, 114, 107, 29, 172, 75, 106, 94, 254, 74, 156, 109, 13, 58, 180, 155, 53, 240, 216, 64, 129, 63, 156, 43, 95, 55, 0, 106, 62, 5, 158, 134, 83, 59, 147, 116, 36, 106, 7, 103, 188, 44, 228, 13, 120, 202, 126, 151, 12, 100, 3, 225, 183, 231, 203, 60, 55, 88, 66, 4, 80, 215, 0, 96, 89, 68, 113, 97, 87, 138, 180, 3, 163, 101, 120, 116, 160, 192, 161, 81, 159, 203, 69, 33, 230, 40, 58, 27, 52, 251, 215, 69, 248, 198, 74, 183, 238, 165, 175, 141, 248, 60, 114, 178, 192, 165, 188, 44, 9, 100, 22, 128, 192, 127, 238, 73, 209, 18, 81, 252, 109, 52, 224, 222, 247, 179, 179, 46, 206, 93, 102, 142, 119, 193, 76, 216, 96, 247, 13, 46, 223, 189, 201, 101, 207, 74, 143, 148, 99, 183, 159, 250, 184, 72, 207, 96, 169, 46, 136, 16, 192, 183, 91, 61, 94, 233, 140, 241, 81, 198, 176, 229, 173, 204, 226, 198, 175, 102, 5, 194, 243, 157, 113, 246, 221, 236, 225, 42, 232, 29, 9, 184, 255, 104, 174, 62, 0, 165, 192, 239, 78, 163, 129, 174, 195, 57, 14, 143, 5, 255, 115, 114, 197, 29, 197, 200, 221, 41, 82, 14, 188, 63, 30, 240, 245, 190, 220, 162, 145, 208, 0, 141, 174, 66, 1, 37, 129, 195, 163, 254, 34, 40, 1, 191, 70, 25, 250, 50, 75, 197, 156, 149, 15, 132, 27, 254, 62, 205, 229, 178, 176, 163, 201, 161, 103, 115, 172, 182, 14, 196, 181, 53, 114, 38, 107, 64, 22, 194, 92, 147, 80, 200, 67, 105, 50, 247, 165, 171, 156, 104, 141, 105, 70, 186, 211, 200, 131, 105, 214, 46, 82, 53, 69, 3, 119, 244, 217, 240, 63, 177, 214, 35, 233, 170, 250, 66, 164, 20, 11, 221, 52, 240, 171, 77, 49, 114, 6, 198, 74, 18, 158, 106, 5, 239, 110, 79, 208, 236, 41, 254, 93, 16, 206, 102, 204, 162, 30, 14, 78, 27, 158, 60, 93, 68, 1, 7, 191, 150, 176, 73, 60, 31, 64, 182, 178, 185, 49, 169, 103, 80, 132, 235, 166, 164, 38, 238, 64, 66, 67, 104, 94, 224, 229, 206, 56, 111, 93, 182, 116, 61, 246, 81, 177, 118, 166, 107, 248, 253, 121, 43, 92, 119, 52, 106, 86, 39, 245, 66, 0, 147, 101, 9, 105, 188, 171, 165, 186, 198, 127, 179, 57, 202, 233, 233, 106, 216, 9, 79, 113, 169, 96, 216, 119, 179, 135, 47, 112, 240, 114, 185, 110, 169, 77, 149, 132, 95, 159, 181, 32, 182, 54, 58, 139, 83, 112, 231, 7, 121, 0, 126, 210, 17, 129, 96, 150, 134, 213, 9, 205, 84, 185, 42, 29, 121, 103, 91, 130, 15, 38, 45, 228, 105, 95, 40, 207, 97, 173, 209, 83, 124, 179, 213, 227, 153, 13, 81, 16, 91, 205, 247, 174, 116, 113, 42, 118, 31, 89, 227, 86, 37, 109, 8, 224, 189, 97, 159, 178, 64, 71, 82, 207, 166, 129, 192, 75, 231, 203, 180, 68, 170, 235, 252, 95, 57, 195, 150, 138, 218, 156, 43, 8, 70, 102, 43, 98, 96, 103, 146, 63, 119, 198, 120, 115, 216, 210, 243, 179, 245, 81, 222, 248, 106, 156, 141, 73, 77, 201, 192, 109, 141, 14, 86, 171, 231, 39, 161, 99, 209, 158, 43, 152, 48, 156, 237, 41, 205, 123, 163, 1, 174, 99, 55, 38, 3, 225, 209, 142, 40, 7, 78, 23, 217, 182, 220, 2, 120, 247, 202, 172, 59, 27, 155, 28, 90, 163, 138, 76, 32, 28, 159, 12, 192, 23, 30, 110, 181, 148, 238, 63, 85, 64, 128, 166, 121, 149, 160, 23, 118, 96, 21, 122, 255, 226, 150, 40, 103, 178, 134, 132, 182, 123, 167, 50, 134, 95, 222, 18, 229, 108, 198, 112, 99, 212, 238, 29, 155, 156, 5, 240, 253, 53, 54, 84, 127, 25, 246, 9, 4, 214, 175, 112, 104, 139, 107, 46, 20, 132, 129, 41, 179, 196, 60, 96, 108, 228, 155, 61, 107, 60, 237, 41, 140, 82, 100, 138, 66, 186, 146, 151, 67, 89, 195, 119, 142, 231, 65, 36, 212, 251, 209, 188, 132, 212, 116, 85, 18, 236, 233, 143, 139, 0, 252, 174, 34, 62, 71, 39, 131, 80, 107, 138, 82, 11, 128, 182, 213, 176, 33, 169, 33, 128, 159, 174, 143, 176, 231, 104, 30, 20, 172, 170, 120, 187, 111, 181, 199, 171, 151, 124, 80, 48, 94, 17, 204, 111, 173, 246, 160, 44, 188, 182, 45, 73, 103, 131, 189, 110, 120, 218, 240, 192, 74, 151, 29, 77, 22, 80, 207, 80, 137, 6, 79, 227, 42, 136, 42, 112, 230, 244, 153, 16, 128, 18, 155, 193, 0, 127, 237, 74, 48, 81, 18, 50, 190, 128, 8, 55, 198, 236, 207, 186, 251, 243, 161, 10, 205, 112, 255, 189, 85, 46, 178, 103, 25, 61, 67, 37, 222, 24, 177, 168, 142, 237, 74, 209, 28, 213, 76, 248, 66, 206, 192, 67, 95, 242, 56, 240, 229, 8, 253, 21, 26, 126, 176, 54, 178, 112, 34, 18, 5, 63, 255, 180, 196, 211, 237, 17, 20, 240, 236, 39, 37, 11, 79, 89, 158, 247, 159, 242, 57, 50, 211, 164, 20, 60, 126, 178, 64, 68, 131, 163, 96, 239, 201, 2, 34, 112, 100, 220, 231, 135, 107, 35, 188, 114, 209, 103, 119, 179, 67, 163, 171, 24, 200, 24, 122, 134, 138, 124, 158, 23, 86, 197, 53, 23, 239, 74, 242, 112, 171, 199, 243, 131, 69, 112, 212, 188, 137, 40, 0, 121, 48, 109, 109, 244, 102, 174, 105, 8, 92, 151, 208, 244, 109, 79, 112, 177, 32, 220, 182, 76, 115, 123, 95, 142, 254, 137, 32, 188, 127, 172, 59, 133, 163, 160, 225, 245, 105, 112, 213, 188, 42, 112, 224, 197, 138, 108, 158, 216, 153, 248, 226, 61, 88, 224, 79, 91, 227, 180, 189, 157, 97, 115, 74, 115, 104, 44, 160, 127, 78, 153, 162, 160, 28, 64, 84, 171, 218, 101, 184, 247, 159, 5, 174, 248, 176, 37, 165, 121, 118, 83, 244, 11, 5, 161, 179, 209, 225, 76, 222, 240, 194, 230, 24, 142, 134, 61, 253, 121, 112, 170, 69, 172, 33, 162, 24, 47, 75, 157, 177, 92, 65, 87, 95, 22, 128, 31, 183, 69, 56, 176, 33, 90, 37, 205, 245, 214, 241, 241, 128, 67, 35, 1, 39, 38, 13, 94, 239, 52, 147, 229, 234, 255, 221, 211, 234, 17, 85, 208, 119, 37, 176, 237, 116, 177, 169, 120, 38, 148, 91, 151, 59, 124, 216, 149, 168, 12, 153, 1, 123, 79, 228, 25, 206, 203, 82, 47, 137, 186, 244, 100, 187, 211, 36, 52, 220, 255, 97, 158, 222, 138, 84, 235, 26, 131, 26, 199, 198, 3, 154, 14, 102, 152, 240, 133, 7, 90, 28, 62, 223, 157, 226, 165, 173, 113, 86, 120, 138, 168, 14, 29, 176, 169, 163, 150, 54, 254, 199, 219, 227, 36, 52, 156, 206, 25, 122, 47, 148, 107, 191, 11, 22, 72, 165, 130, 95, 108, 140, 241, 163, 54, 111, 230, 46, 138, 6, 2, 17, 130, 202, 212, 173, 21, 228, 12, 220, 249, 143, 28, 3, 19, 166, 170, 53, 183, 196, 20, 71, 182, 39, 105, 139, 219, 205, 230, 131, 25, 70, 75, 114, 245, 0, 102, 100, 122, 69, 76, 177, 171, 217, 229, 153, 142, 8, 183, 166, 106, 243, 112, 46, 47, 97, 146, 165, 92, 104, 175, 140, 106, 99, 62, 108, 122, 39, 195, 112, 65, 234, 191, 140, 150, 10, 37, 70, 64, 43, 54, 164, 53, 77, 17, 133, 8, 92, 42, 26, 118, 44, 119, 121, 170, 61, 66, 103, 186, 26, 220, 80, 78, 120, 238, 179, 18, 47, 12, 150, 170, 43, 226, 154, 0, 92, 197, 155, 0, 20, 237, 203, 172, 238, 127, 50, 101, 108, 239, 175, 147, 36, 238, 117, 125, 234, 86, 12, 125, 58, 51, 100, 106, 150, 124, 36, 254, 23, 153, 41, 93, 205, 81, 212, 105, 60, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)
$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
$objForm.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
$objform.forecolor = "blue"

$MyLinkLabel = New-Object System.Windows.Forms.LinkLabel
$MyLinkLabel.Location = New-Object System.Drawing.Size(450,3)
$MyLinkLabel.Size = New-Object System.Drawing.Size(130,15)
$MyLinkLabel.DisabledLinkColor = [System.Drawing.Color]::Red
$MyLinkLabel.VisitedLinkColor = [System.Drawing.Color]::Blue
$MyLinkLabel.LinkBehavior = [System.Windows.Forms.LinkBehavior]::HoverUnderline
$MyLinkLabel.LinkColor = [System.Drawing.Color]::Navy
$MyLinkLabel.TabStop = $False
$MyLinkLabel.Text = "  www.myskypelab.com"
$MyLinkLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomRight    #TopRight
$MyLinkLabel.add_click(
{
	 [system.Diagnostics.Process]::start("http://www.myskypelab.com")
})
$objForm.Controls.Add($MyLinkLabel)


#Draw Status text

$statustext = new-object System.Windows.Forms.Label
$statustext.location = new-object system.drawing.size(20,575)
$statustext.size = new-object system.drawing.size(550,20)
$statustext.forecolor = "red"
$statustext.text = ""
$objform.controls.add($statustext)

#Draw start button
$StartButton = New-Object System.Windows.Forms.Button
$StartButton.Location = New-Object System.Drawing.Size(10,320)
$StartButton.Size = New-Object System.Drawing.Size(90,23)
$StartButton.Text = "Start Tracing"
#When the user selects start tracing, define the variables, and run the StartTracing function
$StartButton.Add_Click({
#Buttons Default as disabled
$StartButton.enabled = $false
$StopButton.enabled = $false
$ExportButton.enabled = $false
$UpdateButton.enabled = $false
$alwaysOnButton.enabled = $false
$startTimeTextBox.Enabled = $false
$stopTimeTextBox.enabled = $false
$durationTextBox.enabled = $false

#Filter components
$levellistbox.enabled = $false
$filterPhonetext.enabled = $false
$filterURItext.enabled = $false
$filterCallIDtext.enabled = $false
$filterIPtext.enabled = $false
$matchTypeCheckBox.enabled = $false
$timeCheckBox.enabled = $false
$complistbox.enabled = $false
$sipContentstext.enabled = $false
$compCheckbox.enabled = $false
$durationTextBox.Enabled = $false
$durationCheckBox.Enabled = $false
$startTimeTextBox.Enabled = $false
$stopTimeTextBox.Enabled = $false

$script:scenariotype=$scenarioListBox.SelectedItem
$script:pooltotrace=$poollistbox.SelectedItem
$script:duration=$durationTextBox.text
StartTracing
})
$objForm.Controls.Add($StartButton)


#Draw start button
$UpdateButton = New-Object System.Windows.Forms.Button
$UpdateButton.Location = New-Object System.Drawing.Size(235,320)
$UpdateButton.Size = New-Object System.Drawing.Size(95,23)
$UpdateButton.Text = "Update Duration"
#When the user selects start tracing, define the variables, and run the StartTracing function
$UpdateButton.Add_Click({

#Buttons Default as disabled
$StartButton.enabled = $false
$StopButton.enabled = $false
$ExportButton.enabled = $false
$UpdateButton.enabled = $false
$alwaysOnButton.enabled = $false
$startTimeTextBox.Enabled = $false
$stopTimeTextBox.enabled = $false
$durationTextBox.enabled = $false

#Filter components
$levellistbox.enabled = $false
$filterPhonetext.enabled = $false
$filterURItext.enabled = $false
$filterCallIDtext.enabled = $false
$filterIPtext.enabled = $false
$matchTypeCheckBox.enabled = $false
$timeCheckBox.enabled = $false
$complistbox.enabled = $false
$sipContentstext.enabled = $false
$compCheckbox.enabled = $false
$durationTextBox.Enabled = $false
$durationCheckBox.Enabled = $false
$startTimeTextBox.Enabled = $false
$stopTimeTextBox.Enabled = $false

$script:duration=$durationTextBox.text
UpdateTracing
})
$objForm.Controls.Add($UpdateButton)


##AlwaysOn Label
$alwaysOnLabel = New-Object System.Windows.Forms.Label
$alwaysOnLabel.Location = New-Object System.Drawing.Size(110,385) 
$alwaysOnLabel.Size = New-Object System.Drawing.Size(150,15) 
$alwaysOnLabel.Text = ""
$alwaysOnLabel.forecolor = "green"
$objForm.Controls.Add($alwaysOnLabel) 

#draw AlwaysOn button
$alwaysOnButton = New-Object System.Windows.Forms.Button
$alwaysOnButton.Location = New-Object System.Drawing.Size(10,380)
$alwaysOnButton.Size = New-Object System.Drawing.Size(90,23)
$alwaysOnButton.Text = "Always On"
$alwaysOnButton.forecolor = "blue"
#When the user selects Stop Tracing, run the StopTracing function
$alwaysOnButton.Add_Click({
$alwaysOnButton.Enabled = $false
$StartButton.enabled = $false
$StopButton.enabled = $false
$ExportButton.enabled = $false
$UpdateButton.enabled = $false

#Disable all GUI components
$StartButton.enabled = $false
$StopButton.enabled = $false
$ExportButton.enabled = $false
$UpdateButton.enabled = $true
$ExportButton.enabled = $false
#Filter components
$levellistbox.enabled = $false
$filterPhonetext.enabled = $false
$filterURItext.enabled = $false
$filterCallIDtext.enabled = $false
$filterIPtext.enabled = $false
$matchTypeCheckBox.enabled = $false
$timeCheckBox.enabled = $false
$complistbox.enabled = $false
$sipContentstext.enabled = $false
$compCheckbox.enabled = $false
$durationTextBox.Enabled = $false
$durationCheckBox.Enabled = $false
$startTimeTextBox.Enabled = $false
$stopTimeTextBox.Enabled = $false
	
AlwaysOn-Change
})
$objForm.Controls.Add($alwaysOnButton)

 ##Duration Label
$durationLabel = New-Object System.Windows.Forms.Label
$durationLabel.Location = New-Object System.Drawing.Size(130,305) 
$durationLabel.Size = New-Object System.Drawing.Size(100,15) 
$durationLabel.Text = "Logging Duration"
$objForm.Controls.Add($durationLabel) 


#Duration Text box
$durationTextBox= new-object System.Windows.Forms.textbox
$durationTextBox.location = new-object system.drawing.size(130,320)
$durationTextBox.size= new-object system.drawing.size(100,23)
$durationTextBox.text = "0.04:00"   #0 days, 4 Hours, 00 minutes
$objform.controls.add($durationTextBox)

#for 2 hours (0 days.02 hours:00 minutes) and then stop:
#-Duration 0.02:00
#This following syntax would specify a duration of 3 hours and 15 minutes:
#-Duration 0.03:15
#The following syntax would specify a duration of 6 days, 5 hours and 12 minutes:
#-Duration 6.05:12


#Match type
$durationCheckBox = New-Object System.Windows.Forms.Checkbox 
$durationCheckBox.Location = New-Object System.Drawing.Size(110,320) 
$durationCheckBox.Size = New-Object System.Drawing.Size(20,20)
$durationCheckBox.Add_Click({
if($durationCheckBox.checked)
{
	$durationTextBox.enabled = $true
}
else
{
	$durationTextBox.enabled = $false
} 
})
$objForm.Controls.Add($durationCheckBox)


#draw stop button
$StopButton = New-Object System.Windows.Forms.Button
$StopButton.Location = New-Object System.Drawing.Size(10,350)
$StopButton.Size = New-Object System.Drawing.Size(90,23)
$StopButton.Text = "Stop Tracing"
#When the user selects Stop Tracing, run the StopTracing function
$StopButton.Add_Click({
#Buttons Default as disabled
$StartButton.enabled = $false
$StopButton.enabled = $false
$ExportButton.enabled = $false
$UpdateButton.enabled = $false
$alwaysOnButton.enabled = $false
$startTimeTextBox.Enabled = $false
$stopTimeTextBox.enabled = $false
$durationTextBox.enabled = $false

#Filter components
$levellistbox.enabled = $false
$filterPhonetext.enabled = $false
$filterURItext.enabled = $false
$filterCallIDtext.enabled = $false
$filterIPtext.enabled = $false
$matchTypeCheckBox.enabled = $false
$timeCheckBox.enabled = $false
$complistbox.enabled = $false
$sipContentstext.enabled = $false
$compCheckbox.enabled = $false
$durationTextBox.Enabled = $false
$durationCheckBox.Enabled = $false
$startTimeTextBox.Enabled = $false
$stopTimeTextBox.Enabled = $false

StopTracing
PopulateComponents
})
$objForm.Controls.Add($StopButton)


 ##Stop Label
$stopLabel = New-Object System.Windows.Forms.Label
$stopLabel.Location = New-Object System.Drawing.Size(110,355) 
$stopLabel.Size = New-Object System.Drawing.Size(250,15) 
$stopLabel.forecolor = "green"
$stopLabel.Text = ""
$objForm.Controls.Add($stopLabel) 


##Time Label
$matchTypeLabel = New-Object System.Windows.Forms.Label
$matchTypeLabel.Location = New-Object System.Drawing.Size(35,470) 
$matchTypeLabel.Size = New-Object System.Drawing.Size(100,15) 
$matchTypeLabel.Text = "MatchAll"
$objForm.Controls.Add($matchTypeLabel) 

#Match type
$matchTypeCheckBox = New-Object System.Windows.Forms.Checkbox 
$matchTypeCheckBox.Location = New-Object System.Drawing.Size(20,467) 
$matchTypeCheckBox.Size = New-Object System.Drawing.Size(20,20)
$matchTypeCheckBox.Add_Click({
if($matchTypeCheckBox.checked)
{
	$matchTypeLabel.forecolor = "red"
	$matchTypeLabel.Text = "MatchAny"
}
else
{
	$matchTypeLabel.forecolor = "blue"
	$matchTypeLabel.Text = "MatchAll"
} 
})
$objForm.Controls.Add($matchTypeCheckBox)


##Time Label
$timeLabel = New-Object System.Windows.Forms.Label
$timeLabel.Location = New-Object System.Drawing.Size(450,325) 
$timeLabel.Size = New-Object System.Drawing.Size(100,15) 
$timeLabel.Text = "Export Start / End Time:"
$objForm.Controls.Add($timeLabel) 

$timeCheckBox = New-Object System.Windows.Forms.Checkbox 
$timeCheckBox.Location = New-Object System.Drawing.Size(555,345) 
$timeCheckBox.Size = New-Object System.Drawing.Size(20,20)
$timeCheckBox.Add_Click({
if($timeCheckBox.checked)
{
	$startTimeTextBox.enabled = $true
	$stopTimeTextBox.enabled = $true
}
else
{
	$startTimeTextBox.enabled = $false
	$stopTimeTextBox.enabled = $false	
} 
})
$objForm.Controls.Add($timeCheckBox)



# documentation says: -StartTime "8/1/2012 8:00AM"
# The standard Get-Date works - "08/09/2013 16:05:56"

$a1 = Get-Date -format (Get-culture).DateTimeFormat.ShortDatePattern
$a2 = ((Get-Date).AddHours(-1)).ToString((Get-culture).DateTimeFormat.ShortTimePattern)
$a = "$a1 $a2"

$b1 = Get-Date -format (Get-culture).DateTimeFormat.ShortDatePattern
$b2 = Get-Date -format (Get-culture).DateTimeFormat.ShortTimePattern
$b = "$b1 $b2"

#Start Time Text box
$startTimeTextBox = new-object System.Windows.Forms.textbox
$startTimeTextBox.location = new-object system.drawing.size(445,345)
$startTimeTextBox.size= new-object system.drawing.size(100,23)
$startTimeTextBox.text = $a
$objform.controls.add($startTimeTextBox)

#Stop Time Text box
$stopTimeTextBox = new-object System.Windows.Forms.textbox
$stopTimeTextBox.location = new-object system.drawing.size(445,370)
$stopTimeTextBox.size= new-object system.drawing.size(100,23)
$stopTimeTextBox.text = $b
$objform.controls.add($stopTimeTextBox)


#draw export button
$ExportButton = New-Object System.Windows.Forms.Button
$ExportButton.Location = New-Object System.Drawing.Size(10,410)
$ExportButton.Size = New-Object System.Drawing.Size(90,23)
$ExportButton.Text = "Export Logs"
#When the user select export logs, define the variables and run the ExportLogs function
$ExportButton.Add_Click(
{
	ExportLogs
})
$objForm.Controls.Add($ExportButton)



#draw Analyse button
$AnalyseButton = New-Object System.Windows.Forms.Button
$AnalyseButton.Location = New-Object System.Drawing.Size(10,440)
$AnalyseButton.Size = New-Object System.Drawing.Size(90,23)
$AnalyseButton.Text = "Analyse Log.."
$AnalyseButton.Add_Click(
{
	if (Test-Path $snooperLocation)
	{
		$Filter = "All Files (*.*)|*.*"
		[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
		$fileForm = New-Object System.Windows.Forms.OpenFileDialog
		$fileForm.InitialDirectory = $pathbox.text
		$fileForm.Filter = $Filter
		$fileForm.Title = "Open Log File"
		$Show = $fileForm.ShowDialog()
		if ($Show -eq "OK")
		{
			[string] $filename = $fileForm.FileName
			#Launch Snooper
			$pinfo = New-Object System.Diagnostics.ProcessStartInfo
			$pinfo.FileName = $snooperLocation
			$pinfo.RedirectStandardError = $false
			$pinfo.RedirectStandardOutput = $false
			$pinfo.UseShellExecute = $true
			$pinfo.Arguments = "`"$filename`""
			$p = New-Object System.Diagnostics.Process
			$p.StartInfo = $pinfo
			$p.Start() | Out-Null
		}
		else
		{
			Write-Host "Operation cancelled by user."
		}
	}
	else
	{
		Write-Host "Snooper not found in location. Please install Lync 2013 debugging tools."
	}	
	
})
$objForm.Controls.Add($AnalyseButton)		
		
		
		
#draw path label
$pathlabel= new-object System.Windows.Forms.Label
$pathlabel.location = new-object system.drawing.size(220,415)
$pathlabel.size = new-object system.drawing.size(70,23)
$pathlabel.text = "Export Path:"
$objform.controls.add($pathlabel)

#file path box
$pathbox= new-object System.Windows.Forms.textbox
$pathbox.location = new-object system.drawing.size(290,410)
$pathbox.size= new-object system.drawing.size(250,23)
$pathbox.text = $pathDialog
$pathbox.enabled = $false
$objform.controls.add($pathbox)


#draw browse button
$BrowseButton = New-Object System.Windows.Forms.Button
$BrowseButton.Location = New-Object System.Drawing.Size(540,410)
$BrowseButton.Size = New-Object System.Drawing.Size(20,18)
$BrowseButton.Text = ".."
$BrowseButton.Add_Click(
{
	$objFolderForm = New-Object System.Windows.Forms.FolderBrowserDialog
	#$objFolderForm.SelectedPath = "$ScriptFolder"
	$objFolderForm.Description = "Select Folder"
	$Show = $objFolderForm.ShowDialog()
	if ($Show -eq "OK")
	{
		[string]$path = $objFolderForm.SelectedPath
		$pathbox.text = "$path\"
		$script:pathDialog = "$path\"
	}
	else
	{
		Write-Host "File Dialog cancelled by user." -foreground "Red"
		return
	}
})
$objForm.Controls.Add($BrowseButton)



#Phone Fitler
$filterPhone = new-object system.windows.forms.label
$filterPhone.location = new-object system.drawing.size(20,490)
$filterPhone.size = new-object system.drawing.size(110,15)
$filterPhone.text = "Phone Export Filter:"
$objform.controls.add($filterPhone)

$filterPhonetext = new-object System.Windows.Forms.TextBox
$filterPhonetext.location = new-object system.drawing.size(20,505)
$filterPhonetext.size = new-object system.drawing.size(120,23)
$filterPhonetext.text = "None"
$objform.controls.add($filterPhonetext)


#URI Filter
$filterURI = new-object system.windows.forms.label
$filterURI.location = new-object system.drawing.size(150,490)
$filterURI.size = new-object system.drawing.size(110,15)
$filterURI.text = "URI Export Filter:"
$objform.controls.add($filterURI)

$filterURItext = new-object System.Windows.Forms.TextBox
$filterURItext.location = new-object system.drawing.size(150,505)
$filterURItext.size = new-object system.drawing.size(120,23)
$filterURItext.text = "None"
$objform.controls.add($filterURItext)


#CallID Fitler
$filterCallID = new-object system.windows.forms.label
$filterCallID.location = new-object system.drawing.size(20,530)
$filterCallID.size = new-object system.drawing.size(110,15)
$filterCallID.text = "CallID Export Filter:"
$objform.controls.add($filterCallID)

$filterCallIDtext = new-object System.Windows.Forms.TextBox
$filterCallIDtext.location = new-object system.drawing.size(20,545)
$filterCallIDtext.size = new-object system.drawing.size(120,23)
$filterCallIDtext.text = "None"
$objform.controls.add($filterCallIDtext)


#IP Filter
$filterIP = new-object system.windows.forms.label
$filterIP.location = new-object system.drawing.size(150,530)
$filterIP.size = new-object system.drawing.size(110,15)
$filterIP.text = "IP Export Filter:"
$objform.controls.add($filterIP)

$filterIPtext = new-object System.Windows.Forms.TextBox
$filterIPtext.location = new-object system.drawing.size(150,545)
$filterIPtext.size = new-object system.drawing.size(120,23)
$filterIPtext.text = "None"
$objform.controls.add($filterIPtext)


#SIP Contents Filter
$sipContents = new-object system.windows.forms.label
$sipContents.location = new-object system.drawing.size(290,530)
$sipContents.size = new-object system.drawing.size(200,15)
$sipContents.text = "SIP Contents Export Filter:"
$objform.controls.add($sipContents)

$sipContentstext = new-object System.Windows.Forms.TextBox
$sipContentstext.location = new-object system.drawing.size(290,545)
$sipContentstext.size = new-object system.drawing.size(270,23)
$sipContentstext.text = "None"
$objform.controls.add($sipContentstext)


#draw logging level label
$levellabel= new-object System.Windows.Forms.Label
$levellabel.location = new-object system.drawing.size(360,315)
$levellabel.size = new-object system.drawing.size(75,23)
$levellabel.text = "Export logging level:"
$objform.controls.add($levellabel)

#DRAW logging level list
$levellistbox = new-object System.Windows.Forms.ListBox
$levellistbox.location = new-object system.drawing.size(360,345)
$levellistbox.size= new-object system.drawing.size(80,20)
$levellistbox.height = 60
[void]$levelListBox.Items.Add("Debug")
[void]$levelListBox.Items.Add("Fatal")
[void]$levelListBox.Items.Add("Error")
[void]$levelListBox.Items.Add("Warning")
[void]$levelListBox.Items.Add("Info")
[void]$levelListBox.Items.Add("Verbose")
[void]$levelListBox.Items.Add("All")
$levellistbox.setselected(0,$true)
$objForm.Controls.Add($levellistBox) 

##Draw Components Search Label
$complabel= new-object System.Windows.Forms.Label
$complabel.Location = new-object system.drawing.size(215,440)
$complabel.Size = new-object system.drawing.size(70,30)
$complabel.Text = "Components Export filter:"
$objform.controls.add($complabel)

##Draw components search list
$complistbox = new-object System.Windows.Forms.ListBox
$complistbox.Location = new-object system.drawing.size(290,440)
$complistbox.Size= new-object system.drawing.size(270,20)
$complistbox.Height = 90
$complistbox.Sorted = $true
$complistbox.SelectionMode = "MultiSimple"
$objForm.Controls.Add($complistbox)
$scenario = Invoke-Expression "Get-CSCLSScenario -Identity Global/AlwaysOn"
foreach ($sc in $scenario.provider) {[void] $complistbox.Items.Add($sc.name)}


##CompCheck Label
$compCheckboxLabel = New-Object System.Windows.Forms.Label
$compCheckboxLabel.Location = New-Object System.Drawing.Size(230,472) 
$compCheckboxLabel.Size = New-Object System.Drawing.Size(40,15) 
$compCheckboxLabel.Text = "List All"
$objForm.Controls.Add($compCheckboxLabel) 

#CompCheck type
$compCheckbox = New-Object System.Windows.Forms.Checkbox 
$compCheckbox.Location = New-Object System.Drawing.Size(270,470) 
$compCheckbox.Size = New-Object System.Drawing.Size(20,20)
$compCheckbox.Checked = $true
$compCheckbox.Add_Click({
	PopulateComponents
})
$objForm.Controls.Add($compCheckbox)


##Populate the list of CLS Scenarios
$scenarioLabel = New-Object System.Windows.Forms.Label
$scenarioLabel.Location = New-Object System.Drawing.Size(290,20) 
$scenarioLabel.Size = New-Object System.Drawing.Size(280,20) 
$scenarioLabel.Text = "Select Scenario to Log:"
$objForm.Controls.Add($scenarioLabel) 

$scenarioListBox = New-Object System.Windows.Forms.ListBox 
$scenarioListBox.Location = New-Object System.Drawing.Size(290,40) 
$scenarioListBox.Size = New-Object System.Drawing.Size(270,20) 
$scenarioListBox.Height = 275
$scenarioListBox.Sorted = $true
$scenarioListBox.Add_Click(
{
	$poolListItemSelected = $poollistbox.SelectedItem
	$scenarioListBoxItemSelected = $scenarioListBox.SelectedItem
	$script:scenariotype=$scenarioListBox.SelectedItem
})


$scenarioBoxCmds = Get-CsClsConfiguration | Select-Object -ExpandProperty Scenarios

foreach($scenarioBoxCmd in $scenarioBoxCmds)
{
	$nameOfScenario = $scenarioBoxCmd.Name
	#Don't include always on in list because it has it's own button
	if($nameOfScenario -ne "AlwaysOn")
	{[void]$scenarioListBox.Items.Add($nameOfScenario)}
}
$objForm.Controls.Add($scenarioListBox) 


##POOL SELECTION

##Populate the list of Servers
$poolLabel = New-Object System.Windows.Forms.Label
$poolLabel.Location = New-Object System.Drawing.Size(10,20) 
$poolLabel.Size = New-Object System.Drawing.Size(280,20) 
$poolLabel.Text = "Select Pool:"
$objForm.Controls.Add($poolLabel) 


$poolListBox = New-Object System.Windows.Forms.ListBox 
$poolListBox.Location = New-Object System.Drawing.Size(10,40) 
$poolListBox.Size = New-Object System.Drawing.Size(270,20) 
$poolListBox.Height = 275
$poolListBox.Add_Click(
{
	$script:pooltotrace=$poollistbox.SelectedItem
	Update-Buttons
})
$poolListBox.add_KeyUp(
{
	if ($_.KeyCode -eq "Up" -or $_.KeyCode -eq "Down") 
	{	
		$script:pooltotrace=$poollistbox.SelectedItem
		Update-Buttons
	}
})

##Add front ends
[void]$poolListBox.Items.Add("-----FRONT END SERVERS-----")
$serverNameArray = @()
Get-CSService -Registrar | where-object {$_.version -eq "6" -or $_.version -eq "7" -or $_.version -eq "8"} | select-object PoolFQDN | ForEach-Object {$serverNameArray += $_.PoolFQDN}
[array]::sort($serverNameArray)
foreach($serverName in $serverNameArray)
{
	[void] $poolListBox.Items.Add($serverName)
	$pools +=  $serverName
}

##Add mediation
[void]$poolListBox.Items.Add("-----MEDIATION SERVERS-----")
$serverNameArray = @()
#Check if the mediation server is co-located before adding it to the list. The pool doesn't need to be added to the list if it is.
Get-CSService -MediationServer  | where-object {$_.version -eq "6" -or $_.version -eq "7" -or $_.version -eq "8"} | ForEach-Object {$sub1 = ([string]$_.Identity).Substring(16);$sub2 = ([string]$_.Registrar).Substring(10);if($sub1 -ne $sub2){$serverNameArray += $_.PoolFQDN}}
[array]::sort($serverNameArray)
foreach($serverName in $serverNameArray)
{
	[void] $poolListBox.Items.Add($serverName)
	$pools +=  $serverName
}


##Add Chat
[void]$poolListBox.Items.Add("-----PERSISTENT CHAT SERVERS-----")
$serverNameArray = @()
#Check if the chat server is co-located before adding it to the list. The pool doesn't need to be added to the list if it is.
Get-CSService -PersistentChatServer | where-object {$_.version -eq "6" -or $_.version -eq "7" -or $_.version -eq "8"} | ForEach-Object {$sub1 = ([string]$_.Identity).Substring(21);$sub2 = ([string]$_.Registrar).Substring(10);if($sub1 -ne $sub2){$serverNameArray += $_.PoolFQDN}}
[array]::sort($serverNameArray)
foreach($serverName in $serverNameArray)
{
	[void] $poolListBox.Items.Add($serverName)
	$pools +=  $serverName
}


##Add edge
[void]$poolListBox.Items.Add("-----EDGE SERVERS-----")
$serverNameArray = @()
Get-CSService -EdgeServer | where-object {$_.version -eq "6" -or $_.version -eq "7" -or $_.version -eq "8"} | select-object PoolFQDN | ForEach-Object {$serverNameArray += $_.PoolFQDN}
[array]::sort($serverNameArray)
foreach($serverName in $serverNameArray)
{
	[void] $poolListBox.Items.Add($serverName)
	$pools +=  $serverName
}

$objForm.Controls.Add($poolListBox) 

Write-Host "--------Getting current status of pools, please wait---------" -Foreground "Green"
Write-Host
Get-AgentStatus


# Set defaults
$poolListBox.SelectedIndex = 0
$scenarioListBox.SelectedIndex = 0

#This setting makes the form be always on top. This can get in the way, so only turn it on if you want it.
#$objForm.Topmost = $True


#Buttons Default as disabled
$StartButton.enabled = $false
$StopButton.enabled = $false
$ExportButton.enabled = $false
$UpdateButton.enabled = $false
$alwaysOnButton.enabled = $false
$startTimeTextBox.Enabled = $false
$stopTimeTextBox.enabled = $false
$durationTextBox.enabled = $false

#Filter components
$levellistbox.enabled = $false
$filterPhonetext.enabled = $false
$filterURItext.enabled = $false
$filterCallIDtext.enabled = $false
$filterIPtext.enabled = $false
$matchTypeCheckBox.enabled = $false
$timeCheckBox.enabled = $false
$complistbox.enabled = $false
$sipContentstext.enabled = $false
$compCheckbox.enabled = $false
$durationCheckBox.Enabled = $false
$statustext.text = "INFO: Please select a pool..."

$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()

##END DRAWING UI
####################################################################################################
####################################################################################################



# SIG # Begin signature block
# MIIcWAYJKoZIhvcNAQcCoIIcSTCCHEUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUn5Vd4EkB4HLHlbSJu/yt0OwG
# MFmggheHMIIFEDCCA/igAwIBAgIQBsCriv7g+QV/64ncHMA83zANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTE4MDExMzAwMDAwMFoXDTE5MDQx
# ODEyMDAwMFowTTELMAkGA1UEBhMCQVUxEDAOBgNVBAcTB01pdGNoYW0xFTATBgNV
# BAoTDEphbWVzIEN1c3NlbjEVMBMGA1UEAxMMSmFtZXMgQ3Vzc2VuMIIBIjANBgkq
# hkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAukBaV5eP8/bHNonSdpgvTK/2iYj9XRl4
# VzpuJE1fK2sk0ZnjIidsaYXhFpL1LUbUlalxnO7cbWY5ok5bHg0vPx9p8IHHBH28
# xrisz7wcXTTjXMrOL+yynDJYUCMpKV5rMkBn5kJJlLUrY5kcT6Y0fa4HKmvLYAVC
# 6T83mvUvwVs0TlLqY5Dcm/eoVzSmv9Frn3A5WNKxElhhUL2W6LEHdikzCRltk0+e
# g6OXRSYHwulwL+HzcQ+83YEVp/YG9GM+v3Ra4UeuSWaOkt4FQI5JGMlKvhQ3wSu4
# 455xAyj56MTul2FQ1s+j2KI/bvJOMwzO86RDwUC+yZuhh8+IYVObpQIDAQABo4IB
# xTCCAcEwHwYDVR0jBBgwFoAUWsS5eyoKo6XqcQPAYPkt9mV1DlgwHQYDVR0OBBYE
# FHXxYdsGH8A4rhw89n7VPGve7xb5MA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAK
# BggrBgEFBQcDAzB3BgNVHR8EcDBuMDWgM6Axhi9odHRwOi8vY3JsMy5kaWdpY2Vy
# dC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNybDA1oDOgMYYvaHR0cDovL2NybDQu
# ZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwTAYDVR0gBEUwQzA3
# BglghkgBhv1sAwEwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQu
# Y29tL0NQUzAIBgZngQwBBAEwgYQGCCsGAQUFBwEBBHgwdjAkBggrBgEFBQcwAYYY
# aHR0cDovL29jc3AuZGlnaWNlcnQuY29tME4GCCsGAQUFBzAChkJodHRwOi8vY2Fj
# ZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRTSEEyQXNzdXJlZElEQ29kZVNpZ25p
# bmdDQS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAQEATz4Xu/3x
# ae3iTkPfYm7uEWpB16eV1Ig+8FMDg6CJ+465oidj2amAjD1n+MwekysJOmcWAiEg
# R7TcQKUpgy5QTTJGSsPm2rwwcBL0jye6hXgs5eD8szZEhdJOnl1txRsdhMtilV2I
# H7X1nQ6S/eRu4WneUUF3YIDreqFYGLIfAobafEEufP7pMk05zgO6lqBM97ee+roR
# eP12IG7CBokmhzoERIDdTjfNEbDtob3OKPKfao2K8MJ079CSoG+NnpieO4CSRQtu
# kaCfg4rK9iCFIksrHq+qSMMRobnVwZq5tDZrkQOjO+lBdL0XWF4nrBavCjs4DjBh
# JHz6nkyqXDNAuTCCBTAwggQYoAMCAQICEAQJGBtf1btmdVNDtW+VUAgwDQYJKoZI
# hvcNAQELBQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZ
# MBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNz
# dXJlZCBJRCBSb290IENBMB4XDTEzMTAyMjEyMDAwMFoXDTI4MTAyMjEyMDAwMFow
# cjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQ
# d3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVk
# IElEIENvZGUgU2lnbmluZyBDQTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoC
# ggEBAPjTsxx/DhGvZ3cH0wsxSRnP0PtFmbE620T1f+Wondsy13Hqdp0FLreP+pJD
# wKX5idQ3Gde2qvCchqXYJawOeSg6funRZ9PG+yknx9N7I5TkkSOWkHeC+aGEI2YS
# VDNQdLEoJrskacLCUvIUZ4qJRdQtoaPpiCwgla4cSocI3wz14k1gGL6qxLKucDFm
# M3E+rHCiq85/6XzLkqHlOzEcz+ryCuRXu0q16XTmK/5sy350OTYNkO/ktU6kqepq
# CquE86xnTrXE94zRICUj6whkPlKWwfIPEvTFjg/BougsUfdzvL2FsWKDc0GCB+Q4
# i2pzINAPZHM8np+mM6n9Gd8lk9ECAwEAAaOCAc0wggHJMBIGA1UdEwEB/wQIMAYB
# Af8CAQAwDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHkGCCsG
# AQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29t
# MEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8v
# Y3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqg
# OKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURS
# b290Q0EuY3JsME8GA1UdIARIMEYwOAYKYIZIAYb9bAACBDAqMCgGCCsGAQUFBwIB
# FhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMAoGCGCGSAGG/WwDMB0GA1Ud
# DgQWBBRaxLl7KgqjpepxA8Bg+S32ZXUOWDAfBgNVHSMEGDAWgBRF66Kv9JLLgjEt
# UYunpyGd823IDzANBgkqhkiG9w0BAQsFAAOCAQEAPuwNWiSz8yLRFcgsfCUpdqgd
# XRwtOhrE7zBh134LYP3DPQ/Er4v97yrfIFU3sOH20ZJ1D1G0bqWOWuJeJIFOEKTu
# P3GOYw4TS63XX0R58zYUBor3nEZOXP+QsRsHDpEV+7qvtVHCjSSuJMbHJyqhKSga
# OnEoAjwukaPAJRHinBRHoXpoaK+bp1wgXNlxsQyPu6j4xRJon89Ay0BEpRPw5mQM
# JQhCMrI2iiQC/i9yfhzXSUWW6Fkd6fp0ZGuy62ZD2rOwjNXpDd32ASDOmTFjPQga
# GLOBm0/GkxAG/AeB+ova+YJJ92JuoVP6EpQYhS6SkepobEQysmah5xikmmRR7zCC
# BmowggVSoAMCAQICEAMBmgI6/1ixa9bV6uYX8GYwDQYJKoZIhvcNAQEFBQAwYjEL
# MAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3
# LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJRCBDQS0x
# MB4XDTE0MTAyMjAwMDAwMFoXDTI0MTAyMjAwMDAwMFowRzELMAkGA1UEBhMCVVMx
# ETAPBgNVBAoTCERpZ2lDZXJ0MSUwIwYDVQQDExxEaWdpQ2VydCBUaW1lc3RhbXAg
# UmVzcG9uZGVyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAo2Rd/Hyz
# 4II14OD2xirmSXU7zG7gU6mfH2RZ5nxrf2uMnVX4kuOe1VpjWwJJUNmDzm9m7t3L
# helfpfnUh3SIRDsZyeX1kZ/GFDmsJOqoSyyRicxeKPRktlC39RKzc5YKZ6O+YZ+u
# 8/0SeHUOplsU/UUjjoZEVX0YhgWMVYd5SEb3yg6Np95OX+Koti1ZAmGIYXIYaLm4
# fO7m5zQvMXeBMB+7NgGN7yfj95rwTDFkjePr+hmHqH7P7IwMNlt6wXq4eMfJBi5G
# EMiN6ARg27xzdPpO2P6qQPGyznBGg+naQKFZOtkVCVeZVjCT88lhzNAIzGvsYkKR
# rALA76TwiRGPdwIDAQABo4IDNTCCAzEwDgYDVR0PAQH/BAQDAgeAMAwGA1UdEwEB
# /wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwggG/BgNVHSAEggG2MIIBsjCC
# AaEGCWCGSAGG/WwHATCCAZIwKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2lj
# ZXJ0LmNvbS9DUFMwggFkBggrBgEFBQcCAjCCAVYeggFSAEEAbgB5ACAAdQBzAGUA
# IABvAGYAIAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBjAGEAdABlACAAYwBvAG4A
# cwB0AGkAdAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBuAGMAZQAgAG8AZgAgAHQA
# aABlACAARABpAGcAaQBDAGUAcgB0ACAAQwBQAC8AQwBQAFMAIABhAG4AZAAgAHQA
# aABlACAAUgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAgAEEAZwByAGUAZQBtAGUA
# bgB0ACAAdwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABpAGEAYgBpAGwAaQB0AHkA
# IABhAG4AZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwByAGEAdABlAGQAIABoAGUA
# cgBlAGkAbgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBlAC4wCwYJYIZIAYb9bAMV
# MB8GA1UdIwQYMBaAFBUAEisTmLKZB+0e36K+Vw0rZwLNMB0GA1UdDgQWBBRhWk0k
# tkkynUoqeRqDS/QeicHKfTB9BgNVHR8EdjB0MDigNqA0hjJodHRwOi8vY3JsMy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNybDA4oDagNIYyaHR0
# cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0EtMS5jcmww
# dwYIKwYBBQUHAQEEazBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2Vy
# dC5jb20wQQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydEFzc3VyZWRJRENBLTEuY3J0MA0GCSqGSIb3DQEBBQUAA4IBAQCdJX4b
# M02yJoFcm4bOIyAPgIfliP//sdRqLDHtOhcZcRfNqRu8WhY5AJ3jbITkWkD73gYB
# jDf6m7GdJH7+IKRXrVu3mrBgJuppVyFdNC8fcbCDlBkFazWQEKB7l8f2P+fiEUGm
# vWLZ8Cc9OB0obzpSCfDscGLTYkuw4HOmksDTjjHYL+NtFxMG7uQDthSr849Dp3Gd
# Id0UyhVdkkHa+Q+B0Zl0DSbEDn8btfWg8cZ3BigV6diT5VUW8LsKqxzbXEgnZsij
# iwoc5ZXarsQuWaBh3drzbaJh6YoLbewSGL33VVRAA5Ira8JRwgpIr7DUbuD0FAo6
# G+OPPcqvao173NhEMIIGzTCCBbWgAwIBAgIQBv35A5YDreoACus/J7u6GzANBgkq
# hkiG9w0BAQUFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5j
# MRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBB
# c3N1cmVkIElEIFJvb3QgQ0EwHhcNMDYxMTEwMDAwMDAwWhcNMjExMTEwMDAwMDAw
# WjBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQL
# ExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElE
# IENBLTEwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDogi2Z+crCQpWl
# gHNAcNKeVlRcqcTSQQaPyTP8TUWRXIGf7Syc+BZZ3561JBXCmLm0d0ncicQK2q/L
# XmvtrbBxMevPOkAMRk2T7It6NggDqww0/hhJgv7HxzFIgHweog+SDlDJxofrNj/Y
# MMP/pvf7os1vcyP+rFYFkPAyIRaJxnCI+QWXfaPHQ90C6Ds97bFBo+0/vtuVSMTu
# HrPyvAwrmdDGXRJCgeGDboJzPyZLFJCuWWYKxI2+0s4Grq2Eb0iEm09AufFM8q+Y
# +/bOQF1c9qjxL6/siSLyaxhlscFzrdfx2M8eCnRcQrhofrfVdwonVnwPYqQ/MhRg
# lf0HBKIJAgMBAAGjggN6MIIDdjAOBgNVHQ8BAf8EBAMCAYYwOwYDVR0lBDQwMgYI
# KwYBBQUHAwEGCCsGAQUFBwMCBggrBgEFBQcDAwYIKwYBBQUHAwQGCCsGAQUFBwMI
# MIIB0gYDVR0gBIIByTCCAcUwggG0BgpghkgBhv1sAAEEMIIBpDA6BggrBgEFBQcC
# ARYuaHR0cDovL3d3dy5kaWdpY2VydC5jb20vc3NsLWNwcy1yZXBvc2l0b3J5Lmh0
# bTCCAWQGCCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQA
# aABpAHMAIABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUA
# dABlAHMAIABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkA
# ZwBpAEMAZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUA
# bAB5AGkAbgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgA
# aQBjAGgAIABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAA
# YQByAGUAIABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAA
# YgB5ACAAcgBlAGYAZQByAGUAbgBjAGUALjALBglghkgBhv1sAxUwEgYDVR0TAQH/
# BAgwBgEB/wIBADB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9v
# Y3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHow
# eDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJl
# ZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDAdBgNVHQ4EFgQUFQASKxOYspkH7R7f
# or5XDStnAs0wHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZI
# hvcNAQEFBQADggEBAEZQPsm3KCSnOB22WymvUs9S6TFHq1Zce9UNC0Gz7+x1H3Q4
# 8rJcYaKclcNQ5IK5I9G6OoZyrTh4rHVdFxc0ckeFlFbR67s2hHfMJKXzBBlVqefj
# 56tizfuLLZDCwNK1lL1eT7EF0g49GqkUW6aGMWKoqDPkmzmnxPXOHXh2lCVz5Cqr
# z5x2S+1fwksW5EtwTACJHvzFebxMElf+X+EevAJdqP77BzhPDcZdkbkPZ0XN1oPt
# 55INjbFpjE/7WeAjD9KqrgB87pxCDs+R1ye3Fu4Pw718CqDuLAhVhSK46xgaTfwq
# Ia1JMYNHlXdx3LEbS0scEJx3FMGdTy9alQgpECYxggQ7MIIENwIBATCBhjByMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQg
# Q29kZSBTaWduaW5nIENBAhAGwKuK/uD5BX/ridwcwDzfMAkGBSsOAwIaBQCgeDAY
# BgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3
# AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEW
# BBQCzxYKf0m8k8AiOxAiWbSzIWktljANBgkqhkiG9w0BAQEFAASCAQAdo1TecvN/
# I0xMLsjBSyYIV9m1gIQd17wmEs6JZRcVExyMitrOZcrLq14jLOPF0h1MODako+1L
# AAvFwTs5NkxrP5SeegGtZYPelkOHcPepUNgllKf7TUL6QWLi5vx3QIA1MtHFeG9d
# Og47EXx1jD4RH7+BbEAc9RCAba9vDsvw9dCWOxgrB1ahORuLuzPxWPQZqxBQm0Kj
# nQgSaZhVgyQfc5EnGwpLgBKXGmkyyCawQUrYBT4cGDOKh/khCyubS/IdnGIyh/cz
# qbXl2DCiKQLRIa5K4o2+/ix7e5urwfVCqe8EM7ZsNFddmbKa3Q2pVtPRxC9CNQhT
# Hnat41kEwTfloYICDzCCAgsGCSqGSIb3DQEJBjGCAfwwggH4AgEBMHYwYjELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJRCBDQS0xAhAD
# AZoCOv9YsWvW1ermF/BmMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTAyMDIwMTMxNTlaMCMGCSqGSIb3DQEJ
# BDEWBBTVGnKPnTQQgdfvEE80Tvhh9hUe8zANBgkqhkiG9w0BAQEFAASCAQB8vcRi
# Pb05VPuFbJmnZsitwbr9L5Flku7HZ78o8cU3y7Rb9KMHhjc4ABhcxfb5nW6SeZIY
# dm783hygFFoVuGyVjKUTzblwZX2/v9Q5msYQu73A7qvcDH8R1hi7HD2b1K98vCmX
# zrO9o8Eg/v/0BmIzBQRdrZu852z0E9H161INZr+lEtfmX+L0DEWBYLOt8XNdN415
# 7QnvLChQmb2IsSbxq1egOKd5cNiJCB6j/tTONYg2a+uE6a18GDH3Y6VwSaA3R1f6
# VuoyrPtCbUrmW5dhGhf1NZevpZCWSWT8BWsKmr/8AboAO+6nzKay05LO2Yq1mq0s
# pKqwZBZKfrVhRaRZ
# SIG # End signature block

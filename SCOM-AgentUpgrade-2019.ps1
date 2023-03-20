#=================================================================================
# By:   Luke Tatlock
# Date: 20/03/23
# Ver:  1.0
# Desc: Checks SCOM agent version and upgrades it to 2019. PART OF SCOM MP 
#=================================================================================

# Constants section - modify stuff here:

# Assign script name variable for use in event logging
$ScriptName = "SCOM.AgentUpdate.2019.Task.WA.ps1"
$EventID = "5053"


# Gather the start time of the script
$StartTime = Get-Date

#Set variable to be used in logging events
$whoami = whoami

# Load MOMScript API
$momapi = New-Object -comObject MOM.ScriptAPI

#Log script event that we are starting task
$Message = "`nGateway Upgrade Script is starting. `nRunning as ($whoami)."
$momapi.LogScriptEvent($ScriptName,$EventID,0,$Message)
Write-Host $Message




#Validate this is SCOM2019 GW installed
$SCOMRegKey = "HKLM:\SOFTWARE\Microsoft\Microsoft Operations Manager\3.0\Setup"
$SCOMPath = (Get-ItemProperty $SCOMRegKey).InstallDirectory
$SCOMPath = $SCOMPath.TrimEnd("\")
$SCOMCorePath = $SCOMPath.TrimEnd("Agent")
$SCOMCorePath = $SCOMCorePath.TrimEnd("\")
			
# Check to see if this is a agent
IF ($SCOMCorePath -match "Agent")
{
  #Gateway Detected
  $Message = "Agent detected."
  $momapi.LogScriptEvent($ScriptName,$EventID,0,$Message)
  Write-Host $Message

  $ServerURFile = Get-Item $SCOMPath\HealthService.dll
  $ServerURFileVersion = $ServerURFile.VersionInfo.FileVersion

  #For SCOM 2019 version should be 10.19.x
  $ServerURFileVersionSplit = $ServerURFileVersion.Split(".")
  [string]$AgentMajorVersion = $ServerURFileVersionSplit[0] + "." + $ServerURFileVersionSplit[1]
  IF ($AgentMajorVersion -ne "10.19")
  {
    #Verified SCOM Agent below 2019 installed - proceed with update
    $Message = "Agent is $AgentMajorVersion. `nStarting Update now. `nThe Healthservice will be restarted."
    $momapi.LogScriptEvent($ScriptName,$EventID,0,$Message)
    Write-Host $Message  

    $UpdateExists = Test-Path C:\SCOM2019Aent\MOMAgent.msi
    IF ($UpdateExists)
    {
      #Do update
	  $Command = 'Start-Sleep -s 5;Start-Process msiexec.exe -Wait -ArgumentList """/i C:\SCOM2019Aent\MOMAgent.msi NOAPM=1 AcceptEndUserLicenceAgreeent=1"""'
      $Process = ([wmiclass]"root\cimv2:Win32_ProcessStartup").CreateInstance()
      $Process.ShowWindow = 0
      $Process.CreateFlags = 16777216
      ([wmiclass]"root\cimv2:Win32_Process").Create("powershell.exe $Command")|Out-Null
    }
    ELSE
    {
      #Update file not found
      $Message = "FATAL ERROR: Agent File is not found.  `nStarting Update now. `nMissing: C:\SCOM2019Aent\MOMAgent.msi"
      $momapi.LogScriptEvent($ScriptName,$EventID,0,$Message)
      Write-Host $Message 
    }
  }
  ELSE
  {
    #Wrong Version
    $Message = "FATAL ERROR: Agent version is already 2019. `nDetected version: ($AgentMajorVersion)."
    $momapi.LogScriptEvent($ScriptName,$EventID,0,$Message)
    Write-Host $Message  
  }
}
ELSE
{
  #Agent not found
  $Message = "FATAL ERROR: Agent is NOT FOUND."
  $momapi.LogScriptEvent($ScriptName,$EventID,0,$Message)
  Write-Host $Message
}

#Log an event for script ending and total execution time.
$EndTime = Get-Date
$ScriptTime = ($EndTime - $StartTime).TotalSeconds
$Message = "Script Completed. `nRuntime: ($ScriptTime) seconds."
$momapi.LogScriptEvent($ScriptName,$EventID,0,$Message)
Write-Host $Message

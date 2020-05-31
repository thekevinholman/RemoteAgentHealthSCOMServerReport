#=================================================================================
#  Remote Agent Health script
#
#  Author: Kevin Holman
#  v1.1
#=================================================================================


# Constants section - modify stuff here:
#=================================================================================
# Assign script name variable for use in event logging.  
# ScriptName should be the same as the ID of the module that the script is contained in
$ScriptName = "RemoteAgentHealthSCOMServerReport.ps1"
$EventID = "2001"
#=================================================================================


# Starting Script section - All scripts get this
#=================================================================================
# Gather the start time of the script
$StartTime = Get-Date
#Set variable to be used in logging events
$whoami = whoami
# Load MOMScript API
$momapi = New-Object -comObject MOM.ScriptAPI
#Log script event that we are starting task
$momapi.LogScriptEvent($ScriptName,$EventID,0,"`nScript is starting. `nRunning as ($whoami).")
#=================================================================================


# Connect to local SCOM Management Group Section - If required
#=================================================================================
# I have found this to be the most reliable method to load SCOM modules for scripts running on Management Servers
# Clear any previous errors
$Error.Clear()
# Import the OperationsManager module and connect to the management group
$SCOMPowerShellKey = "HKLM:\SOFTWARE\Microsoft\System Center Operations Manager\12\Setup\Powershell\V2"
$SCOMModulePath = Join-Path (Get-ItemProperty $SCOMPowerShellKey).InstallDirectory "OperationsManager"
Import-module $SCOMModulePath
New-DefaultManagementGroupConnection "localhost"
IF ($Error) 
{ 
  $momapi.LogScriptEvent($ScriptName,$EventID,1,"`n FATAL ERROR: Unable to load OperationsManager module or unable to connect to Management Server. `n Terminating script. `n Error is: ($Error).")
  EXIT
}
#=================================================================================


# Begin MAIN script section
#=================================================================================
Clear-Host
[array]$AgentArr = @()

#Get the HSW class
$HSWClass = Get-SCOMClass -Name "Microsoft.SystemCenter.HealthServiceWatcher"

#Get the HSW Instances in a critical state
$HSWInstances = $HSWClass | Get-SCOMClassInstance | Where {$_.HealthState -eq "Error"}
$HSWInstancesCount = $HSWInstances.Count
$momapi.LogScriptEvent($ScriptName,$EventID,0,"`nFinished getting all Health Service Watcher Objects that are Critical in SCOM. `nRetrieved ($HSWInstancesCount) objects.")


#######################################
#######################################
#######################################
#######################################
# Limit for testing
# $HSWInstances = $HSWInstances | Select-Object -First 10
#######################################
#######################################
#######################################
#######################################


#Loop through each HSW instance and compare monitor modified time to threshold
FOREACH ($HSWInstance in $HSWInstances)
{
  #Get the agent Name
  $AgentName = $HSWInstance.DisplayName

  #Build a PSObject to hold properties for this agent
  $AgentItem = New-Object PSObject
  $AgentItem | Add-Member -type NoteProperty -Name 'AgentName' -Value $AgentName
  $AgentItem | Add-Member -type NoteProperty -Name 'InMaintenance' -Value 'NA'
  $AgentItem | Add-Member -type NoteProperty -Name 'HeartBeatState' -Value 'NA'
  $AgentItem | Add-Member -type NoteProperty -Name 'HeartBeatLastModified' -Value 'NA'
  $AgentItem | Add-Member -type NoteProperty -Name 'Primary' -Value 'NA'
  $AgentItem | Add-Member -type NoteProperty -Name 'Failover' -Value 'NA'
  $AgentItem | Add-Member -type NoteProperty -Name 'IPfromDNS' -Value 'NA'
  $AgentItem | Add-Member -type NoteProperty -Name 'Ping' -Value 'NA'
  $AgentItem | Add-Member -type NoteProperty -Name 'GetService' -Value 'NA'
  $AgentItem | Add-Member -type NoteProperty -Name 'HealthServiceExists' -Value 'NA'
  $AgentItem | Add-Member -type NoteProperty -Name 'HealthStatus' -Value 'NA'
  $AgentItem | Add-Member -type NoteProperty -Name 'HealthStartupType' -Value 'NA'
  $AgentItem | Add-Member -type NoteProperty -Name 'HSStartTypeFixed' -Value 'NA'
  $AgentItem | Add-Member -type NoteProperty -Name 'HSRecovery' -Value 'NA'
  
  #Get MM.  Do not continue remediation if object is in MM.
  $HSWInstanceMM = $HSWInstance.InMaintenanceMode
  $AgentItem.InMaintenance = $HSWInstanceMM

  IF ($HSWInstanceMM -eq $false)
  {
    #Get the monitor for HB failure for this instance
    $HBFMonitor = Get-SCOMMonitor -Instance $HSWInstance -Recurse | where {$_.DisplayName -eq "Health Service Heartbeat Failure"} 

    #Create and set the monitor collection to empty and create the collection to contain monitors
    $MonitorColl = @()
    $MonitorColl = New-Object "System.Collections.Generic.List[Microsoft.EnterpriseManagement.Configuration.ManagementPackMonitor]"

    #Add this monitor to a collection
    $MonitorColl.Add($HBFMonitor)

    #Get the MonitorData associated with this specific monitor
    $HBFMonitorData = $HSWInstance.getmonitoringstates($MonitorColl)

    #Get the monitor state
    [string]$HBFState = $HBFMonitorData.HealthState
    $HBMonitorLastUpdated = $HBFMonitorData.LastTimeModified

    $AgentItem.HeartBeatState = $HBFState
    $AgentItem.HeartBeatLastModified = $HBMonitorLastUpdated
    Write-Host "Heartbeat Failure Monitor Current State: ($HBFState)"
    Write-Host "Heartbeat Failure Monitor State Last Updated: ($HBMonitorLastUpdated)"

    IF ($HBFState -eq "Error")
    {
      #The heartbeat monitor is in error status.  Continue
      #Get the last modified time stamp for this monitor
      [datetime]$HBFLMT = $HBFMonitorData.LastTimeModified.ToLocalTime()

      # Get Parents
      $Agent = Get-SCOMAgent -DNSHostName $AgentName
      $Primary = $Agent.GetPrimaryManagementServer().DisplayName
      $Failover = $Agent.GetFailoverManagementServers().DisplayName

      $AgentItem.Primary = $Primary
      $AgentItem.Failover = $Failover    
      Write-Host "Primary Parent Server for ($AgentName) is ($Primary)"
      Write-Host "Failover List for ($AgentName) is ($Failover)"

#See if we can find the Computer Account in the domain
# This requires RSAT.  Consider adding it later

#See if we can query SCCM table
#Get ifexists and last inventory

      #See if we can resolve agent DNS name to an IP address
      # Clear any previous errors
      $Error.Clear()
      $ip = ""
      TRY
      {  
        $ip = ([System.Net.Dns]::GetHostAddresses($AgentName)).IPAddressToString
      }
      CATCH
      {
        Write-Host "ERROR: DNS did not resolve IP address for ($AgentName)."
        $AgentItem.IPfromDNS = "DNS Lookup Failed"
      }
  
      IF ($ip)
      {
        $AgentItem.IPfromDNS = $ip
        Write-Host "SUCCESS: DNS resolved IP ($ip) for ($AgentName)."
        $Ping = Test-Connection -ComputerName $AgentName -Count 1 -Quiet
        $AgentItem.Ping = $Ping

        IF ($Ping)
        {
          Write-Host "SUCCESS: Ping response on IP ($ip) for ($AgentName)."
          #Test SCM
          $SvcTest = @()
          TRY
          {    
            $SvcTest = Get-Service -ComputerName $AgentName -ErrorAction SilentlyContinue
          }
          CATCH
          {
            Write-Host "ERROR: Failed to Get-Service"
            $AgentItem.GetService = "Failed"
          }

          IF ($SvcTest)
          {
            $AgentItem.GetService = "Success"           
            Write-Host "SUCCESS: Connected to Service Control Manager. Attempting to get Healthservice."
            $Svc = Get-Service -ComputerName $AgentName -Name HealthService -ErrorAction SilentlyContinue
            IF ($Svc)
            {
              $AgentItem.HealthServiceExists = "True"
              Write-Host "SUCCESS: HealthService exists on ($AgentName)."
              $SvcStatus = $Svc.Status
              $SvcStartType = $Svc.StartType
              Write-Host "HealthService Status: ($SvcStatus)."
              Write-Host "HealthService Startup Type: ($SvcStartType)."
              $AgentItem.HealthStatus = $SvcStatus
              $AgentItem.HealthStartupType = $SvcStartType

              #Check to make sure service is set to Automatic
              IF ($SvcStartType -ne "Automatic")
              {
                $Error.Clear()          
                Write-Host "HealthService has an incorrect startup type of: ($SvcStartType). We will attempt to set this to Automatic startup."
                # Set service to automatic
                Set-Service -ComputerName $AgentName -Name HealthService -StartupType Automatic -ErrorAction SilentlyContinue

                IF ($Error)
                {
                  Write-Host "ERROR: Unable to set Healthservice to automatic. Error is: ($Error)."
                  $AgentItem.HSStartTypeFixed = "Error"
                }
                ELSE
                {
                  #Verify service is now automatic
                  $Svc = Get-Service -ComputerName $AgentName -Name HealthService -ErrorAction SilentlyContinue
                  $SvcStartType = $Svc.StartType
                  IF ($SvcStartType -ne "Automatic")
                  {
                    Write-Host "ERROR: Unable to set Healthservice to automatic. Startuptype was detected as ($SvcStartType). Error is: ($Error)."
                    $AgentItem.HSStartTypeFixed = "Error"
                  }
                  ELSE
                  {
                    Write-Host "SUCCESS: Service StartupType was changed to Automatic startup." -ForegroundColor Green
                    $AgentItem.HSStartTypeFixed = "True"
                  }
                }
              }

              IF ($SvcStatus -ne "Running")
              {
                $Error.Clear()
                #Attempt to Start Service
                Write-Host "1st attempt to start service: (HealthService)."           
                Set-Service -ComputerName $AgentName -Name HealthService -Status Running -ErrorAction SilentlyContinue
                IF ($Error)
                {
                  Write-Host "ERROR: Unable to START HealthService. `nError is: ($Error)."
                }
                #Verify service is now running
                Start-Sleep 20
                $Svc = Get-Service -ComputerName $AgentName -Name HealthService -ErrorAction SilentlyContinue
                $SvcStatus = $Svc.Status

                IF ($SvcStatus -ne "Running")
                {
                  #Attempt to Start Service
                  Write-Host "2nd attempt to start service: (HealthService)."
                  Set-Service -ComputerName $AgentName -Name HealthService -Status Running -ErrorAction SilentlyContinue
                  #Verify service is now running
                  Start-Sleep 20
                  $Svc = Get-Service -ComputerName $AgentName -Name HealthService -ErrorAction SilentlyContinue
                  $SvcStatus = $Svc.Status
                }

                IF ($SvcStatus -ne "Running")
                {              
                  $Error.Clear()
                  #Attempt to Start Service
                  Write-Host "3rd attempt to start service: (HealthService)."
                  Set-Service -ComputerName $AgentName -Name HealthService -Status Running -ErrorAction SilentlyContinue
                  #Verify service is now running
                  Start-Sleep 20
                  $Svc = Get-Service -ComputerName $AgentName -Name HealthService -ErrorAction SilentlyContinue
                  $SvcStatus = $Svc.Status
                }  
         
                IF ($SvcStatus -ne "Running")
                { 
                  Write-Host "ERROR: 3 failed attempts to start Healthservice. `nError is: ($Error)."
                  $AgentItem.HSRecovery = "Failed"
                }
                ELSE
                {
                  Write-Host "SUCCESS: Started the HealthService."
                  $AgentItem.HSRecovery = "Started"  
                }
              }
              ELSE
              {
                #Service is already running.
                Write-Host "Service state for HealthService was initially found to be: ($SvcStatus)."        
              }
            }
            ELSE
            {
              #Throw Healthservice does not exist!
              Write-Host "ERROR: Connected to Service Control Manager but cannot find an existing HealthService on ($AgentName). `nError is: ($Error)."
              $AgentItem.HealthServiceExists = "Missing"
            }
          }
          ELSE
          {
            #Throw error on unable to get services at all
            Write-Host "ERROR: Unable to connect to Service Control Manager on ($AgentName). `nError is: ($Error)."
          }
        }
        ELSE
        {
          Write-Host "ERROR: No ping response on IP ($ip) for ($AgentName)."  
        }
      }
    }
  } #End is MM false
  $AgentArr += $AgentItem
} #End FOREACH loop

#=================================================================================
# End MAIN script section


#Output Section:
#=================================================================================
$AgentArr | Export-Csv -Path C:\Windows\Temp\SCOMAgentDownResults.csv -NoTypeInformation
$AgentArr | Out-GridView
#=================================================================================


# End of script section
#=================================================================================
#Log an event for script ending and total execution time.
$EndTime = Get-Date
$ScriptTime = ($EndTime - $StartTime).TotalSeconds
$momapi.LogScriptEvent($ScriptName,$EventID,0,"`n Script Completed. `n Script Runtime: ($ScriptTime) seconds.")
#=================================================================================
# End of script
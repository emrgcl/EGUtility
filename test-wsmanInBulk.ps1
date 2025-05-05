[CmdletBinding(SupportsShouldProcess=$true)] # Added SupportsShouldProcess for best practice
param (
    [Parameter(Mandatory=$true)]
    [string]$CollectionName,

    [Parameter(Mandatory=$false)]
    [string]$FilePath = "C:\Temp\WSMANResult_$($CollectionName -replace '[^a-zA-Z0-9]','_')_$(Get-Date -Format 'yyyyMMddHHmmss').csv",

    [Parameter(Mandatory=$false)]
    [string]$SiteCode, # Optional: Specify Site Code if needed

    # Allow configuring the maximum number of concurrent threads
    [Parameter(Mandatory=$false)]
    [int]$MaxThreads = [Math]::Min([Environment]::ProcessorCount, 64), # Default to CPU count, capped at 64 for sanity

    #Log File Path
    [Parameter(Mandatory=$true)]
    [string]$LogFilePath
)

Function Write-Log {
     [CmdletBinding()]

     Param (
         [Parameter(Mandatory=$true)]
         [string]$Message,
         [string]$LogFilePath,
         [String]$LogLevel = 'Verbose'
     )  
     if ([string]::IsNullOrEmpty($LogfilePath)) {
        $LogfilePath = $script:LogFilePAth
        }   
     $MessageString = "[$(Get-Date -format G)] [$LogLevel] $Message"
     Add-Content -Value $MessageString -Path $LogFilePath
     Write-Verbose $MessageString
 }

$ExecutionPath = (Get-Location).Path

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
Write-Log "Script started. Testing collection '$CollectionName'." 
Write-Log "Max concurrent threads set to: $MaxThreads." 
Write-Log "Output file path: $FilePath." 

# --- Initial Setup & Validation ---
$ErrorActionPreference = 'Stop' # Exit on critical errors early

try {
    # Construct path to SCCM module
    $AdminUIPath = $Env:SMS_ADMIN_UI_PATH
    if (-not $AdminUIPath) {
        throw "SCCM Console environment variable (SMS_ADMIN_UI_PATH) not found. Please run this script from an SCCM Console machine."
    }
    $ModulePath = Join-Path (Split-Path $AdminUIPath -Parent) "ConfigurationManager.psd1"

    if (-not (Test-Path $ModulePath)) {
        throw "Configuration Manager module not found at expected path: $ModulePath"
    }

    Write-Log "Loading SCCM Module from $ModulePath..." 
    Import-Module $ModulePath -Verbose:$false

    # Determine and set current location to SCCM site drive
    if (-not $SiteCode) {
        # Attempt to auto-detect Site Code if not provided
        $SiteCode = (Get-CMSite -Verbose:$false).SiteCode
        if (-not $SiteCode) {
           throw "Could not automatically determine SCCM Site Code. Please provide the -SiteCode parameter."
        }
        Write-Log "Auto-detected Site Code: $SiteCode" 
    }

    $SiteDrive = "$($SiteCode):"
    if (-not (Test-Path $SiteDrive)) {
        Write-Log "SCCM PSDrive '$SiteDrive' not found. Attempting to establish connection..." 
         # This command might be needed if the drive isn't automatically created upon module import in some environments
         # New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $env:COMPUTERNAME -Verbose:$false
         # Let's try Get-CMSite again to potentially trigger drive creation or confirm connection
         try {
             Get-CMSite -SiteCode $SiteCode -Verbose:$false | Out-Null
             if (-not (Test-Path $SiteDrive)) {
                 throw "Failed to establish PSDrive connection to site '$SiteCode'. Verify permissions and connectivity."
             }
             Write-Log "Successfully connected to site '$SiteCode'." 
         } catch {
            throw "Failed to connect to SCCM site '$SiteCode'. $_"
         }
    }
    Set-Location $SiteDrive -Verbose:$false
    Write-Log "Current location set to '$($SiteDrive)'." 

    # Get Collection and check if it exists
    Write-Log "Verifying collection '$CollectionName'..." 
    $Collection = Get-CMCollection -Name $CollectionName -Verbose:$false
    if (-not $Collection) {
         throw "Collection '$CollectionName' not found."
    }
    Write-Log "Found Collection ID: $($Collection.CollectionID)." 

    # Get collection members
    Write-Log "Retrieving members for collection '$CollectionName'..." 
    $CollectionMembers = Get-CMCollectionMember -CollectionName $CollectionName -Verbose:$false
    $TotalMembers = $CollectionMembers.Count
    Write-Log "Collection '$CollectionName' has $TotalMembers members." 

    if ($TotalMembers -eq 0) {
        Write-Log "Collection '$CollectionName' is empty. No tests to run."  -LogLevel Warning
        # Optional: Still create an empty CSV? Or just exit? Let's create an empty one.
         New-Item -Path $FilePath -ItemType File -Force | Out-Null # Create empty file
         Write-Log "Empty results file created at $FilePath."  -LogLevel Warning
         $stopwatch.Stop()
         Write-Log "Script finished in $([Math]::Round($stopwatch.Elapsed.TotalSeconds)) seconds." 
         return # Exit script gracefully
    }

} catch {
    $log = "[$(Get-Date -format G)] Critical error during initial setup: $($_.Exception.Message)"
    Write-Error $log
    Write-Log $log  -LogLevel Error
    exit 1 # Exit with error code
} finally {
    $ErrorActionPreference = 'Continue' # Reset error action preference
}


# --- Parallel Processing ---
$runspacePool = $null
$TestResults = [System.Collections.Concurrent.ConcurrentBag[PSObject]]::new() # Thread-safe collection
$Jobs = [System.Collections.Generic.List[object]]::new()
$InitialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$WriteLogDefinition = Get-Content Function:\Write-Log
$SessionStateFunction = [System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new('Write-Log',$WriteLogDefinition)
$InitialSessionState.commands.add($SessionStateFunction)
try {
    # Initialize RunspacePool
    Write-Log "Initializing Runspace Pool with Min=1, Max=$MaxThreads threads." 
    $runspacePool = [RunspaceFactory]::CreateRunspacePool(1, $MaxThreads, $InitialSessionState, $Host)
    $runspacePool.Open()

    Write-Log "Submitting $TotalMembers jobs to the Runspace Pool..." 
    # Launch runspaces
    foreach ($member in $CollectionMembers) {
        $ServerName = $member.Name

        $ps = [PowerShell]::Create()
        $ps.RunspacePool = $runspacePool

        $null = $ps.AddScript({
            param ($TargetServerName, $ResultsBag, $LogFilePath)

            # Set error preference within the runspace thread
            $ErrorActionPreference = 'SilentlyContinue'
            $VerbosePreference = $ParentVerbosePreference # Inherit Verbose preference if needed

            #Write-Log "Testing WSMAN on '$TargetServerName'..." 
            
            
            $wsmanSuccess = $false
            $pingSuccess = $false

            # Test WSMAN
            try {
                # Using -ErrorAction Stop to force catch block on failure
                Test-WSMAN $TargetServerName -ErrorAction Stop | Out-Null
                $wsmanSuccess = $true
                $wsmanError = 'N/A'
                $ErrorValue = 0
                #Write-Log "WSMAN Success for '$TargetServerName'." 
            } catch {
                # Capture the specific WSManFault details if available
                $WSMANFAULT = ([xml]$_.ToString()).WSMANFAULT
                $ErrorMessage = $WSMANFAULT.Message # Capture the whole error record
                $ErrorValue = $WSMANFAULT.code
                Write-Log "WSMAN Failed for '$TargetServerName'. Error: $ErrorMessage" 
            }

            # Test Ping only if WSMAN failed, or always if needed (adjust logic here)
            # Let's always test ping for completeness
            #Write-Log "Testing Ping for '$TargetServerName'..." 
            $pingSuccess = Test-Connection -ComputerName $TargetServerName -Count 1 -Quiet -ErrorAction SilentlyContinue
            #Write-Log "Ping result for '$TargetServerName': $pingSuccess" 

            $result = [pscustomobject]@{
                Timestamp    = (Get-Date -format G)
                ServerName   = $TargetServerName
                WSManSuccess = $wsmanSuccess
                PingSuccess  = $pingSuccess
                WSManCode    = $ErrorValue
                WSManMessage = $ErrorMessage
            }

            # Add result to the thread-safe bag
            $ResultsBag.Add($result)

        }).AddArgument($ServerName).AddArgument($TestResults).AddArgument($LogFilePath)

        # Add inherited preferences if needed (e.g., for Write-Verbose inside the script block)
        $null = $ps.AddParameter("ParentVerbosePreference", $VerbosePreference)

        # Store the job handle
        $Job = [PSCustomObject]@{
            PowerShell = $ps
            Handle     = $ps.BeginInvoke()
            ServerName = $ServerName # Store for potential progress reporting
        }
        $Jobs.Add($Job)
    }

    # --- Wait for Completion & Progress ---
    Write-Log "All jobs submitted. Waiting for completion..." 
    $CompletedCount = 0
    $LastProgressUpdate = Get-Date

    while ($Jobs.Handle.IsCompleted -contains $false) {
        # Optional: Progress Reporting
        $Now = Get-Date
        if (($Now - $LastProgressUpdate).TotalSeconds -ge 5) { # Update every 5 seconds
            $CurrentCompleted = ($Jobs.Handle | Where-Object { $_.IsCompleted }).Count
            if ($CurrentCompleted -gt $CompletedCount) {
                $CompletedCount = $CurrentCompleted
                $PercentComplete = [Math]::Round(($CompletedCount / $TotalMembers) * 100)
                 Write-Progress -Activity "Running WSMAN Checks" -Status "$CompletedCount / $TotalMembers Completed ($PercentComplete%)" -PercentComplete $PercentComplete -CurrentOperation "Waiting for threads..."
                 $LastProgressUpdate = $Now
            }
        }
        Start-Sleep -Milliseconds 200 # Check status every 200ms
    }
    Write-Progress -Activity "Running WSMAN Checks" -Completed # Close progress bar

    Write-Log "All jobs completed. Processing results..." 

    # EndInvoke and check for script block errors
    foreach ($Job in $Jobs) {
        try {
            $Job.PowerShell.EndInvoke($Job.Handle)
        } catch {
            Write-Log "Error invoking script for server '$($Job.ServerName)': $($_.Exception.Message)"  -LogLevel Warning
            # Optionally add an error entry to $TestResults here if needed
        } finally {
             $Job.PowerShell.Dispose() # Dispose PowerShell instance
        }
    }

} finally {
    # Ensure RunspacePool is always closed and disposed
    if ($runspacePool -ne $null) {
        #Write-Log "Closing and disposing Runspace Pool..." 
        $runspacePool.Close()
        $runspacePool.Dispose()
        #Write-Log "Runspace Pool disposed." 
    }
}

$ErrorCount = ($TestResults | where {-not $_.WSManSuccess}).Count
$SuccessRate = [Math]::round(($TestResults | where {$_.WSManSuccess}).Count/$TestResults.Count*100)
Write-Log "$Errorcount servers coduldnt pass wsman test." 
Write-Log "Success Rate is $SuccessRate percent."
# --- Output Results ---
Write-Log "Exporting $($TestResults.Count) results to CSV: $FilePath" 
# Check if $WhatIf preference is set
if ($PSCmdlet.ShouldProcess($FilePath, "Export WSMAN Test Results")) {
    # Ensure directory exists
    $DirectoryPath = Split-Path -Path $FilePath -Parent
    if (-not (Test-Path -Path $DirectoryPath)) {
        Write-Log "Creating directory: $DirectoryPath" 
        New-Item -Path $DirectoryPath -ItemType Directory -Force | Out-Null
    }
    # Sort results before exporting for consistency
    $TestResults | Sort-Object ServerName | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
}

$stopwatch.Stop()
set-location $ExecutionPath
Write-Log "Script finished. Total time: $([Math]::Round($stopwatch.Elapsed.TotalSeconds)) seconds. Results exported to $FilePath"

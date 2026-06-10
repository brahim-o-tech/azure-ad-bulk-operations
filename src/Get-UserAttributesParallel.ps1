<#
.SYNOPSIS
    Retrieves Azure AD user attributes in parallel using RunspacePool.

.DESCRIPTION
    Processes a large list of Azure AD users (30,000+) in parallel using a
    PowerShell RunspacePool. For each user, retrieves license status,
    WhenCreated, and LastPasswordChangeDateTime via Microsoft Graph.

    Key pattern: BeginInvoke() dispatches jobs asynchronously.
    EndInvoke() is called exactly once per job to collect results.
    RunspacePool is always disposed in the finally block.

    Performance: ~7 minutes for 30,000+ users vs several hours sequentially.

.PARAMETER InputCsvPath
    Path to CSV file containing users to process.
    Must have a UserPrincipalName column.

.PARAMETER OutputPath
    Directory for output CSV files. Defaults to script root.

.PARAMETER SharedDrivePath
    Network share path where final CSV files are copied.

.PARAMETER OfficeFilterA
    Office attribute filter for store type A. Example: "# Store-A*"

.PARAMETER OfficeFilterB
    Office attribute filter for store type B. Example: "# Store-B*"

.PARAMETER MaxThreads
    Maximum concurrent threads. Defaults to 100.

.EXAMPLE
    .\Get-UserAttributesParallel.ps1 `
        -InputCsvPath    "C:\output\users.csv" `
        -OutputPath      "C:\output" `
        -SharedDrivePath "\\fileserver\shares\output" `
        -OfficeFilterA   "# Store-A*" `
        -OfficeFilterB   "# Store-B*" `
        -MaxThreads      100

.NOTES
    Author      : Brahim O.
    Version     : 1.0
    Requires    : Microsoft.Graph module
    Permissions : User.Read.All, Directory.Read.All
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ })]
    [string]$InputCsvPath,

    [string]$OutputPath = "",

    [Parameter(Mandatory)]
    [string]$SharedDrivePath,

    [string]$OfficeFilterA = "# Store-A*",
    [string]$OfficeFilterB = "# Store-B*",

    [ValidateRange(1, 500)]
    [int]$MaxThreads = 100
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not $OutputPath) { $OutputPath = $PSScriptRoot }

$RunStamp   = Get-Date -Format "yyyy-MM-dd_HH-mm"
$Transcript = Join-Path $OutputPath "Get-UserAttributesParallel-$RunStamp.log"
$TempCsv    = Join-Path $OutputPath "Temp-UserAttributes-$RunStamp.csv"

if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

Start-Transcript -Path $Transcript

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $entry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"
    switch ($Level) {
        "ERROR"   { Write-Host $entry -ForegroundColor Red    }
        "SUCCESS" { Write-Host $entry -ForegroundColor Green  }
        "WARN"    { Write-Host $entry -ForegroundColor Yellow }
        default   { Write-Host $entry }
    }
}

Write-Log "Script started — MaxThreads: $MaxThreads"
Write-Log "Input  : $InputCsvPath"
Write-Log "Output : $OutputPath"

# Import user list
$userList = Import-Csv -Path $InputCsvPath

if ($userList.Count -eq 0) {
    Write-Log "No users found in input file." "WARN"
    Stop-Transcript; exit 0
}

Write-Log "Users to process: $($userList.Count)" "SUCCESS"

# ScriptBlock — executed per user in each thread
$ScriptBlock = {
    param($User)
    try {
        $mgUser = Get-MgUser `
            -UserId   $User.UserPrincipalName `
            -Property DisplayName, AssignedLicenses, WhenCreated,
                      LastPasswordChangeDateTime, OfficeLocation `
            -ErrorAction Stop

        [PSCustomObject]@{
            EmployeeId                 = $User.EmployeeId
            DisplayName                = $mgUser.DisplayName
            UserPrincipalName          = $User.UserPrincipalName
            Office                     = $mgUser.OfficeLocation
            IsLicensed                 = ($mgUser.AssignedLicenses.Count -gt 0)
            WhenCreated                = $mgUser.WhenCreated
            LastPasswordChangeDateTime = $mgUser.LastPasswordChangeDateTime
        }
    }
    catch {
        [PSCustomObject]@{
            EmployeeId                 = $User.EmployeeId
            DisplayName                = $User.DisplayName
            UserPrincipalName          = $User.UserPrincipalName
            Office                     = "ERROR"
            IsLicensed                 = $false
            WhenCreated                = $null
            LastPasswordChangeDateTime = $null
        }
    }
}

$Pool = $null

try {
    # Initialize RunspacePool
    Write-Log "Initializing RunspacePool (1 to $MaxThreads threads)..."
    $Pool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads)
    $Pool.Open()
    Write-Log "RunspacePool opened." "SUCCESS"

    # Dispatch all jobs — BeginInvoke() is non-blocking
    Write-Log "Dispatching jobs..."
    $i = 0
    $Jobs = foreach ($User in $userList) {
        $Job = [System.Management.Automation.PowerShell]::Create()
        $Job.RunspacePool = $Pool
        $Job.AddScript($ScriptBlock).AddArgument($User) | Out-Null
        [PSCustomObject]@{ Pipe = $Job; Result = $Job.BeginInvoke() }
        $i++
        if ($i % 500 -eq 0) {
            Write-Progress -Activity "Dispatching..." -Status "$i of $($userList.Count)" `
                           -PercentComplete ([int](($i / $userList.Count) * 100))
        }
    }
    Write-Progress -Activity "Dispatching..." -Completed
    Write-Log "All $($Jobs.Count) jobs dispatched." "SUCCESS"

    # Monitor completion
    $elapsed = Measure-Command {
        do {
            $done = ($Jobs.Result | Where-Object { $_.IsCompleted }).Count
            Write-Log "Progress: $done / $($Jobs.Count) completed"
            if ($done -lt $Jobs.Count) { Start-Sleep -Seconds 30 }
        } while ($Jobs.Result.IsCompleted -contains $false)
    }
    Write-Log "All jobs completed in $([int]$elapsed.TotalMinutes) min." "SUCCESS"

    # Collect results — EndInvoke called ONCE per job
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($Job in $Jobs) {
        $out = $Job.Pipe.EndInvoke($Job.Result)
        if ($out) { $results.Add($out) }
        $Job.Pipe.Dispose()
    }
    Write-Log "Results collected: $($results.Count)" "SUCCESS"

    # Coherence check
    $countA = ($results | Where-Object { $_.Office -like $OfficeFilterA }).Count
    $countB = ($results | Where-Object { $_.Office -like $OfficeFilterB }).Count
    Write-Log "Store-A: $countA | Store-B: $countB | Total: $($results.Count)"

    if (($countA + $countB) -eq 0) {
        throw "Coherence check failed — no users matched either filter."
    }

    # Export
    $results | Export-Csv -Path $TempCsv -NoTypeInformation -Encoding UTF8

    $results | Where-Object { $_.Office -like $OfficeFilterA } |
        Export-Csv -Path (Join-Path $SharedDrivePath "UserAttributes-StoreA.csv") `
                   -NoTypeInformation -Encoding UTF8

    $results | Where-Object { $_.Office -like $OfficeFilterB } |
        Export-Csv -Path (Join-Path $SharedDrivePath "UserAttributes-StoreB.csv") `
                   -NoTypeInformation -Encoding UTF8

    Write-Log "All exports completed." "SUCCESS"
}
catch {
    Write-Log "FATAL ERROR: $($_.Exception.Message)" "ERROR"
    Write-Error $_.Exception.Message
}
finally {
    if ($null -ne $Pool) {
        $Pool.Close()
        $Pool.Dispose()
        Write-Log "RunspacePool disposed." "SUCCESS"
    }
    Write-Log "Script completed."
    Stop-Transcript
}

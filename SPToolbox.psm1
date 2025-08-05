<##
    SPToolbox – lightweight SharePoint admin toolbox
    Usage:
        Import-Module .\SPToolbox.psm1   # if saved as module
        Get-SPTool                       # list available tools
        Get-SPTool -Name 'Procmon','ULS Viewer' -Download   # download specific tools
        Get-SPTool -DownloadAll                          # download everything
#>

$Script:Tools = @(
    [PSCustomObject]@{
        Name        = "Process Monitor (Procmon)"
        FileName    = "ProcessMonitor.zip"

        OfficialUri = "https://download.sysinternals.com/files/ProcessMonitor.zip"
        InternalUri = "\\fileserver\Share\ProcessMonitor.zip"
        Description = "Captures real-time file, registry and process/thread activity on SharePoint servers to pinpoint permission errors, missing DLLs or unexpected file I/O."
    }
    [PSCustomObject]@{
        Name        = "ULS Viewer"
        FileName    = "ULSViewer.zip"
        OfficialUri = "https://download.microsoft.com/download/f/1/7/f17e2780-9456-479a-8faa-803cebf74b6a/ulsviewer.zip"
        InternalUri = "\\fileserver\Share\ULSViewer.zip"
        Description = "Tails and filters SharePoint's Unified Logging Service logs live-letting you hunt for correlation IDs and narrow in on specific services, categories and severities."
    }
    [PSCustomObject]@{
        Name        = "Fiddler"
        FileName    = "FiddlerSetup.exe"
        OfficialUri = "https://telerik-fiddler.s3.amazonaws.com/fiddler/FiddlerSetup.exe"
        InternalUri = "\\fileserver\Share\FiddlerSetup.exe"
        Description = "Intercepts and inspects HTTP/S traffic between clients and your SharePoint Web Front Ends-ideal for debugging REST/OData calls, cookies, redirects and auth flows."
    }
    [PSCustomObject]@{
        Name        = "Log Parser"
        FileName    = "LogParser.msi"
        OfficialUri = "https://download.microsoft.com/download/1/0/0/100b12a3-0f98-4af8-b63a-52a7de9d6cd3/LogParser.msi"
        InternalUri = "\\fileserver\Share\LogParser.msi"
        Description = "Run SQL-style queries against IIS, ULS and Usage logs to spot trends like most-hit pages, failing requests or slow response times."
    }
    [PSCustomObject]@{
        Name        = "SQL Server Profiler"
        FileName    = "SQLProfiler.msi"
        OfficialUri = "https://download.microsoft.com/download/2/5/2/25233c1c-1f6b-4e77-bf1e-d875c9d9e1bc/SqlProfiler.msi"
        InternalUri = "\\fileserver\Share\SqlProfiler.msi"
        Description = "Traces T-SQL calls from SharePoint into your content databases-helpful when diagnosing locking, blocking or tracking down an expensive query."
    }
    [PSCustomObject]@{
        Name        = "Performance Monitor"
        FileName    = "PerfMon"
        OfficialUri = $null
        InternalUri = $null
        Description = "Monitors Windows counters (CPU, memory, disk I/O) alongside SharePoint-specific counters to baseline and troubleshoot performance bottlenecks."
    }
    [PSCustomObject]@{
        Name        = "Debug Diagnostic Tool"
        FileName    = "DebugDiag.msi"
        OfficialUri = "https://download.microsoft.com/download/b/0/0/b0031952-e4bb-4ae3-93c0-85e8cdff3a0c/DebugDiag.msi"
        InternalUri = "\\fileserver\Share\DebugDiag.msi"
        Description = "Captures and analyzes IIS/COM+/CLR crash dumps on your SharePoint servers, with SharePoint-specific rules for common exceptions."
    }
    [PSCustomObject]@{
        Name        = "SPDiag (SPDiagnosticStudio)"
        FileName    = "SPDiagSetup.msi"
        OfficialUri = $null
        InternalUri = $null
        Description = "Microsoft's diagnostic framework for gathering configuration snapshots, ULS and PerfMon data, health reports and auto-analysis in one package."
    }
    [PSCustomObject]@{
        Name        = "SharePoint Log Viewer (SPLogViewer)"
        FileName    = "SPLogViewer.zip"
        OfficialUri = "https://download.microsoft.com/download/9/f/c/9fcd538e-d97c-4bb3-8ddb-8842f885f58e/SPLogViewer.zip"
        InternalUri = "\\fileserver\Share\SPLogViewer.zip"
        Description = "Lightweight, SharePoint-focused log tail utility with built-in hyperlinking of correlation IDs to Active Directory info."
    }
    [PSCustomObject]@{
        Name        = "SharePoint Manager (SPM)"
        FileName    = "SPM.zip"
        OfficialUri = $null
        InternalUri = $null
        Description = "Tree-view tool for browsing the entire SharePoint configuration database-including hidden properties on Service Applications, Web Applications, Sites and Features."
    }
    [PSCustomObject]@{
        Name        = "PnP PowerShell"
        FileName    = "PnP.PowerShell.msi"
        OfficialUri = "https://github.com/pnp/powershell/releases/latest/download/PnP.PowerShell.msi"
        InternalUri = "\\fileserver\Share\PnP.PowerShell.msi"
        Description = "Community-driven cmdlets that complement the out-of-the-box SharePoint module-great for bulk site provisioning, template extraction and cross-farm operations."
    }
    [PSCustomObject]@{
        Name        = "SPDocKit"
        FileName    = "SPDocKit.exe"
        OfficialUri = $null
        InternalUri = $null
        Description = "Third-party GUI that inventories your farm: permissions reports, feature usage, storage metrics, impact analysis before updates and much more."
    }
    [PSCustomObject]@{
        Name        = "SharePoint Health Analyzer"
        FileName    = "HealthAnalyzer"
        OfficialUri = $null
        InternalUri = $null

        Description = "Built-in farm-wide rule engine; export its reports and drill into them with ULS Viewer or SPDiag."
    }
    [PSCustomObject]@{
        Name        = "Process Explorer"
        FileName    = "ProcessExplorer.zip"
        OfficialUri = "https://download.sysinternals.com/files/ProcessExplorer.zip"
        InternalUri = "\\fileserver\Share\ProcessExplorer.zip"
        Description = "Deep process and handle inspection including SharePoint worker processes such as w3wp.exe and OWSTIMER.EXE."
    }
    [PSCustomObject]@{
        Name        = "Resource Monitor & Task Manager"
        FileName    = "ResourceMonitor"
        OfficialUri = $null
        InternalUri = $null
        Description = "Windows' built-in utilities for quick checks on CPU, RAM, disk and network per process-useful for at-a-glance health checks."
    }
)

function Get-SPTool {
    [CmdletBinding(DefaultParameterSetName = 'List')]
    param(
        [Parameter(Position = 0, ParameterSetName = 'Single')]
        [string[]] $Name,

        [Parameter(ParameterSetName = 'Single')]
        [switch] $Download,

        [Parameter(ParameterSetName = 'All')]
        [switch] $DownloadAll,

        [Parameter(ParameterSetName = 'Single')]
        [Parameter(ParameterSetName = 'All')]
        [ValidateSet('Official','Internal')]
        [string] $Source = 'Official',

        [string] $Destination = 'E:\SpToolbox\'
    )

    # If just listing, skip drive/directory checks
    if ($PSCmdlet.ParameterSetName -eq 'List') {
        return $Tools | Select-Object Name, Description, OfficialUri, InternalUri
    }

    # Validate target drive exists before attempting downloads
    $driveRoot = Split-Path -Path $Destination -Qualifier
    if (-not (Test-Path $driveRoot)) {
        Write-Host "Drive '$driveRoot' not found."
        $available = (Get-PSDrive -PSProvider FileSystem).Name -join ', '
        Write-Host "Available drives: $available"

        $maxAttempts = 3
        for ($i = 0; $i -lt $maxAttempts; $i++) {
            $letter = (Read-Host -Prompt "Enter the drive letter for the toolbox (use -Destination for a full path)").Trim().TrimEnd(':')
    
            if ($letter -notmatch '^[A-Za-z]$') {
                Write-Host "Please enter a single drive letter (e.g. C or D)."
            } else {
                $driveCandidate = "$($letter.ToUpper()):\\"
                if (Test-Path $driveCandidate) {
                    $Destination = Join-Path $driveCandidate 'SpToolbox'
                    break
                }
                Write-Host "Drive '$driveCandidate' does not exist."
            }
    
            if ($i -eq $maxAttempts - 1) {
                throw "Valid drive not specified."
            }
        }
    }

    # Ensure the destination folder exists
    if (-not (Test-Path $Destination)) {
        try {
            New-Item -Path $Destination -ItemType Directory -Force -ErrorAction Stop | Out-Null
        } catch {
            Write-Error "Failed to create destination folder '$Destination': $($_.Exception.Message)"
            return
        }
    }

    if ($DownloadAll) {
        foreach ($tool in $Tools) {
            $uri = if ($Source -eq 'Internal') { $tool.InternalUri } else { $tool.OfficialUri }
            if ($uri) {
                $path = Join-Path $Destination $tool.FileName
                try {
                    Invoke-WebRequest -Uri $uri -OutFile $path -ErrorAction Stop
                } catch {
                    Write-Warning "Could not download $($tool.Name): $($_.Exception.Message)"
                }
            }
        }
        Write-Host "Downloaded all tools to $Destination from '$Source' source."
        return
    }

    # Download specific tools
    foreach ($toolName in $Name) {
        $tool = $Tools | Where-Object { $_.Name -like "*$toolName*" }
        if (-not $tool) {
            Write-Error "Tool '$toolName' not found."
            continue
        }
        $uri = if ($Source -eq 'Internal') { $tool.InternalUri } else { $tool.OfficialUri }
        if (-not $uri) {
            Write-Warning "No '$Source' URI defined for $($tool.Name)."
            continue
        }
        $path = Join-Path $Destination $tool.FileName
        try {
            Invoke-WebRequest -Uri $uri -OutFile $path -ErrorAction Stop
            Write-Host "Downloaded $($tool.Name) to $path"
        } catch {
            Write-Error "Failed to download $($tool.Name): $($_.Exception.Message)"
        }
    }
}

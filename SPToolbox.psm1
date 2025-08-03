<##
    SPToolbox – lightweight SharePoint admin toolbox
    Usage:
        Import-Module .\SPToolbox.psm1   # if saved as module
        Get-SPTool                       # list available tools
        Get-SPTool "Procmon" -Download   # download specific tool
        Get-SPTool -DownloadAll          # download everything
#>

$Script:Tools = @(
    [PSCustomObject]@{
        Name        = "Process Monitor (Procmon)"
        FileName    = "ProcessMonitor.zip"
        Uri         = "https://download.sysinternals.com/files/ProcessMonitor.zip"
        Description = "Captures real-time file, registry and process/thread activity on SharePoint servers to pinpoint permission errors, missing DLLs or unexpected file I/O."
    }
    [PSCustomObject]@{
        Name        = "ULS Viewer"
        FileName    = "ULSViewer.zip"
        Uri         = "https://download.microsoft.com/download/0/7/0/0707d73f-7473-4c6e-aafa-bc69d80dfe39/ULSViewer.zip"
        Description = "Tails and filters SharePoint's Unified Logging Service logs live—letting you hunt for correlation IDs and narrow in on specific services, categories and severities."
    }
    [PSCustomObject]@{
        Name        = "Fiddler"
        FileName    = "FiddlerSetup.exe"
        Uri         = "https://telerik-fiddler.s3.amazonaws.com/fiddler/FiddlerSetup.exe"
        Description = "Intercepts and inspects HTTP/S traffic between clients and your SharePoint Web Front Ends—ideal for debugging REST/OData calls, cookies, redirects and auth flows."
    }
    [PSCustomObject]@{
        Name        = "Log Parser"
        FileName    = "LogParser.msi"
        Uri         = "https://download.microsoft.com/download/1/0/0/100b12a3-0f98-4af8-b63a-52a7de9d6cd3/LogParser.msi"
        Description = "Run SQL-style queries against IIS, ULS and Usage logs to spot trends like most-hit pages, failing requests or slow response times."
    }
    [PSCustomObject]@{
        Name        = "SQL Server Profiler"
        FileName    = "SQLProfiler.msi"
        Uri         = "https://download.microsoft.com/download/2/5/2/25233c1c-1f6b-4e77-bf1e-d875c9d9e1bc/SqlProfiler.msi"
        Description = "Traces T-SQL calls from SharePoint into your content databases—helpful when diagnosing locking, blocking or tracking down an expensive query."
    }
    [PSCustomObject]@{
        Name        = "Performance Monitor"
        FileName    = "PerfMon"
        Uri         = $null
        Description = "Monitors Windows counters (CPU, memory, disk I/O) alongside SharePoint-specific counters to baseline and troubleshoot performance bottlenecks."
    }
    [PSCustomObject]@{
        Name        = "Debug Diagnostic Tool"
        FileName    = "DebugDiag.msi"
        Uri         = "https://download.microsoft.com/download/b/0/0/b0031952-e4bb-4ae3-93c0-85e8cdff3a0c/DebugDiag.msi"
        Description = "Captures and analyzes IIS/COM+/CLR crash dumps on your SharePoint servers, with SharePoint-specific rules for common exceptions."
    }
    [PSCustomObject]@{
        Name        = "SPDiag (SPDiagnosticStudio)"
        FileName    = "SPDiagSetup.msi"
        Uri         = $null
        Description = "Microsoft's diagnostic framework for gathering configuration snapshots, ULS and PerfMon data, health reports and auto-analysis in one package."
    }
    [PSCustomObject]@{
        Name        = "SharePoint Log Viewer (SPLogViewer)"
        FileName    = "SPLogViewer.zip"
        Uri         = "https://download.microsoft.com/download/9/f/c/9fcd538e-d97c-4bb3-8ddb-8842f885f58e/SPLogViewer.zip"
        Description = "Lightweight, SharePoint-focused log tail utility with built-in hyperlinking of correlation IDs to Active Directory info."
    }
    [PSCustomObject]@{
        Name        = "SharePoint Manager (SPM)"
        FileName    = "SPM.zip"
        Uri         = $null
        Description = "Tree-view tool for browsing the entire SharePoint configuration database—including hidden properties on Service Applications, Web Applications, Sites and Features."
    }
    [PSCustomObject]@{
        Name        = "PnP PowerShell"
        FileName    = "PnP.PowerShell.msi"
        Uri         = "https://github.com/pnp/powershell/releases/latest/download/PnP.PowerShell.msi"
        Description = "Community-driven cmdlets that complement the out-of-the-box SharePoint module—great for bulk site provisioning, template extraction and cross-farm operations."
    }
    [PSCustomObject]@{
        Name        = "SPDocKit"
        FileName    = "SPDocKit.exe"
        Uri         = $null
        Description = "Third-party GUI that inventories your farm: permissions reports, feature usage, storage metrics, impact analysis before updates and much more."
    }
    [PSCustomObject]@{
        Name        = "SharePoint Health Analyzer"
        FileName    = "HealthAnalyzer"
        Uri         = $null
        Description = "Built-in farm-wide rule engine; export its reports and drill into them with ULS Viewer or SPDiag."
    }
    [PSCustomObject]@{
        Name        = "Process Explorer"
        FileName    = "ProcessExplorer.zip"
        Uri         = "https://download.sysinternals.com/files/ProcessExplorer.zip"
        Description = "Deep process and handle inspection including SharePoint worker processes such as w3wp.exe and OWSTIMER.EXE."
    }
    [PSCustomObject]@{
        Name        = "Resource Monitor & Task Manager"
        FileName    = "ResourceMonitor"
        Uri         = $null
        Description = "Windows’ built-in utilities for quick checks on CPU, RAM, disk and network per process—useful for at-a-glance health checks."
    }
)

function Get-SPTool {
    [CmdletBinding(DefaultParameterSetName = 'List')]
    param(
        [Parameter(Position = 0, ParameterSetName = 'Single')]
        [string] $Name,

        [Parameter(ParameterSetName = 'Single')]
        [switch] $Download,

        [Parameter(ParameterSetName = 'All')]
        [switch] $DownloadAll,

        [string] $Destination = (Join-Path $PSScriptRoot 'Downloads')
    )

    if ($PSCmdlet.ParameterSetName -eq 'List') {
        return $Tools | Select-Object Name, Description, Uri
    }

    if (-not (Test-Path $Destination)) {
        New-Item -ItemType Directory -Path $Destination | Out-Null
    }

    if ($DownloadAll) {
        foreach ($tool in $Tools) {
            if ($null -ne $tool.Uri) {
                $path = Join-Path $Destination $tool.FileName
                Invoke-WebRequest -Uri $tool.Uri -OutFile $path
            }
        }
        Write-Host "All downloadable tools fetched to $Destination"
        return
    }

    $tool = $Tools | Where-Object { $_.Name -like "*$Name*" }
    if ($tool -and $tool.Uri) {
        $path = Join-Path $Destination $tool.FileName
        Invoke-WebRequest -Uri $tool.Uri -OutFile $path
        Write-Host "Downloaded $($tool.Name) to $path"
    } elseif ($tool) {
        Write-Warning "$($tool.Name) is built into Windows or requires manual download."
    } else {
        Write-Error "Tool '$Name' not found."
    }
}

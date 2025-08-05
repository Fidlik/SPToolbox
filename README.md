# SPToolbox

A personal collection of PowerShell scripts, admin utilities, and automation snippets tailored for daily SharePoint operations – both on-premises and online. Built from real-world troubleshooting, patching, and governance scenarios, this toolbox is meant to save time, reduce human error, and make life easier for SharePoint administrators.

## PowerShell Module

`SPToolbox.psm1` exposes a simple `Get-SPTool` function for listing or downloading common diagnostic utilities.

```powershell
# 1. Set execution policy for the current session only
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# 2. Use TLS 1.2 for GitHub requests
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# 3. Define URL and destination
$repoUrl     = "https://raw.githubusercontent.com/Fidlik/SPToolbox/main/SPToolbox.psm1"
$destination = "$env:TEMP\SPToolbox.psm1"

# 4. Download the file
Invoke-WebRequest -Uri $repoUrl -OutFile $destination -UseBasicParsing

# 5. Unblock the downloaded file (remove Internet zone restriction)
Unblock-File -Path $destination

# 6. Import the module silently
Import-Module $destination -Force
#Import-Module .\SPToolbox.psm1


Get-SPTool               # list tools with descriptions

Get-SPTool -Name 'Procmon','ULS Viewer'                # download specific tools
Get-SPTool -DownloadAll -Source Official               # fetch everything from official sources
Get-SPTool -Name 'Procmon' -Source Internal            # fetch from internal source

By default, downloads are placed in a "E:\SpToolbox\". Override this location with the `-Destination` parameter.

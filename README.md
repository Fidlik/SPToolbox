# SPToolbox

A personal collection of PowerShell scripts, admin utilities, and automation snippets tailored for daily SharePoint operations – both on-premises and online. Built from real-world troubleshooting, patching, and governance scenarios, this toolbox is meant to save time, reduce human error, and make life easier for SharePoint administrators.

## PowerShell Module

`SPToolbox.psm1` exposes a simple `Get-SPTool` function for listing or downloading common diagnostic utilities.

```powershell
Import-Module .\SPToolbox.psm1
Get-SPTool               # list tools with descriptions

Get-SPTool "Procmon" -Download               # download specific tool
Get-SPTool "Procmon" -Download -Source Internal  # fetch from internal share
Get-SPTool -DownloadAll -Source Official      # fetch everything from official sources
 main
```

By default, downloads are placed in a `Downloads` folder next to the module file. Override this location with the `-Destination` parameter.

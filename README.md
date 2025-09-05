# PSCascadeWindows

This PowerShell script cascades all visible windows on the desktop.

## Requirements

- Microsoft Windows
  - Windows PowerShell 5.1+

## Usage

```powershell
# powershell

git clone https://...

Set-Location .\PSCascadeWindows\
.\Invoke-CascadeWindows.ps1
```

## Appendix: How to create shortcut

```powershell
# powershell

Set-Location .\PSCascadeWindows\

$location = Get-Location
$gitDirectoryPath = $location.Path
$shortcutPath = Join-Path $gitDirectoryPath "PSCascadeWindows.lnk"
$scriptPath = Join-Path $gitDirectoryPath "Invoke-CascadeWindows.ps1"

$wshShell = New-Object -ComObject WScript.Shell
$shortcut = $wshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = "powershell.exe"
$shortcut.Arguments = "-NoProfile -File `"$scriptPath`""
$shortcut.WindowStyle = 7  # Minimized, https://learn.microsoft.com/en-us/troubleshoot/windows-client/admin-development/create-desktop-shortcut-with-wsh
$shortcut.Save()

Start-Process $gitDirectoryPath
```

## License

MIT License

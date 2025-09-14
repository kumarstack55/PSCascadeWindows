# PSCascadeWindows

This PowerShell script cascades all visible windows on the desktop.

![screenshot](./images/screenshot.png)

## Motivation

Starting with Windows 11, Microsoft no longer provides a built-in feature for cascading windows.

<https://learn.microsoft.com/en-us/answers/questions/5516749/cascade-windows-11>
> Windows 11 doesn’t include a built-in "Cascade Windows" option like previous versions, but you can use Snap Layouts to quickly organize your open apps on the screen.

Older versions of Windows supported cascading windows.
This feature is useful. For example, it's handy when using a narrow screen.

The following would make it even more convenient:

- Easier selection of the window behind via the top-right corner. This prevents accidentally clicking the close button.
- Displaying the last line of the window behind at the bottom. This allows viewing the latest log line with `tail -f`.

## Requirements

- Microsoft Windows 11
  - Windows PowerShell 5.1+

## Usage

```powershell
# powershell

git clone https://github.com/kumarstack55/PSCascadeWindows.git

Set-Location .\PSCascadeWindows\
.\Invoke-CascadeWindows.ps1
```

## More easy way to cascade windows: Shortcut

You can create a shortcut using the following steps.
Double-clicking the created shortcut will cascade windows.

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

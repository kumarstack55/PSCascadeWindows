Add-Type -AssemblyName System.Windows.Forms

Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public class User32 {
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

    [DllImport("user32.dll")]
    public static extern IntPtr GetTopWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

    [DllImport("user32.dll")]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool IsZoomed(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
}

public class Dwm {
    [DllImport("dwmapi.dll")]
    public static extern int DwmGetWindowAttribute(IntPtr hWnd, int dwAttribute, out int pvAttribute, int cbAttribute);
}

[StructLayout(LayoutKind.Sequential)]
public struct RECT {
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;
}
"@

function Get-WindowHandleList {
    param()

    $list = [System.Collections.Generic.List[IntPtr]]::new()

    $callback = [User32+EnumWindowsProc]{
        param($hWnd, $lParam)
        $list.Add($hWnd)
        return $true
    }

    $isOk = [User32]::EnumWindows($callback, [IntPtr]::Zero)
    if (-not $isOk) {
        throw "Failed to enumerate windows."
    }

    return $list
}

filter Select-VisibleWindowHandles {
    process {
        if ([User32]::IsWindowVisible($_)) {
            $_
        }
    }
}

function Get-WindowTitle {
    param(
        [Parameter(Mandatory)]
        [IntPtr]$hWnd
    )

    $maxCount = 256
    $sb = New-Object System.Text.StringBuilder $maxCount
    $n = [User32]::GetWindowText($hWnd, $sb, $sb.Capacity)
    if ($n -ge $sb.Capacity) {
        throw "Window title is too long. (n: $n)"
    }

    return $sb.ToString()
}

function Get-WindowPosition {
    param(
        [Parameter(Mandatory)]
        [IntPtr]$hWnd
    )

    $rect = New-Object RECT
    $isOk = [User32]::GetWindowRect($hWnd, [ref]$rect)
    if (-not $isOk) {
        throw "Failed to get window rect for handle $Handle."
    }

    return @{
        Left   = $rect.Left
        Top    = $rect.Top
        Right  = $rect.Right
        Bottom = $rect.Bottom
        Width  = $rect.Right - $rect.Left
        Height = $rect.Bottom - $rect.Top
    }
}

function Get-ProcessId {
    param(
        [Parameter(Mandatory)]
        [IntPtr]$hWnd
    )

    $processId = 0
    $threadId = [User32]::GetWindowThreadProcessId($hWnd, [ref]$processId)
    if ($threadId -eq 0) {
        throw "Failed to get process ID for handle $hWnd."
    }

    return $processId
}

function Test-TooSmallWindow {
    param(
        [Parameter(Mandatory)]
        [IntPtr]$hWnd,
        [int]$minWidth = 80,
        [int]$minHeight = 80
    )

    $rect = New-Object RECT
    $isOk = [User32]::GetWindowRect($hWnd, [ref]$rect)
    if (-not $isOk) {
        throw "Failed to get window rect for handle $Handle."
    }

    $width = $rect.Right - $rect.Left
    $height = $rect.Bottom - $rect.Top

    return ($width -lt $minWidth) -or ($height -lt $minHeight)
}

function Test-WindowHasOwner {
    param(
        [Parameter(Mandatory)]
        [IntPtr]$hWnd
    )

    $GW_OWNER = 4
    $ownerHandle = [User32]::GetWindow($hWnd, $GW_OWNER)
    $hasOwner = $ownerHandle -ne [IntPtr]::Zero
    return $hasOwner
}

function Test-IsCloakedWindow {
    param(
        [Parameter(Mandatory)]
        [IntPtr]$hWnd
    )

    $DWMWA_CLOAKED = 14
    $isCloaked = 0
    $hresult = [Dwm]::DwmGetWindowAttribute($hWnd, $DWMWA_CLOAKED, [ref]$isCloaked, 4)
    if ($hresult -ne 0) {
        throw "Failed to get DWM window attribute for handle $hWnd. HRESULT: $hresult"
    }

    return $isCloaked -ne 0
}

function Get-ClassName {
    param(
        [Parameter(Mandatory)]
        [IntPtr]$hWnd
    )

    $maxCount = 256
    $sb = New-Object System.Text.StringBuilder $maxCount
    $n = [User32]::GetClassName($hWnd, $sb, $sb.Capacity)
    if ($n -ge $sb.Capacity) {
        throw "Class name is too long. (n: $n)"
    }

    return $sb.ToString()
}

filter Select-NormalWindow {
    process {
        $hWnd = $_

        # Skip invisible windows
        if (-not [User32]::IsWindowVisible($hWnd)) {
            return
        }

        # Skip cloaked windows (e.g., UWP apps)
        if (Test-IsCloakedWindow -hWnd $hWnd) {
            return
        }

        # Skip windows that have an owner (e.g., dialog boxes)
        if (Test-WindowHasOwner -hWnd $hWnd) {
            return
        }

        # Skip windows that are too small
        if (Test-TooSmallWindow -hWnd $hWnd) {
            return
        }

        $className = Get-ClassName -hWnd $hWnd
        $excludeClassNameExists = [hashtable]@{
            "Progman" = $true            # explorer
            #"Shell_TrayWnd" = $true      # Taskbar
            #"Button" = $true             # Start button
            #"MsgrIMEWindowClass" = $true # IME window
        }
        if ($excludeClassNameExists.ContainsKey($className)) {
            return
        }

        $hWnd
    }
}

function Get-ZOrderedHandleToOrderMap {
    $zOrderedHandles = [System.Collections.Generic.List[IntPtr]]::new()
    $handle = [User32]::GetTopWindow([IntPtr]::Zero)
    while ($handle -ne [IntPtr]::Zero) {
        $zOrderedHandles.Add($handle)
        $GW_HWNDNEXT = 2
        $handle = [User32]::GetWindow($handle, $GW_HWNDNEXT)
    }

    $zOrderedHandleToOrderMap = @{}
    for ($index = 0; $index -lt $zOrderedHandles.Count; $index++) {
        $handle = $zOrderedHandles[$index]
        $zOrderedHandleToOrderMap[$handle] = $index
    }

    return $zOrderedHandleToOrderMap
}

function Get-WindowPsoList {
    param()

    # Get all window handles
    $allHandles = Get-WindowHandleList
    write-host "allHandles.Count: $($allHandles.Count)"

    # Filter to get only normal visible windows
    $handles = @($allHandles | Select-NormalWindow)
    write-host "visibleHandles.Count: $($handles.Count)"

    # Get z-order
    $zOrderedHandleToOrderMap = Get-ZOrderedHandleToOrderMap

    # Create PSObject list
    $windowPsoList = [System.Collections.Generic.List[PSObject]]::new()
    $handleToPsoMap = @{}
    foreach ($hWnd in $handles) {
        $title = Get-WindowTitle -hWnd $hWnd
        $processId = Get-ProcessId -hWnd $hWnd
        $className = Get-ClassName -hWnd $hWnd
        $process = Get-Process -Id $processId
        $processName = $process.ProcessName

        $isMaximized = [User32]::IsZoomed($hWnd)

        $zOrder = $zOrderedHandleToOrderMap[$hWnd]

        $hashtable = @{
            Title  = $title
            Handle = $hWnd
            ZOrder = $zOrder
            ProcessName = $processName
            ProcessId   = $processId
            ClassName = $className
            IsMaximized = $isMaximized
        }

        $pso = [PSCustomObject]$hashtable
        $windowPsoList.Add($pso)
        $handleToPsoMap[$hWnd] = $pso
    }

    # Sort by z-order
    $windowPsoList2 = $windowPsoList | Sort-Object -Property ZOrder

    return $windowPsoList2
}

function Move-Windows {
    param(
        [Parameter(Mandatory)]
        [System.Drawing.Rectangle]$WorkingArea,

        [Parameter(Mandatory)]
        [System.Collections.Generic.List[PSObject]]$WindowPsoList,

        [int]$Margin = 16,
        [int]$CascadeStepX = 160,
        [int]$CascadeStepY = 40
    )

    # Define cascade area
    $Margin = 20
    $cascadeRect = [PSCustomObject]@{
        Left = $WorkingArea.X + $Margin
        Top = $WorkingArea.Y + $Margin
        Right = $WorkingArea.Right - $Margin * 2
        Bottom = $WorkingArea.Bottom - $Margin * 2
    }
    $cascadeWidth = $cascadeRect.Right - $cascadeRect.Left
    $cascadeHeight = $cascadeRect.Bottom - $cascadeRect.Top

    # Calculate window dimensions
    $windowsCount = $WindowPsoList.Count
    $windowWidth = [int]($cascadeWidth - $CascadeStepX * ($windowsCount - 1))
    $windowHeight = [int]($cascadeHeight - $CascadeStepY * ($windowsCount - 1))

    # Resize and move windows
    for ($index = 0; $index -lt $windowsCount; $index++) {
        $pso = $WindowPsoList[$index]

        $hWnd = $pso.Handle

        $x = $cascadeRect.Left + $CascadeStepX * $index
        $y = $cascadeRect.Top + $CascadeStepY * $index

        $isOk = [User32]::SetWindowPos($hWnd, [IntPtr]::Zero, $x, $y, $windowWidth, $windowHeight, 0)
        if (-not $isOk) {
            throw "Failed to move window: $($pso.Handle) $($pso.Title)"
        }
    }
}

function Invoke-CascadeWindows {
    param()

    # Get window list
    $windowPsoList = Get-WindowPsoList
    $windowPsoList

    # Restore maximized windows
    $SW_RESTORE = 9
    foreach ($pso in $windowPsoList) {
        $hWnd = $pso.Handle
        if ($hWnd -ne [IntPtr]::Zero -and [User32]::IsZoomed($hWnd)) {
            [User32]::ShowWindow($hWnd, $SW_RESTORE) | Out-Null
        }
    }

    # Get working area of the primary screen
    $primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
    $workingArea = $primaryScreen.WorkingArea
    Move-Windows -WorkingArea $workingArea -WindowPsoList $windowPsoList
}

if ($MyInvocation.InvocationName -eq '.') {
    # Dot-sourced
    ; # do nothing
} elseif ($MyInvocation.InvocationName -match '\.ps1$') {
    # Run directly as script
    Invoke-CascadeWindows
}

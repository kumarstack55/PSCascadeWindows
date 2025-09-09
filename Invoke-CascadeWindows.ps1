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

function Get-ClassName {
    param(
        [Parameter(Mandatory)]
        [IntPtr]$Handle
    )

    $maxCount = 256
    $sb = New-Object System.Text.StringBuilder $maxCount
    $n = [User32]::GetClassName($Handle, $sb, $sb.Capacity)
    if ($n -ge $sb.Capacity) {
        throw "Class name is too long. (n: ${n})"
    }

    return $sb.ToString()
}

function Get-ProcessId {
    param(
        [Parameter(Mandatory)]
        [IntPtr]$Handle
    )

    $processId = 0
    $threadId = [User32]::GetWindowThreadProcessId($Handle, [ref]$processId)
    if ($threadId -eq 0) {
        throw "Failed to get process ID for handle ${Handle}."
    }

    return $processId
}

function Get-ProcessName {
    param(
        [Parameter(Mandatory)]
        [IntPtr]$Handle
    )

    $processId = Get-ProcessId -Handle $Handle
    try {
        $process = Get-Process -Id $processId -ErrorAction Stop
        return $process.ProcessName
    } catch {
        return ""
    }
}
function Get-WindowTitle {
    param(
        [Parameter(Mandatory)]
        [IntPtr]$Handle
    )

    $maxCount = 256
    $sb = New-Object System.Text.StringBuilder $maxCount
    $n = [User32]::GetWindowText($Handle, $sb, $sb.Capacity)
    if ($n -ge $sb.Capacity) {
        throw "Window title is too long. (n: ${n})"
    }

    return $sb.ToString()
}
function Test-WindowIsMaximized {
    param(
        [Parameter(Mandatory)]
        [IntPtr]$Handle
    )

    return [User32]::IsZoomed($Handle)
}

class Window {
    [IntPtr]$Handle
    [string]$ClassName
    [int]$ProcessId
    [string]$ProcessName
    [string]$Title
    [int]$ZOrder
    [bool]$IsMaximized
}

class WindowFactory {
    hidden [hashtable]$ZOrderedHandleToOrderMap

    WindowFactory([hashtable]$ZOrderedHandleToOrderMap) {
        $this.ZOrderedHandleToOrderMap = $ZOrderedHandleToOrderMap
    }
    [Window] Create([IntPtr]$Handle) {
        $window = [Window]::new()

        if ([IntPtr]::Zero -eq $Handle) {
            throw "Handle is zero."
        }
        $window.Handle = $Handle

        $window.ClassName = Get-ClassName -Handle $Handle
        $window.ProcessId = Get-ProcessId -Handle $Handle
        $window.ProcessName = Get-ProcessName -Handle $Handle
        $window.Title = Get-WindowTitle -Handle $Handle
        $window.ZOrder = $this.ZOrderedHandleToOrderMap[$Handle]
        $window.IsMaximized = Test-WindowIsMaximized -Handle $Handle

        return $window
    }
}

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

function Test-TooSmallWindow {
    param(
        [Parameter(Mandatory)]
        [IntPtr]$Handle,
        [int]$minWidth = 80,
        [int]$minHeight = 80
    )

    $rect = New-Object RECT
    $isOk = [User32]::GetWindowRect($Handle, [ref]$rect)
    if (-not $isOk) {
        throw "Failed to get window rect for handle ${Handle}."
    }

    $width = $rect.Right - $rect.Left
    $height = $rect.Bottom - $rect.Top

    return ($width -lt $minWidth) -or ($height -lt $minHeight)
}

function Test-WindowHasOwner {
    param(
        [Parameter(Mandatory)]
        [IntPtr]$Handle
    )

    $GW_OWNER = 4
    $ownerHandle = [User32]::GetWindow($Handle, $GW_OWNER)
    $hasOwner = $ownerHandle -ne [IntPtr]::Zero

    return $hasOwner
}

function Test-IsCloakedWindow {
    param(
        [Parameter(Mandatory)]
        [IntPtr]$Handle
    )

    $DWMWA_CLOAKED = 14
    $isCloaked = 0
    $hResult = [Dwm]::DwmGetWindowAttribute($Handle, $DWMWA_CLOAKED, [ref]$isCloaked, 4)
    if ($hResult -ne 0) {
        throw "Failed to get DWM window attribute for handle ${Handle}. HRESULT: ${hResult}"
    }

    return $isCloaked -ne 0
}

filter Select-NormalWindow {
    process {
        $handle = $_

        # Skip invisible windows
        if (-not [User32]::IsWindowVisible($handle)) {
            return
        }

        # Skip cloaked windows (e.g., UWP apps)
        if (Test-IsCloakedWindow -Handle $handle) {
            return
        }

        # Skip windows that have an owner (e.g., dialog boxes)
        if (Test-WindowHasOwner -Handle $handle) {
            return
        }

        # Skip windows that are too small
        if (Test-TooSmallWindow -Handle $handle) {
            return
        }

        $className = Get-ClassName -Handle $handle
        $excludeClassNameExists = [hashtable]@{
            "Progman" = $true            # explorer
            #"Shell_TrayWnd" = $true      # Taskbar
            #"Button" = $true             # Start button
            #"MsgrIMEWindowClass" = $true # IME window
        }
        if ($excludeClassNameExists.ContainsKey($className)) {
            return
        }

        return $handle
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

function Get-WindowList {
    param()

    # Get all window handles
    $allHandles = Get-WindowHandleList
    Write-Host "allHandles.Count: $($allHandles.Count)"

    # Filter to get only normal visible windows
    $handles = @($allHandles | Select-NormalWindow)
    Write-Host "visibleHandles.Count: $($handles.Count)"

    # Create window list
    $windowList = [System.Collections.Generic.List[Window]]::new()
    $zOrderedHandleToOrderMap = Get-ZOrderedHandleToOrderMap
    $windowFactory = [WindowFactory]::new($zOrderedHandleToOrderMap)
    foreach ($handle in $handles) {
        $window = $windowFactory.Create($handle)
        $windowList.Add($window)
    }

    # Sort by z-order
    #
    # The process returns an Array instead of a List. Since it's not a problem, we'll choose not to worry about it.
    $windowList2 = $windowList | Sort-Object -Property ZOrder

    return $windowList2
}

function Move-WindowsCascaded {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [System.Drawing.Rectangle]$WorkingArea,

        [Parameter(Mandatory)]
        [System.Collections.Generic.List[Window]]$WindowList,

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
    $windowsCount = $WindowList.Count
    $windowWidth = [int]($cascadeWidth - $CascadeStepX * ($windowsCount - 1))
    $windowHeight = [int]($cascadeHeight - $CascadeStepY * ($windowsCount - 1))

    # Resize and move windows
    #
    # Some applications such like google chrome, etc are changed z-order when resizing/moving.
    # So, we move windows in z-order from bottom to top.
    for ($index = $windowsCount - 1; $index -ge 0; $index--) {
        $window = $WindowList[$index]
        $handle = $window.Handle
        $title = $window.Title

        $target = "${handle} (title: ${title})"

        # Restore maximized windows
        $SW_RESTORE = 9
        $operation1 = "Restore window"
        if ([User32]::IsZoomed($handle)) {
            if ($PSCmdlet.ShouldProcess($target, $operation1)) {
                [User32]::ShowWindow($handle, $SW_RESTORE) | Out-Null
            }
        }

        # Move window
        $x = $cascadeRect.Left + $CascadeStepX * $index
        $y = $cascadeRect.Top + $CascadeStepY * $index
        $operation2 = "Move to (x: ${x}, y: ${y}, w: ${windowWidth}, h: ${windowHeight})"
        if ($PSCmdlet.ShouldProcess($target, $operation2)) {
            $NO_FLAGS = 0
            $isOk = [User32]::SetWindowPos($handle, [IntPtr]::Zero, $x, $y, $windowWidth, $windowHeight, $NO_FLAGS)
            if (-not $isOk) {
                throw "Failed to move window: ${handle} ${title}"
            }
        }
    }
}

function Invoke-CascadeWindows {
    [CmdletBinding(SupportsShouldProcess)]
    param()

    # Get window list
    $windowList = Get-WindowList
    $windowList

    # Get working area of the primary screen
    $primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
    $workingArea = $primaryScreen.WorkingArea

    # Cascade windows
    Move-WindowsCascaded -WorkingArea $workingArea -WindowList $windowList
}

if ($MyInvocation.InvocationName -eq '.') {
    # Dot-sourced
    ; # do nothing
} elseif ($MyInvocation.InvocationName -match '\.ps1$') {
    # Run directly as script
    Invoke-CascadeWindows
}

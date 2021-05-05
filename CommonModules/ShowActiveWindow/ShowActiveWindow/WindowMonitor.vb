Imports System.Runtime.InteropServices
Imports System.Threading

Module WindowMonitor
    Private Const MAXIMIZED_OFFSET = 8

    Private Delegate Sub WinEventDelegate(
        ByVal hWinEventHook As IntPtr,
        ByVal eventType As UInteger,
        ByVal hwnd As IntPtr,
        ByVal idObject As Integer,
        ByVal idChild As Integer,
        ByVal dwEventThread As UInteger,
        ByVal dwmsEventTime As UInteger
    )

    <DllImport("user32.dll")>
    Private Function SetWinEventHook(eventMin As UInteger, eventMax As UInteger, hmodWinEventProc As IntPtr,
                                            lpfnWinEventProc As WinEventDelegate, idProcess As UInteger, idThread As UInteger,
                                            dwFlags As UInteger) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Function UnhookWinEvent(hWinEventHook As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Function GetWindowThreadProcessId(ByVal hwnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    End Function

    Enum ShowWindowCommand As Integer
        Hide = 0
        Normal = 1
        ShowMinimized = 2
        Maximize = 3
        ShowMaximized = 3
        ShowNoActivate = 4
        Show = 5
        Minimize = 6
        ShowMinNoActive = 7
        ShowNA = 8
        Restore = 9
        ShowDefault = 10
        ForceMinimize = 11
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Private Structure WINDOWPLACEMENT
        Public length As Integer
        Public flags As Integer
        Public showCmd As ShowWindowCommand
        Public minPosition As Point
        Public maxPosition As Point
        Public normalPosition As RECT
    End Structure

    <DllImport("user32.dll")>
    Private Function GetWindowPlacement(ByVal hWnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
    End Function

    Private Const WINEVENT_OUTOFCONTEXT As UInteger = 0
    Private Const WINEVENT_SKIPOWNPROCESS As UInteger = 2
    Private Const EVENT_SYSTEM_FOREGROUND As UInteger = 3
    Private Const EVENT_SYSTEM_MOVESIZESTART As UInteger = &HA
    Private Const EVENT_SYSTEM_MOVESIZEEND As UInteger = &HB
    Private Const EVENT_OBJECT_SHOW As UInteger = &H8002
    Private Const EVENT_OBJECT_HIDE As UInteger = &H8003
    Private Const EVENT_OBJECT_IME_SHOW As UInteger = &H8027
    Private Const EVENT_OBJECT_IME_HIDE As UInteger = &H8028
    Private Const EVENT_OBJECT_LOCATIONCHANGE As UInteger = &H800B

    Private objSysDelegate As New WinEventDelegate(AddressOf SystemEventProc)
    Private handleSystemEventHookFore As IntPtr = IntPtr.Zero
    Private handleSystemEventHookSize As IntPtr = IntPtr.Zero
    Private handleObjectEventHookState As IntPtr = IntPtr.Zero

    Private handleForeground As IntPtr

    Public Sub InitializeSystemMonitor()
        UninitializeSystemMonitor()
        handleSystemEventHookFore = SetWinEventHook(EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND, IntPtr.Zero, objSysDelegate, 0, 0, WINEVENT_OUTOFCONTEXT Or WINEVENT_SKIPOWNPROCESS)
        handleSystemEventHookSize = SetWinEventHook(EVENT_SYSTEM_MOVESIZESTART, EVENT_SYSTEM_MOVESIZEEND, IntPtr.Zero, objSysDelegate, 0, 0, WINEVENT_OUTOFCONTEXT Or WINEVENT_SKIPOWNPROCESS)
    End Sub

    Public Sub UninitializeSystemMonitor()
        If handleSystemEventHookFore <> IntPtr.Zero Then
            UnhookWinEvent(handleSystemEventHookFore)
            handleSystemEventHookFore = IntPtr.Zero
        End If

        If handleSystemEventHookSize <> IntPtr.Zero Then
            UnhookWinEvent(handleSystemEventHookSize)
            handleSystemEventHookSize = IntPtr.Zero
        End If

        If handleObjectEventHookState <> IntPtr.Zero Then
            UnhookWinEvent(handleObjectEventHookState)
            handleObjectEventHookState = IntPtr.Zero
        End If

    End Sub

    Public Sub SystemEventProc(hWinEventHook As IntPtr, eventType As UInteger, hwnd As IntPtr, idObject As Integer,
                            idChild As Integer, dwEventThread As UInteger, dwmsEventTime As UInteger)

        If hwnd = frmMain.Handle Then
            Exit Sub
        End If

        If Not (hWinEventHook = handleSystemEventHookFore Or hWinEventHook = handleSystemEventHookSize Or hWinEventHook = handleObjectEventHookState) Then
            Exit Sub
        End If

        Select Case eventType
            Case EVENT_SYSTEM_FOREGROUND
                handleForeground = hwnd

                ' Listen to window move events
                If handleObjectEventHookState <> IntPtr.Zero Then
                    UnhookWinEvent(handleObjectEventHookState)
                    handleObjectEventHookState = IntPtr.Zero
                End If

                ' Get process ID
                Dim processID As Integer
                GetWindowThreadProcessId(hwnd, processID)
                handleObjectEventHookState = SetWinEventHook(EVENT_OBJECT_LOCATIONCHANGE, EVENT_OBJECT_LOCATIONCHANGE, IntPtr.Zero, objSysDelegate, processID, 0, WINEVENT_OUTOFCONTEXT Or WINEVENT_SKIPOWNPROCESS)

                RepositionMainForm(hwnd)

            Case EVENT_SYSTEM_MOVESIZESTART
                frmMain.Hide()

            Case EVENT_SYSTEM_MOVESIZEEND
                RepositionMainForm(hwnd)
                frmMain.Show()

            Case EVENT_OBJECT_LOCATIONCHANGE
                If hwnd = handleForeground Then
                    RepositionMainForm(hwnd)
                    frmMain.Show()
                End If
        End Select

    End Sub

    Public Function IsWindowMaximized(handle As IntPtr) As Boolean
        Dim placement As WINDOWPLACEMENT
        GetWindowPlacement(handle, placement)

        If placement.showCmd = ShowWindowCommand.ShowMaximized Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure


    <DllImport("user32.dll")>
    Private Function GetWindowRect(ByVal hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
    End Function

    Public Sub RepositionMainForm(mimicHwnd As IntPtr)

        ' Get size of target window
        Dim aRect As RECT
        GetWindowRect(mimicHwnd, aRect)

        ' Do special sizing for maximized windows
        Dim maximizedOffset As Integer = 0

        If IsWindowMaximized(mimicHwnd) Then
            ' Maximized.  Move it in a little
            maximizedOffset = MAXIMIZED_OFFSET
        End If

        frmMain.Size = New Size(aRect.Right - aRect.Left - maximizedOffset * 2, aRect.Bottom - aRect.Top - maximizedOffset * 2)
        frmMain.Location = New Point(aRect.Left + maximizedOffset, aRect.Top + maximizedOffset)
        frmMain.Invalidate()
    End Sub


End Module


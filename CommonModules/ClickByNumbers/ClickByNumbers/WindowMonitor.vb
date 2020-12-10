Imports System.Runtime.InteropServices
Imports System.Threading

Module WindowMonitor
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

    Private Const WINEVENT_OUTOFCONTEXT As UInteger = 0
    Private Const WINEVENT_SKIPOWNPROCESS As UInteger = 2
    Private Const EVENT_SYSTEM_FOREGROUND As UInteger = 3
    Private Const EVENT_OBJECT_SHOW As UInteger = &H8002
    Private Const EVENT_OBJECT_HIDE As UInteger = &H8003
    Private Const EVENT_OBJECT_IME_SHOW As UInteger = &H8027
    Private Const EVENT_OBJECT_IME_HIDE As UInteger = &H8028

    Private objSysDelegate As New WinEventDelegate(AddressOf SystemEventProc)
    Private handleSystemEventHook As IntPtr
    Private handleWinEventHook(1) As IntPtr

    Public Sub InitializeSystemMonitor()
        UninitializeSystemMonitor()
        handleSystemEventHook = SetWinEventHook(EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND, IntPtr.Zero, objSysDelegate, 0, 0, WINEVENT_OUTOFCONTEXT Or WINEVENT_SKIPOWNPROCESS)
    End Sub

    Public Sub UninitializeSystemMonitor()
        If handleSystemEventHook <> Nothing Then
            UnhookWinEvent(handleSystemEventHook)
            handleSystemEventHook = Nothing
        End If
    End Sub

    Public Sub SystemEventProc(hWinEventHook As IntPtr, eventType As UInteger, hwnd As IntPtr, idObject As Integer,
                            idChild As Integer, dwEventThread As UInteger, dwmsEventTime As UInteger)

        If Not hWinEventHook.Equals(handleSystemEventHook) Then
            Exit Sub
        End If

        If hwnd = frmMain.Handle Or
           hwnd = frmOptions.Handle Or
           hwnd = frmProcess.Handle Or
           hwnd = frmHelp.Handle Or
           hwnd = frmSplash.Handle Then
            Exit Sub
        End If

        If eventType = EVENT_SYSTEM_FOREGROUND Then
            frmMain.ResizeToTargetWindow(hwnd)
        End If

    End Sub

    Public Function GetProcessNameForWindow(handle As IntPtr) As String
        Dim pid As Integer
        GetWindowThreadProcessId(handle, pid)
        Dim proc As Process
        Try
            proc = Process.GetProcessById(pid)
        Catch ex As Exception
            Return ""
        End Try

        Return proc.ProcessName
    End Function


    'Public Sub MonitorWindowChanges(handle As IntPtr)
    '    UnmonitorWindowChanges()

    '    ' Process ID as not sure
    '    Dim processId As UInteger
    '    GetWindowThreadProcessId(handle, processId)
    '    'handleWinEventHook(0) = SetWinEventHook(EVENT_OBJECT_SHOW, EVENT_OBJECT_HIDE, IntPtr.Zero, objWinDelegate, processId, 0, WINEVENT_OUTOFCONTEXT Or WINEVENT_SKIPOWNPROCESS)
    '    'handleWinEventHook(1) = SetWinEventHook(EVENT_OBJECT_IME_SHOW, EVENT_OBJECT_IME_HIDE, IntPtr.Zero, objWinDelegate, processId, 0, WINEVENT_OUTOFCONTEXT Or WINEVENT_SKIPOWNPROCESS)
    'End Sub

    'Public Sub UnmonitorWindowChanges()
    '    If handleWinEventHook.Length = 0 Then
    '        Exit Sub
    '    End If

    '    For Each hook As IntPtr In handleWinEventHook
    '        If hook <> Nothing Then
    '            UnhookWinEvent(hook)
    '        End If
    '    Next
    '    Array.Clear(handleWinEventHook, 0, handleWinEventHook.Length)
    'End Sub

    'Public Sub WinEventProc(hWinEventHook As IntPtr, eventType As UInteger, hwnd As IntPtr, idObject As Integer,
    '                            idChild As Integer, dwEventThread As UInteger, dwmsEventTime As UInteger)


    '    Select Case eventType
    '        Case EVENT_OBJECT_SHOW, EVENT_OBJECT_IME_SHOW
    '            Dim rectangle As EnumWindows.RECT
    '            EnumWindows.GetWindowRect(hwnd, rectangle)
    '            Dim winRectangle As New Windows.Rect(rectangle.Left, rectangle.Top,
    '                                                 rectangle.Right - rectangle.Left, rectangle.Bottom - rectangle.Top)
    '            frmMain.MakeCallout(hwnd, winRectangle, EnumWindows.GetWindowTextByHandle(hwnd))

    '        Case EVENT_OBJECT_HIDE, EVENT_OBJECT_IME_HIDE
    '            frmMain.RemoveCallout(hwnd)
    '    End Select
    'End Sub
End Module


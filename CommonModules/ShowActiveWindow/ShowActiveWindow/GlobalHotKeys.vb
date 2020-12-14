Imports System.Runtime.InteropServices
Imports System.Windows.Input

Module GlobalHotKeys
    Public Const WM_HOTKEY As Integer = &H312

    Enum KeyModifier
        None = 0
        Alt = &H1
        Control = &H2
        Shift = &H4
        Winkey = &H8
    End Enum

    Enum HotKeys
        Refresh = 0
    End Enum

    <DllImport("user32.dll", SetLastError:=True)>
    Private Function GetForegroundWindow() As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Function RegisterHotKey(ByVal handle As IntPtr, ByVal id As Integer, ByVal fsModifier As Integer, ByVal vk As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Function UnregisterHotKey(ByVal handle As IntPtr, ByVal id As Integer) As Integer
    End Function

    Public Sub RegisterHotkeys()
        RegisterHotKey(frmMain.Handle, HotKeys.Refresh, KeyModifier.Alt Or KeyModifier.Control Or KeyModifier.Shift, KeyInterop.VirtualKeyFromKey(Key.F5))
    End Sub

    Public Sub UnRegisterHotkeys()
        UnregisterHotKey(frmMain.Handle, HotKeys.Refresh)
    End Sub

    Public Sub HandleGlobalHotKeyEvent(ByVal hotkeyId As IntPtr)
        Select Case hotkeyId
            Case HotKeys.Refresh
                RepositionMainForm(GetForegroundWindow())
        End Select

    End Sub
End Module

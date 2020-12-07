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
        BringAppToForeground = 0
        ToggleStickyFlags = 1
        ShowHideFlags = 2
        RefreshFlags = 3
        FlagOptions = 4
    End Enum

    <DllImport("user32.dll")>
    Private Function RegisterHotKey(ByVal handle As IntPtr, ByVal id As Integer, ByVal fsModifier As Integer, ByVal vk As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Function UnregisterHotKey(ByVal handle As IntPtr, ByVal id As Integer) As Integer
    End Function

    Public Sub RegisterHotkeys()
        RegisterHotKey(frmMain.Handle, HotKeys.BringAppToForeground, KeyModifier.Alt Or KeyModifier.Control Or KeyModifier.Shift, KeyInterop.VirtualKeyFromKey(Key.F))
        RegisterHotKey(frmMain.Handle, HotKeys.FlagOptions, KeyModifier.Alt Or KeyModifier.Control Or KeyModifier.Shift, KeyInterop.VirtualKeyFromKey(Key.F2))
        RegisterHotKey(frmMain.Handle, HotKeys.ToggleStickyFlags, KeyModifier.Alt Or KeyModifier.Control Or KeyModifier.Shift, KeyInterop.VirtualKeyFromKey(Key.F3))
        RegisterHotKey(frmMain.Handle, HotKeys.ShowHideFlags, KeyModifier.Alt Or KeyModifier.Control Or KeyModifier.Shift, KeyInterop.VirtualKeyFromKey(Key.F4))
        RegisterHotKey(frmMain.Handle, HotKeys.RefreshFlags, KeyModifier.Alt Or KeyModifier.Control Or KeyModifier.Shift, KeyInterop.VirtualKeyFromKey(Key.F5))
    End Sub

    Public Sub UnRegisterHotkeys()
        UnregisterHotKey(frmMain.Handle, HotKeys.BringAppToForeground)
        UnregisterHotKey(frmMain.Handle, HotKeys.ToggleStickyFlags)
        UnregisterHotKey(frmMain.Handle, HotKeys.ShowHideFlags)
        UnregisterHotKey(frmMain.Handle, HotKeys.RefreshFlags)
        UnregisterHotKey(frmMain.Handle, HotKeys.FlagOptions)
    End Sub

    Public Sub HandleGlobalHotKeyEvent(ByVal hotkeyId As IntPtr)
        Select Case hotkeyId
            Case HotKeys.BringAppToForeground
                frmMain.BringToForeground()
            Case HotKeys.ToggleStickyFlags
                frmMain.HandleToggleStickySetting()

            Case HotKeys.ShowHideFlags
                frmMain.HandleToggleShowHideCallouts()

            Case HotKeys.RefreshFlags
                frmMain.RefreshCallouts()

            Case HotKeys.FlagOptions
                frmMain.ShowOptions()
        End Select

    End Sub
End Module

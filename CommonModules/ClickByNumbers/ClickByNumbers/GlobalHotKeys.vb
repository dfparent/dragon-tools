Imports System.Runtime.InteropServices
Imports System.Windows.Input
Imports Microsoft.Win32

Module GlobalHotKeys
    Public Const WM_HOTKEY As Integer = &H312
    Public Const DEFAULT_BRING_TO_FOREGROUND_HOTKEY As Key = Key.F1

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
        RegisterHotKey(frmMain.Handle, HotKeys.BringAppToForeground,
                       KeyModifier.Alt Or KeyModifier.Control Or KeyModifier.Shift,
                       KeyInterop.VirtualKeyFromKey(GetBringToForegroundHotkey()))
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

    Public Function GetBringToForegroundHotkey() As Key
        Dim key As RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey(REGISTRY_PATH_GLOBAL_SETTINGS, False)
        Dim converter As KeyConverter = New KeyConverter()

        If key Is Nothing Then
            ' Key does not exist.  Add it.
            key = My.Computer.Registry.CurrentUser.CreateSubKey(REGISTRY_PATH_GLOBAL_SETTINGS, True)
            If key Is Nothing Then
                MsgBox("Click By Numbers failed to create hot key registry key.")
                Return DEFAULT_BRING_TO_FOREGROUND_HOTKEY
            End If

            Try
                key.SetValue(REGISTRY_VALUE_BRING_TO_FOREGROUND_HOTKEY, converter.ConvertToString(DEFAULT_BRING_TO_FOREGROUND_HOTKEY))
                key.Close()
                key = My.Computer.Registry.CurrentUser.OpenSubKey(REGISTRY_PATH_GLOBAL_SETTINGS, False)
            Catch ex As Exception
                MsgBox("Click My Numbers failed to save new hotkey registry key.")
                Return DEFAULT_BRING_TO_FOREGROUND_HOTKEY
            End Try
        End If

        GetBringToForegroundHotkey = converter.ConvertFromString(key.GetValue(REGISTRY_VALUE_BRING_TO_FOREGROUND_HOTKEY, converter.ConvertToString(DEFAULT_BRING_TO_FOREGROUND_HOTKEY)))
        key.Close()
    End Function

    Public Function GetBringToForegroundHotKeyString() As String
        Dim converter As KeyConverter = New KeyConverter()
        Dim hotkey As Key
        hotkey = GetBringToForegroundHotkey()

        Return converter.ConvertToString(hotkey)
    End Function
End Module

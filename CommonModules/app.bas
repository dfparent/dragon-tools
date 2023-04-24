'#uses "imports.bas"
'#Uses "window.bas"
'#Language "WWB.NET"
Option Explicit On

Private activeWindowHandle As Long

'''''''''''''''''''''''''''''''''
''' THIS DOES NOT WORK

Public Sub SwitchToApp(switchText As String)

    activeWindowHandle = GetForegroundWindow()

    Try
        Dim myDelegate As EnumWindowsDelegateCallBack
        myDelegate = New EnumWindowsDelegateCallBack(AddressOf EnumWindowCallback)
        EnumWindows(myDelegate, 0)
    Catch ex As System.Exception
        MsgBox(ex.Message)
    End Try

End Sub


Private Function EnumWindowCallback(ByVal hwnd As Integer, ByVal lParam As Integer) As Boolean

    Dim dragonResultsBoxWindowHandle As System.IntPtr

    'working vars
    Dim className As New StringBuilder(255)
    Dim windowText As New StringBuilder(255)

    Try

        GetClassName(hwnd, className, 255)
        GetWindowText(hwnd, windowText, 255)

    Catch ex As System.Exception
        MsgBox(ex.Message)
        EnumWindowCallback = False
        Exit Function
    End Try

    EnumWindowCallback = True

End Function


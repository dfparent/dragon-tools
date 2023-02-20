Public Class frmMain


    Dim borderWidth As Single = 2  ' Even numbers work best
    Dim borderColor As Color = Color.Red
    Dim firstRefreshInterval As Integer = 1000  ' Refresh highlight after 1 second to compensate for some missed window switch events
    Dim secondRefreshInterval As Integer = 5000  ' Refresh highlight after 10 seconds to compensate for Some splash screens that 
    ' Go away but don't generate a window switch event

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializeSystemMonitor()
        RegisterHotkeys()
        timRefresh.Interval = firstRefreshInterval
    End Sub

    Private Sub frmMain_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        UninitializeSystemMonitor()
        UnRegisterHotkeys()
    End Sub

    Protected Overrides Sub OnPaintBackground(ByVal e As PaintEventArgs)
        MyBase.OnPaintBackground(e)

        Dim rect As New Rectangle(borderWidth / 2, borderWidth / 2, Me.ClientSize.Width - borderWidth, Me.ClientSize.Height - borderWidth)
        Dim thePen As New Pen(borderColor, borderWidth)

        e.Graphics.DrawRectangle(thePen, rect)
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles trayIcon.MouseDoubleClick

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        If m.Msg = GlobalHotKeys.WM_HOTKEY Then
            GlobalHotKeys.HandleGlobalHotKeyEvent(m.WParam)
        End If
        MyBase.WndProc(m)
    End Sub

    Private Declare Function GetForegroundWindow Lib "user32" () As Int32

    Public Sub InitializeTimer()
        timRefresh.Interval = firstRefreshInterval
        timRefresh.Start()
    End Sub

    Private Sub timRefresh_Tick(sender As Object, e As EventArgs) Handles timRefresh.Tick
        Dim hwnd As IntPtr
        hwnd = GetForegroundWindow()
        If hwnd <> GetHandleForeground() Then
            ClingToWindow(hwnd)
            timRefresh.Interval = firstRefreshInterval
        Else
            If timRefresh.Interval = firstRefreshInterval Then
                ' Set timer again but for longer interval
                timRefresh.Interval = secondRefreshInterval
            Else
                timRefresh.Stop()
            End If
        End If
    End Sub

    ' Trap "Alt+F4" which closes the app and prompt user if that is what he wants to do.
    ' Used to have this in the "form closeing" event, but then the system is prompted during shutdown and it halts shut down
    Private Sub frmMain_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.Alt And e.KeyCode = Keys.F4 Then
            If MsgBox("Do you want to close the Show Active Window app?", MsgBoxStyle.YesNo, MsgBoxStyle.Question) = MsgBoxResult.No Then
                e.Handled = True
            End If
        End If
    End Sub
End Class

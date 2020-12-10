'Imports System.Runtime.InteropServices

Public Class frmMain

    '<DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    'Private Shared Function GetForegroundWindow() As IntPtr
    'End Function

    Dim borderWidth As Single = 2  ' Even numbers work best
    Dim borderColor As Color = Color.Red

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializeSystemMonitor()
        RegisterHotkeys()

        'SendKeys.Send("%{tab}")
        'Dim handle As IntPtr
        'Handle = GetForegroundWindow()
        'WindowMonitor.RepositionMainForm(handle)
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

End Class

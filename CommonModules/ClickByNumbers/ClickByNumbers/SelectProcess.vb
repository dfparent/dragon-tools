Public Class frmProcess
    Private selectedProcess As WinProcess

    Private Sub frmProcess_Load(sender As Object, e As EventArgs) Handles Me.Load

        ' Populate unique process list
        Dim winProc As WinProcess
        Dim distinctProcesses As New Dictionary(Of String, WinProcess)
        Dim processes() As Process
        processes = Process.GetProcesses()

        For Each p As Process In processes
            If p.MainWindowHandle = IntPtr.Zero Then
                Continue For
            End If

            winProc = New WinProcess(p)
            Try
                distinctProcesses.Add(winProc.ProcessName, winProc)
            Catch ex As Exception
                ' Skip duplicates
            End Try

        Next

        cboProcess.Items.AddRange(distinctProcesses.Values.ToArray())
    End Sub

    Public Function GetSelectedProcess() As WinProcess
        Return selectedProcess
    End Function

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Hide()
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        selectedProcess = cboProcess.SelectedItem
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Hide()
    End Sub
End Class
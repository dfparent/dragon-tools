Public Class frmOptions
    Private Sub frmOptions_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        chkDisable.Checked = Not frmMain.GetCalloutsEnabled()
        chkSticky.Checked = frmMain.GetCalloutsSticky()
    End Sub

    Private Sub chkSticky_CheckedChanged(sender As Object, e As EventArgs) Handles chkSticky.CheckedChanged
        frmMain.SetCalloutsSticky(chkSticky.Checked)
    End Sub

    Private Sub chkDisable_CheckedChanged(sender As Object, e As EventArgs)
        frmMain.SetCalloutsEnabled(Not chkDisable.Checked)
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub
End Class
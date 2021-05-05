Imports System.Windows.Input

Public Class frmSplash
    Private Sub frmSplash_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblPrompt.Text = "For Help, press ctrl+shift+alt+" & GetBringToForegroundHotKeyString() & " and then F1."
        lblPrompt3.Text = "HKCU\" & REGISTRY_PATH_GLOBAL_SETTINGS
    End Sub

    Private Sub FlowLayoutPanel1_Paint(sender As Object, e As PaintEventArgs)

    End Sub
End Class
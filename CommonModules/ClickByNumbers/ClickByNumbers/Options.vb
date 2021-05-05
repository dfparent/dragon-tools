Imports Microsoft.Win32
Imports System.Runtime.InteropServices

Public Class frmOptions

    Private currentSettings As AppSettings

    Private Sub frmOptions_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MakeDirty(False)

        ' Load the names of all the processes for which we have settings, plus <default>
        cboApplyTo.Items.Clear()

        Dim key As RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey(REGISTRY_PATH_APP_SETTINGS)
        If key IsNot Nothing Then
            cboApplyTo.Items.AddRange(key.GetSubKeyNames())
            key.Close()
        End If

        If Not cboApplyTo.Items.Contains(DEFAULT_PROCESS_NAME) Then
            cboApplyTo.Items.Add(DEFAULT_PROCESS_NAME)
        End If

        ' Select settings for current target window, if possible
        Dim processName As String
        processName = GetProcessNameForWindow(frmMain.GetCurrentTargetWindowHandle())
        If processName <> "" Then
            If cboApplyTo.Items.Contains(processName) Then
                cboApplyTo.SelectedItem = processName
            Else
                cboApplyTo.SelectedItem = DEFAULT_PROCESS_NAME
            End If
        End If

    End Sub

    Private Sub cboApplyTo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboApplyTo.SelectedIndexChanged
        If IsDirty() Then
            If MsgBox("Do you want to save your changes?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                currentSettings.Save()
            End If
        End If
        LoadSettings(cboApplyTo.SelectedItem())
        MakeDirty(False)
    End Sub

    Private Sub LoadSettings(processName As String)
        currentSettings = New AppSettings()
        If Not currentSettings.Load(processName) Then
            MsgBox("Error loading settings for " & processName & ".")
            Exit Sub
        End If

        chkSticky.Checked = currentSettings.IsSticky
        chkDisable.Checked = currentSettings.IsDisabled
        updnOpacity.Value = currentSettings.Opacity
        chkUsesWindowHandles.Checked = currentSettings.UsesWindowHandleDiscovery
        chkUsesUIAutomation.Checked = currentSettings.UsesUIAutomationDiscovery
        chkUsesMSAA.Checked = currentSettings.UsesMSAADiscovery

    End Sub


    Private Sub chkSticky_CheckedChanged(sender As Object, e As EventArgs) Handles chkSticky.CheckedChanged
        currentSettings.IsSticky = chkSticky.Checked
        MakeDirty(True)
    End Sub

    Private Sub chkDisable_CheckedChanged(sender As Object, e As EventArgs) Handles chkDisable.CheckedChanged
        currentSettings.IsDisabled = chkDisable.Checked
        MakeDirty(True)
    End Sub

    Private Sub updnOpacity_ValueChanged(sender As Object, e As EventArgs) Handles updnOpacity.ValueChanged
        currentSettings.Opacity = updnOpacity.Value
        MakeDirty(True)
    End Sub

    Private Sub chkUsesWindowHandles_CheckedChanged(sender As Object, e As EventArgs) Handles chkUsesWindowHandles.CheckedChanged
        currentSettings.UsesWindowHandleDiscovery = chkUsesWindowHandles.Checked
        MakeDirty(True)
    End Sub

    Private Sub chkUsesUIAutomation_CheckedChanged(sender As Object, e As EventArgs) Handles chkUsesUIAutomation.CheckedChanged
        currentSettings.UsesUIAutomationDiscovery = chkUsesUIAutomation.Checked
        MakeDirty(True)
    End Sub

    Private Sub chkUsesMSAA_CheckedChanged(sender As Object, e As EventArgs) Handles chkUsesMSAA.CheckedChanged
        currentSettings.UsesMSAADiscovery = chkUsesMSAA.Checked
        MakeDirty(True)
    End Sub

    Private Sub MakeDirty(isDirty As Boolean)
        cmdSave.Enabled = isDirty
    End Sub

    Private Function IsDirty() As Boolean
        Return cmdSave.Enabled
    End Function

    Private Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
        currentSettings.Save()
        MakeDirty(False)
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        If IsDirty() Then
            If MsgBox("Do you want to save your changes?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                currentSettings.Save()
            End If
        End If
        Me.Close()
    End Sub


    Private Sub btnAddProcess_Click(sender As Object, e As EventArgs) Handles btnAddProcess.Click
        If frmProcess.Visible Then
            frmProcess.Focus()
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()
        frmProcess.ShowDialog()
        Me.Cursor = Cursors.Default
        If frmProcess.DialogResult = DialogResult.OK Then
            cboApplyTo.Items.Add(frmProcess.GetSelectedProcess().ProcessName)
            cboApplyTo.SelectedItem = frmProcess.GetSelectedProcess().ProcessName
        End If
    End Sub

    Private Sub btnRemoveProcess_Click(sender As Object, e As EventArgs) Handles btnRemoveProcess.Click
        If MsgBox("Remove the settings for " & cboApplyTo.SelectedItem & "?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            Exit Sub
        End If

        If currentSettings.Delete() Then
            cboApplyTo.Items.Remove(cboApplyTo.SelectedItem)
            cboApplyTo.SelectedItem = DEFAULT_PROCESS_NAME
        Else
            MsgBox("Error removing settings.")
        End If

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub
End Class
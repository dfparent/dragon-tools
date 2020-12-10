Imports System.ComponentModel

Public Class frmOpacity

    Private m_parent As Form
    Private m_Opacity As Double
    Private m_originalOpacity As Double

    Public Property GridOpacity As Double
        Get
            Return m_Opacity
        End Get
        Set(value As Double)
            m_Opacity = value
            updnOpacity.Value = value * 100
        End Set
    End Property

    Public Sub Init(parent As Form)
        m_parent = parent
        GridOpacity = m_parent.Opacity
        m_originalOpacity = m_parent.Opacity
        updnOpacity.Focus()
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        Me.DialogResult = DialogResult.OK
        Me.Hide()
    End Sub

    Private Sub updnOpacity_ValueChanged(sender As Object, e As EventArgs) Handles updnOpacity.ValueChanged
        m_Opacity = updnOpacity.Value / 100
        frmMain.Opacity = m_Opacity
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Hide()
    End Sub


    Private Sub frmOpacity_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If Me.DialogResult <> DialogResult.OK Then
            frmMain.Opacity = m_originalOpacity
        End If
    End Sub
End Class
Public Class frmPrompt

    Public Sub SetPrompt(text As String)
        lblPrompt.Text = text
        lblPrompt.Left = (Me.Width / 2) - lblPrompt.Width / 2
        lblPrompt.Top = (Me.Height / 2) - lblPrompt.Height / 2

    End Sub

    Public Function GetPrompt() As String
        Return lblPrompt.Text
    End Function
End Class
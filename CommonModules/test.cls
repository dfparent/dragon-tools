'#Language "WWB.NET"

Public Class Test
    Private m_String As String

    Public Sub New()
        m_String = "Hi"
    End Sub

    Public Sub Speak()
        Msgbox(m_String)
    End Sub
End Class
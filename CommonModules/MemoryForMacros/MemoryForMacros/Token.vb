Public Class Token
    Private m_name As String
    Private m_length As Integer

    Public Sub New(name As String)
        m_name = name
        m_length = name.Length
    End Sub

    Public ReadOnly Property Name As String
        Get
            Return m_name
        End Get
    End Property

    Public ReadOnly Property Length As Integer
        Get
            Return m_length
        End Get
    End Property
End Class

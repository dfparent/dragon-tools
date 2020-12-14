Public Class WinProcess
    Private p As Process
    Private mainHandle As IntPtr
    Private mainTitle As String
    Private pid As Integer
    Private pName As String

    Public Sub New(p As Process)
        Me.p = p
        mainHandle = p.MainWindowHandle
        mainTitle = p.MainWindowTitle
        pid = p.Id
        pName = p.ProcessName
    End Sub

    Public ReadOnly Property ProcessObj() As Process
        Get
            Return p
        End Get
    End Property

    Public ReadOnly Property ID() As Integer
        Get
            Return pid
        End Get
    End Property

    Public ReadOnly Property ProcessName() As String
        Get
            Return pName
        End Get
    End Property

    Public ReadOnly Property MainWindowHandle() As IntPtr
        Get
            Return mainHandle
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return pName & " (" & mainTitle & ")"
    End Function
End Class

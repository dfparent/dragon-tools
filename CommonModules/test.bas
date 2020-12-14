'#Uses "test.cls"
'#Language "WWB.NET"
Option Explicit On

Dim aString As String

Public Sub DoTest()
    Dim obj As Object
    obj = ModuleLoadThis("Public Class Test" & vbcrlf &
                         "    Private m_String As String" & vbcrlf &
                         "    Public Sub New()" & vbcrlf &
                         "        m_String = ""Hi""" & vbcrlf &
                         "    End Sub" & vbcrlf &
                         "    Public Sub Speak()" & vbcrlf &
                         "        Msgbox(m_String)" & vbcrlf &
                         "    End Sub" & vbcrlf &
                         "End Class")
    Dim aTest As New Test
    aTest.Speak()
End Sub

Public Sub DoDelayTest()

End Sub
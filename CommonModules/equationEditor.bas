'#Language "WWB.NET"
'#Uses "utilities.bas"
'option explicit

' If dictation contains "underscore", then use the following notation:
'   i_* for input files
'   e_* for equation files
'   o_* for output files
Public Sub OpenFile(dictation As String)
    Dim words() As String
    words = Split(dictation, "_")

    SendKeys("{F6}")

    If UBound(words) >= 1 Then
        Select Case words(0).ToLower()
            Case "echo", "e", "the"
                SendKeys("e_")
            Case "india", "i"
                SendKeys("i_")
            Case "oscar", "oh", "all"
                SendKeys("o_")
            Case Else
                SendKeys(words(0).ToLower)
                SendKeys("_")
        End Select

		if words(1) = "avoid" then
			words(1) = "void"
		end if
        SendKeys(ArrayToString(words, 1))

    Else
        SendKeys(dictation)
    End If

    SendKeys("~")

End Sub
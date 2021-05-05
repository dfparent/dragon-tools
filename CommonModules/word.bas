'#Uses "keyboard.bas"

'#Language "WWB.NET"
Option Explicit On

' If "dictation" Includes the word "through", Then This macro will find a block of text
' Starting with the text before "through" and ending with the text after "through".
' For example,If dictation is "this through that", then this will find
'   this Is lot of text before get around to saying that
'   this or that
'   this is a bit shorter than that
' This Macro requires the use of my WordUtilities addin
Public Sub SelectTextBlock(dictation As String, forward As Boolean)
    Dim parts(2) As String
    Dim index As Integer
    index = dictation.IndexOf(" through ")
    If index >= 0 Then
        parts(0) = dictation.Substring(0, index)
        parts(1) = dictation.Substring(index + 9)
    Else
        parts(0) = dictation
        parts(1) = ""
    End If

    ' Add regular expression to make the first char either upper or lower case
    Dim first As String
    first = parts(0).Substring(0, 1)
    first = "[" & first.ToLower() & "|" & first.ToUpper() & "]"
    parts(0) = first & parts(0).Substring(1)

    Dim findText As String
    If parts(1).Length > 0 Then
        findText = parts(0) & "*" & parts(1)
    Else
        findText = parts(0)
    End If

    'msgbox findText

    ' Run macro
    ShowKeyTips
    SendKeys("xmn")

    If forward Then
        SendKeys("n")
    Else
        SendKeys("c")
    End If

    Wait 0.1

    SendKeys(findText)
    SendKeys("~")

End Sub
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

    ' Add regular expression to make case insensitive.
    ' This will be used by the Word Find method which does not support both wildcards and case insensitive
    ' because it is not "normal" regular expressions but Microsoft's dumbed down version of it.
    ' Use Microsoft's notation to make each letter either uppercase or lowercase (i.e. [a|A])

    ' NOTE: Microsoft oftentimes complains that these Regular expressions are "too complicated".  Boo-hoo.  Cry baby.
    'parts(0) = MakeCaseInsensitiveRegEx(parts(0))
    'parts(1) = MakeCaseInsensitiveRegEx(parts(1))

    Dim findText As String
    If parts(1).Length > 0 Then
        findText = parts(0) & "*" & parts(1)
    Else
        findText = parts(0)
    End If

    'msgbox findText
    'Clipboard(findText)
    'Exit Sub


    ' Run macro
    ShowKeyTips
    SendKeys("xmn")

    If forward Then
        SendKeys("n")
    Else
        SendKeys("c")
    End If

    Wait(0.1)

    SendKeys(findText)
    SendKeys("~")

End Sub

' This can easily produce patterns that are considered too complex for Word
Public Function MakeCaseInsensitiveRegEx(text As String) As String
    ' Make each letter Case insensitive using Microsoft's Find Wildcard Notation (i.e. [a|A])
    ' So "this is text" becomes "[t|T][h|H][i|I][s|S] [i|I][s|S] [t|T][e|E][x|X][t|T]"
    Dim out As System.Text.StringBuilder
    out = New System.Text.StringBuilder()
    Dim aChar As Char
    Dim i As Integer
    For i = 0 To text.Length - 1
        aChar = text(i)
        If Char.IsSeparator(aChar) Then
            out.Append("[").Append(aChar).Append("]")
        ElseIf Char.IsLetter(aChar) Then
            out.Append("[").Append(Char.ToLower(aChar)).Append("|").Append(Char.ToUpper(aChar)).Append("]")
        Else
            out.Append(aChar)
        End If
    Next

    Return out.ToString()
End Function
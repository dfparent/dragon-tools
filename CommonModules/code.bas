'#Uses "utilities.bas"
'#Uses "window.bas"
'#Language "WWB.NET"
'Option Explicit On

Imports System
Imports Microsoft.VisualBasic
Private Const REGISTRY_CODING_LANGUAGE = "CodingLanguage"
Private Const LANGUAGE_VB = "VB"
Private Const LANGUAGE_CSHARP = "CSharp"

Public Enum Language
    VB
    CSHARP
End Enum

''''''''''''''''''''''''''''''''''''''''''''''
' Output code
''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoCodeBlock(ListVar1 As String, Optional injectSelection As Boolean = False)
    Dim titles(1) As String
    titles(0) = "Visual Basic"
    titles(1) = "KnowBrainer Command Editor"

    Dim selectedText As String

    If CheckWindowText(titles) Then
        If injectSelection Then
            SendKeys("^x")
            selectedText = GetClipboard()
        Else
            selectedText = ""
        End If

        If ListVar1 = "if" Then
            SendKeys("if 1 then~~End If{Up}")
            InjectCode(selectedText)
            SendKeys("{Up}{Home}{Right 3}+{Right}")
        ElseIf ListVar1 = "select" Then
            SendKeys "select case ~~end select{up}"
            InjectCode(selectedText)
            SendKeys "{up}{end}"
        ElseIf ListVar1 = "for" Then
            SendKeys "for i = 1 to "
            SendKeys "{Enter 2}next i{Up}"
            InjectCode(selectedText)
            SendKeys "{up}{End}"
        ElseIf ListVar1 = "for each" Then
            SendKeys "for each x in y"
            SendKeys "~~next~{Up 2}"
            InjectCode(selectedText)
            SendKeys("{Up}{Home}^{Right 2}+{Right}")
        ElseIf ListVar1 = "do while" Then
            SendKeys("do while 1~~loop~{Up 2}")
            InjectCode(selectedText)
            SendKeys("{Up}{Home}^{Right 2}+{Right}")
        ElseIf ListVar1 = "do until" Then
            SendKeys("do until 1~~loop~{Up 2}")
            InjectCode(selectedText)
            SendKeys("{Up }{Home}^{Right 2}+{Right}")
        ElseIf ListVar1 = "do loop" Then
            SendKeys("do~~loop ~{Up 2}")
            InjectCode(selectedText)
        ElseIf ListVar1 = "do loop until" Then
            SendKeys("do~~loop until 1~{Up 2}")
            InjectCode(selectedText)
        ElseIf ListVar1 = "do loop while" Then
            SendKeys("do~~loop while 1~{Up 2}")
            InjectCode(selectedText)
        ElseIf ListVar1 = "with" Then
            SendKeys("With Name~~End With~{Up 2}")
            InjectCode(selectedText)
            SendKeys("{Up}{End}^+{Left}")
        Else
            MsgBox("Code block '" & ListVar1 & "' not yet implemented.")
        End If
    Else
        ' Dev studio
        Dim lang As Language
        lang = GetCodingLanguage()

        If injectSelection Then
            SendKeys("^c")
            selectedText = GetClipboard()
        Else
            selectedText = ""
        End If

        If ListVar1 = "if" Then
            If injectSelection Then
                SendKeys("^k^s")
            End If
            SendKeys("if")
            Wait(0.1)
            SendKeys("{tab 2}")
        ElseIf ListVar1 = "select" Or ListVar1 = "switch" Then
            If injectSelection Then
                SendKeys("^k^s")
            End If
            SendKeys("switch")
            Wait(0.1)
            SendKeys("{Tab 2}")
        ElseIf ListVar1 = "for" Then
            If injectSelection Then
                SendKeys("^k^s")
            End If
            SendKeys("for")
            Wait(0.1)
            SendKeys("{Tab 2}")
        ElseIf ListVar1 = "for each" Then
            If injectSelection Then
                SendKeys("^k^s")
            End If
            Select Case lang
                Case Language.VB
                    SendKeys("for each x in y")
                    SendKeys("~~next~{Up 3}{Home}^{Right 2}+{Right}")
                Case Language.CSHARP
                    SendKeys("foreach")
                    Wait(0.1)
                    SendKeys("{tab 2}")
                Case Else
            End Select
        ElseIf ListVar1 = "do while" Then
            If injectSelection Then
                SendKeys("^k^s")
            End If
            Select Case lang
                Case Language.VB
                    SendKeys("do while {(}1{)}~~loop~{Up 3}^{Right 2}{Right}+{Right}")
                Case Language.CSHARP
                    SendKeys("do")
                    Wait(0.1)
                    SendKeys("{tab 2}")
                Case Else
            End Select
        ElseIf ListVar1 = "do until" Then
            If injectSelection Then
                SendKeys("^k^s")
            End If
            If lang = Language.VB Then SendKeys "do until {(}1{)}~~loop~{Up 3}^{Right 2}{Right}+{Right}"
        ElseIf ListVar1 = "do loop" Then
            If injectSelection Then
                SendKeys("^k^s")
            End If
            If lang = Language.VB Then SendKeys("do~~loop ~{Up}{End}")
        ElseIf ListVar1 = "do loop until" Then
            If injectSelection Then
                SendKeys("^k^s")
            End If
            If lang = Language.VB Then SendKeys("do~~loop until 1~{Up}{End}^+{Left}")
        ElseIf ListVar1 = "do loop while" Then
            If injectSelection Then
                SendKeys("^k^s")
            End If
            If lang = Language.VB Then SendKeys("do~~loop while 1~{Up}{End}^+{Left}")
        ElseIf ListVar1 = "while" Then
            If injectSelection Then
                SendKeys("^k^s")
            End If
            If lang = Language.CSHARP Then
                SendKeys("while")
                Wait(0.1)
                SendKeys("{tab 2}")
            End If
        ElseIf ListVar1 = "try" Then
            If injectSelection Then
                SendKeys("^k^s")
            End If
            Select Case lang
                Case Language.VB
                    SendKeys("try{Tab}{Enter}")
                Case Language.CSHARP
                    SendKeys("try")
                    Wait(0.1)
                    SendKeys("{tab 2}")
            End Select
        ElseIf ListVar1 = "catch" Then
            SendKeys("^x")
            SendKeys("catch {(}{)} {{}{Enter}")
            InjectCode(selectedText)
            SendKeys("{Up 2}{End}{Left}")
        ElseIf ListVar1 = "finally" Then
            SendKeys("^x")
            SendKeys("finally {{}{Enter}")
            InjectCode(selectedText)
        ElseIf ListVar1 = "class" Then
            If injectSelection Then
                SendKeys("^k^s")
            End If
            SendKeys("class")
            Wait(0.1)
            SendKeys("{Tab 2}")
        ElseIf ListVar1 = "namespace" Then
            If injectSelection Then
                SendKeys("^k^s")
            End If
            SendKeys("namespace")
            Wait(0.1)
            SendKeys("{Tab 2}")
        ElseIf ListVar1 = "using" Then
            If injectSelection Then
                SendKeys("^k^s")
            End If
            SendKeys("using")
            Wait(0.1)
            SendKeys("{Tab 2}")
        ElseIf ListVar1 = "property" Then
            SendKeys("prop{Tab 2}")

        ElseIf ListVar1 = "full property" Then
            SendKeys("public void Name {{}~")
            SendKeys("get {{}~")
            SendKeys("{Down}~set {{}~")
            SendKeys("{Up 8}{End}+^{Left}")

        Else
            MsgBox("Code block '" & ListVar1 & "' not yet implemented.")
        End If
    End If
End Sub

Private Sub InjectCode(code As String)

    SendKeys(code)

    ' Move cursor back up to top line of injected code
    Dim numLines As Integer
    numLines = 1
    Dim index As Integer
    index = -1
    Do
        index = code.IndexOf(vblf)
        If index >= 0 Then
            numLines = numLines + 1
        End If

    Loop While index >= 0

    SendKeys("{up " & CStr(numLines - 1) & "}")

End Sub

Public Sub DoCodeStatement(ListVar1 As String, Optional dictation As String = "")
    Dim lang As Language
    lang = GetCodingLanguage()

    Select Case LCase(ListVar1)
        Case "comment", "comment line"
            Select Case lang
                Case Language.VB
                    SendKeys("{'} ")
                Case Language.CSHARP
                    SendKeys("{/}{/} ")
            End Select
        Case "if"
            Select Case lang
                Case Language.VB
                    SendKeys("If " & dictation)
                Case Language.CSHARP
                    SendKeys("if {(}")
            End Select
        Case "then"
            If lang = Language.VB Then SendKeys(" Then ")
        Case "else"
            Select Case lang
                Case Language.VB
                    SendKeys("Else~")
                Case Language.CSHARP
                    SendKeys("else{tab 2}")
            End Select
        Case "else if"
            Select Case lang
                Case Language.VB
                    SendKeys("ElseIf " & dictation)
                Case Language.CSHARP
                    SendKeys("else if{tab 2}")
            End Select

        Case "end if"
            If lang = Language.VB Then SendKeys("End If")
        Case "case"
			select case lang
				case Language.VB
					If IsNumeric(dictation) Then
						SendKeys("Case " & dictation & "~")
					ElseIf LCase(dictation) = "else" Then
						SendKeys("Case Else~")
					ElseIf dictation = "" Then
						SendKeys("Case """"{Left}")
					Else
						SendKeys("Case """"{Left}" & dictation)
						SendKeys("{End}~")
					End If
				case Language.CSHARP
					If IsNumeric(dictation) Then
                        SendKeys("Case " & dictation & "{:}~")
                    ElseIf LCase(dictation) = "default" Then
                        SendKeys("default{:}~")
                    ElseIf dictation = "" Then
                        SendKeys("Case """"{:}{Left 2}")
                    Else
                        SendKeys("Case """ & dictation & """{:}~")
                    End If
			end select
            If lang = Language.VB Then
                
            End If
        Case "go to"
            SendKeys("GoTo " & formatCapsNoSpaces(dictation, True))

        Case "or"
            Select Case lang
                Case Language.VB
                    SendKeys(" Or ")
                Case Language.CSHARP
                    SendKeys(" || ")
            End Select
        Case "and"
            Select Case lang
                Case Language.VB
                    SendKeys(" And ")
                Case Language.CSHARP
                    SendKeys(" && ")
            End Select
        Case "not"
            Select Case lang
                Case Language.VB
                    SendKeys(" Not ")
                Case Language.CSHARP
                    SendKeys(" !")
            End Select
        Case "throw"
            If dictation = "" Or LCase(dictation) = "new exception" Then
                dictation = "exception"
            End If
            SendKeys("throw new " & formatCapsNoSpaces(dictation, True) & "{Tab}{(}{""}{""}{Left}")
        Case "finally"
            SendKeys("finally~{{}~")
    End Select
End Sub

Public Sub DoMessageBox(Optional dictation As String = "")
    Dim lang As Language
    lang = GetCodingLanguage()
    Select Case lang
        Case Language.VB
            SendKeys("MsgBox{(}""")
            If dictation <> "" Then
                SendKeys(dictation)
            End If
            SendKeys("""{)}{Left 2}")
        Case Language.CSHARP
            SendKeys("MessageBox.Show{(}""")
            If dictation <> "" Then
                SendKeys(dictation)
            End If
            SendKeys("""{)}{;}{Left 3}")
        Case Else

    End Select
End Sub
Public Sub DoLogicalOperator(theOperator As String, Optional dictation As String = "")
    Dim lang As Language
    lang = GetCodingLanguage()

    Select Case theOperator
        Case "equals", "equal to"
            SendKeys " = "
        Case "not equals", "not equal to"
            Select Case lang
                Case Language.VB
                    SendKeys " <> "
                Case Language.CSHARP
                    SendKeys(" != ")
            End Select
        Case "equal equals"
            SendKeys(" == ")
        Case "greater than"
            SendKeys " > "
        Case "greater than or equal to"
            SendKeys " >= "
        Case "less than"
            SendKeys " < "
        Case "less than or equal to"
            SendKeys " <= "
    End Select

    If LCase(dictation) = "true" Or LCase(dictation) = "false" Then
        SendKeys(dictation)
        SendKeys("{Tab}")
    ElseIf lcase(dictation) = "empty string" Then
        SendKeys("{""}{""}")
    ElseIf lcase(dictation) = "no" Then
        SendKeys("null{Escape}")
    Else
        SendKeys(formatCapsNoSpaces(dictation, False))
        SendKeys("^{Space}")
    End If
End Sub

Public Sub DoMathOperator(ListVar1 As String, Optional number As String = "")
    SendKeys(" {" & Left(ListVar1, 1) & "} ")
    SendKeys(number)
End Sub

Public Sub DoMathAssignment(ListVar1 As String)
    SendKeys(" {" & ListVar1 & "}{=} ")
End Sub

public sub DoArrayIndexEmpty()
    Dim lang As Language
    lang = GetCodingLanguage()
    Select Case lang
        Case Language.VB
            SendKeys("{(}{)}")
        Case Language.CSHARP
            SendKeys("{[}{]}{Left}")
    End Select

end sub

Public Sub DoArrayIndex(index As Integer)
    Dim lang As Language
    lang = GetCodingLanguage()
    Select Case lang
        Case Language.VB
            SendKeys("{(}" & CStr(index) & "{)}")
        Case Language.CSHARP
            SendKeys("{[}" & CStr(index) & "{]}")
    End Select

End Sub

Public Sub DoArrayIndexWithVar(variable As String)
    Dim lang As Language
    lang = GetCodingLanguage()
    Select Case lang
        Case Language.VB
            SendKeys("{(}" & formatCapsNoSpaces(variable, False) & "{)}")
        Case Language.CSHARP
            SendKeys("{[}" & formatCapsNoSpaces(variable, False) & "{]}")
    End Select

End Sub

Public Sub DoDeclare(scope As String, dictation As String)
    If scope = "local" Then
        scope = ""
    End If
    SendKeys(scope & " " & formatCapsNoSpaces(dictation, False))
End Sub

Public Sub DoDefineVariableName(dictation As String)
    SendKeys("{Tab}{Space}")
    SendKeys(formatCapsNoSpaces(dictation, False))
End Sub

Public Sub DoSet(ListVar1 As String)
    SendKeys("set " & formatCapsNoSpaces(ListVar1, False) & "^{Space}")
End Sub

Public Sub DoOnError(dictation As String)
    SendKeys("On Error ")
    dictation = dictation.ToLower()
    If dictation = "resume next" Then
        SendKeys("Resume Next")
    ElseIf dictation.StartsWith("go to") Then
        Dim goToLocation As String
        goToLocation = dictation.Substring("go to".Length)
        If goToLocation.Trim = "0" Then
            SendKeys("GoTo 0")
        Else
            SendKeys("GoTo ")
            SendKeys(formatCapsNoSpaces(goToLocation, True))
        End If
    End If
End Sub

Public Sub DoSymbol(dictation As String)
    SendKeys(formatCapsNoSpaces(dictation, False))
    SendKeys("^{Space}")
End Sub

Public Sub DoAs(Dictation As String)
    Dim words() As String

    If Dictation.ToLower.StartsWith("new ") Then
        SendKeys(" as New ")
        Dictation = Dictation.Substring(4)
    Else
        SendKeys(" as ")
    End If

    SendKeys formatCapsNoSpaces(Dictation, True) 
    SendKeys "^{Space}"
End Sub

Public Sub DoAsSpell(ListVar1 As String, Optional ListVar2 As String = "", Optional ListVar3 As String = "")
    SendKeys "as " & Left(ListVar1, 1) & Left(ListVar2, 1) & Left(ListVar3, 1) & "^{Space}"
End Sub

Public Sub DoDot(Optional Dictation As String = "")
    SendKeys "." & formatCapsNoSpaces(Dictation, False) 
End Sub

Public Sub DoDotSpell(ListVar1 As String, Optional ListVar2 As String = "", Optional ListVar3 As String = "")
    SendKeys "." & Left$(ListVar1, 1) & Left$(ListVar2, 1) & Left$(ListVar3, 1)
End Sub

Public Sub DoCodeProcedureWithName(scope As String, procedure As String, Optional dictation As String = "")
    Dim lang As Language
    lang = GetCodingLanguage()

    If lang = Language.CSHARP Then
        SendKeys(scope & " void ")
        If dictation <> "" Then
            SendKeys(formatCapsNoSpaces(dictation, True))
        Else
            SendKeys("Name")
        End If

        SendKeys("{(}{)}~{{}~")

        If dictation = "" Then
            SendKeys("{Up 2}{Home}^{right 2}^+{Right}")
        Else
            SendKeys("{Up 2}{End}^{Left}")
        End If

    Else
        SendKeys(scope & "{Space}" & procedure & " ")
        If dictation <> "" Then
            SendKeys formatCapsNoSpaces(dictation, True)
        Else
            SendKeys("Name")
        End If

        If procedure = "property let" Then
            SendKeys("{(}value{)}")
        End If

        SendKeys("~{Up}{End}")

        If dictation = "" Then
            SendKeys("{Left 2}^+{Left}")
        Else
            SendKeys("{Left}")
        End If
    End If

End Sub

Public Sub DoCompleteLine()
    SendKeys("{End};~")
End Sub

Public Sub DoConcatenate()
    Dim lang As Language
    lang = GetCodingLanguage()

    Select Case lang
        Case Language.VB
            SendKeys(" & ")
        Case Language.CSHARP
            SendKeys(" {+} ")
        Case Else

    End Select

End Sub

Public Sub DoDoubleQuotes()
    ' Sets up an embedded quote pair in a VB string (escaped with double quotes)
    SendKeys " {""}{""}{""}{""} "
    SendKeys "{Left 3}"
End Sub

Public Sub DoDoubleQuote()
    ' Sets up an embedded quote in a VB string (escaped with double quotes)
    SendKeys "{""}{""}"
End Sub

Public Sub DoTripleQuote()
    ' Sets up an embedded quote in a VB string (escaped with double quotes)
    SendKeys "{""}{""}{""}"
End Sub

Public Sub DoCaseString(theString As String)
    SendKeys "case """ & theString & """~{tab}"
End Sub

Public Sub DoConvert(ListVar1 As String)
    Select Case ListVar1
        Case "string"
            SendKeys "CStr{(}"

    Case "long"
            SendKeys "CLng{(}"

    Case "double"
            SendKeys "CDbl{(}"

    Case "float"
            SendKeys "CSng{(}"

    Case "date"
            SendKeys "CDate{(}"

    Case Else
            Beep

    End Select
End Sub

Public Sub DoTabOpenParen()
    SendKeys "{Tab}{(}"
End Sub

Public Sub DoTabDot()
    SendKeys "{Tab}{.}"
End Sub

Public Sub DoIncrementThat()
    SendKeys("^{Left}^+{Right}^c")
    Dim theVariable As String
    theVariable = GetClipboard()
    SendKeys("{Right}")
    SendKeys(" = ")
    SendKeys(theVariable)
    SendKeys(" {+} 1")
End Sub

''''''''''''''''''''''''''''''''''''''''''''''
' Manipulate environment
''''''''''''''''''''''''''''''''''''''''''''''

Public Sub DoRenameThat()
    Dim titles(0) As String
    titles(0) = "Microsoft Visual Studio"
    If CheckWindowText(titles) Then
        SendKeys "+{F10}"
        Wait 0.1
        SendKeys "r"
    Else
        ' VBA
        SendKeys("^h")
        Wait(0.1)
        SendKeys("^c{Tab}^v{Home}")
    End If
End Sub

Public Sub DoGoToLine(ListVar1 As String)
    SendKeys "%egl"
    Wait 0.1
    SendKeys ListVar1
    SendKeys "{Enter}"
End Sub

Public Sub DoListMembers()
    SendKeys "^j"
End Sub

Public Sub DoParameterInfo()
    SendKeys "^+i"
End Sub

Public Sub DoQuickInfo()
    SendKeys "^i"
End Sub

Public Sub DoCompleteWord()
    SendKeys "^{Space}"
End Sub

Public Sub DoGoToDefinition()
    SendKeys "+{F2}"
End Sub

Public Sub DoPeekDefinition()
    SendKeys "%{F12}"
End Sub

Public Sub DoStartDebugging()
    SendKeys "{F5}"
End Sub

Public Sub DoStopDebugging()
    SendKeys "%de"
End Sub

Public Sub DoErrorList()
    SendKeys "^w^e"
End Sub

Public Sub DoGoToCode()
    SendKeys "{F7}"
End Sub

Public Sub DoFindReferences()
    SendKeys "%{F2}"
End Sub

Public Sub DoSelectBlock()
    SendKeys "%+{]}"
End Sub

Public Sub DoCollapseAll()
    SendKeys "^m^o"
End Sub

Public Sub DoExpandAll()
    SendKeys "^m^l"
End Sub

Public Sub DoGrowShrinkCurrent()
    SendKeys "^m^m"
End Sub

Public Sub DoGoTo(ListVar1 As String, ListVar2 As String)
    Dim titles(0) As String
    titles(0) = "Microsoft Visual Studio"
    If CheckWindowText(titles) Then
        Select Case ListVar1
            Case "procedure", "member"
                SendKeys "%{\}"
        Case "symbol"
                SendKeys "^1^s"
        Case "type"
                SendKeys "^1^t"
        Case "file"
                SendKeys "^1^r"
        End Select

        Wait 0.1
        SendKeys ListVar2
    Else
        ' VBA
        SendKeys("^f")
        Wait(0.1)
        SendKeys("{Delete}" & formatCapsNoSpaces(ListVar2))
        SendKeys("%m%n")
        Wait(0.1)
        SendKeys("{Escape}")
    End If

End Sub

Public Sub DoExtendSelection()
    SendKeys "%+{=}"
End Sub

Public Sub DoSetNextStatement()
    SendKeys("+{F10}n")
End Sub

Public Sub DoSetCodingLanguage(theLanguage As String)
    Select Case theLanguage
        Case "Visual Basic", "VB"
            SaveValue(REGISTRY_CODING_LANGUAGE, LANGUAGE_VB)
            TTSPlayString("Coding in VB.")
        Case "C#"
            SaveValue(REGISTRY_CODING_LANGUAGE, LANGUAGE_CSHARP)
            TTSPlayString("Coding in C sharp.")
        Case Else
            MsgBox("The """ & theLanguage & """ language is not currently supported.")
            Exit Sub
    End Select

End Sub

Public Function GetCodingLanguage() As Language
    Dim theLanguageStr As String
    theLanguageStr = GetValue(REGISTRY_CODING_LANGUAGE)
    If theLanguageStr = "" Then
        theLanguageStr = LANGUAGE_VB
    End If
    Select Case theLanguageStr
        Case LANGUAGE_VB
            Return Language.VB
        Case LANGUAGE_CSHARP
            Return Language.CSHARP
    End Select
End Function
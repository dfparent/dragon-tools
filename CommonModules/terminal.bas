'#Uses "utilities.bas"
'#Uses "paths.bas"
'#Uses "cache.bas"
'#Language "WWB.NET"
Option Explicit On

' Some macros in this file require the following function to be created in the ~/.bashrc (or some similar) file:
'   
'cdd()
'{
'  cd "$@" && ls;
'}

Private Const TERMINAL_TYPE_KEY = "TerminalType"
Private Const TERMINAL_TYPE_CYGWIN = "Cygwin"
Private Const TERMINAL_TYPE_OTHER = "Other"

Public Sub doFileCommand(command As String)
    ButtonClick(1, 2)
    SendKeys("~")
    SendKeys command
    SendKeys " +{ins}"
End Sub

Public Function translateCommand(command As String)
    command = Split(command, "\")(0)
    Select Case command
        Case "edit"
            Return "vim"
        Case "delete"
            Return "rm"
        Case "tail"
            Return "tail -f"
        Case "copy"
            Return "cp"
        Case "move"
            Return "mv"
        Case "rename"
            Return "mv"
        Case "system control start"
            Return "systemctl start"
        Case "system control stop"
            Return "systemctl stop"
        Case "system control restart"
            Return "systemctl restart"
        Case "system control status"
            Return "systemctl status"
        Case "sytem control enable"
            Return "systemctl enable"
        Case "system control"
            Return "systemctl"
        Case "make directory"
            Return "mkdir"
        Case Else
            Return command
    End Select
End Function

Public Sub SetTerminalTypeToCygwin(value As Boolean)
    If value Then
        ' Save in registry
        SaveValue(TERMINAL_TYPE_KEY, TERMINAL_TYPE_CYGWIN)
        UpdateCacheValue(TERMINAL_TYPE_KEY, {TERMINAL_TYPE_CYGWIN})
        TTSPlayString("Terminal type sig win.")
    Else
        SaveValue(TERMINAL_TYPE_KEY, TERMINAL_TYPE_OTHER)
        UpdateCacheValue(TERMINAL_TYPE_KEY, {TERMINAL_TYPE_OTHER})
        TTSPlayString("Terminal type other.")
    End If
End Sub

Public Function IsTerminalTypeCygwin() As Boolean
    Dim terminalTypes() As String

    If Not CacheValueExists(TERMINAL_TYPE_KEY) Then
        Dim terminalType As String
        terminalType = GetValue(TERMINAL_TYPE_KEY)
        If terminalType = "" Then
            terminalType = TERMINAL_TYPE_OTHER
        End If
        UpdateCacheValue(TERMINAL_TYPE_KEY, {terminalType})
    End If

    terminalTypes = GetCacheValue(TERMINAL_TYPE_KEY)

    If terminalTypes(0) = TERMINAL_TYPE_CYGWIN Then
        Return True
    Else
        Return False
    End If
End Function

Public Function ConvertPathToCygwin(path As String) As String
    Dim newPath As String
    newPath = path
    If Not IsTerminalTypeCygwin() Then
        Return newPath
    End If

    newPath = newPath.Trim()

    Dim index As Integer
    index = newPath.IndexOf(":")
    If index >= 0 Then
        ' replace drive with cygdrive
        Dim drive As String
        drive = newPath.Substring(0, index)
        newPath = "/cygdrive/" & drive & "/" & newPath.Substring(index + 1)
    End If

    ' Replace backslash with slash
    newPath = newPath.Replace("\", "/")

    Return newPath
End Function

Public Sub DoClearLine()
    SendKeys "^u"
End Sub

Public Sub DoStartOfLine()
    SendKeys "^a"
End Sub

Public Sub DoEndOfLine()
    SendKeys "^e"
End Sub

Public Sub DoCancel()
    SendKeys "^c"
End Sub


Public Sub DoComplete(ListVar1 As String)
    SendKeys formatNoCapsNoSpaces(ListVar1)
    SendKeys "{tab}"
End Sub

Public Sub DoClearToEnd()
    SendKeys "^k"
End Sub

Public Sub DoGoThere()
    doFileCommand("cdd")
    SendKeys "~"
End Sub

Public Sub DoSub(ListVar1 As String)
    SendKeys "cdd "
    'SendKeys formatNoCapsNoSpaces(ListVar1)
    SendKeys(ListVar1)
    SendKeys "{tab}"
End Sub

Public Sub DoSubSpell(ListVar1 As String, ListVar2 As String, ListVar3 As String)
    SendKeys "cdd " & split(ListVar1, "\")(0) & split(ListVar2, "\")(0) & split(ListVar3, "\")(0)
    SendKeys "{tab}"
End Sub

Public Sub DoGoUp(Optional ListVar1 As String = "1")
    SendKeys "cdd "
    Dim i As Integer
    For i = 1 To CInt(ListVar1)
        SendKeys "../"
    Next i
    SendKeys "{enter}"
End Sub

Public Sub DoGoBack()
    SendKeys "cdd -"
    SendKeys "{enter}"
End Sub

Public Sub DoFolderDollar(ListVar1 As String)
    SendKeys "cdd $"
    SendKeys formatNoCapsNoSpaces(ListVar1)
End Sub

Public Sub GoHome()
    SendKeys "cdd {~}"
    SendKeys "{enter}"
End Sub

Public Sub DoScroll(ListVar1 As String, Optional ListVar2 As String = "40")
    Select Case LCase(ListVar1)
        Case "up"
            SendKeys("+{Up " & ListVar2 & "}")
        Case "down"
            SendKeys("+{Down " & ListVar2 & "}")
    End Select

End Sub

Public Sub DoList()
    SendKeys "ll~"
End Sub

Public Sub DoCopyPaste()
    ButtonClick(1, 2)
    ButtonClick(2, 1)

End Sub

Public Sub DoExitShells(ListVar1 As String)
    Dim i As Integer
    For i = 1 To CInt(ListVar1)
        SendKeys "exit~"
    Next i
End Sub

Public Sub DoPipe()
    SendKeys " | "
End Sub

Public Sub DoRecursiveGrep()
    SendKeys "grep -r --include="""" "
    SendKeys "{Left 2}"
End Sub

Public Sub DoPsGrep(ListVar1 As String)
    SendKeys "ps -ef | grep " & ListVar1
End Sub

Public Sub DoLinuxCommandThat(ListVar1 As String)
    doFileCommand(translateCommand(ListVar1))
    Select Case ListVar1
        Case "delete", "move", "rename"
            ' Do not auto press enter

        Case Else
            SendKeys "{enter}"
    End Select
End Sub

Public Sub DoLinuxCommand(ListVar1 As String)
    SendKeys translateCommand(ListVar1)
    SendKeys " "
End Sub

Public Sub DoLinuxCommandWithParameter(ListVar1 As String, ListVar2 As String)
    SendKeys translateCommand(ListVar1)
    SendKeys " "
    SendKeys formatNoCapsNoSpaces(ListVar2)
    SendKeys "{Tab}"
End Sub

Public Sub DoHyphen(ListVar1 As String, Optional ListVar2 As String = "", Optional ListVar3 As String = "")
    SendKeys " -"
    SendKeys Split(ListVar1, "\")(0)
    If ListVar2 <> "" Then
        SendKeys Split(ListVar2, "\")(0)
    End If
    If ListVar3 <> "" Then
        SendKeys Split(ListVar3, "\")(0)
    End If
End Sub

Public Sub DoDashDash(ListVar1 As String)
    SendKeys " --"
    SendKeys formatNoCapsNoSpaces(ListVar1)
End Sub

Public Sub DoPipeCommand(ListVar1 As String)
    SendKeys " | "
    SendKeys translateCommand ListVar1
End Sub

Public Sub DoFolderLinuxCommand(ListVar1 As String)
    SendKeys "cdd "
    Dim thePath As String
    thePath = ConvertPathToCygwin(getPath(ListVar1))
    If thePath <> "" Then
        SendKeys thePath
        SendKeys "{Enter}"
    Else
        SendKeys formatNoCapsNoSpaces(ListVar1)
    End If
End Sub

Public Sub DoSudoRoot()
    SendKeys("sudo su - ~")
End Sub

Public Sub DoFolder(ListVar1 As String)

    SendKeys("cdd ")
    SendKeys(DictationToKeystrokes(ListVar1))
End Sub

Public Sub DoListMore()
    SendKeys("ll | more~")
End Sub

Public Sub DoSudoCommand(ListVar1 As String, Optional ListVar2 As String = "")
    SendKeys "sudo "
    SendKeys translateCommand(Split(ListVar1)(0))
    SendKeys " "
    If ListVar2 <> "" Then
        SendKeys formatNoCapsNoSpaces(ListVar2)
        SendKeys "{Tab}"
    End If
End Sub

Public Sub DoPaste()
    SendKeys "+{ins}"
End Sub

Public Sub DoEnterList()
    SendKeys("~ll~")
End Sub

Public Sub DoPage(ListVar1 As String, Optional ListVar2 As String = "1")
    Select Case LCase(ListVar1)
        Case "up"
            SendKeys("+{PgUp " & ListVar2 & "}")
        Case "down"
            SendKeys("+{PgDn " & ListVar2 & "}")
    End Select
End Sub




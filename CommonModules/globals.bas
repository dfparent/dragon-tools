'#Language "WWB.NET"
'#Uses "utilities.bas"
'#Uses "window.bas"
'#Uses "imports.bas"
'#Uses "cache.bas"
'option explicit

''' I have lost way too much work in macros from Knowbrainer's stupid save bug, that I am putting any macros
''' that have an appreciable amount of code in them in external files which KnowBrainer WON'T lose.
''' 

Public Enum ModifierKey
    None = 0
    Control = 1
    Shift = 2
    Alternate = 4
    Windows = 8
End Enum

' Command string format:
'   <type of command>:<command content>:delay;
' e.g. Command:app next:0.1;
Public Sub DoSaveCommands()
    ' Prompt the user for each command
    If MsgBox("Do you want to save a new set of commands for playback later (this will erase any previously saved commands)?", vbYesNo, "Save Commands") = vbNo Then
        Exit Sub
    End If

    Dim more As Boolean
    more = True

    Dim value As String
    Dim commands As New System.Text.StringBuilder()

    Dim delay As String
    delay = InputBox("How much of a delay between commands?", "Record Commands", "0.1")
    If delay = "" Then
        delay = "0.1"
    End If

    Do While (more)
        Dim newCommand As String
        '        If PromptForCommandDialog(newCommand) = False Then
        newCommand = InputBox("Enter your commands. Usage:" & vbcrlf &
                              "For keystrokes, use ""SendKeys:"" with Dragon notation: (i.e. SendKeys:{ctrl+v} or {ctrl+v}).  " & vbcrlf &
                              "For special keystrokes, use ""SendSystemKeys:"" ." & vbcrlf &
                              "For spoken commands, use ""Command:"" (i.e. Command:Minimize Window).  ", "Record Commands")
        If newCommand = "" Then
            more = False
        Else
            If commands.Length > 0 Then
                commands.Append(";")
            End If
            commands.Append(newCommand)
            commands.Append(":")
            commands.Append(delay)
            If MsgBox("Add another command?", vbYesNo) = vbNo Then
                more = False
            End If
        End If
    Loop

    SaveCommand(commands.ToString())

End Sub

' Delay is in seconds
Public Sub DoRepeatCommonCommand(theCommand As String, optional delay as string = "")
    Dim waitTime As Double
	
	if delay = "" then
		delay = InputBox("Say 'Stop Commands' to stop." & vbcrlf & "How much of a delay between commands (in seconds)?", "Repeat common command", "0.1")
		If delay = "" Then
			Exit Sub
		End If
	end if
    waitTime = CDbl(delay)

    ClearDelayedCommands()

    AddDelayedCommand("{" & theCommand & "}", DelayedCommandType.UseSendKeys, waitTime * 1000, -1)

    RunDelayedCommands()

End Sub

' Repeats commands saved using "save commands" macro
' If count < 0, run the first command repeatedly until told to stop
' Command string format:
'   <type of command>:<command content>:delay;
' e.g. Command:app next:0.1;
Public Sub DoRepeatCommands(Optional count As Integer = 1)

    Dim theCommand As String
    theCommand = GetSavedCommand()

    If theCommand = "" Then
        MsgBox("No commands saved.  First save some commands by saying 'Record Commands'.")
        Exit Sub
    End If

    ' Multiple commands are separated by ;
    Dim commands() As String
    commands = Split(theCommand, ";")

    Dim delay As String
    Dim waitTime As Double
    Dim i As Integer, j As Integer
    Dim command As String
    Dim delayedCommandRepeat As Integer
    delayedCommandRepeat = 1

    If count < 0 Then
        ' Will hand off the repeating to the MemoryForMacros process
        count = 1
        delayedCommandRepeat = -1
    End If

    ClearDelayedCommands()

    Dim commandParts() As String

    If count > 1 Then
        If MsgBox("Repeat current commands " & CStr(count) & " times (say 'stop commands' to stop)?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If

    For i = 1 To count
        For j = 0 To UBound(commands)
            'msgbox("Command: " & commands(j))
            commandParts = Split(commands(j), ":")
            If UBound(commandParts) = 2 Then
                ' Includes command type
                If commandParts(0) = "Command" Then
                    command = commandParts(1)
                    waitTime = CDbl(commandParts(2))
                    AddDelayedCommand(command, DelayedCommandType.Spoken, waitTime * 1000, delayedCommandRepeat)
                ElseIf commandParts(0) = "SendKeys" Then
                    command = commandParts(1)
                    waitTime = CDbl(commandParts(2))
                    AddDelayedCommand(command, DelayedCommandType.UseSendKeys, waitTime * 1000, delayedCommandRepeat)
                ElseIf commandParts(0) = "SendSystemKeys" Then
                    command = commandParts(1)
                    waitTime = CDbl(commandParts(2))
                    AddDelayedCommand(command, DelayedCommandType.UseSendSystemKeys, waitTime * 1000, delayedCommandRepeat)
                End If
            Else
                ' Default command type
                command = commandParts(0)
                'MsgBox(command)
                waitTime = CDbl(commandParts(1))
                AddDelayedCommand(command, DelayedCommandType.UseSendKeys, waitTime * 1000, delayedCommandRepeat)
            End If

            If delayedCommandRepeat < 0 Then
                Exit For
            End If
        Next j
        'Wait(waitTime)

    Next i

    RunDelayedCommands()

End Sub

Public Sub DoStopCommands()
    KillDelayedCommands()
End Sub

' Allows user to edit the current saved play commands
Public Sub DoEditPlayCommands()
    Dim theCommand As String
    theCommand = GetSavedCommand()

    If theCommand = "" Then
        MsgBox("No commands saved.  First save some commands by saying 'Record Commands'.")
        Exit Sub
    End If

    Dim newCommand As String
    newCommand = InputBox("Please edit the command.  Commands are separated by "";"".  Use Dragon SendKeys notation (e.g. use {ctrl+c}, not ^c).", "Record Commands", theCommand)
    If newCommand = "" Or newCommand = theCommand Then
        Exit Sub
    End If


    SaveCommand(newCommand)

End Sub

' Dictation consists of a sequence of commands, each separated by the word "then".
' For example: "got up then go down then press enter"
' using this, the user cana series of short commands without needing to pause in between
Public Sub DoCommandSequence(dictation As String)
    Dim commands() As String
    commands = dictation.Split(New String() {" then ", " Then "}, System.StringSplitOptions.None)

    ClearDelayedCommands()

    Dim aCommand As String
    For Each aCommand In commands
        'msgbox(aCommand)
        AddDelayedCommand(aCommand, DelayedCommandType.Spoken, 100)
    Next

    RunDelayedCommands()

End Sub

' "plain" prints out Text in all lowercase with no spaces
' "plain space" inserts a space before printing out the Text in all lowercase with no spaces
' "plain uppercase" prints out the text in all uppercase With no spaces
Public Sub DoMacroPlain(dictation As String)
    Dim words() As String
    words = split(dictation)
    Dim out As String

    If words(0) = "space" Then
        SendKeys(" ")
        SendKeys(formatNoCapsNoSpaces(ArrayToString(words, 1)))
    ElseIf words(0) = "uppercase" Then
        SendKeys(formatAllCapsNoSpaces(ArrayToString(words, 1)))
    Else
        SendKeys(formatNoCapsNoSpaces(dictation))
    End If
End Sub

Public Sub DoMacroSpell(dictation() As String, Optional allCaps As Boolean = False)
    Dim letter() As String
    ReDim letter(UBound(dictation))

    For i As Integer = 0 To UBound(dictation)
        letter(i) = ExtractDictationListToken(dictation(i))
        If letter(i) = "backslash" Then letter(i) = "\"
        If letter(i) = "space" Then letter(i) = " "
        If allCaps Then letter(i) = UCase(letter(i))
        SendKeys letter(i)
    Next i

    'TTSPlayString("Spelled out.")
End Sub
' Dictation options are:
'   vba jump <dictation>
'   VBA jump back dictation
Public Sub DoMacroVbaJump(ListVar1 As String)
    SendKeys "^f"

    Dim words() As String, text As String
    words = split(ListVar1)
    Dim directionUp As Boolean
    If words(0) = "back" Then
        directionUp = True
        text = ArrayToString(words, 1, UBound(words), "")
    Else
        directionUp = False
        text = formatNocapsNoSpaces(ListVar1)
    End If

    Dim saveClipboard As String
    saveClipboard = saveClipboard
    Clipboard(text)
    Wait 0.2

    SendKeys "%f"
    Wait 0.1
    SendKeys "^v"
    Wait 0.1
    SendKeys "%m"
    Wait 0.1
    If directionUp Then
        SendKeys "%du"
    Else
        SendKeys "%dd"
    End If
    Wait 0.1
    SendKeys "%n"
    Wait 0.3
    SendKeys "{Esc}"

End Sub

' Prints text for performing a "SendKeys" on a special key
' Key can be a dictation string like "PgDn\Page Down"
' For modifiers, pass in a string containing the modifier symbols, i.e. %^ or ^+.  Use "w" for the windows key.
Public Sub DoMacroCodeKeys(dictationKey As String, Optional count As Integer = 1, Optional modifiers As String = "", Optional includeSendKeys As Boolean = False)
    If includeSendKeys Then
        SendKeys("SendKeys{(}""""{)}")
        SendKeys("{left 2}")
    End If

    If modifiers.Contains("%") Then SendKeys("{%}")
    If modifiers.Contains("^") Then SendKeys("{^}")
    If modifiers.Contains("+") Then SendKeys("{+}")
    If modifiers.Contains("w") Then SendKeys("{{}WindowsHold}")

    SendKeys("{{}" & Split(dictationKey, "\")(0))

    If count > 1 Then
        SendKeys(" " & CStr(count))
    End If

    SendKeys("}")

End Sub

'Public Sub DoMacroCodeKeys(dictationKey As String, Optional count As Integer = 1, Optional modifiers As Long = 0, Optional includeSendKeys As Boolean = False)
'    If includeSendKeys Then
'        SendKeys("SendKeys{(}""""{)}")
'        SendKeys("{left 2}")
'    End If

'    If modifiers And ModifierKey.Alternate Then SendKeys("{%}")
'    If modifiers And ModifierKey.Control Then SendKeys("{^}")
'    If modifiers And ModifierKey.Shift Then SendKeys("{+}")
'    If modifiers And ModifierKey.Windows Then SendKeys("{{}WindowsHold}")

'    'Dim modifier As ModifierKey
'    'For Each modifier In modifiers
'    '    'msgbox CStr(modifier)
'    '    Select Case modifier
'    '        Case ModifierKey.Alternate
'    '            SendKeys("{%}")
'    '        Case ModifierKey.Control
'    '            SendKeys("{^}")
'    '        Case ModifierKey.Shift
'    '            SendKeys("{+}")
'    '        Case ModifierKey.Windows
'    '            SendKeys("{{}WindowsHold}")
'    '    End Select
'    'Next

'    SendKeys("{{}" & Split(dictationKey, "\")(0))

'    If count > 1 Then
'        SendKeys(" " & CStr(count))
'    End If

'    SendKeys("}")

'End Sub

Public Sub DoMacroTrimLastChar(theChar As String)
    ' Get working text
    'Dim saveClipboard As String
    'saveClipboard = GetClipboard

    SendKeys("+{Home}^c")

    Dim fullText As String
    fullText = GetClipboard()

    ' Get insertion point
    Dim index As Integer
    index = InStrRev(fullText, theChar)
    If index = 0 Then
        Beep()
        SendKeys("{Right}")
        Exit Sub
    End If

    Dim out As String
    out = Mid(fullText, 1, index - 1) & Mid(fullText, index + 1)

    PutClipboard(out)
    SendKeys("^v")
    'Wait 0.1
    'Clipboard(saveClipboard)

End Sub

Public Sub DoGetWindowPosition()

    Dim windowRect As RECT
    GetActiveWindowRect(windowRect)

    Dim moveCommand As String
    moveCommand = "'#Uses ""C:\Users\KnowBrainer\CommonModules\window.bas""" & vbcrlf & vbcrlf &
                    "Sub Main" & vbCrLf &
                    "MoveActiveWindow(" & windowRect.Left & ", " & windowRect.Top & ", " &
                                          windowRect.Right - windowRect.Left & ", " &
                                          windowRect.Bottom - windowRect.Top & ")" & vbcrlf &
                    "End Sub" & vbCrLf
    PutClipboard(moveCommand)

End Sub


' say "<number> <plus|minus|times|divided by> <second number>"
' or "<plus|minus|times|divided by> <second number>" where the first number is the result of the last math operation (in the clipboard)

Public Sub DoCalculate(ListVar1 As String)
    Dim dictation As String
    dictation = ListVar1
    dictation = Replace(dictation, "plus", "+")
    dictation = Replace(dictation, "minus", "-")
    dictation = Replace(dictation, "times", "×")
    dictation = Replace(dictation, "divided by", "÷")

    Dim op As String
    Dim index As Integer
    index = InStr(dictation, "+")
    If index Then
        op = "+"
    Else
        index = InStr(dictation, "-")
        If index Then
            op = "-"
        Else
            index = InStr(dictation, "×")
            If index Then
                op = "x"
            Else
                index = InStr(dictation, "÷")
                If index Then
                    op = "/"
                Else
                    MsgBox("Please use a phrase that contains plus, minus, times or divided by.")
                    Exit Sub
                End If
            End If
        End If
    End If


    'Separate out the numbers
    On Error Resume Next
    err.Clear

    Dim firstnum As Double
    Dim secondnum As Double
    Dim result As String
    Dim firstNumStr As String
    firstNumStr = Mid(dictation, 1, index - 1)
    firstNumStr = Replace(firstNumStr, " ", "")
    If firstNumStr = "" Then
        firstnum = CDbl(Clipboard)
        If err.Number <> 0 Then
            msgbox("Not a number: " & Clipboard)
            Exit Sub
        End If
    Else
        firstnum = CDbl(firstNumStr)
        If err.Number <> 0 Then
            msgbox("Not a number: " & firstNumStr)
            Exit Sub
        End If
    End If

    Dim secondNumStr As String
    secondNumStr = Mid(dictation, index + 1, Len(dictation) - index + 1)
    secondnum = CDbl(secondNumStr)
    If err.Number <> 0 Then
        msgbox("Not a number: " & secondNumStr)
        Exit Sub
    End If


    Dim answer As Double
    Select Case op
        Case "+"
            answer = firstnum + secondnum

        Case "-"
            answer = firstnum - secondnum

        Case "x"
            answer = firstnum * secondnum

        Case "/"
            answer = firstnum / secondnum

    End Select

    msgbox(firstNumStr & " " & op & " " & secondNumStr & " = " & vbcrlf & vbcrlf & CStr(answer))
    clipboard(CStr(answer))

End Sub

Public Enum MemoryOperation
    Add
    Subtract
    Multiply
    Divide
    Clear
    Store
    Recall
End Enum

Private Const CALCULATOR_MEMORY_COMMAND = "CalculatorMemory"

Public Sub DoCalculateMemory(operation As MemoryOperation)
    Dim value As String
    value = GetClipboard()
    If Not IsNumeric(value) Then
        If (operation = MemoryOperation.Add Or operation = MemoryOperation.Subtract Or operation = MemoryOperation.Multiply Or
            operation = MemoryOperation.Divide Or operation = MemoryOperation.Store) Then
            MsgBox("To modify the calculator memory, first do a calulation using the ""do math"" command.")
            Exit Sub
        Else
            ' Recall or clear.  The value does not matter, but it must be Numeric
            value = "0"
        End If
    End If

    Dim valueNum As Double
    valueNum = CDbl(value)

    Dim memoryStr As String
    Dim memoryNum As Double
    memoryStr = GetValue(CALCULATOR_MEMORY_COMMAND)

    If memoryStr = "" Or Not IsNumeric(memoryStr) Then
        memoryNum = 0
    Else
        memoryNum = CDbl(memoryStr)
    End If

    Select Case operation
        Case MemoryOperation.Add
            memoryNum = memoryNum + valueNum

        Case MemoryOperation.Subtract
            memoryNum = memoryNum - valueNum

        Case MemoryOperation.Multiply
            memoryNum = memoryNum * valueNum

        Case MemoryOperation.Divide
            If valueNum = 0.0 Then
                MsgBox("You cannot divide the memory contents by zero.")
                Exit Sub
            End If

            memoryNum = memoryNum / valueNum

        Case MemoryOperation.Clear
            memoryNum = 0

        Case MemoryOperation.Store
            memoryNum = valueNum
            MsgBox("Current memory contents is " & CStr(memoryNum))

        Case MemoryOperation.Recall
            MsgBox("Current memory contents is " & CStr(memoryNum))
            Exit Sub
    End Select

    SaveValue(CALCULATOR_MEMORY_COMMAND, CStr(memoryNum))

End Sub

Public Sub DoMoveCursor(dir1 As String, count1 As Integer, Optional dir2 As String = "", Optional count2 As Integer = 0)

    Dim i As Integer
    Dim theDirection As String
    Dim theCount As Integer

    For i = 1 To 2
        Select Case i
            Case 1
                theDirection = dir1
                theCount = count1
            Case 2
                theDirection = dir2
                theCount = count2
        End Select

        While theCount > 50
            Select Case theDirection
                Case "Left"
                    SendKeys("{Left 50}")
                Case "Right"
                    SendKeys("{Right 50}")
                Case "Up"
                    SendKeys("{Up 50}")
                Case "Down"
                    SendKeys("{Down 50}")
            End Select
            theCount = theCount - 50
        End While

        Select Case theDirection
            Case "Left"
                SendKeys("{Left " + CStr(theCount) + "}")
            Case "Right"
                SendKeys("{Right " + CStr(theCount) + "}")
            Case "Up"
                SendKeys("{Up " + CStr(theCount) + "}")
            Case "Down"
                SendKeys("{Down " + CStr(theCount) + "}")
        End Select
    Next i
End Sub

' Separate commands with "then"
' THIS DOES NOT WORK
Public Sub DoMultipleCommands(dictation As String)
    ' Split up into multiple commands
    Dim commands() As String
    commands = dictation.Split(New String() {" then "}, System.StringSplitOptions.RemoveEmptyEntries)
    Dim last As Integer
    last = UBound(commands)
    Dim i As Integer
    For i = 0 To last
        EmulateRecognition(commands(i))
        Wait(0.1)
    Next

End Sub
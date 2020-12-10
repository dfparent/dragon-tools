'#Language "WWB.NET"
'#Uses "imports.bas"
Imports System.String
Imports System.Collections.Generic
'option explicit

private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As System.IntPtr, ByRef lpdwProcessId As Integer) As Integer

Public Const CURRENT_WINDOW_HANDLE_VALUE = "Menus.CurrentWindowHandle"
Public Const CURRENT_PROCESS_VALUE = "Menus.CurrentProcess"

Dim jiraUrl = "https://mc10inc.atlassian.net/browse/"

' Location for saving settings in registry
' Starting location is HKEY_CURRENT_USER\ Software\ VB and VBA Program Settings
Dim registryAppName = "..\Douglas Parent\KBCommands"
Dim registrySaveCommand = "SaveCommand"
Dim registryCommandName = "CommandName"
dim registryCommandValues = "CommandValues"

Public Sub SaveValue(name As String, value as string)
    SaveSetting(registryAppName, registryCommandValues, name, value)
End Sub

Public Function GetValue(name as string) As String
    Return GetSetting(registryAppName, registryCommandValues, name)
End Function

Public Function DeleteValue(name as string)
    On Error Resume Next
    DeleteSetting(registryAppName, registryCommandValues, name)

End Function

Public Sub RunPowerShellCommand(command As String)
    ShellExecute("powershell.exe -Command ""& {" & command & "}""")
End Sub

' If you don't wait long at between clipboard actions, you can get errors
Public Sub ClipboardWait()
	Wait 0.5
end Sub

Public Function GetClipboard() As String
    ' Sometimes it takes a little while for the clipboard to respond.
    On Error Resume Next
    Dim maxWait As Long
    maxWait = 100
    Dim waitCount As Long
    waitCount = 0
    Do
        Wait(0.1)
        Err.Clear
        GetClipboard = Clipboard
        waitCount = waitCount + 1
    Loop Until err.number = 0 Or waitCount > maxWait
    On Error GoTo 0
End Function

Public Sub PutClipboard(value As String)
    ' Sometimes it takes a little while for the clipboard to respond.
    On Error Resume Next
    Dim maxWait As Long
    maxWait = 100
    Dim waitCount As Long
    waitCount = 0
    Do
        Wait(0.1)
        Err.Clear
        Clipboard(value)
        waitCount = waitCount + 1
    Loop Until err.number = 0 Or waitCount > maxWait
    On Error GoTo 0
End Sub

Public function getMachineName() as string
	return System.Environment.MachineName
end Function

Public Function IsProcessRunning(processName As String) As Boolean
	Dim processes() As System.Diagnostics.Process
    processes = System.Diagnostics.Process.GetProcessesByName(processName)
    Return processes.Length > 0
End Function

Public Function GetForegroundProcessName() As String
    Dim handle As Long
    handle = GetForegroundWindow()
    If handle = 0 Then
        msgbox("Can't get foreground Window.")
        Exit Function
    End If

    Dim currentWindowHandle As String
    Dim name As String
    currentWindowHandle = GetValue(CURRENT_WINDOW_HANDLE_VALUE)
    If CStr(handle) <> currentWindowHandle Then
        currentWindowHandle = CStr(handle)
        SaveValue(CURRENT_WINDOW_HANDLE_VALUE, currentWindowHandle)

        Dim processID As Integer
        Dim handlePtr As System.IntPtr = New System.IntPtr(handle)
        GetWindowThreadProcessId(handlePtr, processID)

        Dim theProcess As System.Diagnostics.Process
        theProcess = System.Diagnostics.Process.GetProcessById(processID)
        'msgbox(theProcess.ProcessName)
        name = theProcess.ProcessName
        SaveValue(CURRENT_PROCESS_VALUE, name)
    Else
        name = GetValue(CURRENT_PROCESS_VALUE)
    End If

    Return name


End Function

' This inserts a space if there is not already one at the insertion point.
Public sub ensureSpace
    dim saveData as String
    saveData = GetClipboard()
    SendKeys "+{left}^c"
    'ClipboardWait()
    Dim aChar as String
    aChar = GetClipboard()
    If aChar <> "" then
        ' Move cursor back
        SendKeys "{right " & CStr(Len(aChar)) & "}"
		Wait 0.1
    end if

	if aChar <> "" and aChar <> vbcrlf and InStr(" (@#$_{[", aChar) = 0 then
        ' Not start of doc or start of line and no preceding space desired
        SendKeys " "
    end If
    PutClipboard(saveData)
End Sub

' "This is some Text" -> "thisissometext"
Public function formatNoCapsNoSpaces(text as string)
    text = lcase$(text)
    return Replace(text, " ", "")
End Function

' "This is some Text" ->  "THISISSOMETEXT"
Public function formatAllCapsNoSpaces(text as string)
    text = Ucase$(text)
    return Replace(text, " ", "")
End Function

' "This is some Text" -> "ThisissomeText"
Public Function formatNoSpaces(text As String)
    Return Replace(text, " ", "")
End Function

Public Function formatCamelCase(text As String) As String
    Return formatCapsNoSpaces(text, False)
End Function

Public Function formatPascalCase(text As String) As String
    Return formatCapsNoSpaces(text, True)
End Function

' False for camel case:  "This is some Text" ->  "thisIsSomeText"
' True for Pascal case: "This is some Text" ->  "ThisIsSomeText"
Public Function formatCapsNoSpaces(text As String, Optional capitalizeFirst As Boolean = True) As String
    Dim words() As String
    words = Split(text)
    Dim out As String
    Dim i As Long
    For i = 0 To UBound(words)
        words(i) = lcase$(words(i))
        If i <> 0 Or i = 0 And capitalizeFirst Then
            words(i) = UCase$(Left$(words(i), 1)) & Mid$(words(i), 2)
        End If
        out = out & words(i)
    Next
    Return out
End Function

' allCaps = true: This is some Text" -> THIS_IS_SOME_TEXT
' allCaps = false: This is some Text" -> this_is_some_text
Public Function formatUnderscores(text As String, Optional allCaps As Boolean = False) As String
    Dim words() As String
    words = Split(text)
    Dim out As String
    Dim i As Long
    For i = 0 To UBound(words)
        If allCaps Then
            words(i) = Ucase$(words(i))
        Else
            words(i) = LCase$(words(i))
        End If
        If i > 0 Then
            out = out & "_"
        End If
        out = out & words(i)
    Next
    Return out
End Function

' This is some Text" -> this-is-some-text
Public function formatHyphens(text as string)
    text = Replace(text, " ", "-")
    text = LCase(text)
    return text
end Function

' Does a SendKeys after cleaning up some characters which otherwise might not get printed,
' but might affect other key strokes because of the special meeting in the Sendkeys command.
Public Sub TypeString(dictation As String, Optional allCaps As Boolean = False)
    Dim text As String
    If allCaps Then
        text = UCase(dictation)
    Else
        text = dictation
    End If
    text = Replace(text, "(", "{(}")
    text = Replace(text, """", "{""}")
    text = Replace(text, "~", "{~}")
    text = Replace(text, "%", "{%}")
    text = Replace(text, vbcrlf, "~")
    SendKeys(text)
End Sub

Public Sub SlowTypeString(dictation As String)
    Dim i As Integer
    For i = 1 To Len(dictation)
        SendKeys(Mid(dictation, i, 1))
        Wait(0.1)
    Next i
End Sub

Public Function CapitalizeString(text As String, Optional firstWordOnly As Boolean = False) As String
    Dim words() As String
    words = Split(text)
    Dim out As String
    Dim i As Long
    For i = 0 To UBound(words)
        If firstWordOnly And i > 0 Then
            out = out & words(i)
        Else
            out = out & UCase$(Left$(words(i), 1)) & Mid$(words(i), 2)
        End If
        If i < UBound(words) Then
            out = out & " "
        End If
    Next
    Return out
End Function

Public Function ArrayToString(theArray, fromIndex As Integer, Optional toIndex As Integer = -1, Optional separator As String = " ") As String
    If Not IsArray(theArray) Then
        MsgBox("Passed in variable is not an array in ArrayToString.")
        Return ""
    End If

    If fromIndex < 0 Or fromIndex > Ubound(theArray) Then
        MsgBox("fromIndex is not valid for the passed in array in ArrayToString.")
        Return ""
    End If

    If toIndex <> -1 And (toIndex < 0 Or toIndex > UBound(theArray)) Then
        MsgBox("toIndex is not valid for the passed in array in ArrayToString.")
        Return ""
    End If

    If toIndex = -1 Then
        toIndex = Ubound(theArray)
    End If

    Dim out As String
    For i As Integer = fromIndex To toIndex
        If out <> "" Then
            out = out & separator
        End If
        out = out & theArray(i)
    Next i

    Return out
End Function

' Searches for the given value in the given array and returns the 1st index number of the value, If found in the array
' If not found in the array, or if an error, returns -1
Public Function FindArrayElement(theValue, theArray) As Integer
    FindArrayElement = -1
    If Not IsArray(theArray) Then
        MsgBox("FindArrayElement: Please provide an array and the 2nd parameter.")
        Exit Function
    End If

    If IsArray(theValue) Then
        MsgBox("FindArrayElement: Please provide a non-array value in the 1st parameter.")
        Exit Function
    End If

    Dim element
    Dim index As Integer
    For index = 0 To UBound(theArray)
        If theValue = theArray(index) Then
            Return index
        End If
    Next
    Return -1
End Function

' Allows you to split a string and retain the delimiters
Public Function SplitString(source As String, delimiters() As String, Optional options As System.StringSplitOptions = System.StringSplitOptions.RemoveEmptyEntries, Optional retainDelimiters As Boolean = True) As String()
    If source = "" Then
        Return Nothing
    End If

    If Not retainDelimiters Then
        Return source.Split(delimiters, options)
    End If

    Dim outArray As New List(Of String)
    Dim startSplit As Integer = 0
    Dim delimiterString As String
    delimiterString = ArrayToString(delimiters, 0, -1, "")
    Dim i As Integer
    For i = 0 To source.Length - 1
        'msgbox(source.Substring(i, 1))
        If delimiterString.Contains(source.Substring(i, 1)) Then
            outArray.Add(source.Substring(startSplit, i - startSplit))
            outArray.Add(source.Substring(i, 1))
            startSplit = i + 1
        End If
    Next

    If startSplit < source.Length Then
        outArray.Add(source.Substring(startSplit))
    End If

    Return outArray.ToArray()

End Function

Public Function DigitizeNumbers(dictation As String) As String
    ' Convert whole "spelled-out" numbers into digits, ie. one -> 1
    Dim words() As String
    words = Split(dictation)
    Dim out As String
    Dim fixedWord As String
    Dim lastWordWasNumber As Boolean
    For Each word As String In words
        fixedWord = numberNameToDigit(word)
        If fixedWord <> "" Then
            ' Is a number.  Add leading space if this is the first digit of a new number
            If out <> "" And Not lastWordWasNumber Then
                out = out & " "
            End If
            out = out & fixedWord
            lastWordWasNumber = True
        Else
            ' Not a number.  Just add it
            If out <> "" Then
                out = out & " "
            End If
            out = out & word
            lastWordWasNumber = False
        End If
    Next

    DigitizeNumbers = out
End Function

' Examines dictation and converts spelled out numbers to digits and phoenetic letters to letters
' If first 2 words are "all upper" then the remaining text is returned in all uppercase
' Does not print out words, but keystrokes, so no added spaces
' Any actual words thrown in are assumed to be spelled out and are printed without spaces
' Due to limitation with KnowBrainer, to capitalize letters, use "upper" rather than "cap". :(
'
Public Function DictationToKeystrokes(dictation As String, Optional allcaps As Boolean = False) As String
    'msgbox(dictation & vbcrlf & "All caps: " & CStr(allcaps))
    Dim words() As String
    ' Split on all punctuation and retain delimiters
    ' Do not split on - or you will split hyphenated words like x-ray
    Dim separators() As String = {"\", "|", "!", """", "'", "#", "*", ",", ".", "..", "...", "/", ":", ";", "<", "=", ">", "?", "@", " ", "_", "(", ")", "+", "[", "]", "{", "}", "~", "$", "%", "&", "^"}
    words = SplitString(dictation, separators)
    Dim out As String
    Dim fixedWord As String
    Dim capitalizeNext As Boolean = False

    ' With dictation: "type all upper some text", the dictation is "all upper some text"
    ' After split this is an array with UBound = 6 (Includes space delimiters)
    If UBound(words) >= 2 Then
        If LCase(words(0)) = "all" And words(1) = " " And LCase(words(2)) = "upper" Then
            allcaps = True
            ' Do not return the words "all upper"
            words(0) = " "
            words(1) = " "
            words(2) = " "
        End If
    End If

    For Each word As String In words
        ' Skip space delimiter
        If word = " " Then
            Continue For
        End If

        If LCase(word) = "upper" Then
            capitalizeNext = True
            Continue For
        End If

        ' Is it a number?
        fixedWord = numberNameToDigit(word)
        If fixedWord = "" Then
            ' Is it a phonetic letter?
            fixedWord = PhoneticToLetter(word, capitalizeNext)
            If fixedWord = "" Then
                ' Is it punctuation?
                fixedWord = PunctuationNameToCharacter(word)
                If fixedWord = "" Then
                    If capitalizeNext Then
                        fixedWord = UCase$(Left$(word, 1)) & LCase(Mid$(word, 2))
                    Else
                        fixedWord = LCase(word)
                    End If
                End If
            End If
        End If

        If allcaps Then
            fixedWord = fixedWord.ToUpper()
        End If

        out = out & fixedWord

        capitalizeNext = False
    Next

    DictationToKeystrokes = out

End Function

' Returns the number as a character if you pass in the number name, i.e., "one" -> "1"
' if the string being passed in is not a number, the function returns An empty string.
Public function numberNameToDigit(number as string) as string
	dim num as string
	num = lcase(number)
	select case num
        Case "one", "1"
            Return "1"

        Case "two", "to", "too", "2"
            Return "2"

        Case "three", "3"
            Return "3"

        Case "four", "for", "fore", "4"
            Return "4"

        Case "five", "5"
            Return "5"

        Case "six", "6"
            Return "6"

        Case "seven", "7"
            Return "7"

        Case "eight", "8"
            Return "8"

        Case "nine", "9"
            Return "9"

        Case "zero", "0"
            Return "0"

        Case Else
            Return ""

    End select
end function

Public Function monthNameToNumber(name As String) As Integer
    Select Case name
        Case "January"
            Return 1
        Case "February"
            Return 2
        Case "March"
            Return 3
        Case "April"
            Return 4
        Case "May"
            Return 5
        Case "June"
            Return 6
        Case "July"
            Return 7
        Case "August"
            Return 8
        Case "September"
            Return 9
        Case "October"
            Return 10
        Case "November"
            Return 11
        Case "December"
            Return 12
        Case Else
            Return 0
    End Select
End Function

public Function monthNumberToShortName(month as integer) as string
    Select Case month
        Case 1
            Return "Jan"
        Case 2
            Return "Feb"
        Case 3
            Return "Mar"
        Case 4
			Return "Apr"
        Case 5
            Return "May"
        Case 6
            Return "Jun"
        Case 7
            Return "Jul"
        Case 8
            Return "Aug"
        Case 9
            Return "Sep"
        Case 10
            Return "Oct"
        Case 11
            Return "Nov"
        Case 12
            Return "Dec"
        Case Else
            Return ""
    End Select
end Function

' Returns an empty string if not a letter or phonetic letter
Public Function PhoneticToLetter(dictation As String, Optional capitalize As Boolean = False) As String
    Dim text As String
    text = lcase(dictation)
    'msgbox text
    Select Case text
        Case "alpha", "alfa", "bravo", "be", "bee", "charlie", "delta", "echo", "foxtrot", "golf", "hotel", "india", "juliet", "jay",
             "kilo", "lima", "mike", "november", "oscar", "papa", "pop", "québec", "quebec", "romeo", "sierra",
             "tango", "tee", "uniform", "victor", "whiskey", "x-ray", "xray", "yankee", "zulu"
            PhoneticToLetter = Left(text, 1)
        Case "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"
            PhoneticToLetter = text
        Case "see"
            PhoneticToLetter = "c"
        Case "eye"
            PhoneticToLetter = "i"
        Case "and", "in"
            PhoneticToLetter = "n"
		case "popper"
			PhoneticToLetter = "p"
        Case "are"
            PhoneticToLetter = "r"
        Case "you"
            PhoneticToLetter = "u"
        Case "why"
            PhoneticToLetter = "y"
        Case Else
            PhoneticToLetter = ""
    End Select

    If capitalize Then
        PhoneticToLetter = UCase(PhoneticToLetter)
    End If
End Function

' Provides the zero-based sequence number within the alphabet for a particular letter. 
' For example, A = 0, B = 1, etc...
' Returns -1 if the passed in string is not a letter.
Public Function LetterToNumber(letter As String) As Integer
    LetterToNumber = -1

    If letter.Length <> 1 Then
        MsgBox("Not a letter: " & letter)
        Exit Function
    End If

    Dim alphabet As String = "abcdefghijklmnopqrstuvwxyz"
    Dim index As Integer
    index = alphabet.IndexOf(letter.ToLower())

    If index = -1 Then
        MsgBox("Not a letter: " & letter)
        Exit Function
    End If

    LetterToNumber = index

End Function

' Returns an empty string if not a punctuation name or character
Public Function PunctuationNameToCharacter(dictation As String) As String
    Dim text As String
    text = lcase(dictation)
    Dim out As String
    'msgbox text

    Select Case text
        Case "\", "backslash"
            out = "\"
        Case "|", "vertical bar"
            out = "|"
        Case "!", "exclamation mark"
            out = "!"
        Case """", "close quote", "open quote", "double quote", "quote"
            out = """"
        Case "'", "apostrophe"
            out = "'"
        Case "#", "hash sign", "pound sign"
            out = "#"
        Case "*", "asterisk"
            out = "*"
        Case ",", "comma"
            out = ","
        Case "-", "hyphen", "minus sign"
            out = "-"
        Case ".", "full stop", "dot", "point", "period"
            out = "."
		case ".."
			out = ".."
		case "...", "…", "elipsis"
			out = "..."
        Case "/", "slash"
            out = "/"
        Case ":", "colon"
            out = ":"
        Case ";", "semi colon"
            out = ";"
        Case "<", " less than", " less than sign"
            out = "<"
        Case "=", " equals", " equal sign"
            out = "="
        Case ">", "greater than", "greater than sign"
            out = ">"
        Case "?", "question mark"
            out = "?"
        Case "@", "at sign"
            out = "@"
        Case " ", "space", "space bar"
            out = " "
        Case "_", "underscore"
            out = "_"
        Case "(", "open paren"
            out = "{(}"
        Case ")", "close paren"
            out = "{)}"
        Case "+", "plus sign"
            out = "{+}"
        Case "[", "open square bracket"
            out = "{[}"
        Case "]", "close square bracket"
            out = "{]}"
        Case "{", "open curly brace"
            out = "{{}"
        Case "}", "close curly brace"
            out = "}"
        Case "~", "tilde"
            out = "{~}"
        Case "$", "dollar sign"
            out = "$"
        Case "%", "percent sign"
            out = "{%}"
        Case "&", "ampersand"
            out = "&"
        Case "^", "caret"
            out = "{^}"
        Case Else
            out = ""
    End Select

    PunctuationNameToCharacter = out
End Function

Public Sub SaveCommand(command As String)
    SaveSetting(registryAppName, registrySaveCommand, registryCommandName, command)
End Sub

Public Function GetSavedCommand() As String
    Return GetSetting(registryAppName, registrySaveCommand, registryCommandName)
End Function

Public Function DeleteSavedCommand()
    On Error Resume Next
    DeleteSetting(registryAppName, registrySaveCommand, registryCommandName)

End Function

' Repeats keystrokes with waits in between each
public sub RepeatKeyStrokes(keys as string, count as integer, optional delay as double = 0.1)
	dim i as integer
	for i = 1 to count
		SendKeys keys
		Wait delay
	next i
end Sub

Public Sub lowercaseSelection()

    SendKeys "^c"

    Dim fullText As String
    fullText = GetClipboard()
    fullText = LCase(fullText)
    PutClipboard(fullText)

    SendKeys("^v")
End Sub

' Single text string.  The "find" and replace parts should be in the same utterance, separated by the word "with"
Public Sub lineReplace(dictation As String)

    ' Parse dictation
    Dim lowerListVar As String
    lowerListVar = LCase$(dictation)
    Dim withIndex As Integer
    withIndex = Instr(lowerListVar, " with ")
    If withIndex = 0 Then
        TTSPlayString("To do a replace, use the 'with' trigger word. As in 'line replace this with that'.")
        Exit Sub
    End If

    ' Find the find text and replace text
    Dim findText As String
    Dim replaceText As String
    findText = Mid(lowerListVar, 1, withIndex - 1)
    findText = PunctuationNameToCharacter(findText)
    replaceText = Mid(dictation, withIndex + 6)
    replaceText = PunctuationNameToCharacter(replaceText)

    ' Get working text
    'dim saveClipboard as String
    'saveClipboard = Clipboard

    SendKeys "{Home}+{End}^c"

    'ClipboardWait()


    Dim fullText As String
    'fullText = Clipboard
    fullText = GetClipboard()

    ' Get insertion point
    Dim index As Integer
    index = InStr(LCase(fullText), LCase(findText))
    If index = 0 Then
        Beep()
        SendKeys "{Right}"
        Exit Sub
    End If

    Dim out As String
    out = fullText
    Dim lastOut As String
    Do
        lastOut = out
        out = Mid(out, 1, index - 1) & replaceText & Mid(out, index + Len(findText))
        index = InStr(LCase(out), LCase(findText))
    Loop Until index = 0 Or lastOut = out

    'Clipboard(out)
    PutClipboard(out)
    'ClipboardWait()

    SendKeys "^v"

    If out.EndsWith(vbcrlf) Then
        SendKeys("~{Left}")
    End If

    ' Clipboard(saveClipboard)


End Sub

Public Sub LineInsertRemove(ListVar1 As String)
    ' Say "line insert cows before donkeys" to insert "cows" before "donkeys" in the current line.
    ' You can also specify "after" instead of "before".
    ' Say "line insert before blah" to move insertion point only.
    ' Only works on the current line.

    ' Parse dictation
    Dim lowerListVar As String
    lowerListVar = LCase$(ListVar1)
    Dim bBefore As Boolean
    bBefore = False
    Dim beforeAfterIndex As Integer
    beforeAfterIndex = Instr(lowerListVar, "before")
    If beforeAfterIndex <> 0 Then
        bBefore = True
    Else
        beforeAfterIndex = Instr(lowerListVar, "after")
        If beforeAfterIndex = 0 Then
            ' No before or after.  Simply print the text and exit.
            SendKeys(ListVar1)
            Exit Sub
        End If
    End If

    ' Find the text prior to before/after and following before/after
    ' the "insert what" will include a trailing space
    Dim insertWhat As String
    insertWhat = Mid(ListVar1, 1, beforeAfterIndex - 1)

    ' The "insert where" will not include a trailing space
    Dim insertWhere As String
    If bBefore Then
        insertWhere = mid(lowerListVar, beforeAfterIndex + 6)
    Else
        insertWhere = mid(lowerListVar, beforeAfterIndex + 5)
    End If
    insertWhere = insertWhere.TrimStart()

    ' Get working text
    Dim saveClipboard As String
    saveClipboard = GetClipboard()

    SendKeys("{Home}+{End}^c")

    Dim fullText As String
    fullText = GetClipboard()

    ' Get insertion point
    Dim index As Integer
    index = InStr(LCase(fullText), insertWhere)
    If index = 0 Then
        SendKeys("{Right}")
        Exit Sub
    End If

    If Not bBefore Then
        index = index + len(insertWhere)

        If len(insertWhat) > 0 Then
            ' Need to fix spaces around insertwhat
            insertWhat = " " & Trim(insertWhat)
        End If
    End If


    Dim out As String
    out = Mid(fullText, 1, index - 1) & insertWhat & Mid(fullText, index)

    PutClipboard(out)
    SendKeys("^v{Home}")
    SendKeys("{Right " & CStr(index - 1 + Len(insertWhat)) & "}")
    PutClipboard(saveClipboard)
End Sub

Public Sub LineSelect(ListVar1 As String)
    ' Get working text
    Dim saveClipboard As String
    saveClipboard = GetClipboard()

    SendKeys("{Home}+{End}^c")

    Dim fullText As String
    fullText = GetClipboard()

    fullText = fullText.Replace(vbcrlf, "")
    fullText = fullText.Replace(vbcr, "")
    fullText = fullText.Replace(vblf, "")

    ' Get selection point
    Dim index As Integer
    index = InStr(LCase(fullText), LCase(ListVar1))
    If index = 0 Then
        Beep()
        Exit Sub
    End If

    ' Select from start or end?
    Dim moveKey As String
    If index < Len(fullText) / 2 Then
        SendKeys("{end}{Home}")
        'SendKeys("{left}")
        moveKey = "{Right "
    Else
        index = Len(fullText) - index - len(ListVar1) + 2
        'SendKeys("{right}")
		SendKeys("{end}")
        moveKey = "{Left "
    End If

    ' If you send too many keys, you get an error.
    ' Need to break the next command up into multiple commands
    While index > 0
        If index > 10 Then
            SendKeys(moveKey & "10}")
            index = index - 10
        Else
            SendKeys(moveKey & CStr(index - 1) & "}")
            index = 0
        End If
    End While
    SendKeys("+" & moveKey & len(ListVar1) & "}")

    PutClipboard(saveClipboard)
End Sub

' Listvar Specifies enclosing punctuation like ( ) or [ ]
' it must include 2 characters
Public Sub LineSelectInside(ListVar1 As String, Optional count As Integer = 1)
    Dim saveClipboard As String
    saveClipboard = GetClipboard()

    ' Get working text
    SendKeys("{Home}+{End}^c")

    Dim fullText As String
    fullText = GetClipboard()

    fullText = fullText.Replace(vbcrlf, "")
    fullText = fullText.Replace(vbcr, "")
    fullText = fullText.Replace(vblf, "")

    Dim startChar As String, endChar As String, Chars As String
    Chars = Split(ListVar1, "\")(0)
    startChar = Left(chars, 1)
    endChar = Mid(chars, 2, 1)

    ' Get selection start point
    Dim startIndex As Integer
    startIndex = 0
    Dim i As Integer
    For i = 1 To count
        startIndex = InStr(startIndex + 1, LCase(fullText), startChar)
        If startIndex = 0 Then
            Beep()
            Exit Sub
        End If
    Next

    startIndex = startIndex + 1

    ' Get selection end point
    Dim endIndex As Integer
    endIndex = InStr(startIndex, LCase(fullText), endChar)
    If endIndex = 0 Then
        endIndex = Len(fullText) + 1
    End If


    Dim lenSelection As Integer
    lenSelection = endIndex - startIndex

    ' Select from start or end?
    Dim moveKey As String
    If startIndex < Len(fullText) / 2 Then
        SendKeys("{End}{Home}")
        moveKey = "{Right "
    Else
        startIndex = Len(fullText) - startIndex - lenSelection + 2
        SendKeys("{End}")
        moveKey = "{Left "
    End If

    ' If you send too many keys, you get an error.
    ' Need to break the next command up into multiple commands
    Dim index As Integer
    index = startIndex
    While index > 0
        If index > 10 Then
            SendKeys(moveKey & "10}")
            index = index - 10
        Else
            SendKeys(moveKey & CStr(index - 1) & "}")
            index = 0
        End If
    End While

    SendKeys("+" & moveKey & lenSelection & "}")

    PutClipboard(saveClipboard)
End Sub

' Sometimes dictation list items can be segmented by "\" such as
'   A\cap alpha
' where the text after the backslash is what you say and the text before the backslash is what you type
' This function returns the first token (what you type) given the entire dictation list item
Public Function ExtractDictationListToken(dictationList As String) As String
    Dim index As Integer
    index = dictationList.IndexOf("\")
    If index >= 0 Then
        Return dictationList.Substring(0, index)
    Else
        Return dictationList
    End If
End Function

Public Function DictationVarsToArray(dictation1 As String, dictation2 As String,
                Optional dictation3 As String = "", Optional dictation4 As String = "", Optional dictation5 As String = "",
                Optional dictation6 As String = "", Optional dictation7 As String = "", Optional dictation8 As String = "",
                Optional dictation9 As String = "", Optional dictation10 As String = "", Optional dictation11 As String = "",
                Optional dictation12 As String = "") As String()
    Dim theArray() As String
    If dictation12 <> "" Then
        ReDim theArray(11)
    ElseIf dictation11 <> "" Then
        ReDim theArray(10)
    ElseIf dictation10 <> "" Then
        ReDim theArray(9)
    ElseIf dictation9 <> "" Then
        ReDim theArray(8)
    ElseIf dictation8 <> "" Then
        ReDim theArray(7)
    ElseIf dictation7 <> "" Then
        ReDim theArray(6)
    ElseIf dictation6 <> "" Then
        ReDim theArray(5)
    ElseIf dictation5 <> "" Then
        ReDim theArray(4)
    ElseIf dictation4 <> "" Then
        ReDim theArray(3)
    ElseIf dictation3 <> "" Then
        ReDim theArray(2)
    ElseIf dictation2 <> "" Then
        ReDim theArray(1)
    ElseIf dictation1 <> "" Then
        ReDim theArray(0)
    End If

    For i As Integer = 0 To UBound(theArray)
        Select Case i
            Case 0
                theArray(i) = dictation1
            Case 1
                theArray(i) = dictation2
            Case 2
                theArray(i) = dictation3
            Case 3
                theArray(i) = dictation4
            Case 4
                theArray(i) = dictation5
            Case 5
                theArray(i) = dictation6
            Case 6
                theArray(i) = dictation7
            Case 7
                theArray(i) = dictation8
            Case 8
                theArray(i) = dictation9
            Case 9
                theArray(i) = dictation10
            Case 10
                theArray(i) = dictation11
            Case 11
                theArray(i) = dictation12
        End Select
    Next

    DictationVarsToArray = theArray
End Function

'''''''''''''''''''''''''''''''''''''''
' JIRA ticket functions
'''''''''''''''''''''''''''''''''''''''
' Returns a ticket number
'Public function getTicketNumber(project as string, _
'                                digit1 as string, _
'                                optional digit2 as string = "", _
'                                optional digit3 as string = "", _
'                                optional digit4 as string = "", _
'                                optional digit5 as string = "")

''    if project = "oh help" then
''        project = "OHELP"
''    elseif project = "oak" then
''        project= "OAC"
''	elseif project = "O N E" then
''		project = "ONE"
''    end if

'    dim ticket as String
'    ticket = digit1
'    if digit2 <> "" then ticket = ticket & digit2
'    if digit3 <> "" then ticket = ticket & digit3
'    if digit4 <> "" then ticket = ticket & digit4
'    if digit5 <> "" then ticket = ticket & digit5

'	' Project may include a backslash for an alternate pronounciation.  
'	' Pick out whatever is in front of the backslash

'	return Split(project, "\")(0) & "-" & ticket

'end function

'' Types out a ticket number						
'public sub generateTicketNumber(project as string, _
'                                digit1 as string, _
'                                optional digit2 as string = "", _
'                                optional digit3 as string = "", _
'                                optional digit4 as string = "", _
'                                optional digit5 as string = "", _
'								optional ensureLeadingSpace as Boolean = True)

'    if ensureLeadingSpace then
'		ensureSpace
'	end if
'	dim ticket as string
'	ticket = getTicketNumber(project, digit1, digit2, digit3, digit4, digit5)
'    SendKeys ticket
'end sub

'' Types out a link to a ticket						
'public sub generateTicketLink(project as string, _
'                                digit1 as string, _
'                                optional digit2 as string = "", _
'                                optional digit3 as string = "", _
'                                optional digit4 as string = "", _
'                                optional digit5 as string = "")

'    ensureSpace
'	dim ticket as string
'	ticket = getTicketNumber(project, digit1, digit2, digit3, digit4, digit5)
'    SendKeys jiraUrl & ticket

'end sub

'' Brings up a JIRA page for the given ticket number
'public sub bringUpJiraTicket(project as string, _
'                                digit1 as string, _
'                                optional digit2 as string = "", _
'                                optional digit3 as string = "", _
'                                optional digit4 as string = "", _
'                                optional digit5 as string = "")

'	dim ticket as string
'	ticket = getTicketNumber(project, digit1, digit2, digit3, digit4, digit5)
'	ShellExecute jiraUrl & ticket
'	Wait 0.5
'end sub

'public sub EditJiraField(theField as string)
'	'SendKeys "."
'	'Wait 0.1
'	'SendKeys "edit"
'	'Wait 0.3
'	'SendKeys "{Enter}"
'	SendKeys "e"
'	Wait 1			
'	SendKeys "^f"
'	Wait 0.3
'	SendKeys theField
'	Wait 0.1
'	SendKeys "{Esc}{Tab}"

'	'SendKeys "."
'	'Wait 0.3
'	'SendKeys theField
'	'Wait 0.2
'	'SendKeys "{Enter}"
'end Sub


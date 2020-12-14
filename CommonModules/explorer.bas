'#Uses "paths.bas"
'#Uses "utilities.bas"
'#Uses "cache.bas"
'#Uses "window.bas"
'#Language "WWB.NET"
Option Explicit On

Private Const COMPARE_FILE_ONE_VALUE = "CompareFileOne"
Private DIALOGUE_TITLES() As String = {"Save As", "Save All", "Save Attachment", "Save Print Output As", _
									   "Change Source:", "Open", "Insert File", "Browse"}

Public sub doAddressBar
    SendKeys("%d")
    Wait(0.1)
End sub

public sub doLeftSide
	doAddressBar()
    SendKeys("{Tab 2}")
End Sub

Public Sub doRightSide()
    doAddressBar()
    RepeatKeyStrokes("{Tab}", 3)
	SendKeys("{Down}{Up}")
End Sub

Public Sub doDialogLeftSide()
    SendKeys("%d%n")
    SendKeys("+{Tab 3}")
    Wait(0.1)
End Sub

Public Sub doDialogRightSide()

    ' Save current name.  Sometimes get screwed up when doing this
    SendKeys("%n^c")


    SendKeys("%d{tab}%n")
    SendKeys("+{Tab 2}")
    SendKeys("{Down}{Up}")
    Wait(0.1)
End Sub

Public sub doDialogFolder(dictation as string)

    If CheckWindowText("Outlook") Then
        Exit Sub
    End If

    If Not CheckWindowText(DIALOGUE_TITLES) Then
        Exit Sub
    End If

    ' Is this a known path?
    Dim pathName as string
    pathName = Trim$(dictation)
    dim thePath as string
    thePath = getPath(pathName)

    if thePath = "" Then
        ' Not a pre known path
        TTSPlayString("Let me find the " & dictation & " folder.")
        doDialogLeftSide()
        SendKeys(pathName)
        Wait(0.1)
        SendKeys("{Enter}")
        Wait(0.1)
        SendKeys("{tab}%n")
    Else
        ' Pre known
        'TTSPlayString("Pre set path")
        doAddressBar()
        SendKeys(thePath)
        SendKeys("{Enter}")
        Wait(0.5)
        SendKeys("{tab}%n")
    End if
End Sub

Public Sub DoDialogueSubName(dictation As String)
    If Not CheckWindowText(DIALOGUE_TITLES) Then
        Exit Sub
    End If


    doDialogRightSide()
    SendKeys(formatNoCapsNoSpaces(dictation))
    SendKeys("{Enter}")
End Sub

Public Sub DoDialogueSubNumber(dictation As String)
    If Not CheckWindowText(DIALOGUE_TITLES) Then
        Exit Sub
    End If

    doDialogRightSide()
    SendKeys("{Home}")

    Dim count As Integer
    count = CInt(numberNameToDigit(dictation))
    If count > 1 Then
        SendKeys("{Down " & CStr(count - 1) & "}")
    End If
    SendKeys("{Enter}")

End Sub

Public Sub DoFolder(ListVar1 As String)
    ' Is this a known path?
    Dim pathName As String
    pathName = Trim$(ListVar1)
    Dim thePath As String

    thePath = getPath(pathName)

    If thePath = "" Then
        ' Not a pre known path
        TTSPlayString("Unknown path.  Let me find that.")
        doLeftSide()
        SendKeys pathName
        SendKeys "{Enter}"
        Wait 0.1
        doRightSide()
    Else
        ' Pre known
        'TTSPlayString("Pre set path")
        doAddressBar()
        SendKeys thePath
        SendKeys "{Enter}"
        Wait 0.1
        SendKeys "{Tab 3}"
    End If
End Sub

Public Sub DoSubNumber(listvar1 As String)
    doRightSide()
    SendKeys("{home}")
    Dim count As Long
    count = CInt(listvar1) - 1
    SendKeys("{Down " & CStr(count) & "}")
    SendKeys("{Enter}")
End Sub

Public Sub DoSubName(dictation As String)
    doRightSide()
    SendKeys "{home}"
    SendKeys Trim(dictation)
    Wait(0.1)
    SendKeys "{Enter}"
End Sub

Public Sub CopyFileName()
    SendKeys("{F2}^a")
    Wait(0.1)
    SendKeys("^c")
    SendKeys("{esc}")
End Sub

Public Function GetFilePath() As String
    SendKeys("{F2}^a")
    Wait(0.1)
    SendKeys("^c")
    Wait(0.1)
    Dim name As String
    name = GetClipboard()
    SendKeys("{Esc}")
    'msgbox name

    doAddressBar()
    SendKeys("^c")
    Wait(0.1)
    Dim path As String
    path = GetClipboard()
    'msgbox path

    SendKeys("{Tab 3}")

    GetFilePath = path & "\" & name
End Function

Public Function GetFolderPath() As String
    doAddressBar()
    SendKeys("^c")
    Dim path As String
    path = GetClipboard()

    SendKeys("{Tab 3}")

    GetFolderPath = path
End Function

Public Function CopyFilePath() As String
    'msgbox path
    Dim path As String
    path = GetFilePath()
    PutClipboard(path)

    CopyFilePath = path
End Function

Public Function CopyFolderPath() As String
    CopyFolderPath = GetFolderPath()
End Function

Public Function CopyUNCPath() As String
    Dim path As String
    path = GetFilePath()
    Dim uncPath As String
    uncPath = ConvertToUncPath(path)
    'msgbox uncPath
    PutClipboard(uncPath)
    CopyUNCPath = uncPath
End Function

Public Function CopyFolderUNCPath() As String
    Dim path As String
    path = GetFolderPath()
    Dim uncPath As String
    uncPath = ConvertToUncPath(path)
    'msgbox uncPath
    PutClipboard(uncPath)
    CopyFolderUNCPath = uncPath
End Function

Public Function ConvertToUncPath(path As String)
    If path.StartsWith("\\") Then
        Return path
    End If
    Dim key As Microsoft.Win32.RegistryKey
    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Network\" & path.Substring(0, 1))
    If Not key Is Nothing Then
        'msgbox "Network\" & path.Substring(0, 1)
        path = key.GetValue("RemotePath").ToString() & path.Substring(2)
    End If
    Return path
End Function


Public Sub DoCompareFileOne(Optional compareFolder As Boolean = False)
    Dim fileOne As String
    If compareFolder Then
        fileOne = CopyFolderPath()
    Else
        fileOne = CopyFilePath()
    End If

    Dim values As String() = {fileOne}
    UpdateCacheValue(COMPARE_FILE_ONE_VALUE, values)
    'MsgBox fileOne
End Sub

Public Sub DoCompareFiles(Optional compareFolders As Boolean = False)
    If Not CacheValueExists(COMPARE_FILE_ONE_VALUE) Then
        Msgbox("Please select a file (or folder) to compare to and say ""Compare file one"" or ""Compare folder one"".  Then select the second file and say ""Compare files"" or ""Compare folders""")
        Exit Sub
    End If

    Dim fileOne As String
    fileOne = GetCacheValue(COMPARE_FILE_ONE_VALUE)(0)

    Dim fileTwo As String
    If compareFolders Then
        fileTwo = CopyFolderPath()
    Else
        fileTwo = CopyFilePath()
    End If


    Dim compareExe As String
    Dim compareExeFileKey As String
    compareExeFileKey = "compare executable"
    compareExe = getFile(compareExeFileKey)
    If compareExe = "" Then
        msgbox("There is no file defined with key " + compareExeFileKey)
        Exit Sub
    End If

    Dim commandLine As String
    commandLine = """" & compareExe & """ """ & fileOne & """ """ & fileTwo & """"
    'Msgbox commandLine
    'PutClipboard(commandLine)
    Shell commandLine
End Sub
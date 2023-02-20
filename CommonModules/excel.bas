'#Uses "utilities.bas"
'#Uses "imports.bas"
'option explicit

Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Sub ReplaceCellContents(dictation As String)
    SendKeys("{F2}^+{Home}^c")

    Dim text As String
    text = GetClipboard()

    Dim findText As String
    Dim replaceText As String

    If Not ParseReplaceDictation(dictation, findText, replaceText) Then
        SendKeys("{escape}")
        Exit Sub
    End If

    text = Replace(text, findText, replaceText)

    PutClipboard(text)
    SendKeys("^v~")
End Sub

Sub SetFocusNameBox()
    Dim Res As Long
	dim hExcel, hExcel2, hNameBox as long
	hExcel = GetForegroundWindow()
	hExcel2 = FindWindowEx(hExcel, 0, "EXCEL;", vbNullString)
	hNameBox = FindWindowEx(hExcel2, 0, "combobox", vbNullString)
    Res = SetFocus(hNameBox)
End Sub

' Gets the "A1" designation for the current selection
Public Function GetCellAddress() As String

    SendKeys("+{F10}")
    Wait(0.1)
    SendKeys "a"
    Wait 0.1
    SendKeys "%r^c"
    Wait 0.1

    Dim address As String
    address = GetClipboard

    ' Remove sheet designation
    'address = address.Substring(address.IndexOf("!") + 1)

    ' Remove dollar signs
    'address = address.Replace("$", "")

    'MsgBox(address)
    GetCellAddress = address
End Function

Public Sub GoToCellAddress(address As String)
    SendKeys("^g")
    Wait(0.1)

    SendKeys(address)
    SendKeys("~")
End Sub

Public Function GetCurrentRow() As Integer

End Function

Public Function GetCurrentColumn() As Integer

End Function

Public Sub goToCell(columnLetter1, rowNumber1, Optional columnLetter2 = "", Optional rowNumber2 = "", Optional rowNumber3 = "", Optional rowNumber4 = "")
    SendKeys("^g")
    Wait(0.1)

    SendKeys(GetCellAddressFromDictation(columnLetter1, rowNumber1, columnLetter2, rowNumber2, rowNumber3, rowNumber4))
    SendKeys("~")
    Wait(0.1)
End Sub

Private Function GetCellAddressFromDictation(columnLetter1, rowNumber1, Optional columnLetter2 = "", Optional rowNumber2 = "", Optional rowNumber3 = "", Optional rowNumber4 = "") As String
    Dim column, row As String
    column = Split(columnLetter1, "\")(0)
    If columnLetter2 <> "" Then
        column = column & Split(columnLetter2, "\")(0)
    End If

    row = numberNameToDigit(Split(rowNumber1, "\")(0))

    If rowNumber2 <> "" Then
        row = row & numberNameToDigit(Split(rowNumber2, "\")(0))
    End If
    If rowNumber3 <> "" Then
        row = row & numberNameToDigit(Split(rowNumber3, "\")(0))
    End If
    If rowNumber4 <> "" Then
        row = row & numberNameToDigit(Split(rowNumber4, "\")(0))
    End If

    GetCellAddressFromDictation = column & row

End Function

Public Sub GoToLastCell()
    SendKeys("^g")
    Wait(0.1)
    SendKeys("~")
End Sub

Public Sub SelectThrough(columnLetter1, rowNumber1, Optional columnLetter2 = "", Optional rowNumber2 = "", Optional rowNumber3 = "", Optional rowNumber4 = "")
    goToCell(columnLetter1, rowNumber1, columnLetter2, rowNumber2, rowNumber3, rowNumber4)

    SendKeys("^g")
    Wait(0.1)
    SendKeys("{Right}:")
    SendKeys(GetCellAddressFromDictation(columnLetter1, rowNumber1, columnLetter2, rowNumber2, rowNumber3, rowNumber4))
    SendKeys("~")

End Sub

Public Sub AutoSeries(direction As String, startValue As Integer, stopValue As Integer)

    Dim stepValue As Integer
    If startValue < stopValue Then
        stepValue = 1
    Else
        stepValue = -1
    End If

    Dim i As Integer
    For i = startValue To stopValue Step stepValue
        SendKeys(CStr(i))
        Select Case LCase(direction)
            Case "up"
                SendKeys("{Up}")
            Case "down"
                SendKeys("{Down}")
            Case "left"
                SendKeys("{Left}")
            Case "right"
                SendKeys("{Right}")
            Case Else
                Exit Sub
        End Select
        'Wait(0.1)
    Next

	' Select last entered number
	SendKeys("{up}")
	
    ' Go Back
'    For i = startValue To stopValue Step stepValue
'        Select Case LCase(direction)
'            Case "up"
'                SendKeys("{Down}")
'            Case "down"
'                SendKeys("{Up}")
'            Case "left"
'                SendKeys("{Right}")
'            Case "right"
'                SendKeys("{Left}")
'            Case Else
'                Exit Sub
'        End Select
'        'Wait(0.1)
'    Next
End Sub

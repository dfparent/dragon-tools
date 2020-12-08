'#Uses "utilities.bas"
'#Uses "imports.bas"
'option explicit

Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

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
    address = Clipboard

    ' Remove sheet designation
    'address = address.Substring(address.IndexOf("!") + 1)

    ' Remove dollar signs
    'address = address.Replace("$", "")

    'MsgBox(address)

End Function

Public Function GetCurrentRow() As Integer

End Function

Public Function GetCurrentColumn() As Integer

End Function

Public Sub goToCell(columnLetter1, rowNumber1, Optional columnLetter2 = "", Optional rowNumber2 = "", Optional rowNumber3 = "", Optional rowNumber4 = "")
    SendKeys("^g")
    Wait(0.1)

    SendKeys(GetCellAddressFromDictation(columnLetter1, rowNumber1, columnLetter2, rowNumber2, rowNumber3, rowNumber4))
    SendKeys("~")
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

Public Sub PickColor(color As String)
    SendKeys("{Home}")

    Select Case color
        Case "Black"
            SendKeys("{Right}~")
            Exit Sub
        Case "Grey"
            SendKeys("{Right 2}~")
            Exit Sub
        Case "White"
            SendKeys("{Enter}")
            Exit Sub
    End Select

    SendKeys("{Home}{Down 6}")

    Select Case color
        Case "Green"
            SendKeys("{Right 5}{Enter}")
        Case "Light Green"
            SendKeys("{Right 4}{Enter}")
        Case "Blue", "Light Blue", "Teal"
            SendKeys("{Right 6}{Enter}")
        Case "Dark Blue"
            SendKeys("{Right 7}{Enter}")
        Case "Navy Blue"
            SendKeys("{Right 8}{Enter}")
        Case "Maroon"
            SendKeys("{Enter}")
        Case "Yellow"
            SendKeys("{Right 3}{Enter}")
        Case "Purple"
            SendKeys("{Right 9}{Enter}")
        Case "Red"
            SendKeys("{Right}{Enter}")
        Case "Orange"
            SendKeys("{Right 2}{Enter}")
    End Select
End Sub
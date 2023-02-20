'#Uses "utilities.bas"
Option Explicit On

' dataPairs is an object from the MemoryForMacros utility.  It contains a collection of name value pairs and the values contain arrays.
Public Function SelectDataDialog(attemptedMatch As String, dataPairs As Object, dataName As String, ByRef selectedData As String) As Boolean

    SelectDataDialog = True

    ' convert from MemoryForMacros object to formatted array
    Dim showData() As String
    Dim pair As Object
    Dim i As Integer
    i = 0

    Dim bestMatchIndex As Integer
    Dim bestMatchCharCount As Integer
    Dim factor As Integer
    bestMatchIndex = -1
    bestMatchCharCount = 0
    factor = 0

    For Each pair In dataPairs
        ReDim Preserve showData(i)
        showData(i) = pair.Key & " => " & pair.Value(0)
        factor = GetBestMatchFactor(attemptedMatch, pair.Key)
        If factor > bestMatchCharCount Then
            bestMatchCharCount = factor
            bestMatchIndex = i
        End If

        i = i + 1
    Next

    Begin Dialog UserDialog 1000, 371, "Select " & dataName, .DialogFunc ' %GRID:10,7,1,1
    ListBox 20, 56, 950, 273, showData(), .lstPaths, 3
  Text 20, 35, 280, 14, "Which " & dataName & " do you want to use?", .Text1
  Text 20, 14, 60, 14, "I heard:  ", .Text2
  Text 90, 14, 870, 14, attemptedMatch, .txtDictation
  OKButton 410, 343, 90, 21, .cmdOK
  CancelButton 520, 343, 80, 21, .cmdCancel
 End Dialog
        
    Dim dlg As UserDialog
    dlg.lstPaths = bestMatchIndex
    If Dialog(dlg) = 0 Then
        SelectDataDialog = False
        Exit Function
    End If

    ' Update return value
    Dim selectedValue As String
    selectedValue = showData(dlg.lstPaths)

    ' Parse out path
    'msgbox(selectedValue)
    selectedData = Mid(selectedValue, InStr(selectedValue, " => ") + 4)
    'msgbox(selectedData)

    SelectDataDialog = True
End Function

' Returns true if perfect match
Private Function GetBestMatchFactor(attemptedMatch As String, matchString As String) As Integer
    Dim i As Integer

    attemptedMatch = LCase(attemptedMatch)
    matchString = LCase(matchString)

    Dim attemptedMatchLen As Integer
    Dim matchStringLen As Integer

    attemptedMatchLen = Len(attemptedMatch)
    matchStringLen = Len(matchString)

    Dim factor As Integer
    factor = 0

    For i = 1 To attemptedMatchLen

        If i > matchStringLen Then
            GetBestMatchFactor = factor
            Exit Function
        End If

        If Mid(attemptedMatch, i, 1) <> Mid(matchString, i, 1) Then
            GetBestMatchFactor = factor
            Exit Function
        End If

        factor = i
    Next

    GetBestMatchFactor = factor

End Function

Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
    Select Case Action
        Case 1 ' Initialization
            Exit Function

        Case 2 ' Listbox changed, button click, etc.
            Select Case DlgItem
                Case "cmdOK"
                    DialogFunc = False

                Case "cmdCancel"
                    DialogFunc = False

                Case Else
                    DialogFunc = True ' Don't close the dialog box yet

            End Select


        Case 3 ' text changed
            ' N/A
            Exit Function

        Case 4 ' Receive focus
            Exit Function

        Case 5 ' idle processing
            DialogFunc = False  ' Stop receiving idle calls
            Exit Function

        Case 6 ' Function key pressed
            Exit Function
    End Select

End Function

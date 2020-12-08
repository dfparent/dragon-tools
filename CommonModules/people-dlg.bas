Option Explicit On

' showpeople must be an array
Public Function ShowPersonDialog(match As String, showPeople() As String, ByRef selectedName As String) As Boolean
    Dim prompt As String
    prompt = "Multiple people match """ & UCase(match) & """"
    ShowPersonDialog = True

    Dim displayPeople() As String
    ReDim displayPeople(UBound(showPeople))
    Dim i As Integer
    For i = 0 To UBound(showPeople)
        displayPeople(i) = CStr(i + 1) & ": " & showPeople(i)
    Next i

    Begin Dialog UserDialog 340, 238, "Choose a Person", .DialogFunc ' %GRID:10,7,1,1
    Text 40, 7, 240, 14, prompt, .lblPrompt
      Text 70, 28, 200, 14, "Which person do you want?", .lblPrompt2
      Text 20, 49, 300, 14, "Say ""Choose"" followed by the person's number", .lblPrompt3
      ListBox 20, 84, 300, 105, displayPeople(), .lstPeople
      OKButton 70, 210, 90, 21, .cmdOK
      CancelButton 180, 210, 90, 21, .cmdCancel
    End Dialog
    
    Dim dlg As UserDialog
    If Dialog(dlg) = 0 Then
        ShowPersonDialog = False
        Exit Function
    End If

    selectedName = showPeople(dlg.lstPeople)
    ShowPersonDialog = True
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

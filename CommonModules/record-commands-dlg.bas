Option Explicit On

Private commandText As String

Public Function PromptForCommandDialog(ByRef newCommand As String) As Boolean

    PromptForCommandDialog = True

    Begin Dialog UserDialog 520, 105 ' %GRID:10,7,1,1
    Text 10, 7, 250, 14, "Please say or type in a command:", .lblPrompt
    TextBox 10, 28, 490, 21, .txtCommand
    PushButton 170, 70, 80, 21, "OK", .btnOK
    PushButton 260, 70, 80, 21, "Cancel", .btnCancel
    End Dialog
    
    Dim dlg As UserDialog
    If Dialog(dlg) = 0 Then
        PromptForCommandDialog = False
        Exit Function
    End If

    newCommand = commandText

End Function

Private Function DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
    Select Case Action
        Case 1 ' Initialization
            Exit Function

        Case 2 ' Listbox changed, button click, etc.
            Select Case DlgItem
                Case "btnOK"
                    commandText = DlgText("txtCommand")
                    DialogFunc = False

                Case "btnCancel"
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

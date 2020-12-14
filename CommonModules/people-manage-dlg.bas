'#Uses "utilities.bas"
Option Explicit On

Dim thePeople() As String

' showpeople must be an array
Public Function ShowManagePeopleDialog(showPeople() As String, ByRef newPeople() As String) As Boolean

    ShowManagePeopleDialog = True

    thePeople = showPeople

    Begin Dialog UserDialog 400, 371, "Manage People", .DialogFunc ' %GRID:10,7,1,1
    ListBox 20, 35, 360, 252, thePeople(), .lstPeople, 2
      Text 20, 14, 200, 14, "People:", .Text1
      OKButton 110, 343, 90, 21, .cmdOK
      CancelButton 210, 343, 80, 21, .cmdCancel
      PushButton 110, 294, 90, 21, "&Add", .cmdAdd
      PushButton 210, 294, 80, 21, "&Remove", .cmdRemove
   End Dialog
 
    Dim dlg As UserDialog
    If Dialog(dlg) = 0 Then
        ShowManagePeopleDialog = False
        Exit Function
    End If

    ' Update people list
    newPeople = thePeople

    ShowManagePeopleDialog = True
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

                Case "cmdAdd"
                    DialogFunc = True
                    Dim value As String
                    value = InputBox("Enter the first and last name of the new person:", "Add Person")
                    If value = "" Then
                        Exit Function
                    End If
                    ReDim Preserve thePeople(UBound(thePeople) + 1)
                    thePeople(UBound(thePeople)) = value
                    DlgListBoxArray "lstPeople", thePeople

                Case "cmdRemove"
                    DialogFunc = True
                    If MsgBox("Remove " & DlgText("lstPeople") & "?", vbYesNo, "Remove Person") = vbYes Then
                        Dim index As Integer
                        index = FindArrayElement(DlgText("lstPeople"), thePeople)
                        If index = -1 Then
                            MsgBox("Something went wrong.  Could not remove person.  Please close the dialog box and try again.")
                            Exit Function
                        End If
                        thePeople(index) = ""
                        DlgListBoxArray "lstPeople", thePeople
                    End If

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

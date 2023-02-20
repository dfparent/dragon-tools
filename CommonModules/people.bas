'#Uses "people-dlg.bas"
'#Uses "people-manage-dlg.bas"
'#Uses "cache.bas"
'#Uses "utilities.bas"
'#Language "WWB.NET"
Option Explicit On

Private people As Object
Private fileName As String = ""

' Search can be initials or first name
Public Function getPerson(search As String) As String

    ' On Error GoTo ErrorHandle
    people = GetPeople()

    ' Search by initials.  Did the user speak initials?
    Dim testName As String
    testName = getPersonByInitials(search)
    If testName <> "" Then
        Return testName
    End If

    ' Try searching for the name as spoken
    Dim pair As Object
    Dim aName As String
    Dim names As New System.Collections.Generic.List(Of String)

    Dim nameParts() As String
    nameParts = search.Split()
    Dim firstNameLength As Integer
    firstNameLength = nameParts(0).Length

    Dim aNameParts() As String

    For Each pair In people
        For Each aName In pair.Value
            ' Search for name as is
            If aName = search Then
                ' Dictation was correct.  No op
                Return search
            End If

            aNameParts = aName.Split()

            ' Does first name match?
            If aNameParts(0) = nameParts(0) Then
                names.Add(aName)
            ElseIf aNameParts.Length >= 2 And nameParts.Length >= 2 Then
                ' Do initials match?
                If aNameParts(0).Chars(0) = nameParts(0).Chars(0) And aNameParts(1).Chars(0) = nameParts(1).Chars(0) Then
                    names.Add(aName)
                End If
            End If
        Next
    Next

    If names.Count = 0 Then
        'MsgBox("There is no person with initials or name matching """ & search & """.  Try again.  You can say a person's initials, first name or full name.")
        Beep
        Return search
    End If

    If names.Count = 1 Then
        Return names(0)
        Exit Function
    End If

    ' User needs to choose
    Dim name As String
    Dim namesArray() As String
    namesArray = names.ToArray()
    If Not ShowPersonDialog(search, namesArray, name) Then
        ' user cancelled
        Return ""
    Else
        Return name
    End If

    Exit Function

ErrorHandle:
    Msgbox("Error: " & err.description)
    Return ""

End Function

Public Function GetFirstName(search As String) As String
    Dim name As String
    name = getPerson(search)
    If name = "" Then
        GetFirstName = ""
        Exit Function
    End If

    GetFirstName = Split(name, " ")(0)

End Function

' Returns Either a string if a single person, or an array if multiple
Private Function getPersonByInitials(ByVal initials As String) As String
    On Error GoTo ErrorHandler
    people = GetPeople()
    initials = DictationToKeystrokes(initials)
    'msgbox(initials)
    If Not people.ContainsKey(initials) Then
        'MsgBox("There is no person with initials '" & initials & "'.")
        Return ""
    End If

    Dim name As String
    Dim matchingPeople() As String
    matchingPeople = people(initials)
    If Ubound(matchingPeople) > 0 Then
        If ShowPersonDialog(initials, matchingPeople, name) Then
            Return name
        Else
            Return ""
        End If
    Else
        Return matchingPeople(0)
    End If


ErrorHandler:
    Msgbox("Error: " & err.description)
    Return ""
End Function

Public Sub ManagePeople()
    On Error GoTo ErrorHandler

    people = GetPeople()

    Dim newPeople() As String
    Dim peopleList As New System.Collections.Generic.List(Of String)
    Dim peopleArray() As String
    Dim pair As Object
    Dim aName As String

    For Each pair In people
        For Each aName In pair.Value
            peopleList.Add(aName)
        Next
    Next

    peopleArray = peopleList.ToArray()
    If Not ShowManagePeopleDialog(peopleArray, newPeople) Then
        ' User Cancelled
        Exit Sub
    End If

    'msgbox(ArrayToString(newPeople, 0))

    ' Update people list and save
    people.Clear()

    Dim initials As String
    Dim parts() As String
    Dim part As String
    Dim valueArray() As String
    For Each aName In newPeople
        If aName = "" Then
            Continue For
        End If

        parts = Split(aName, " ")
        initials = ""
        For Each part In parts
            initials = initials & part.Substring(0, 1).ToLower()
        Next
        If Not people.ContainsKey(initials) Then
            people.Add(initials, {aName})
        Else
            '  Add to existing array
            valueArray = people(initials)
            ReDim Preserve valueArray(UBound(valueArray) + 1)
            valueArray(UBound(valueArray)) = aName
            'msgbox(ArrayToString(valueArray, 0, -1, ","))
            people(initials) = valueArray
        End If
    Next

    SavePeople


    Exit Sub

ErrorHandler:
    Msgbox("Error: " & err.description)
    Return
End Sub


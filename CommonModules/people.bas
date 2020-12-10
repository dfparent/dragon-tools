'#Uses "people-dlg.bas"
'#Uses "people-manage-dlg.bas"
'#Uses "cache.bas"
'#Uses "utilities.bas"
'#Language "WWB.NET"
Option Explicit On

Private fileName As String = ""

' Search can be initials or first name
Public Function getPerson(search As String) As String
    Dim name As String
    name = getPersonByInitials(search)
    If name <> "" Then
        Return name
    End If

    name = getPersonByFirstName(search)
    If name = "" Then
        MsgBox("There is no person with initials or first name matching """ & search & """.")
        Return ""
    End If

    Return name

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
    Dim people As Object
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

Private Function getPersonByFirstName(ByVal first As String) As String
    On Error GoTo ErrorHandler
    Dim people As Object
    people = GetPeople()
    Dim pair As Object
    Dim aName As String
    Dim names As New System.Collections.Generic.List(Of String)
    Dim theLength As Integer
    theLength = Len(first)
    For Each pair In people
        For Each aName In pair.Value
            If Left(aName, theLength) = first Then
                names.Add(aName)
            End If
        Next
    Next

    If names.Count = 0 Then
        'MsgBox("There is no person with first name '" & first & "'.")
        getPersonByFirstName = ""
        Exit Function
    End If

    If names.Count = 1 Then
        getPersonByFirstName = names(0)
        Exit Function
    End If

    ' User needs to choose
    Dim name As String
    If ShowPersonDialog(first, names.ToArray(), name) = 0 Then
        ' user cancelled
        getPersonByFirstName = ""
    Else
        getPersonByFirstName = name
    End If

    Exit Function

ErrorHandler:
    Msgbox("Error: " & err.description)
    Return ""
End Function

Public Sub ManagePeople()
    On Error GoTo ErrorHandler

    Dim people As Object
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


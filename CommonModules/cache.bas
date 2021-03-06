'#Uses "mouse.bas"
'#Uses "paths.bas"
'#Language "WWB.NET"
Option Explicit On

Private Const DATA_PATH = "C:\Users\KnowBrainer\CommonModules\Data"
Private memoryApp As Object

Private people As Object
Private peopleFilename As String = DATA_PATH & "\people.txt"

Private paths As Object
Private files As Object
Private pathsFilename As String = DATA_PATH & "\paths-" & System.Environment.MachineName & ".txt"
Private commonPathsFilename As String = DATA_PATH & "\paths.txt"

Private urls As Object
Private urlsFilename As String = DATA_PATH & "\urls.txt"

Private menus As Object
Private menusFilename As String = DATA_PATH & "\menus.txt"

Private touches As Object
Private touchesFileRoot As String = DATA_PATH & "\touch-locations"

Private apps As Object
Private appsFileName As String = DATA_PATH & "\apps-" & System.Environment.MachineName & ".txt"

Private snippets As Object
Private snippetsFileName As String = DATA_PATH & "\snippets.txt"

Private values As Object

Private commandManager As Object

Private peopleDictName As String = "people"
Private pathsDictName As String = "paths"
Private filesDictName As String = "files"
Private urlsDictName As String = "urls"
Private menusDictName As String = "menus"
Private touchesDictName As String = "touch-locations"
Private appsDictName As String = "apps"
Private snippetsDictName As String = "snippets"

' These methods utilize the MemoryForMacros application which loads data and keeps it resident in memory for use across macro calls.
Private Function LoadCache() As Boolean
    On Error GoTo ErrorHandler

    ' The Memory cache keeps loaded data in memory in between macro calls, even though KnowBrainer releases all references in between calls.
    If memoryApp Is Nothing Then
        memoryApp = CreateObject("MemoryForMacros.Application")
        If memoryApp Is Nothing Then
            MsgBox("Failed to create MemoryForMacros Application object.")
            Return False
        End If
        memoryApp.KeepAlive = True

        ' People
        If Not memoryApp.IsDictionaryLoaded(peopleDictName) Then
            memoryApp.LoadDictionaries(peopleFilename)
        End If

        people = memoryApp.GetDictionary(peopleDictName)
        If people.Count = 0 Then
            MsgBox("Failed to load people from " & peopleFilename)
        End If

        ' paths & files
        If Not (memoryApp.IsDictionaryLoaded(pathsDictName) And memoryApp.IsDictionaryLoaded(filesDictName)) Then
            memoryApp.LoadDictionaries(commonPathsFilename)
            memoryApp.LoadDictionaries(pathsFilename, True) ' Load with append
        End If

        paths = memoryApp.GetDictionary(pathsDictName)
        files = memoryApp.GetDictionary(filesDictName)
        If paths.Count = 0 Then
            MsgBox("Failed to load paths from " & pathsFilename & " and/or " & commonPathsFilename)
        End If

        ' URLs
        If Not memoryApp.IsDictionaryLoaded(urlsDictName) Then
            memoryApp.LoadDictionaries(urlsFilename)
        End If

        urls = memoryApp.GetDictionary(urlsDictName)
        If urls.Count = 0 Then
            MsgBox("Failed to load URLs from " & urlsFilename)
        End If

        ' Menus

        ' The names for dictionaries for menus is data-driven (process name)
        ' We can use the common menus dictionary, Though.
        If Not memoryApp.IsDictionaryLoaded(menusDictName) Then
            memoryApp.LoadDictionaries(menusFilename)
        End If

        menus = memoryApp.GetDictionary(menusDictName)
        If menus.Count = 0 Then
            MsgBox("Failed to load Menus from " & menusFilename)
        End If

        ' Apps
        If Not memoryApp.IsDictionaryLoaded(appsDictName) Then
            memoryApp.LoadDictionaries(appsFileName)
        End If

        apps = memoryApp.GetDictionary(appsDictName)
        If apps.Count = 0 Then
            MsgBox("Failed to load apps from " & appsFileName)
        End If

        ' Snippets
        If Not memoryApp.IsDictionaryLoaded(snippetsDictName) Then
            memoryApp.LoadDictionaries(snippetsFileName)
        End If

        snippets = memoryApp.GetDictionary(snippetsDictName)
        If snippets.Count = 0 Then
            MsgBox("Failed to load snippets from " & snippetsFileName)
        End If

        ' Values
        values = memoryApp.GetValueDictionary()

        commandManager = memoryApp.GetDelayedCommandManager()


    End If

    Return True

ErrorHandler:
    MsgBox("Error in LoadCache: " & err.Description)
    Return False
End Function

Public Sub UnloadCache()
    If LoadCache() Then
        memoryApp.Unload()
    End If

    memoryApp = Nothing
End Sub

' Use this to store individual values in between macro runs.
' These values are not persisted and will be lost the next time the MemoryForMacros process is terminated.
' Any number of values for any purpose can be stored in this dictionary.
Public Sub AddCacheValue(key As String, valueArray() As String)
    On Error GoTo ErrorHandler

    If Not LoadCache() Then
        Exit Sub
    End If

    If values.ContainsKey(key) Then
        MsgBox("Cannot add values for key """ & key & """.  There is already a value array keyed to that name.")
        Exit Sub
    End If

    values.Add(key, valueArray)

    Exit Sub

ErrorHandler:
    MsgBox("Error in AddValue: " & err.description)


End Sub

' Updates the value if it exists.  Adds the value if it does not exist.
Public Sub UpdateCacheValue(key As String, valueArray() As String)
    If CacheValueExists(key) Then
        RemoveCacheValue(key)
    End If

    AddCacheValue(key, valueArray)

End Sub

' Checks to see if the value dictionary contains values tied to a particular key.
Public Function CacheValueExists(key As String) As Boolean
    On Error GoTo ErrorHandler

    If Not LoadCache() Then
        CacheValueExists = False
        Exit Function
    End If

    CacheValueExists = values.ContainsKey(key)

    Exit Function

ErrorHandler:
    MsgBox("Error in CacheValueExists: " & err.description)
End Function

Public Function GetCacheValue(key As String) As String()
    On Error GoTo ErrorHandler

    If Not LoadCache() Then
        GetCacheValue = Nothing
        Exit Function
    End If

    If Not values.ContainsKey(key) Then
        MsgBox("Cannot retrieve values for key """ & key & """.  There is no value keyed to that name.")
        Exit Function
    End If

    GetCacheValue = values(key)

    Exit Function

ErrorHandler:
    MsgBox("Error in GetCacheValue: " & err.description)


End Function

Public Sub RemoveCacheValue(key As String)
    On Error GoTo ErrorHandler

    If Not LoadCache() Then
        Exit Sub
    End If

    values.Remove(key)

    Exit Sub

ErrorHandler:
    MsgBox("Error in RemoveCacheValue: " & err.description)
End Sub

Public Function GetPeople() As Object
    On Error GoTo ErrorHandler

    If Not LoadCache() Then
        GetPeople = Nothing
        Exit Function
    End If

    GetPeople = people
    Exit Function

ErrorHandler:
    MsgBox("Error in GetPeople: " & err.description)

End Function

Public Sub SavePeople()
    On Error GoTo ErrorHandler

    If Not LoadCache() Then
        Exit Sub
    End If

    'msgbox(ArrayToString(people("kl"), 0))

    memoryApp.SaveDictionaries(peopleFilename, {peopleDictName})

    Exit Sub

ErrorHandler:
    MsgBox("Error in SavePeople: " & err.description)
End Sub

Public Function GetPaths() As Object
    On Error GoTo ErrorHandler

    If Not LoadCache() Then
        GetPaths = Nothing
        Exit Function
    End If

    GetPaths = paths
    Exit Function

ErrorHandler:
    MsgBox("Error in GetPaths: " & err.description)

End Function

Public Function GetFiles() As Object
    On Error GoTo ErrorHandler

    If Not LoadCache() Then
        GetFiles = Nothing
        Exit Function
    End If

    GetFiles = files
    Exit Function

ErrorHandler:
    MsgBox("Error in GetFiles: " & err.description)

End Function

Public Function GetURLs() As Object
    On Error GoTo ErrorHandler

    If Not LoadCache() Then
        GetURLs = Nothing
        Exit Function
    End If

    GetURLs = urls
    Exit Function

ErrorHandler:
    MsgBox("Error in GetURLs: " & err.description)

End Function

Public Function GetCommonMenus() As Object
    On Error GoTo ErrorHandler

    If Not LoadCache() Then
        GetCommonMenus = Nothing
        Exit Function
    End If

    GetCommonMenus = menus
    Exit Function

ErrorHandler:
    MsgBox("Error in GetCommonMenus: " & err.description)

End Function

Public Function GetMenus(processName As String) As Object
    On Error GoTo ErrorHandler

    If Not LoadCache() Then
        GetMenus = Nothing
        Exit Function
    End If

    ' Get menus by process name
    Dim dictionaryName As String
    dictionaryName = menusDictName & "-" & processName
    If memoryApp.IsDictionaryLoaded(dictionaryName) Then
        ' Process specific menus
        GetMenus = memoryApp.GetDictionary(dictionaryName)
    Else
        ' Common menus
        Msgbox("No menus for process " & processName)
        GetMenus = menus
    End If

    Exit Function

ErrorHandler:
    MsgBox("Error in GetMenus: " & err.description)

End Function

Public Function GetTouchesFileName(processName As String) As String
    GetTouchesFileName = touchesFileRoot & "-" & processName & ".txt"
End Function

Public Function GetTouches(processName As String) As Object
    On Error GoTo ErrorHandler

    If Not LoadCache() Then
        GetTouches = Nothing
        Exit Function
    End If

    ' Get clicks by process name
    Dim dictionaryName As String
    dictionaryName = touchesDictName & "-" & processName
    If Not memoryApp.IsDictionaryLoaded(dictionaryName) Then
        ' Try loading it
        If Len(Dir(GetTouchesFileName(processName), vbNormal)) > 0 Then
            memoryApp.LoadDictionaries(GetTouchesFileName(processName))
        End If
        If Not memoryApp.IsDictionaryLoaded(dictionaryName) Then
            ' Add empty dictionary.  Does not persist to file.
            memoryApp.AddDictionary(dictionaryName)
        End If
    End If

    GetTouches = memoryApp.GetDictionary(dictionaryName)

    Exit Function

ErrorHandler:
    MsgBox("Error in GetTouches: " & err.description)

End Function

Public Sub SaveTouchLocations(processName As String)
    On Error GoTo ErrorHandler

    If Not LoadCache() Then
        Exit Sub
    End If


    Dim dictionaryName As String
    dictionaryName = touchesDictName & "-" & processName

    If Not memoryApp.IsDictionaryLoaded(dictionaryName) Then
        ' Does dictionary file exist?
        If Len(Dir(GetTouchesFileName(processName), vbNormal)) > 0 Then
            ' Try loading it
            memoryApp.LoadDictionaries(GetTouchesFileName(processName))
        End If
        If Not memoryApp.IsDictionaryLoaded(dictionaryName) Then
            ' Add it
            memoryApp.AddDictionary(dictionaryName)
        End If
    End If


    ' Save dictionary
    memoryApp.SaveDictionaries(touchesFileRoot & "-" & processName & ".txt", {dictionaryName})

    Exit Sub

ErrorHandler:
    MsgBox("Error in SaveTouchLocations: " & err.description)
End Sub

' For spoken commands, set spoken = true
' For typed commands, set spoken = false
Public Sub AddDelayedCommand(command As String, spoken As Boolean, delay As Integer)
    On Error GoTo ErrorHandler

    If Not LoadCache() Then
        Exit Sub
    End If

    If spoken Then
        commandManager.AddCommand(command, 0, delay)
    Else
        commandManager.AddCommand(command, 1, delay)
    End If


    Exit Sub

ErrorHandler:
    MsgBox("Error in AddDelayedCommand: " & err.description)
End Sub

Public Sub ClearDelayedCommands()
    On Error GoTo ErrorHandler

    If Not LoadCache() Then
        Exit Sub
    End If

    commandManager.ClearCommands()

    Exit Sub

ErrorHandler:
    MsgBox("Error in ClearDelayedCommands: " & err.description)

End Sub

Public Sub RunDelayedCommands()
    On Error GoTo ErrorHandler

    If Not LoadCache() Then
        Exit Sub
    End If

    commandManager.StartCommands()

	System.Threading.Thread.Sleep(100)
    Exit Sub

ErrorHandler:
    MsgBox("Error in StartCommands: " & err.description)

End Sub

Public Sub KillDelayedCommands()
    On Error GoTo ErrorHandler

    If Not LoadCache() Then
        Exit Sub
    End If

    commandManager.StopCommands()

    Exit Sub

ErrorHandler:
    MsgBox("Error in KillDelayedCommands: " & err.description)

End Sub

Public Function GetApps() As Object
    On Error GoTo ErrorHandler

    If Not LoadCache() Then
        GetApps = Nothing
        Exit Function
    End If

    GetApps = apps
    Exit Function

ErrorHandler:
    MsgBox("Error in GetApps: " & err.description)

End Function

Public Function GetSnippets() As Object
    On Error GoTo ErrorHandler

    If Not LoadCache() Then
        GetSnippets = Nothing
        Exit Function
    End If

    GetSnippets = snippets

    Exit Function

ErrorHandler:
    MsgBox("Error in GetSnippets: " & err.description)

End Function
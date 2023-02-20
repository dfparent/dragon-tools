#Region "Imports directives"

Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Collections
Imports System.Text
Imports MemoryForMacros

#End Region

''' <summary>
''' Provides top level access to all MemoryForMacros functionality.  
''' To use MemoryForMacros, create an instance of <c>Application</c>.
''' </summary>
<ComClass(Application.ClassId, Application.InterfaceId, Application.EventsId)>
Public Class Application
    Inherits ReferenceCountedObject

#Region "COM Registration"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "432b9da2-4792-449e-8f2a-3754e84aec5c"
    Public Const InterfaceId As String = "e7f8fdb6-ad0c-49e9-afb9-e12f66450aba"
    Public Const EventsId As String = "15078b6d-99ff-453c-b6fa-5cf544516ec4"

    ' These routines perform the additional COM registration needed by 
    ' the service.

    <ComRegisterFunction(), EditorBrowsable(EditorBrowsableState.Never)>
    Public Shared Sub Register(ByVal t As Type)
        Try
            COMHelper.RegasmRegisterLocalServer(t)
        Catch ex As Exception
            Console.WriteLine(ex.Message) ' Log the error
            Throw ex ' Re-throw the exception
        End Try
    End Sub

    <EditorBrowsable(EditorBrowsableState.Never), ComUnregisterFunction()>
    Public Shared Sub Unregister(ByVal t As Type)
        Try
            COMHelper.RegasmUnregisterLocalServer(t)
        Catch ex As Exception
            Console.WriteLine(ex.Message) ' Log the error
            Throw ex ' Re-throw the exception
        End Try
    End Sub
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Private Shared keepAliveValue As Boolean = False
    Private Shared dictionaries As New System.Collections.Generic.Dictionary(Of String, MemoryForMacros.Dictionary)
    Private Shared valueDictionary As New Dictionary
    Private Shared m_commandManager As New DelayedCommandManager

    Friend Shared COMMENT_CHAR As New Token(";")
    Friend Shared START_REFERENCE As New Token("${")
    Friend Shared END_REFERENCE As New Token("}")
    Friend Shared REFERENCE_SEPARATOR As New Token(":")
    Friend Shared START_DICT As New Token("[")
    Friend Shared END_DICT As New Token("]")
    Friend Shared KEY_VALUE_SEPARATOR As New Token("=")
    Friend Shared VALUE_SEPARATOR As New Token(",")
    Friend Shared NEW_LINE As New Token("\n")

    Friend Shared Function GetDictionaries() As Dictionary(Of String, MemoryForMacros.Dictionary)
        Return dictionaries
    End Function

    ''' <summary>
    ''' Used to keep MemoryForMacros running in between macro runs.
    ''' </summary>
    ''' <seealso cref="Application.Unload()"/>
    ''' <value>True or False</value>>
    ''' <remarks>
    ''' Normally, the EXE will self terminate after the last reference to Application is released.  
    ''' When KeepAlive is TRUE, the EXE will remain active even after the last reference is released.
    ''' It will remain in memory indefinitely.  To allow MemoryForMacros to shutdown, set Keep Alive to False
    ''' and release all references.  
    ''' </remarks>
    Public Property KeepAlive As Boolean
        Get
            Return keepAliveValue
        End Get
        Set(value As Boolean)
            keepAliveValue = value
            ExeCOMServer.Instance.KeepAlive = value
        End Set
    End Property

    ''' <summary>
    ''' Returns a dictionary that can be used by macros to store individual values.  
    ''' </summary>
    ''' <returns><see cref="Dictionary"/></returns>
    ''' <remarks>
    ''' Returns a special dictionary that can be used to store individual values.  This dictionary is not loaded from a file, nor saved to a file.
    ''' It is simply used to store any value in between macro runs, such as state, IDs, status, etc.
    ''' </remarks>
    Public Function GetValueDictionary() As Dictionary
        Return valueDictionary
    End Function

    ''' <summary>
    ''' Returns the <c>Dictionary</c> object keyed to the string provided in the <c>name</c> parameter.
    ''' </summary>
    ''' <param name="name">A unique string identifying the desired <c>Dictionary</c>.  Each Dictionary has a unique "name".  
    ''' You can name the <c>Dictionary</c> according to its contents, e.g., "people", "paths", "files", etc.  The <c>name</c> value is the
    ''' value provided when the dictionary was loaded or added.
    ''' </param>
    ''' <seealso cref="Application.LoadDictionaries(String, Boolean)"/>
    ''' <seealso cref="Application.AddDictionary(String)"/>
    ''' <returns>A <see cref="Dictionary"/> object, or Null/Nothing is there is no dictionary keyed to the given name.</returns>
    Public Function GetDictionary(name As String) As Dictionary
        If Not dictionaries.ContainsKey(name) Then
            Return Nothing
        Else
            Return dictionaries(name)
        End If
    End Function

    ''' <summary>
    ''' Add a new dictionary and key it to the string provided in the <c>name</c> the parameter.
    ''' </summary>
    ''' <param name="name">
    ''' The unique name identifying the new Dictionary.  The value provided in <c>name</c> must be unique among all dictionaries.
    ''' </param>
    ''' <returns>
    ''' The new <see cref="Dictionary"/> object.  If a dictionary already exists with the given <c>name</c> the existing <see cref="Dictionary"/> is returned.
    ''' </returns>
    Public Function AddDictionary(name As String) As Dictionary
        If Not dictionaries.ContainsKey(name) Then
            dictionaries.Add(name, New Dictionary())
        End If
        Return dictionaries(name)
    End Function

    ''' <summary>
    ''' Indicates whether a dictionary with the given name is currently loaded in MemoryForMacros.
    ''' </summary>
    ''' <param name="dictionaryName">The name of the <see cref="Dictionary"/> being checked.</param>
    ''' <returns>
    ''' True if the indicated Dictionary is currently loaded.  False if the indicated dictionary is not currently loaded.
    ''' </returns>
    ''' <remarks>
    ''' This method tests to see if the given Dictionary is currently loaded in memory.  It does not check to see if the Dictionary has been persisted to disk.
    ''' One possible response to a "<see langword="False"/>" return value is to load the desired Dictionary from a file.
    ''' </remarks>
    ''' <seealso cref="Application.LoadDictionaries(String, Boolean)"/>
    ''' 
    Public Function IsDictionaryLoaded(dictionaryName As String) As Boolean
        Return dictionaries.ContainsKey(dictionaryName)
    End Function

    ''' <summary>
    ''' Reads the given file and loads the Dictionary definitions contained within the file.
    ''' </summary>
    ''' <param name="fileName">The fully qualified path of the file containing Dictionary definitions.</param>
    ''' <param name="append">Indicates whether loaded dictionary entries should replace currently loaded dictionaries, or be added to currently loaded Dictionaries.
    ''' If any dictionary being loaded has the same name as an already loaded dictionary,  
    ''' the newly loaded entries will be ADDED to the existing Dictionary if this parameter is "<see langword="True"/> 
    ''' and will REPLACE the existing Dictionary if this parameter is "False".
    ''' </param>Have files
    ''' <returns>
    ''' An array of Strings containing the names of the dictionaries that were found in <c>fileName</c>.
    ''' </returns>
    ''' <exception cref="ArgumentException">
    ''' Thrown if the file provided in <c>fileName</c> does not exist.
    ''' </exception>
    ''' <exception cref="FormatException">
    ''' Thrown if the provided file contains invalidly formatted content.
    ''' </exception>
    ''' <exception cref="ApplicationException">
    ''' Thrown if any other errors occur while reading the file.
    ''' </exception>
    ''' <remarks>
    ''' In the loaded files, Dictionaries are identified By name with [...], E.g., [people] identifies a dictionary named "people".
    ''' Entries for the dictionary will follow and consist of a name/value pair, separated by "=".  The values can themselves be a comma separated list of values.  
    ''' You can comment an entire line by starting it with a ";".
    ''' 
    ''' Here is an example of a Dictionary file:
    ''' <para>
    ''' ; The "words" dictionary
    ''' [words]
    ''' cat=An animal with four legs and pointy ear.
    ''' dog=An animal with four legs and lots of drool.
    ''' </para>
    ''' <para>
    ''' ; This is a dictionary called "people"
    ''' [people]
    ''' dp=Douglas Parent
    ''' mj=Michael Jackson,Mihailia Jackson
    ''' </para>
    ''' The format of the key and the format of the value can vary, depending on the contents of the Dictionary,  as long as they follow the rules listed above.
    ''' See generated comments in a saved dictionary file for more details.
    ''' </remarks>
    Public Function LoadDictionaries(fileName As String, Optional append As Boolean = False) As String()

        Dim e As System.Exception
        Dim names As New List(Of String)()

        'MsgBox("Loading " & fileName)

        If Not My.Computer.FileSystem.FileExists(fileName) Then
            ' Return empty array
            Throw New ArgumentException("Can't load dictionary file """ & fileName & """. File does not exist.")
            'Return names.ToArray()
        End If

        Try
            Using fileReader As New System.IO.StreamReader(fileName)

                Dim fields As New List(Of String)()
                'Dim separators() As String = {KEY_VALUE_SEPARATOR.Name, VALUE_SEPARATOR.Name}
                Dim separators() As String = {VALUE_SEPARATOR.Name}
                Dim theKey As String
                Dim line As String
                Dim dict As Dictionary

                Do While fileReader.Peek() >= 0
                    Try
                        line = fileReader.ReadLine()

                        line = line.Trim()

                        If line = "" Then
                            ' Blank line
                            Continue Do
                        End If

                        ' Dictionaries defined like this:
                        ' [mydict]

                        If line.Substring(0, 1) = COMMENT_CHAR.Name Then
                            ' Comment line
                            Continue Do
                        ElseIf line.Substring(0, 1) = START_DICT.Name And line.Contains(END_DICT.Name) Then
                            ' Get dictionary name
                            line = line.Replace(START_DICT.Name, "")
                            line = line.Replace(END_DICT.Name, "")
                            names.Add(line)

                            ' Look for existing dictionary
                            If Not dictionaries.ContainsKey(line) Then
                                dict = New Dictionary()
                                dictionaries.Add(line, dict)
                            Else
                                dict = dictionaries(line)
                                If Not append Then
                                    dict.Clear()
                                End If
                            End If

                            Continue Do
                        End If

                        ' If you've gotten this far, you need to have a dictionary
                        If dict Is Nothing Then
                            Throw New FormatException("MemoryForMacros:  This dictionary file is improperly formatted (entries without a dictionary declaration, i.e. [My Dictionary]): " & fileName)
                        End If

                        ' Entry line is formatted like this:
                        '   <key>=<value 1>,<value 2>,...

                        fields.Clear()

                        'Split at first equal sign
                        Dim equalSignIndex As Integer
                        equalSignIndex = line.IndexOf(KEY_VALUE_SEPARATOR.Name)
                        If equalSignIndex < 0 Then
                            Throw New FormatException("MemoryForMacros:  The following line is formatted incorrectly (no equal sign):  " & line)
                            Continue Do
                        End If

                        theKey = line.Substring(0, equalSignIndex)

                        ' Get values
                        fields.AddRange(line.Substring(equalSignIndex + 1).Split(separators, System.StringSplitOptions.RemoveEmptyEntries))


                        If fields.Count < 1 Then
                            Throw New FormatException("MemoryForMacros:  The following line is formatted incorrectly (no values):  " & line)
                            Continue Do
                        End If

                        ' Replace NEW_LINEs
                        For i As Integer = 0 To fields.Count - 1
                            fields(i) = fields(i).Replace(NEW_LINE.Name, vbCrLf)
                        Next

                        dict.Add(theKey, fields.ToArray())

                    Catch e
                        Throw New ApplicationException("MemoryForMacros:  Failed to read line: " & e.Message)
                    End Try
                Loop
            End Using
        Catch e
            Throw New ApplicationException("MemoryForMacros:  Error processing data file: " & e.Message)
        End Try

        ' Resolve dictionary references
        For Each name As String In names
            dictionaries(name).ResolveReferences()
        Next

        Return names.ToArray()

    End Function

    ''' <summary>
    ''' Scans all dictionaries for references and replaces the references with the referred values if possible
    ''' </summary>
    ''' <remarks>
    ''' Dictionary entries can contain references to other dictionary entries.  References use the format "${&lt;dictionary name>:&lt;entry key>}".  
    ''' For example, given the following dictionary definitions:
    ''' <para>
    ''' [paths]
    ''' my path=C:\Mine
    ''' your_path=C:\Yours
    ''' my docs=${my path}\Documents
    ''' </para>
    ''' <para>
    ''' [files]
    ''' my document=${paths:my docs}\Document1.docx
    ''' </para>
    ''' The value "my docs" will resolve to "C:\Mine\Documents".
    ''' The value "my document" will resolve to "C:\Mine\Documents\Document1.docx".
    ''' As you can see the "dictionary name" part of the reference is optional.  If omitted, the referenced entry must appear in the current dictionary.
    ''' </remarks>
    Public Sub ResolveDictionaries()
        Dim pair As KeyValuePair(Of String, Dictionary)
        For Each pair In dictionaries
            pair.Value.ResolveReferences()
        Next
    End Sub

    ''' <summary>
    ''' Save the dictionaries with the given names to the given file. 
    ''' </summary>
    ''' <param name="fileName">The fully qualified path to the file in which the dictionaries will be saved.</param>
    ''' <param name="dictionaryNames">
    ''' An array of string values containing the names of the dictionaries to save.  
    ''' If a provided name is not the name of a loaded dictionary, an exception will be thrown.</param>
    ''' <exception cref="ApplicationException">If any errors occur while attempting to save the dictionaries, exception will be thrown.</exception>
    ''' <remarks>
    ''' Saved values are the raw values with reference tokens, not the dereferenced values. If the given file exists, a backup copy will be created and 
    ''' the original file's contents will be overwritten by the newly saved definitions.  NOTE: if the existing file currently has a comment section at the beginning of the file, 
    ''' the comments will be retained in the newly saved file.  All comments after the initial comment section at the top of the file will be deleted.
    ''' </remarks>
    Public Sub SaveDictionaries(fileName As String, dictionaryNames() As String)

        Dim header As New StringBuilder
        Dim backupFileName As String
        Dim e As System.Exception

        backupFileName = fileName & ".bak"

        Try
            If My.Computer.FileSystem.FileExists(backupFileName) Then
                System.IO.File.Copy(fileName, backupFileName, True)
            End If
        Catch e
            Throw New ApplicationException("Memory For Macros: Failed to backup dictionary: " & e.Message)
            Exit Sub
        End Try


        Try
            ' Read and save existing header comments
            If My.Computer.FileSystem.FileExists(fileName) Then
                Using fileReader As New System.IO.StreamReader(fileName)
                    Dim line As String
                    Dim done As Boolean = False

                    Do While Not done And fileReader.Peek() >= 0
                        line = fileReader.ReadLine()

                        If line = "" Then
                            ' Blank line.  Have we reached any comment lines yet?
                            If header.Length = 0 Then
                                ' Ignore blank line
                                Continue Do
                            Else
                                ' 1st blank line after header block: All done
                                done = True
                                Continue Do
                            End If
                        End If

                        If line.Substring(0, 1) = COMMENT_CHAR.Name Then
                            ' Comment line: add to saved header
                            header.AppendLine(line)
                        Else
                            ' Non comment line.  All done saved header
                            done = True
                        End If
                    Loop
                End Using
            End If

            Using fileWriter As System.IO.StreamWriter = New System.IO.StreamWriter(fileName, False)
                Dim line As New System.Text.StringBuilder

                fileWriter.WriteLine(header)
                fileWriter.WriteLine()

                Dim dict As Dictionary
                Dim dictName As String

                For Each dictName In dictionaryNames

                    If Not dictionaries.ContainsKey(dictName) Then
                        Throw New FormatException("Can't save dictionary """ & dictName & """.  The dictionary does not exist.")
                    End If

                    dict = dictionaries(dictName)
                    fileWriter.WriteLine(START_DICT.Name & dictName & END_DICT.Name)

                    ' Save RAW values
                    Dim pair As DictionaryPair
                    Dim value As String
                    Dim raw_values() As String

                    For Each pair In dict
                        raw_values = dict.Raw_Item(pair.Key)
                        line.Clear()
                        line.Append(pair.Key)
                        line.Append(KEY_VALUE_SEPARATOR.Name)
                        For Each value In raw_values
                            line.Append(value)
                            line.Append(VALUE_SEPARATOR.Name)
                        Next

                        ' Remove last comma
                        line.Remove(line.Length - VALUE_SEPARATOR.Length, VALUE_SEPARATOR.Length)

                        fileWriter.WriteLine(line.ToString())
                    Next

                    fileWriter.WriteLine()

                Next

            End Using
        Catch e
            Throw New ApplicationException("Failed to create the data file: " & e.Message)
            Exit Sub
        End Try

    End Sub

    ''' <summary>
    ''' Returns an object that allows you to run delayed commands.  This gets around a problem in Knowbrainer that prevents some commands from running completely.
    ''' </summary>
    ''' <param name="delay">The amount of time to delay in milliseconds.</param>
    ''' <param name="command">
    ''' The Dragon command to execute.  The command will execute as though it had been spoken.</param>
    Public Function GetDelayedCommandManager() As DelayedCommandManager
        Return m_commandManager
    End Function

    ''' <summary>
    ''' Shuts down MemoryForMacros immediately, regardless of whether any clients are still using the service.
    ''' </summary>
    ''' <remarks>
    ''' Works similarly to setting KeepAlive to false, except that it will cause the process to shut down even if there are active clients.
    ''' To unload the process more politely, set KeepAlive to false and release any references to the Application object in all clients.
    ''' </remarks>
    Public Sub Unload()
        Dim count As Integer
        KeepAlive = False
        Do
            count = ExeCOMServer.Instance.Unlock()
        Loop Until count = 0
    End Sub
End Class

''' <summary>
''' Class factory for the class SimpleObject.
''' </summary>
Friend Class ApplicationClassFactory
    Implements IClassFactory

    Public Function CreateInstance(ByVal pUnkOuter As IntPtr, ByRef riid As Guid,
                                   <Out()> ByRef ppvObject As IntPtr) As Integer _
                                   Implements IClassFactory.CreateInstance
        ppvObject = IntPtr.Zero

        If (pUnkOuter <> IntPtr.Zero) Then
            ' The pUnkOuter parameter was non-NULL and the object does 
            ' not support aggregation.
            Marshal.ThrowExceptionForHR(COMNative.CLASS_E_NOAGGREGATION)
        End If

        If ((riid = New Guid(Application.ClassId)) OrElse
            (riid = New Guid(COMNative.IID_IDispatch)) OrElse
            (riid = New Guid(COMNative.IID_IUnknown))) Then
            ' Create the instance of the .NET object
            ppvObject = Marshal.GetComInterfaceForObject(
            New Application, GetType(Application).GetInterface("_Application"))
        Else
            ' The object that ppvObject points to does not support the 
            ' interface identified by riid.
            Marshal.ThrowExceptionForHR(COMNative.E_NOINTERFACE)
        End If

        Return 0  ' S_OK
    End Function


    Public Function LockServer(ByVal fLock As Boolean) As Integer _
    Implements IClassFactory.LockServer
        Return 0  ' S_OK
    End Function

End Class


''' <summary>
''' Reference counted object base.
''' </summary>
''' <remarks></remarks>
<ComVisible(False)>
Public Class ReferenceCountedObject

    Public Sub New()
        ' Increment the lock count of objects in the COM server.
        ExeCOMServer.Instance.Lock()
    End Sub

    Protected Overrides Sub Finalize()
        Try
            ' Decrement the lock count of objects in the COM server.
            ExeCOMServer.Instance.Unlock()
        Finally
            MyBase.Finalize()
        End Try
    End Sub

End Class
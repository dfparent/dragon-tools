Imports System.Runtime.InteropServices
Imports MemoryForMacros
Imports System.Collections.Generic
Imports System.Text

<ComClass(Dictionary.ClassId, Dictionary.InterfaceId, Dictionary.EventsId)>
Public Class Dictionary
    Implements IEnumerable(Of DictionaryPair)

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "32bd6c6b-d899-4d9c-95a1-311d12f0a17a"
    Public Const InterfaceId As String = "8369b404-5bed-4b4c-a212-582e6ce288ed"
    Public Const EventsId As String = "ad0d4e4e-1268-434c-a937-fa7ce882d24b"
#End Region

    Private data As New Dictionary(Of String, String())
    Private raw_data As New Dictionary(Of String, String())

    ''' <summary>
    ''' Returns the number of entries in the dictionary.
    ''' </summary>
    ''' <returns>Integer</returns>
    Public Function Count() As Integer
        Return data.Count
    End Function

    ''' <summary>
    ''' Deletes all entries in the Dictionary.
    ''' </summary>
    Public Sub Clear()
        raw_data.Clear()
        data.Clear()
    End Sub

    ''' <summary>
    ''' Checks to see if there is an entry in the dictionary corresponding to the given "key".
    ''' </summary>
    ''' <param name="key">A unique string identifying a single dictionary value.</param>
    ''' <returns>True or False</returns>
    ''' <exception cref="ArgumentNullException"></exception>
    Public Function ContainsKey(key As String) As Boolean
        Return data.ContainsKey(key.ToLower())
    End Function

    ''' <summary>
    ''' Adds a new entry to the dictionary with the given "key" and "value".
    ''' </summary>
    ''' <param name="Key">A unique string identifying the given value.</param>
    ''' <param name="Value">An array of strings.</param>
    ''' <exception cref="ArgumentException"></exception>
    ''' <exception cref="ArgumentNullException"></exception>
    Public Sub Add(Key As String, Value As String())
        Key = Key.ToLower()
        raw_data.Add(Key, Value)
        data.Add(Key, Value)
    End Sub

    ''' <summary>
    ''' Removes the entry identified with the given "key".
    ''' </summary>
    ''' <param name="key">A unique string identifying the value to remove.</param>
    ''' <remarks>
    ''' If evaluate exist in the dictionary associated with the given "key", the key and value will be removed from the dictionary.
    ''' If there is no entry corresponding to the given "key", this method will do nothing.
    ''' </remarks>
    ''' <exception cref="ArgumentNullException"></exception>
    Public Sub Remove(key As String)
        key = key.ToLower()
        If raw_data.ContainsKey(key) Then
            raw_data.Remove(key)
        End If

        If data.ContainsKey(key) Then
            data.Remove(key)
        End If
    End Sub

    ''' <summary>
    ''' Returns an array of strings containing the keys of the dictionary.
    ''' </summary>
    ''' <returns>An array of strings.</returns>
    ''' <remarks>This method returns only the keys.  It does not return any of the entry values.  To access a collection of the values, use the Dictionary enumeration.</remarks>
    Public Function Keys() As String()
        Return data.Keys.ToArray()
    End Function

    ''' <summary>
    ''' Provides access to a single value in the Dictionary.
    ''' </summary>
    ''' <param name="key">The key associated with the value which is being retrieved or stored.</param>
    ''' <value>An array of strings.</value>
    ''' <remarks>
    ''' This property provides access to a single entry in the Dictionary.  
    ''' When reading a value, it will return the array of strings associated with the given key.
    ''' When writing it will save the given array of strings in the dictionary and key it to the provided key.
    ''' You cannot use this method to Add dictionary entries.  If you provide an invalid key, an exception will be thrown.
    ''' </remarks>
    ''' <exception cref="KeyNotFoundException"></exception>
    ''' <exception cref="ArgumentNullException"></exception>
    Default Property Item(ByVal key As String) As String()
        Get
            key = key.ToLower()
            If data.ContainsKey(key) Then
                Return data.Item(key)
            Else
                Throw New KeyNotFoundException
            End If
        End Get
        Set(ByVal value As String())
            key = key.ToLower()
            If data.ContainsKey(key) Then
                data.Item(key) = value
            Else
                Throw New KeyNotFoundException
            End If

            ' need to also update the raw array.  
            ' NOTE: This will wipe out any tokens in the raw data for this value
            '       Unless the calling code manages the tokens!
            Raw_Item(key) = value
        End Set
    End Property

    ' Provides access to a single item in the "raw" data collection
    Friend Property Raw_Item(ByVal key As String) As String()
        Get
            key = key.ToLower()
            If raw_data.ContainsKey(key) Then
                Return raw_data.Item(key)
            Else
                Throw New ArgumentOutOfRangeException
            End If
        End Get
        Set(ByVal value As String())
            key = key.ToLower()
            If raw_data.ContainsKey(key) Then
                raw_data.Item(key) = value
            Else
                Throw New ArgumentOutOfRangeException
            End If
        End Set
    End Property

    ' Resolves all references in the raw data dictionary and saves changes to the main data dictioanry.
    ' The RAW dictionary is left untouched.
    Friend Sub ResolveReferences()
        Dim startIndex As Integer = 0, endIndex As Integer = 0, dictIndex As Integer = 0
        Dim referredToken As String
        Dim referredKey As String
        Dim referredDictname As String
        Dim referredDict As Dictionary
        Dim referredValue As String
        Dim referenceError As String
        Dim dictionaries As Dictionary(Of String, MemoryForMacros.Dictionary) = Application.GetDictionaries()
        Dim pairs() As KeyValuePair(Of String, String())
        Dim pair As KeyValuePair(Of String, String())
        Dim workingValue() As String
        Dim circuitBreaker As Integer = 100 ' Rediculously large number of references in a single value.
        Dim count As Integer = 0

        ' Example references:
        '   my key=my value1,my ${reference here} value2,my value3
        '   my key2=another value and a dict reference: ${another dict:ref value}. plus a second ref: ${my key}.

        ' Need to work on a copy to avoid errors when modifying the collection
        ReDim pairs(raw_data.Count - 1)
        raw_data.ToList().CopyTo(pairs)

        Dim value As String
        Dim i As Integer
        For Each pair In pairs
            ' Make deep copy of value and work with copy
            ReDim workingValue(UBound(pair.Value))
            For i = 0 To UBound(pair.Value)
                workingValue(i) = pair.Value(i)
            Next

            For i = 0 To UBound(workingValue)
                ' Loop through value to look for possible multiple references. Each loop handles 1 reference.
                count = 0
                Do Until count > circuitBreaker
                    ' Each iteration we start over at the beginning because any changes are implemented using
                    ' A replace.
                    startIndex = workingValue(i).IndexOf(Application.START_REFERENCE.Name)
                    If startIndex = -1 Then
                        ' No more references
                        Continue For
                    End If

                    endIndex = workingValue(i).IndexOf(Application.END_REFERENCE.Name, startIndex + Application.START_REFERENCE.Length)

                    If endIndex = -1 Then
                        ' Not a reference: no closing char
                        Exit Do
                    End If

                    ' Found reference.  This extracts the token without the delimiters.
                    referredToken = workingValue(i).Substring(startIndex + Application.START_REFERENCE.Length,
                                                          endIndex - startIndex - Application.START_REFERENCE.Length)

                    ' Default to current dictionary
                    referredDict = Me
                    referredKey = referredToken
                    referenceError = ""

                    ' Is this a reference to another dictionary?
                    dictIndex = referredToken.IndexOf(Application.REFERENCE_SEPARATOR.Name)
                    If dictIndex >= 0 Then
                        ' Includes dictionary reference
                        referredDictname = referredToken.Substring(0, dictIndex)
                        If Not dictionaries.ContainsKey(referredDictname) Then
                            ' Referred dictionary does not exist. 
                            referenceError = "<Dictionary does not exist: " & referredDictname & ">"
                        Else
                            referredDict = dictionaries(referredDictname)
                            referredKey = referredToken.Substring(dictIndex + Application.REFERENCE_SEPARATOR.Length)
                        End If
                    Else
                        ' Make sure it is not a self-reference which will cause an infinite loop
                        If pair.Key = referredKey Then
                            ' Self reference!  
                            referenceError = "<Invalid self-reference>"
                        End If
                    End If

                    If referenceError <> "" Then
                        referredValue = referenceError
                    Else
                        ' Get referred value
                        If Not referredDict.ContainsKey(referredKey) Then
                            referredValue = "<Invalid entry reference: " & referredKey & ">"
                        Else
                            referredValue = referredDict.Item(referredKey)(0)
                        End If
                    End If

                    workingValue(i) = workingValue(i).Replace(Application.START_REFERENCE.Name & referredToken & Application.END_REFERENCE.Name,
                                                          referredValue)
                    count += 1
                Loop
            Next ' Each value in value array for a single dictionary item

            ' Save actual dictionary item
            Item(pair.Key) = workingValue

        Next ' Each item in dictionary
    End Sub

    ''' <summary>
    ''' An iterator that can be used to access all of the entries in the dictionary.
    ''' </summary>
    ''' <returns>IEnumerable</returns>
    ''' <remarks>You can use this enumerator by accessing the Dictionary in a "for each" block.</remarks>
    Public Function GetEnumerator() As IEnumerator(Of DictionaryPair) Implements IEnumerable(Of DictionaryPair).GetEnumerator
        Return New DictionaryEnumerator(data.GetEnumerator())
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function
End Class



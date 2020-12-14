<ComClass(DictionaryEnumerator.ClassId, DictionaryEnumerator.InterfaceId, DictionaryEnumerator.EventsId)> _
Public Class DictionaryEnumerator
    Implements IEnumerator(Of DictionaryPair)
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "db1d5f73-b214-4da7-b604-33f0d335c29c"
    Public Const InterfaceId As String = "a0198323-57d1-4bdb-b308-42e1ca98abf7"
    Public Const EventsId As String = "e5401ab8-8633-459f-aa5b-51138ed8a994"
#End Region

    Public Sub New(theEnum As Dictionary(Of String, String()).Enumerator)
        MyBase.New()
        theEnumerator = theEnum
    End Sub

    Private theEnumerator As Dictionary(Of String, String()).Enumerator

    Public ReadOnly Property Current As DictionaryPair Implements IEnumerator(Of DictionaryPair).Current
        Get
            Return New DictionaryPair(theEnumerator.Current)
        End Get
    End Property

    Private ReadOnly Property IEnumerator_Current As Object Implements IEnumerator.Current
        Get
            Return Me.Current
        End Get
    End Property

    Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext
        Return theEnumerator.MoveNext()
    End Function

    Public Sub Reset() Implements IEnumerator.Reset
        ' Does nothing Because the dictionary enumerator does not support this method
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        theEnumerator.Dispose()
    End Sub

    Protected Overrides Sub Finalize()
        Dispose()
    End Sub
End Class



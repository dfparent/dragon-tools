Imports System.Runtime.InteropServices

<ComClass(DictionaryPair.ClassId, DictionaryPair.InterfaceId, DictionaryPair.EventsId)>
Public Class DictionaryPair

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "90f2e764-8dec-41e3-bc16-5f3d259872f7"
    Public Const InterfaceId As String = "78e38698-6225-412b-a577-e2aa02e0a783"
    Public Const EventsId As String = "14415690-76c4-48f5-8784-d5a3cabf709f"
#End Region

    Public Sub New(aPair As KeyValuePair(Of String, String()))
        MyBase.New()
        thePair = aPair
    End Sub

    Private thePair As KeyValuePair(Of String, String())

    Public ReadOnly Property Key() As String
        Get
            Return thePair.Key
        End Get
    End Property

    Public ReadOnly Property Value() As String()
        Get
            Return thePair.Value
        End Get
    End Property

End Class



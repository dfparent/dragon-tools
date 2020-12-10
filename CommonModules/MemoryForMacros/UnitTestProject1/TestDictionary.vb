Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class TestDictionary

    Private m_testApp As TestApplication

    Public Sub New()
        MyBase.New()

        m_testApp = New TestApplication

    End Sub

    <TestMethod()> Public Sub TestRemove()
        Dim people As MemoryForMacros.Dictionary
        people = m_testApp.GetPeopleDictionary()

        people.Add("zz", {"Zaphod Beeblebrox"})
        people.Remove("zz")

        ' Should quietly complete when trying to remove a non existant key
        people.Remove("zz")

    End Sub
    <TestMethod()> Public Sub TestEnumeration()
        Dim people As MemoryForMacros.Dictionary
        people = m_testApp.GetPeopleDictionary()

        Dim pair As MemoryForMacros.DictionaryPair
        For Each pair In people
            For Each aValue As String In pair.Value
                If aValue Is Nothing Or aValue = "" Then
                    Assert.Fail("Bad value found in enumeration.")
                End If
            Next
        Next

    End Sub

End Class
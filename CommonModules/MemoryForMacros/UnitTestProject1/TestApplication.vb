Imports System.Text
Imports System.Threading

<TestClass()> Public Class TestApplication

    Private m_app As MemoryForMacros.Application
    Private m_appObject As Object


    Public Function GetAppObject() As MemoryForMacros.Application

        If m_app Is Nothing Then
            m_app = New MemoryForMacros.Application
            If m_app Is Nothing Then
                Assert.Fail("Failed to create memory for macros application object.")
            End If
        End If

        Return m_app

    End Function

    Public Function GetAppObjectLateBound() As Object
        If m_appObject Is Nothing Then
            m_appObject = CreateObject("MemoryForMacros.Application")
            If m_appObject Is Nothing Then
                Assert.Fail("Failed to create late bound memory for macros application object.")
            End If
        End If
        Return m_appObject
    End Function

    Public Function GetPeopleDictionary() As MemoryForMacros.Dictionary
        Dim app As MemoryForMacros.Application
        Dim people As MemoryForMacros.Dictionary
        app = GetAppObject()
        app.LoadDictionaries("..\..\data1.txt")
        people = app.GetDictionary("people")
        If people Is Nothing Then
            Assert.Fail("Failed to load/get people collection.")
        End If

        Return people
    End Function

    Public Sub ValueDictionary()
        Dim app As MemoryForMacros.Application
        app = GetAppObject()

        Dim testKey As String = "Test Key"
        Dim testValues() As String = {"test value!"}
        Dim valueDictionary As MemoryForMacros.Dictionary
        valueDictionary = app.GetValueDictionary()
        valueDictionary.Add(testKey, testValues)

        Dim values() As String
        values = valueDictionary(testKey)

        Assert.AreEqual(testValues.Count, values.Count())
        Assert.AreEqual(testValues(0), values(0))
    End Sub
    <TestMethod()> Public Sub LoadDictionaries()
        Dim app As MemoryForMacros.Application
        app = GetAppObject()

        ' Test load dictionary/ get dictionary
        Dim people As MemoryForMacros.Dictionary
        people = GetPeopleDictionary()

    End Sub


    <TestMethod()>
    <ExpectedException(GetType(ArgumentException), "Exception not properly thrown when reading a nonexistent file.")>
    Public Sub LoadBadDictionaries()
        Dim app As MemoryForMacros.Application
        app = GetAppObject()
        ' Test bad input file name
        app.LoadDictionaries("..\..\data1-bad.txt")
    End Sub

    <TestMethod()>
    Public Sub LoadMultipleDictionaries()
        Dim app As MemoryForMacros.Application
        app = GetAppObject()

        Dim paths As MemoryForMacros.Dictionary, files As MemoryForMacros.Dictionary
        app.LoadDictionaries("..\..\data2.txt")
        app.LoadDictionaries("..\..\data2-2.txt", True)
        app.ResolveDictionaries()

        paths = app.GetDictionary("paths")
        files = app.GetDictionary("files")
        If paths Is Nothing Or files Is Nothing Then
            Assert.Fail("Failed to load paths or files dictionaries from the same file.")
        End If

        ' Test proper reference resolution
        Dim expectedPaths As New StringBuilder
        expectedPaths.AppendLine("C:\Users\Doug")
        expectedPaths.AppendLine("C:\Users\Doug\AppData")
        expectedPaths.AppendLine("C:\Users\Doug\AppData\Roaming")
        expectedPaths.AppendLine("C:\Users\Doug\Downloads")
        expectedPaths.AppendLine("C:\Users\Doug\Documents")
        expectedPaths.AppendLine("C:\Users\Doug\Documents")
        expectedPaths.AppendLine("C:\Users\Doug\Documents\Personal")
        expectedPaths.AppendLine("C:\Users\Doug\Desktop")
        expectedPaths.AppendLine("C:\Users\Doug\Application Data\Microsoft\Word\STARTUP")
        expectedPaths.AppendLine("C:\Users\Doug\AppData\Roaming\Microsoft\AddIns")
        expectedPaths.AppendLine("C:\Users\Doug\Documents\Personal\Macros")
        expectedPaths.AppendLine("C:\Users\Doug\AppData\Roaming\KnowBrainer\KnowBrainerCommands")
        expectedPaths.AppendLine("personal")
        expectedPaths.AppendLine("C:\Second File\Path")

        Dim expectedFiles As New StringBuilder
        expectedFiles.AppendLine("C:\Users\Doug\Documents\mydoc.txt")
        expectedFiles.AppendLine("C:\Users\Doug\Documents\personal")
        expectedFiles.AppendLine("C:\Users\Doug\Documents\personal.extra")
        expectedFiles.AppendLine("C:\Users\Doug\Desktop\Douglas Parent\contact")
        expectedFiles.AppendLine("<Invalid entry reference: desktop2>\Douglas Parent\contact")
        expectedFiles.AppendLine("C:\Users\Doug\Documents\mydoc.txt\<Dictionary does not exist: people2>\contact")
        expectedFiles.AppendLine("C:\Users\Doug\Desktop\<Invalid entry reference: dp2>\contact")
        expectedFiles.AppendLine("This is a <Invalid self-reference> reference.")
        expectedFiles.AppendLine("C:\Second File\Path\my file.txt")
        expectedFiles.AppendLine("C:\Users\Doug\AppData\Roaming\KnowBrainer\KnowBrainerCommands\MyKBCommands.xml")



        Dim actualPaths As New StringBuilder
        Dim actualFiles As New StringBuilder

        For Each pair In paths
            For Each aValue As String In pair.Value
                actualPaths.AppendLine(aValue)
            Next
        Next

        Debug.Print("Actual Paths:")
        Debug.Print(actualPaths.ToString())

        Assert.AreEqual(expectedPaths.ToString(), actualPaths.ToString())

        For Each pair In files
            For Each aValue As String In pair.Value
                actualFiles.AppendLine(aValue)
            Next
        Next

        Debug.Print("Actual Files:")
        Debug.Print(actualFiles.ToString())

        Assert.AreEqual(expectedFiles.ToString(), actualFiles.ToString())

    End Sub

    <TestMethod()>
    Public Sub TestDelayedCommand()
        Dim app As MemoryForMacros.Application
        app = GetAppObject()

        Dim commandManager As MemoryForMacros.DelayedCommandManager
        commandManager = app.GetDelayedCommandManager()
        commandManager.AddCommand("Computer", 2000)
        commandManager.AddCommand("Bad Computer", 1000)
        commandManager.AddCommand("What Time Is It?")

        commandManager.StartCommands()

        Thread.Sleep(400000)
    End Sub


End Class
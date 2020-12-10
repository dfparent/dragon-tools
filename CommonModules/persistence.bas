'#Language "WWB.NET"
Option Explicit On


Public Sub LoadPeople(fileName As String, data As System.Collections.Generic.Dictionary(Of String, String()))

    If data.Count > 0 Then
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    Try
        Using fileReader As New System.IO.StreamReader(fileName)

            Dim fields As New System.Collections.Generic.List(Of String)()
            Dim separators() As String = {":", ","}
            Dim theKey As String
            Dim aValue As String
            Dim values() As String
            Dim line As String
            Dim e As System.Exception
            data.Clear()

            Do While fileReader.Peek() >= 0
                Try
                    line = fileReader.ReadLine()

                    If line = "" Then
                        Continue Do
                    End If

                    line.Trim()

                    If line.Substring(0, 1) = "'" Then
                        Continue Do
                    End If

                    ' Line is formatted like this:
                    '   <key>:<value 1>,<value 2>,...
                    fields.Clear()
                    fields.AddRange(line.Split(separators, System.StringSplitOptions.RemoveEmptyEntries))

                    If fields.Count < 2 Then
                        MsgBox("The following line is formatted incorrectly:  " & line)
                        Continue Do
                    End If

                    theKey = fields(0)
                    fields.RemoveAt(0)
                    data.Add(theKey, fields.ToArray())

                Catch e
                    msgbox("Failed to read line: " & e.Message)
                End Try
            Loop
        End Using
    Catch e
        msgbox("Error processing data file: " & e.Message)
        Exit Sub
    End Try

    Exit Sub

ErrorHandler:
    Msgbox("Error: " & err.description)
End Sub

Public Sub SavePeople(fileName As String, data As System.Collections.Generic.Dictionary(Of String, String()), Optional header As String = "")
    On Error GoTo HandleError
    Err.clear()

    If data.Count = 0 Then
        Exit Sub
    End If

    Dim backupFileName As String
    Dim e As System.Exception

    backupFileName = fileName & ".bak"

    Try
        System.IO.File.Copy(fileName, backupFileName, True)
    Catch e
        Msgbox("Failed to backup people list: " & e.Message)
        Exit Sub
    End Try


    Try
        Using fileWriter As System.IO.StreamWriter = New System.IO.StreamWriter(fileName, False)
            Dim line As New System.Text.StringBuilder

            fileWriter.WriteLine(header)
            fileWriter.WriteLine()

            Dim value As String
            Dim pair As System.Collections.Generic.KeyValuePair(Of String, String())

            For Each pair In data
                line.Clear()
                line.Append(pair.Key)
                line.Append(":")
                For Each value In pair.Value
                    line.Append(value)
                    line.Append(",")
                Next

                ' Remove last comma
                line.Remove(line.Length - 1, 1)

                Try
                    fileWriter.WriteLine(line.ToString())
                Catch e
                    msgbox("Failed to write line: " & e.Message)
                    GoTo HandleError
                End Try
            Next

        End Using
    Catch e
        msgbox("Failed to create the data file: " & e.Message)
        GoTo HandleError
    End Try

    Exit Sub

HandleError:
    If err.number <> 0 Then
        MsgBox("Error while writing data file:" & Err.Description)
    End If

    If fileName <> "" And backupFileName <> "" Then
        System.IO.File.Copy(backupFileName, fileName, True)
    End If

End Sub


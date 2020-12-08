Imports Microsoft.VisualBasic.ApplicationServices
Imports System.Runtime.InteropServices

Namespace My
    ' The following events are available for MyApplication:
    ' Startup: Raised when the application starts, before the startup form is created.
    ' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally.
    ' UnhandledException: Raised if the application encounters an unhandled exception.
    ' StartupNextInstance: Raised when launching a single-instance application and the application is already active. 
    ' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected.


    Partial Friend Class MyApplication

        Public cmdLineArgs As New CommandLineArgs
        Public clientHandle As IntPtr
        Public positionMode As Boolean

        <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Private Shared Function GetForegroundWindow() As IntPtr
        End Function

        Private Sub MyApplication_Startup(sender As Object, e As StartupEventArgs) Handles Me.Startup
            ' Get the handle of the foreground window
            clientHandle = GetForegroundWindow()
            If clientHandle = IntPtr.Zero Then
                Throw New Exception("Can't find currently active window.")
            End If

            If Not ProcessCommandLineArgs() Then
                e.Cancel = True
                Exit Sub
            End If
            'MsgBox(cmdLineArgs.ToString())

            ' Get foreground window before we create the form for this app
            If cmdLineArgs.relativeToClient Then
                If cmdLineArgs.clientHandle <> 0 Then
                    clientHandle = cmdLineArgs.clientHandle
                End If
            End If
        End Sub



        Private Function ProcessCommandLineArgs() As Boolean
            Dim args() As String
            args = Environment.GetCommandLineArgs()
            Dim nextArg As String

            If UBound(args) = 0 Then
                ' No arguments provided.  Enable position mode.
                positionMode = True
                Return True
            End If

            ' Skip first element which has exe file name
            For i = 1 To UBound(args)
                If i < UBound(args) Then
                    nextArg = args(i + 1)
                Else
                    nextArg = ""
                End If
                Try
                    cmdLineArgs.ProcessArg(args(i), nextArg)
                Catch ex As Exception
                    MsgBox(ex.Message)
                    Return False
                End Try
            Next i
            Return True
        End Function

    End Class
End Namespace

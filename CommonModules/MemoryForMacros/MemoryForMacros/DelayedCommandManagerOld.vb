Imports System.Collections
Imports System.Collections.Concurrent
Imports System.Timers
Imports DNSTools

Public Class DelayedCommandManagerolOld

    Private m_commands As New ConcurrentQueue(Of DelayedCommand)
    Private m_timer As New Timer

    Public Sub New()
        m_timer.AutoReset = False
        m_timer.Enabled = False
        AddHandler m_timer.Elapsed, New ElapsedEventHandler(AddressOf Me.HandleElapsedTimer)
    End Sub

    Public Sub AddCommand(command As DelayedCommand)
        m_commands.Enqueue(command)
    End Sub

    Public Sub ClearCommands()
        StopCommands()
        m_commands = New ConcurrentQueue(Of DelayedCommand)
    End Sub

    Public Sub StartCommands()
        If m_timer.Enabled Then
            ' Already running
            Exit Sub
        End If
        NextCommand()
    End Sub

    Public Sub StopCommands()
        m_timer.Enabled = False
    End Sub

    Private Sub NextCommand()
        Dim nextCommand As DelayedCommand
        If Not m_commands.TryDequeue(nextCommand) Then
            ' No more commands
            Exit Sub
        End If

        Dim dragon As New DgnEngineControl
        dragon.RecognitionMimic(nextCommand.Command)

        m_timer.Interval = nextCommand.Delay
        m_timer.Enabled = True

    End Sub

    Public Sub HandleElapsedTimer(sender As Object, e As ElapsedEventArgs)
        m_timer.Enabled = False
        NextCommand()
    End Sub

End Class

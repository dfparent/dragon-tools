﻿Imports System.Collections
Imports System.Collections.Concurrent
Imports System.Threading
Imports System.Timers
Imports System.Windows.Forms
Imports DNSTools

<ComClass(DelayedCommandManager.ClassId, DelayedCommandManager.InterfaceId, DelayedCommandManager.EventsId)>
Public Class DelayedCommandManager

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "ad22aaed-59e6-478b-8305-bf9a456a9faa"
    Public Const InterfaceId As String = "717f8421-97c6-4228-a9fd-aec665c0039e"
    Public Const EventsId As String = "2c66d2e4-6b72-4e77-a5ab-4f96ad417c5d"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
        m_timer.AutoReset = False
        m_timer.Enabled = False
        AddHandler m_timer.Elapsed, New ElapsedEventHandler(AddressOf Me.HandleElapsedTimer)
        'Debug.AutoFlush = True
    End Sub

    Private m_commands As New ConcurrentQueue(Of DelayedCommand)
    Private m_repeatCommand As DelayedCommand
    Private m_timer As New System.Timers.Timer
    Private m_dragon As DgnEngineControl
    Private m_dragonTools As DgnVoiceCmd

    Public Enum CommandType
        Spoken
        UseSendKeys
        UseSendSystemKeys
    End Enum

    ''' <summary>
    ''' Add a command to the command queue.  This does not start command execution.  You can queue multiple commands before starting execution.
    ''' </summary>
    ''' <param name="command">A string containing the Dragon command statement that you wish to execute.</param>
    ''' <param name="type">The type of command.  If CommandType.Spoken (0), Then this is intended to be issued as a spoken command.
    '''                     If CommandType.UseSendkeys (1), then this Command is intended to be typed out Using Sendkeys
    '''                     If CommandType.UseSendSystemKeys (2), then this command contains keystrokes that may not work using the faster SendKeys command
    '''                     And need to use the slower SendSystemKeys command to get them to work.</param>
    ''' <param name="delay">The number of milliseconds to wait before executing the next command.  Optional parameter.  Default is 100 ms.</param>
    ''' <param name="repeatCount">The number of times to repeat the command.  Optional parameter.  Default is 1.  Value of < 0 will cause the command to repeat
    '''                           indefinitely.  Subsequent commands will not run. Repeat count of 0 skips the command. </param>
    Public Sub AddCommand(command As String, type As CommandType, Optional delay As Integer = 100, Optional repeatCount As Integer = 1)
        m_commands.Enqueue(New DelayedCommand(command, type, delay, repeatCount))
    End Sub

    ''' <summary>
    ''' Stops command execution and deletes all commands from the queue.
    ''' </summary>
    Public Sub ClearCommands()
        StopCommands()
        m_commands = New ConcurrentQueue(Of DelayedCommand)
    End Sub

    ''' <summary>
    ''' Starts executing the commands in the queue.  You must first call AddCommand to add commands to the queue.
    ''' </summary>
    Public Sub StartCommands()
        m_repeatCommand = Nothing
        If m_timer.Enabled Then
            ' Already running
            Exit Sub
        End If
        If m_dragon Is Nothing Then
            m_dragon = New DgnEngineControl
        End If
        If m_dragonTools Is Nothing Then
            m_dragonTools = New DgnVoiceCmd
        End If
        NextCommand()
    End Sub

    ''' <summary>
    ''' Stops command execution.  You can later restart execution using StartCommands.
    ''' </summary>
    Public Sub StopCommands()
        m_timer.Enabled = False
        Try
            If Not m_dragonTools Is Nothing Then
                m_dragonTools = Nothing
            End If
            If Not m_dragon Is Nothing Then
                m_dragon = Nothing
            End If
        Catch
            ' Do nothing
        End Try
    End Sub

    Private Sub NextCommand()
        Dim theNextCommand As DelayedCommand

        If Not m_repeatCommand Is Nothing Then
            theNextCommand = m_repeatCommand
        ElseIf Not m_commands.TryDequeue(theNextCommand) Then
            ' No more commands
            StopCommands()
            Exit Sub
        End If


        If theNextCommand.RepeatCount = 0 Then
            NextCommand()
            Exit Sub
        End If

        If theNextCommand.CommandType = CommandType.Spoken Then
            'MsgBox("command")
            m_dragon.Register()
            m_dragon.RecognitionMimic(theNextCommand.Command)
            m_dragon.UnRegister(False)

        Else
            ' Typed
            Dim errorMessage As String = ""
            Dim errorLine As Integer
            Dim command As String
            If theNextCommand.CommandType = CommandType.UseSendKeys Then
                'MsgBox("send keys")
                command = "SendKeys """ & theNextCommand.Command & """"
            Else
                ' CommandType.UseSendSystemKeys
                'MsgBox("send system keys")
                command = "SendSystemKeys """ & theNextCommand.Command & """"
            End If


            m_dragonTools.Register("")
            errorLine = m_dragonTools.CheckScript(command, errorMessage)
            If errorLine <> 0 Then
                MsgBox("Can't execute command: " & command & ".  " & errorMessage)
            Else
                m_dragonTools.ExecuteScript(command, 0)
            End If
            m_dragonTools.UnRegister()
        End If

        ' Repeat command?
        If theNextCommand.RepeatCount < 0 Then
            ' Repeat indefinitely
            m_repeatCommand = theNextCommand
        ElseIf theNextCommand.RepeatCount > 1 Then
            theNextCommand.RepeatCount = theNextCommand.RepeatCount - 1
            m_repeatCommand = theNextCommand
        Else
            ' Don't repeat
            m_repeatCommand = Nothing
        End If

        m_timer.Interval = theNextCommand.Delay
        m_timer.Enabled = True

    End Sub

    Private Sub HandleElapsedTimer(sender As Object, e As ElapsedEventArgs)
        m_timer.Enabled = False
        NextCommand()
    End Sub

End Class



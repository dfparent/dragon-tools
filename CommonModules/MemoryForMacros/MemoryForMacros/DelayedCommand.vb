﻿Imports System.Timers
Imports DNSTools
Imports MemoryForMacros

Public Class DelayedCommand

    Private m_delay As Integer
    Private m_command As String
    Private m_commandType As DelayedCommandManager.CommandType
    Private m_repeatCount As Integer

    Public Property Command As String
        Get
            Return m_command
        End Get
        Set(value As String)
            m_command = value
        End Set
    End Property

    Public Property Delay As Integer
        Get
            Return m_delay
        End Get
        Set(value As Integer)
            m_delay = value
        End Set
    End Property

    Public Property CommandType As DelayedCommandManager.CommandType
        Get
            Return m_commandType
        End Get
        Set(value As DelayedCommandManager.CommandType)
            m_commandType = value
        End Set
    End Property

    Public Property RepeatCount As Integer
        Get
            Return m_repeatCount
        End Get
        Set(value As Integer)
            m_repeatCount = value
        End Set
    End Property

    Public Sub New(TheCommand As String, TheCommandType As DelayedCommandManager.CommandType, TheDelay As Integer, TheRepeatCount As Integer)
        Delay = TheDelay
        Command = TheCommand
        CommandType = TheCommandType
        RepeatCount = TheRepeatCount
    End Sub

End Class

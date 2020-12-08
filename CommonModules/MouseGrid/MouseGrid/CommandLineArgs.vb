Imports System.Text

Public Class CommandLineArgs

    Public Enum DockLocation
        center
        top
        left
        right
        bottom
        none
    End Enum

    Public Const WIDTH_ARG = "--width"
    Public Const HEIGHT_ARG = "--height"
    Public Const LOCATION_X_ARG = "--location-x"
    Public Const LOCATION_Y_ARG = "--location-y"
    Public Const SIZE_TO_CLIENT_ARG = "--size-to-client"
    Public Const DOCK_TO_CLIENT_ARG = "--dock-to-client"
    Public Const ROW_HEIGHT_ARG = "--row-height"
    Public Const COLUMN_WIDTH_ARG = "--column-width"
    Public Const STICKY_ARG = "--sticky"
    Public Const OPACITY_ARG = "--opacity"
    Public Const ALWAYS_ON_TOP_ARG = "--always-on-top"

    Public Const DEFAULT_ROW_HEIGHT = 20
    Public Const DEFAULT_COLUMN_WIDTH = 20
    Public Const DEFAULT_OPACITY = 0.5

    Private m_width As Long
    Private m_height As Long
    Private m_locationX As Long
    Private m_locationY As Long
    Private m_sizeToClient As Boolean
    Private m_sizeToClientHandle As IntPtr
    Private m_dockToClient As DockLocation
    Private m_rowHeight As Long
    Private m_columnWidth As Long
    Private m_isSticky As Boolean
    Private m_isAlwaysOnTop As Boolean
    Private m_opacity As Double

    Public Sub New(Optional width As Long = 0,
                   Optional height As Long = 0,
                   Optional locationX As Long = -1,
                   Optional locationY As Long = -1,
                   Optional sizeToClient As Boolean = False,
                   Optional sizeToClientHandle As Long = 0,
                   Optional dockToClient As DockLocation = DockLocation.none,
                   Optional rowHeight As Long = DEFAULT_ROW_HEIGHT,
                   Optional columnWidth As Long = DEFAULT_COLUMN_WIDTH,
                   Optional isSticky As Boolean = False,
                   Optional isAlwaysOnTop As Boolean = False,
                   Optional opacity As Double = DEFAULT_OPACITY)
        Me.Width = width
        Me.Height = height
        Me.LocationX = locationX
        Me.LocationY = locationY
        Me.relativeToClient = sizeToClient
        Me.clientHandle = sizeToClientHandle
        Me.dockToClient = dockToClient
        Me.RowHeight = rowHeight
        Me.ColumnWidth = columnWidth
        Me.IsSticky = isSticky
        Me.IsAlwaysOnTop = isAlwaysOnTop
        Me.Opacity = opacity

    End Sub

    ' The nextArgument parameter is for arguments that take values
    Public Sub ProcessArg(argument As String, nextArgument As String)
        Select Case argument
            Case CommandLineArgs.ALWAYS_ON_TOP_ARG
                Me.IsAlwaysOnTop = True

            Case CommandLineArgs.COLUMN_WIDTH_ARG
                If Not IsNumeric(nextArgument) Then
                    Throw New Exception("Column width is not a number.")
                End If
                Me.ColumnWidth = CLng(nextArgument)

            Case CommandLineArgs.HEIGHT_ARG
                If Not IsNumeric(nextArgument) Then
                    Throw New Exception("Height is not a number.")
                End If
                Me.Height = CLng(nextArgument)

            Case CommandLineArgs.LOCATION_X_ARG
                If Not IsNumeric(nextArgument) Then
                    Throw New Exception("LocationX is not a number.")
                End If
                Me.LocationX = CLng(nextArgument)

            Case CommandLineArgs.LOCATION_Y_ARG
                If Not IsNumeric(nextArgument) Then
                    Throw New Exception("LocationY is not a number.")
                End If
                Me.LocationY = CLng(nextArgument)

            Case CommandLineArgs.OPACITY_ARG
                If Not IsNumeric(nextArgument) Then
                    Throw New Exception("Opacity is not a number.")
                End If
                Me.Opacity = CDbl(nextArgument)

            Case CommandLineArgs.ROW_HEIGHT_ARG
                If Not IsNumeric(nextArgument) Then
                    Throw New Exception("Row height is not a number.")
                End If
                Me.RowHeight = CLng(nextArgument)

            Case CommandLineArgs.SIZE_TO_CLIENT_ARG
                Me.relativeToClient = True
                If IsNumeric(nextArgument) Then
                    Me.clientHandle = CLng(nextArgument)
                End If

            Case CommandLineArgs.DOCK_TO_CLIENT_ARG
                Me.relativeToClient = True
                Select Case nextArgument
                    Case DockLocation.none.ToString()
                        Me.dockToClient = DockLocation.none

                    Case DockLocation.center.ToString()
                        Me.dockToClient = DockLocation.center

                    Case DockLocation.left.ToString()
                        Me.dockToClient = DockLocation.left

                    Case DockLocation.right.ToString()
                        Me.dockToClient = DockLocation.right

                    Case DockLocation.top.ToString()
                        Me.dockToClient = DockLocation.top

                    Case DockLocation.bottom.ToString()
                        Me.dockToClient = DockLocation.bottom

                    Case Else
                        Throw New Exception("Invalid dock-to-client argument: " & nextArgument)
                End Select

            Case CommandLineArgs.STICKY_ARG
                Me.IsSticky = True

            Case CommandLineArgs.WIDTH_ARG
                If Not IsNumeric(nextArgument) Then
                    Throw New Exception("Width is not a number.")
                End If
                Me.Width = CLng(nextArgument)

        End Select
    End Sub
    ' Makes sure the combination of args provided are valid
    Public Sub ValidateArgs()
        ' If relative to client is not provide, must provide width and height
        If Not relativeToClient And (Width <= 0 Or Height <= 0 Or LocationX < 0 Or LocationY < 0) Then
            Throw New Exception("You must specify width, height, locationx and locationy when not including relativeToClient.")
        End If
    End Sub
    Public Property Width As Long
        Get
            Return m_width
        End Get
        Set(value As Long)
            If value < 0 Then
                m_width = 0
            Else
                m_width = value
            End If
        End Set
    End Property

    Public Property Height As Long
        Get
            Return m_height
        End Get
        Set(value As Long)
            m_height = value
        End Set
    End Property

    Public Property LocationX As Long
        Get
            Return m_locationX
        End Get
        Set(value As Long)
            m_locationX = value
        End Set
    End Property

    Public Property LocationY As Long
        Get
            Return m_locationY
        End Get
        Set(value As Long)
            m_locationY = value
        End Set
    End Property

    Public Property relativeToClient As Boolean
        Get
            Return m_sizeToClient
        End Get
        Set(value As Boolean)
            m_sizeToClient = value
        End Set
    End Property

    Public Property clientHandle As IntPtr
        Get
            Return m_sizeToClientHandle
        End Get
        Set(value As IntPtr)
            m_sizeToClientHandle = value
        End Set
    End Property

    Public Property dockToClient As DockLocation
        Get
            Return m_dockToClient
        End Get
        Set(value As DockLocation)
            m_dockToClient = value
        End Set
    End Property

    Public Property RowHeight As Long
        Get
            Return m_rowHeight
        End Get
        Set(value As Long)
            m_rowHeight = value
        End Set
    End Property

    Public Property ColumnWidth As Long
        Get
            Return m_columnWidth
        End Get
        Set(value As Long)
            m_columnWidth = value
        End Set
    End Property

    Public Property IsSticky As Boolean
        Get
            Return m_isSticky
        End Get
        Set(value As Boolean)
            m_isSticky = value
        End Set
    End Property

    Public Property IsAlwaysOnTop As Boolean
        Get
            Return m_isAlwaysOnTop
        End Get
        Set(value As Boolean)
            m_isAlwaysOnTop = value
        End Set
    End Property

    Public Property Opacity As Double
        Get
            Return m_opacity
        End Get
        Set(value As Double)
            If value < 0 Or value > 1 Then
                Throw New Exception("Opacity must be between 0 and 1, inclusive.")
            Else
                m_opacity = value
            End If
        End Set
    End Property

    Public Overrides Function ToString() As String
        Dim out As New StringBuilder()

        out.AppendLine("Width:  " & CStr(Width))
        out.AppendLine("Height:  " & CStr(Height))
        out.AppendLine("LocationX:  " & CStr(LocationX))
        out.AppendLine("LocationY:  " & CStr(LocationY))
        out.AppendLine("Size to Client:  " & CStr(relativeToClient))
        out.AppendLine("Size to Client Handle:  " & CStr(clientHandle))
        out.AppendLine("Row Height:  " & CStr(RowHeight))
        out.AppendLine("Column Width:  " & CStr(ColumnWidth))
        out.AppendLine("Is sticky:  " & CStr(IsSticky))
        out.AppendLine("Is always on top:  " & CStr(IsAlwaysOnTop))
        out.AppendLine("Opacity:  " & CStr(Opacity))

        Return out.ToString()
    End Function
End Class

Public Class CalloutData
    Private number As String
    Private displayText As String
    Private tooltipText As String
    Private clickPoint As Point     ' In screen coordinates

    Public Function getNumber() As String
        Return Me.number
    End Function

    Public Sub setNumber(number As String)
        Me.number = number
    End Sub

    Public Function getDisplayText() As String
        Return displayText
    End Function

    Public Sub setDisplayText(displayText As String)
        Me.displayText = displayText
    End Sub

    Public Function getClickPoint() As Point
        Return Me.clickPoint
    End Function

    Public Sub setClickPoint(clickPoint As Point)
        Me.clickPoint = clickPoint
    End Sub

    Public Function getTooltipText() As String
        Return Me.tooltipText
    End Function

    Public Sub setTooltipText(text As String)
        Me.tooltipText = text
    End Sub
End Class

Public Class CalloutData
    Private number As String
    Private displayText As String
    Private tooltipText As String
    Private clickPoint As Point     ' In screen coordinates
    Private elementBoundingRect As Windows.Rect

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

    Public Sub setElementBoundingRect(boundingRect As Windows.Rect)
        elementBoundingRect = boundingRect
    End Sub

    Public Function getElementBoundingRect() As Windows.Rect
        Return elementBoundingRect
    End Function

End Class

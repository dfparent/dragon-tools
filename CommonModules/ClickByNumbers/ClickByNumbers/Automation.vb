Imports System.Windows.Automation
Imports System.Windows

Module Automation
    Private currentWindowAutomationElement As AutomationElement

    Public Sub processAutomationChildren(windowAE As AutomationElement)

        'System.Windows.Automation.Automation.RemoveAllEventHandlers()

        currentWindowAutomationElement = windowAE

        ' Get cached properties
        Dim cacheRequest As New CacheRequest
        cacheRequest.AutomationElementMode = AutomationElementMode.None
        cacheRequest.Add(AutomationElement.BoundingRectangleProperty)
        'cacheRequest.Add(ValuePattern.Pattern)
        'cacheRequest.Add(ValuePattern.ValueProperty)
        'cacheRequest.Add(TextPattern.Pattern)
        'cacheRequest.Add(AutomationElement.NameProperty)
        cacheRequest.Add(AutomationElement.IsOffscreenProperty)
        'cacheRequest.Add(AutomationElement.NativeWindowHandleProperty)

        Dim colChildren As AutomationElementCollection
        Dim visibleCondition = New PropertyCondition(AutomationElement.IsOffscreenProperty, False)
        Dim searchCondition As New AndCondition(visibleCondition, Windows.Automation.Automation.ControlViewCondition)
        Using cacheRequest.Activate()
            colChildren = windowAE.FindAll(TreeScope.Subtree, searchCondition)
        End Using

        Dim boundingRect As System.Windows.Rect

        'Dim propertyEventHandlerDelegate As New AutomationPropertyChangedEventHandler(AddressOf OnPropertyChange)

        For Each child As AutomationElement In colChildren
            Try
                boundingRect = child.GetCachedPropertyValue(AutomationElement.BoundingRectangleProperty, False)
                If Not boundingRect.IsEmpty Then
                    frmMain.MakeCallout(boundingRect, boundingRect.Left & "," & boundingRect.Top & ", " & boundingRect.Right & ", " & boundingRect.Bottom)
                    'controls.Add(frmMain.MakeCallout(boundingRect, ""))
                End If
            Catch ex As Exception
                ' Do nothing
                Debug.WriteLine("Exception processing automation children:  " & ex.Message)
                Throw ex
            End Try
        Next

        ' Subscribe to events
        'Dim eventHandlerDelegate As New AutomationEventHandler(AddressOf OnAutomationEvent)
        'System.Windows.Automation.Automation.
        'AddStructureChangedEventHandler(currentWindowAutomationElement, TreeScope.Descendants, New StructureChangedEventHandler(AddressOf OnStructureChanged))

    End Sub

    'Private Sub OnPropertyChange(ByVal src As Object, ByVal e As AutomationPropertyChangedEventArgs)
    '    MsgBox("Property change")
    'frmMain.RefreshCallouts()
    'End Sub

    'Private Sub OnAutomationEvent(ByVal sender As Object, ByVal e As AutomationEventArgs)
    'If e.EventId Is AutomationElement.LayoutInvalidatedEvent Then
    '       frmMain.ShowPrompt("Layout invalidated", 2000)
    'End If
    'frmMain.RefreshCallouts()

    'End Sub

    Public Function getText(element As AutomationElement) As String
        Dim patternObj As Object
        If (element.TryGetCachedPattern(ValuePattern.Pattern, patternObj)) Then
            Dim vPattern As ValuePattern = patternObj
            Return vPattern.Cached.Value
        ElseIf (element.TryGetCachedPattern(TextPattern.Pattern, patternObj)) Then
            Dim tPattern As TextPattern = patternObj
            Return tPattern.DocumentRange.GetText(-1).TrimEnd("\r")  ' often there is an extra '\r' hanging off the end.
        Else
            Return element.Cached.Name
        End If
    End Function
End Module
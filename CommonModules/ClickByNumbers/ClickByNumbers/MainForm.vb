Imports System.Threading
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Automation
Imports System.Runtime.InteropServices
Imports System.Drawing.Drawing2D


Public Class frmMain

    Private Structure RECT
        Dim Left As Integer
        Dim Top As Integer
        Dim Right As Integer
        Dim Bottom As Integer
    End Structure

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function GetClientRect(ByVal hWnd As System.IntPtr,
                                          ByRef lpRECT As RECT) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function ClientToScreen(ByVal hWnd As IntPtr, ByRef lpPoint As Drawing.Point) As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
    End Function

    Private Const SPLASH_TIME = 5000
    Private currentTargetWindowHandle As IntPtr
    Private currentSettings As AppSettings
    Private animationOpacity As Integer
    Private animationFadeIn As Boolean
    Private animationStep As Integer = 5
    Private animationMinimumOpacity As Integer = 5
    Private transparentColor As Color = Color.Magenta
    Private lineWidth As Integer = 4
    Private calloutFont As Font = New Font("Verdana", 8, FontStyle.Bold)
    Private callouts As New Dictionary(Of String, PictureBox)
    Private Const ClickLocationOffsetX = 5
    Private Const ClickLocationOffsetY = 5
    Private usedLocations As New Dictionary(Of Drawing.Point, CalloutData)
    Private tt As ToolTip = New ToolTip()
    Private userCalloutEntry As String
    Private controlsUnderPointer As New Collection()
    Private lockObject As New Object()

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        frmSplash.Show()
        timSplash.Interval = SPLASH_TIME
        timSplash.Start()

        ' Make background transparent
        Me.TransparencyKey = transparentColor
        Me.BackColor = transparentColor

        GlobalHotKeys.RegisterHotkeys()
        WindowMonitor.InitializeSystemMonitor()
    End Sub

    ' Use force = true to force a refresh on the current window
    Public Sub ResizeToTargetWindow(handle As IntPtr, Optional force As Boolean = False)
        If handle = IntPtr.Zero Then
            'ShowPrompt("Zero handle on resize to target.", 2000)
            Exit Sub
        End If

        If currentTargetWindowHandle = handle And Not force Then
            Exit Sub
        End If

        SyncLock lockObject
            If currentTargetWindowHandle <> handle Then
                currentTargetWindowHandle = handle
                ReloadSettings()
            End If

            Dim startingTargetWindowHandle = currentTargetWindowHandle

            ' Clear out old callouts
            usedLocations.Clear()
            Dim pb As PictureBox
            Dim deleteCalloutList As New List(Of PictureBox)
            Dim pbEnum As IEnumerable(Of PictureBox) = Me.Controls.OfType(Of PictureBox)
            For Each pb In pbEnum
                deleteCalloutList.Add(pb)
            Next
            For Each pb In deleteCalloutList.ToArray()
                Me.Controls.Remove(pb)
            Next

            callouts.Clear()

            If currentSettings.IsDisabled Then
                ' Callouts are not enabled, so don't load them yet.
                Exit Sub
            End If

            ' Position window over the target window
            Dim clientRect As New RECT
            GetClientRect(handle, clientRect)
            Dim clientPoint As New Drawing.Point
            ClientToScreen(handle, clientPoint)
            Me.Location = New Drawing.Point(clientPoint.X, clientPoint.Y)
            Me.Width = clientRect.Right
            Me.Height = clientRect.Bottom

            ShowPrompt("Loading numbers...")

            If currentSettings.UsesWindowHandleDiscovery Then
                ' Create callouts for child windows with handles
                EnumWindows.ProcessWindowChildren(handle)
            End If

            If currentSettings.UsesUIAutomationDiscovery Then
                ' create callouts for child objects using automation
                Dim targetWindowAE As AutomationElement
                Try
                    targetWindowAE = AutomationElement.FromHandle(handle)
                    Automation.processAutomationChildren(targetWindowAE)
                Catch ex As Exception
                    ShowPrompt("Error getting automation children", 5000)
                End Try
            End If

            If currentSettings.UsesMSAADiscovery Then
                MSAAUtils.ProcessAccessibleWindow(handle)
            End If

            If startingTargetWindowHandle = currentTargetWindowHandle Then
                ' Target window has not changed, so proceed
                RenderCallouts()
                ' Check for window change again.  Can happen when clicking results in another window launching
                If startingTargetWindowHandle <> currentTargetWindowHandle Then
                    SetCalloutsEnabled(False)
                End If
            End If


            HidePrompt()

        End SyncLock


    End Sub

    Private Sub ReloadSettings()
        currentSettings = New AppSettings()
        currentSettings.Load(currentTargetWindowHandle)
        Me.Opacity = currentSettings.Opacity / 100
    End Sub

    Public Sub ShowPrompt(text As String, Optional timeout As Integer = 0)
        If timeout > 0 Then
            timPrompt.Interval = timeout
            timPrompt.Start()
        End If

        ' Center over form 
        lblPrompt.Text = text
        lblPrompt.Left = (Me.Width / 2) - lblPrompt.Width / 2
        lblPrompt.Top = (Me.Height / 2) - lblPrompt.Height / 2
        lblPrompt.Visible = True

        Application.DoEvents()
    End Sub

    Public Sub AddToPrompt(text As String, Optional timeout As Integer = 0)
        If lblPrompt.Visible Then
            lblPrompt.Text = lblPrompt.Text & text
            If timeout > 0 Then
                timPrompt.Stop()
                timPrompt.Start()
            End If
        Else
            ShowPrompt(text, timeout)
        End If
    End Sub

    Public Sub HidePrompt()
        lblPrompt.Visible = False
        timPrompt.Stop()
    End Sub

    Private Sub TimPrompt_Tick(sender As Object, e As EventArgs) Handles timPrompt.Tick
        timPrompt.Stop()
        lblPrompt.Visible = False
    End Sub

    Private Sub timSplash_Tick(sender As Object, e As EventArgs) Handles timSplash.Tick
        frmSplash.Close()
        timSplash.Stop()
    End Sub

    Private Sub timAnimation_Tick(sender As Object, e As EventArgs) Handles timAnimation.Tick
        If currentSettings.IsDisabled Then
            timAnimation.Stop()
            Exit Sub
        End If

        ' Check for reversal
        If animationFadeIn Then
            If animationOpacity >= currentSettings.Opacity Then
                ' Fade out
                animationFadeIn = False
            End If
        Else
            If animationOpacity <= animationMinimumOpacity Then
                animationFadeIn = True
            End If
        End If

        ' Change opacity
        If animationFadeIn Then
            animationOpacity = animationOpacity + animationStep
        Else
            animationOpacity = animationOpacity - animationStep
        End If

        Me.Opacity = animationOpacity / 100

    End Sub

    ' Makes a data object representing a callout.  Render happens later because it is expensive
    Public Sub MakeCallout(elementBoundingRectangle As System.Windows.Rect, tooltipText As String)

        ' Get display location for callout 
        ' Place the callout in the top left of the bounding rectangle for the element
        ' Convert to client coordinates
        Dim screenLocationPoint As New Drawing.Point(elementBoundingRectangle.Left, elementBoundingRectangle.Top)
        Dim locationPoint As Drawing.Point = Me.PointToClient(screenLocationPoint)

        ' Is there a callout there already?
        If usedLocations.Keys.Contains(locationPoint) Then
            ' Replace it.  Last one wins
            usedLocations.Remove(locationPoint)
        End If

        ' Setup objects
        Dim data As New CalloutData
        data.setTooltipText(tooltipText)

        ' Set click point near the callout.  Putting the click location in the center of the element sometimes miss clicks
        data.setClickPoint(New Drawing.Point(locationPoint.X + ClickLocationOffsetX,
                                             locationPoint.Y + ClickLocationOffsetY))
        'data.setClickPoint(New Drawing.Point(locationPoint.X + (elementBoundingRectangle.Width / 2),
        'locationPoint.Y + (elementBoundingRectangle.Height / 2)))

        data.setElementBoundingRect(elementBoundingRectangle)

        usedLocations.Add(locationPoint, data)

    End Sub

    ' Once you have created the callouts, create the Picture Boxes and render them.
    ' We do this all at once for performance reasons
    Public Sub RenderCallouts()

        Dim controlRange As New List(Of Control)
        callouts.Clear()

        Dim calloutNextNumber As Integer = 0

        For Each kvp As KeyValuePair(Of Drawing.Point, CalloutData) In usedLocations
            Dim locationPoint As Drawing.Point = kvp.Key
            Dim data As CalloutData = kvp.Value

            ' Setup objects
            calloutNextNumber = calloutNextNumber + 1
            Dim countText As String
            countText = CStr(calloutNextNumber)
            data.setNumber(countText)

            Dim displayText As String
            'displayText = " " & countText & " "
            displayText = countText
            data.setDisplayText(displayText)

            Dim pb As New PictureBox
            pb.Top = locationPoint.Y
            pb.Left = locationPoint.X
            callouts.Add(data.getNumber, pb)

            Dim g As Graphics = pb.CreateGraphics
            Dim theSize As SizeF = g.MeasureString(data.getDisplayText, Me.calloutFont)

            ' Add size for width of line
            'pb.Width = theSize.Width + lineWidth
            'pb.Height = theSize.Height + lineWidth
            pb.Width = theSize.Width
            pb.Height = theSize.Height
            pb.Tag = data

            AddHandler pb.Paint, AddressOf Callout_paint
            AddHandler pb.MouseEnter, AddressOf Callout_OnPictureMouseEnter
            AddHandler pb.MouseLeave, AddressOf Callout_OnPictureMouseLeave

            'Me.Controls.Add(pb)
            controlRange.Add(pb)

        Next

        ' Show them
        Controls.AddRange(controlRange.ToArray())

        ' Start animation
        Me.Opacity = currentSettings.Opacity / 100
        animationOpacity = currentSettings.Opacity
        timAnimation.Start()
    End Sub

    Public Sub SetCalloutsEnabled(enabled As Boolean)
        currentSettings.IsDisabled = Not enabled
        If enabled Then
            ResizeToTargetWindow(currentTargetWindowHandle, True)
        Else
            timAnimation.Stop()
            Dim pb As PictureBox
            For Each kvp As KeyValuePair(Of String, PictureBox) In callouts
                pb = kvp.Value
                pb.Visible = False
            Next
        End If
    End Sub

    Public Function GetCalloutsEnabled() As Boolean
        Return Not currentSettings.IsDisabled
    End Function

    Public Sub SetCalloutsSticky(sticky As Boolean)
        currentSettings.IsSticky = sticky
    End Sub

    Public Function GetCalloutsSticky() As Boolean
        Return currentSettings.IsSticky
    End Function

    Private Sub Callout_paint(sender As Object, e As PaintEventArgs)
        Dim pb As PictureBox = sender
        Dim data As CalloutData = pb.Tag
        'Dim width As Integer = pb.Width
        'Dim height As Integer = pb.Height

        ' Draw callout, leave space for width of line
        'Dim points(4) As Drawing.Point
        'points(0) = New Drawing.Point(lineWidth, lineWidth)
        'points(1) = New Drawing.Point(width - lineWidth, lineWidth)
        'points(2) = New Drawing.Point(width - lineWidth, height - lineWidth)
        'points(3) = New Drawing.Point(lineWidth, height - lineWidth)
        'points(4) = New Drawing.Point(lineWidth, lineWidth)

        Dim g As Graphics = e.Graphics
        'g.Clear(transparentColor)
        g.Clear(Color.White)
        g.SmoothingMode = SmoothingMode.AntiAlias
        'g.DrawPolygon(New Pen(Color.Black, lineWidth), points)
        'g.FillPolygon(New SolidBrush(Color.Yellow), points, FillMode.Winding)

        'Draw drop shadow
        'g.DrawString(data.getDisplayText, Me.calloutFont, New SolidBrush(Color.DarkGray), lineWidth / 2 + 1, lineWidth / 2 + 1)
        'g.DrawString(data.getDisplayText, Me.calloutFont, New SolidBrush(Color.Red), lineWidth / 2, lineWidth / 2)
        'g.DrawString(data.getDisplayText, Me.calloutFont, New SolidBrush(Color.DarkGray), 0, 0)
        g.DrawString(data.getDisplayText, Me.calloutFont, New SolidBrush(Color.Red), 0, 0)


    End Sub

    Private Sub Callout_OnPictureMouseEnter(ByVal sender As Object, ByVal e As EventArgs)
        Dim data As CalloutData
        data = sender.Tag
        If data Is Nothing Then
            Exit Sub
        End If

        Dim pb As PictureBox
        pb = sender

        'Debug.WriteLine("Enter Callout on mouse enter")

        'tt.Show(data.getTooltipText(), sender)

        Dim elementBoundingRect As Windows.Rect
        elementBoundingRect = data.getElementBoundingRect
        Dim locationPointTopLeft As Drawing.Point = Me.PointToClient(New Drawing.Point(elementBoundingRect.Left, elementBoundingRect.Top))
        Dim locationPointBottomRight As Drawing.Point = Me.PointToClient(New Drawing.Point(elementBoundingRect.Right, elementBoundingRect.Bottom))

        txtOutput.Left = locationPointTopLeft.X
        txtOutput.Top = locationPointTopLeft.Y
        txtOutput.Width = locationPointBottomRight.X - locationPointTopLeft.X
        txtOutput.Height = locationPointBottomRight.Y - locationPointTopLeft.Y

        'txtOutput.Visible = True


    End Sub

    Private Sub Callout_OnPictureMouseLeave(ByVal sender As Object, ByVal e As EventArgs)
        'tt.Hide(Me)
        txtOutput.Visible = False
    End Sub

    Private Sub FrmMain_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        GlobalHotKeys.UnRegisterHotkeys()
        WindowMonitor.UninitializeSystemMonitor()
    End Sub

    Private Sub FrmMain_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyData = Keys.F5 Then
            ' Refresh
            RefreshCallouts()

        ElseIf e.KeyCode = Keys.F4 Then
            HandleToggleShowHideCallouts()

        ElseIf e.KeyCode = Keys.F2 Then
            ShowOptions()

        ElseIf e.KeyCode = Keys.F1 Then
            If frmHelp.Visible Then
                frmHelp.Focus()
                Exit Sub
            End If
            frmHelp.ShowDialog()
            Thread.Sleep(100)
            SetForegroundWindow(currentTargetWindowHandle)

        ElseIf e.KeyCode = Keys.Escape Then
            FinishCalloutEntry()
            SetForegroundWindow(currentTargetWindowHandle)

            '        ElseIf e.KeyCode = Keys.Pause Then
            '           calloutsAreSticky = Not calloutsAreSticky

        ElseIf (e.KeyCode >= Keys.D0 And e.KeyCode <= Keys.D9) Or (e.KeyCode >= Keys.NumPad0 And e.KeyCode <= Keys.NumPad9) Then
            ' Get number and add to prompt
            Dim theChar As String
            theChar = Chr(e.KeyCode)
            AddToPrompt(theChar)
            userCalloutEntry = userCalloutEntry & theChar

        ElseIf e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Up Or e.KeyCode = Keys.Down Then
            ' Do some sort of mouse action

            Dim clickPoint As Drawing.Point = GetCallOutClickPoint()
            HideCalloutsUnderPointer(clickPoint)
            Thread.Sleep(100)
            SetForegroundWindow(currentTargetWindowHandle)
            Thread.Sleep(100)

            ' Wait for window
            WaitForWindow(currentTargetWindowHandle, 5)

            If e.KeyData = Keys.Enter Then
                ' Simple mouseclick
                InputUtils.MouseLeftClick(clickPoint)

            ElseIf e.KeyData = (Keys.Shift + Keys.Enter) Then
                ' Shift click
                InputUtils.KeyDown(Keys.Shift)
                InputUtils.MouseLeftClick(clickPoint)
                InputUtils.KeyUp(Keys.Shift)

            ElseIf e.KeyData = (Keys.Control + Keys.Enter) Then
                ' Control click
                InputUtils.KeyDown(Keys.Control)
                InputUtils.MouseLeftClick(clickPoint)
                InputUtils.KeyUp(Keys.Control)

            ElseIf e.KeyData = (Keys.Alt + Keys.Enter) Then
                ' Double-click
                InputUtils.MouseDoubleClick(clickPoint)

            ElseIf e.KeyData = (Keys.Control + Keys.Shift + Keys.Enter) Then
                ' Right-click
                InputUtils.MouseRightClick(clickPoint)

            ElseIf e.KeyCode = Keys.Down Then
                InputUtils.MouseLeftDown(clickPoint)

            ElseIf e.KeyCode = Keys.Up Then
                InputUtils.MouseLeftUp(clickPoint)

            ElseIf e.KeyCode = Keys.Tab Then
                Cursor.Position = clickPoint
            End If

            FinishCalloutEntry()
            ShowHiddenCallouts()
            If Not currentSettings.IsSticky Then
                SetCalloutsEnabled(False)
            ElseIf e.KeyCode = Keys.Enter Then
                RefreshCallouts()
            End If

        End If

    End Sub

    ' For Global hot keys
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        If m.Msg = GlobalHotKeys.WM_HOTKEY Then
            GlobalHotKeys.HandleGlobalHotKeyEvent(m.WParam)
        End If
        MyBase.WndProc(m)
    End Sub

    ' Returns true if Desired window is in the foreground
    ' Returns false if the desired window never entered the foreground during the maximum amount of wait time
    Private Function WaitForWindow(handle As IntPtr, Optional timeoutSeconds As Integer = 10) As Boolean
        ' Set maximum wait time
        Dim stopTime As Date = Date.Now()
        stopTime = DateAdd(DateInterval.Second, timeoutSeconds, stopTime)

        Do
            If GetForegroundWindow() = handle Then
                ' Wait a bit more
                Thread.Sleep(100)
                Return True
            End If

            ' Check your time
            If Date.Compare(stopTime, Date.Now) < 0 Then
                ' Out of time
                Return False
            End If

            ' wait a bit
            Thread.Sleep(100)
        Loop

    End Function

    Public Sub ShowOptions()
        If frmOptions.Visible Then
            frmOptions.Focus()
            Exit Sub
        End If
        frmOptions.ShowDialog()
        Thread.Sleep(100)
        ReloadSettings()
        RefreshCallouts()
        frmOptions.Close()
        SetForegroundWindow(currentTargetWindowHandle)
    End Sub

    Public Sub RefreshCallouts()
        ResizeToTargetWindow(currentTargetWindowHandle, True)
        FinishCalloutEntry()
        SetForegroundWindow(currentTargetWindowHandle)
    End Sub

    Private Sub FinishCalloutEntry()
        userCalloutEntry = ""
        HidePrompt()
    End Sub

    ' Moves the mouse to the Click location for the callout whose number is stored in userCalloutEntry
    ' This returns a location in system coordinates
    Private Function GetCallOutClickPoint() As Drawing.Point
        If userCalloutEntry = "" Or Not Integer.TryParse(userCalloutEntry, Nothing) Then
            Return Cursor.Position
        End If

        If Not callouts.Keys.Contains(userCalloutEntry) Then
            Return Cursor.Position
        End If

        Dim pb As PictureBox = callouts.Item(userCalloutEntry)
        Dim calloutData As CalloutData = pb.Tag
        Return Me.PointToScreen(calloutData.getClickPoint())
    End Function

    ' Hides the callouts that are under the pointer for click action
    Private Sub HideCalloutsUnderPointer(cursorLocation As Drawing.Point)
        Dim clientLocation As Drawing.Point = Me.PointToClient(cursorLocation)
        Dim theControl As Control

        theControl = Me.GetChildAtPoint(clientLocation)

        While Not theControl Is Nothing
            controlsUnderPointer.Add(theControl)
            theControl.Visible = False
            Application.DoEvents()
            theControl = Me.GetChildAtPoint(clientLocation, GetChildAtPointSkip.Invisible)
        End While

    End Sub

    Private Sub ShowHiddenCallouts()
        Dim aControl As Control
        For Each aControl In controlsUnderPointer
            aControl.Visible = True
        Next
        controlsUnderPointer.Clear()
    End Sub

    ' Hides the callouts that are under the pointer for click action
    Private Sub RestoreCalloutsUnderPointer()
        For Each kvp In controlsUnderPointer
            kvp.Value.Visible = False
        Next
    End Sub

    Public Sub BringToForeground()
        SetForegroundWindow(Me.Handle)
    End Sub

    Public Sub HandleToggleStickySetting()
        SetCalloutsSticky(Not GetCalloutsSticky())
    End Sub

    Public Sub HandleToggleShowHideCallouts()
        SetCalloutsEnabled(Not GetCalloutsEnabled())
    End Sub

    Public Function GetCurrentTargetWindowHandle() As IntPtr
        Return currentTargetWindowHandle
    End Function

    Private Sub ExitTrayIconMenuStripItem_Click(sender As Object, e As EventArgs) Handles ExitTrayIconMenuStripItem.Click
        Me.Close()
    End Sub

End Class



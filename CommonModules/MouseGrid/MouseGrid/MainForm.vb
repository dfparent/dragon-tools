Imports System.Threading
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Automation
Imports System.Runtime.InteropServices
Imports System.Drawing.Drawing2D
Imports System.ComponentModel
Imports System.Text


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
    Private Shared Function GetWindowRect(ByVal hWnd As System.IntPtr, ByRef lpRect As RECT) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function ClientToScreen(ByVal hWnd As IntPtr, ByRef lpPoint As Drawing.Point) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function ScreenToClient(ByVal hWnd As IntPtr, ByRef lpPoint As Drawing.Point) As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
    End Function

    Private labelReferenceFont As Font = New Font("Verdana", 8, FontStyle.Bold)
    Private labelActualFont As Font
    Private labelActualSize As SizeF
    Private labelBackNormalBrush As Brush = Brushes.Black
    Private labelBackHighlightBrush As Brush = Brushes.Red
    Private gridLineColor As Color = Color.Black
    Private gridLineColorSticky As Color = Color.Maroon
    Private gridLineWidth As Integer = 4

    Private recalculateLabelSize As Boolean = True

    Private currentTargetWindowHandle As IntPtr
    Private rowEntry As String
    Private columnEntry As String
    Private isSticky As Boolean = False

    Private Enum EntryMode
        Row
        Column
        None
    End Enum
    Private theEntryMode As EntryMode = EntryMode.None


    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Position mode is enabled when the application is started without any commandline arguments
        If My.Application.positionMode Then
            ' Normally, we display the application in "grid mode".p
            ' When Not in grid mode, we enter "position mode".
            ' Need to display the form title bar to allow positioning.
            Me.FormBorderStyle = FormBorderStyle.Sizable
            SetRelativeToClient(False)
            Me.mnuMain.Visible = True
            Me.statusBarLabel.Text = ""
            Me.statusBar.Visible = True
            Me.Opacity = 0.7
            UpdateStatusBar()
        Else
            PositionGridForm()
            ' Me.TopMost = True
        End If

        Me.Opacity = My.Application.cmdLineArgs.Opacity

    End Sub

    ''' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' Used in position mode only
    ''' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' Used in position mode to display and deliver the parameter string that will
    ' display the grid in the current position grid's orientation
    Public Function GetParameterString() As String
        Dim out As New StringBuilder
        Dim pt As New Drawing.Point
        Dim cmdLineArgs As CommandLineArgs
        cmdLineArgs = My.Application.cmdLineArgs

        Dim clientArea As Rectangle
        clientArea = pboxCanvas.ClientRectangle

        pt.X = clientArea.X
        pt.Y = clientArea.Y

        ' Compensate for size of menu which overlaps picture box.  Probably should figure out how to lay this out better.
        ' You would think menu would move client area down, but it doesn't.
        pt.Y = pt.Y + mnuMain.Height

        ' Convert Location 2 screen coordinates
        If Not ClientToScreen(Me.Handle, pt) Then
            Throw New Exception("Unable to convert grid location to screen coordinates.")
        End If

        If cmdLineArgs.relativeToClient Then
            ' Relative to the other client.
            ' Need to convert the origin of the client of the grid to client coordinates for the other application
            If Not ScreenToClient(My.Application.clientHandle, pt) Then
                Throw New Exception("Unable to convert grid location to client coordinates.")
            End If
            out.Append(CommandLineArgs.SIZE_TO_CLIENT_ARG)
            out.Append(" ")
        End If

        out.Append(CommandLineArgs.LOCATION_X_ARG)
        out.Append(" ")
        out.Append(pt.X)
        out.Append(" ")
        out.Append(CommandLineArgs.LOCATION_Y_ARG)
        out.Append(" ")
        out.Append(pt.Y)
        out.Append(" ")

        out.Append(CommandLineArgs.WIDTH_ARG)
        out.Append(" ")
        out.Append(clientArea.Width)
        out.Append(" ")

        out.Append(CommandLineArgs.HEIGHT_ARG)
        out.Append(" ")
        out.Append(clientArea.Height)
        out.Append(" ")

        If cmdLineArgs.RowHeight <> CommandLineArgs.DEFAULT_ROW_HEIGHT Then
            out.Append(CommandLineArgs.ROW_HEIGHT_ARG)
            out.Append(" ")
            out.Append(cmdLineArgs.RowHeight)
            out.Append(" ")
        End If

        If cmdLineArgs.ColumnWidth <> CommandLineArgs.DEFAULT_COLUMN_WIDTH Then
            out.Append(CommandLineArgs.COLUMN_WIDTH_ARG)
            out.Append(" ")
            out.Append(cmdLineArgs.ColumnWidth)
            out.Append(" ")
        End If

        If cmdLineArgs.Opacity <> CommandLineArgs.DEFAULT_OPACITY Then
            out.Append(CommandLineArgs.OPACITY_ARG)
            out.Append(" ")
            out.Append(cmdLineArgs.Opacity)
            out.Append(" ")
        End If

        If cmdLineArgs.IsSticky Then
            out.Append(CommandLineArgs.STICKY_ARG)
            out.Append(" ")
        End If

        If cmdLineArgs.IsAlwaysOnTop Then
            out.Append(CommandLineArgs.ALWAYS_ON_TOP_ARG)
            out.Append(" ")
        End If

        Return out.ToString()

    End Function

    Public Sub UpdateStatusBar()
        statusBarLabel.Text = GetParameterString()
    End Sub

    Public Sub SetRelativeToClient(relativeToClient As Boolean)
        My.Application.cmdLineArgs.relativeToClient = relativeToClient
        mnuUseClient.Checked = relativeToClient
        mnuUseScreen.Checked = Not relativeToClient
    End Sub

    ''' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' Used in Grid mode only
    ''' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Sub PositionGridForm()

        ' We need to calculate
        '   location x in screen coordinates
        '   Location y in screen coordinates
        '   Grid width
        '   Grid height
        ' and position/resize the form accordingly

        Dim screenOriginPoint As New Drawing.Point
        Dim clientRect As New RECT
        If My.Application.cmdLineArgs.relativeToClient Then
            ' Get the target window handle

            GetClientRect(My.Application.clientHandle, clientRect)
            'GetWindowRect(My.Application.clientHandle, clientRect)
            Dim dockFactor As Double = 0.25

            ' Are location x and y specified?
            If My.Application.cmdLineArgs.dockToClient <> CommandLineArgs.DockLocation.none Then
                Select Case My.Application.cmdLineArgs.dockToClient
                    Case CommandLineArgs.DockLocation.left, CommandLineArgs.DockLocation.top
                        screenOriginPoint.X = 0
                        screenOriginPoint.Y = 0

                    Case CommandLineArgs.DockLocation.right
                        screenOriginPoint.X = clientRect.Right - clientRect.Right * dockFactor
                        screenOriginPoint.Y = 0

                    Case CommandLineArgs.DockLocation.bottom
                        screenOriginPoint.X = 0
                        screenOriginPoint.Y = clientRect.Bottom - clientRect.Bottom * dockFactor

                    Case CommandLineArgs.DockLocation.center
                        screenOriginPoint.X = clientRect.Right * dockFactor
                        screenOriginPoint.Y = clientRect.Bottom * dockFactor
                End Select

            ElseIf My.Application.cmdLineArgs.LocationX >= 0 Or My.Application.cmdLineArgs.LocationY >= 0 Then
                ' use provided values.  Values are in client coordinates
                screenOriginPoint.X = My.Application.cmdLineArgs.LocationX
                screenOriginPoint.Y = My.Application.cmdLineArgs.LocationY
            Else
                ' Use client window origin point
                screenOriginPoint.X = 0
                screenOriginPoint.Y = 0
            End If

            'screenOriginPoint = Me.PointToScreen(screenOriginPoint)
            If Not ClientToScreen(My.Application.clientHandle, screenOriginPoint) Then
                Throw New Exception("Unable to convert grid location to screen coordinates.")
            End If

            Me.Top = screenOriginPoint.Y
                Me.Left = screenOriginPoint.X

                ' Get grid size
                ' Are width and height specified?
                If My.Application.cmdLineArgs.dockToClient <> CommandLineArgs.DockLocation.none Then
                    Select Case My.Application.cmdLineArgs.dockToClient
                        Case CommandLineArgs.DockLocation.top, CommandLineArgs.DockLocation.bottom
                            Me.Width = clientRect.Right - clientRect.Left
                            Me.Height = (clientRect.Bottom - clientRect.Top) * dockFactor

                        Case CommandLineArgs.DockLocation.left, CommandLineArgs.DockLocation.right
                            Me.Width = (clientRect.Right - clientRect.Left) * dockFactor
                            Me.Height = clientRect.Bottom - clientRect.Top

                        Case CommandLineArgs.DockLocation.center
                            Me.Width = (clientRect.Right - clientRect.Left) * (dockFactor * 2)
                            Me.Height = (clientRect.Bottom - clientRect.Top) * (dockFactor * 2)

                    End Select
                ElseIf My.Application.cmdLineArgs.Width > 0 And My.Application.cmdLineArgs.Height > 0 Then
                    Me.Width = My.Application.cmdLineArgs.Width
                    Me.Height = My.Application.cmdLineArgs.Height
                Else
                    Me.Width = clientRect.Right - clientRect.Left
                    Me.Height = clientRect.Bottom - clientRect.Top
                End If

            Else
                ' Position grid relative to the screen
                Me.Left = My.Application.cmdLineArgs.LocationX
            Me.Top = My.Application.cmdLineArgs.LocationY

            Me.Width = My.Application.cmdLineArgs.Width
            Me.Height = My.Application.cmdLineArgs.Height
        End If


        '' If the current grid location and size put it outside of the screen working area,
        '' shrink it to fit
        'Dim screenWorkingArea As Rectangle
        'screenWorkingArea = Screen.GetWorkingArea(Me)
        'If screenOriginPoint.X < screenWorkingArea.Left Then
        '    screenOriginPoint.X = screenWorkingArea.Left
        'End If

        'If screenOriginPoint.Y < screenWorkingArea.Top Then
        '    screenOriginPoint.Y = screenWorkingArea.Top
        'End If
        'Me.Top = screenOriginPoint.Y
        'Me.Left = screenOriginPoint.X

        'If screenOriginPoint.X + Me.Width > screenWorkingArea.Right Then
        '    Me.Width = screenWorkingArea.Right - screenOriginPoint.X
        'End If

        'If screenOriginPoint.Y + Me.Height > screenWorkingArea.Bottom Then
        '    Me.Height = screenWorkingArea.Bottom - screenOriginPoint.Y
        'End If

    End Sub

    ''' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' Used in all modes
    ''' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' We have already resized the grid.  
    ' Now we have to draw it.
    ' The current size of the form includes space for labels.
    ' Draw labels on all 4 sides and fill remaining space with grid lines
    Public Sub RenderGrid(g As Graphics)

        If recalculateLabelSize Then
            CalculateLabelFont(g)
            recalculateLabelSize = False
        End If

        ' Clear canvas
        Dim theBitmap As New Bitmap(pboxCanvas.Width, pboxCanvas.Height)
        theBitmap.MakeTransparent()
        g.DrawImage(theBitmap, 0, 0)

        Dim thePen As Pen
        If isSticky Then
            thePen = New Pen(gridLineColorSticky)
        Else
            thePen = New Pen(gridLineColor)
        End If


        Dim textBrush As Brush = Brushes.White
        g.SmoothingMode = SmoothingMode.AntiAlias


        ' Determine height/width of labels
        Dim cmdLineArgs As CommandLineArgs = My.Application.cmdLineArgs

        ' All coordinates are expressed in client coordinates of the picture box.
        Dim x As Integer, y As Integer, height As Integer, width As Integer, number As Integer
        Dim theSize As New SizeF
        Dim boxRectangle1 As Rectangle, boxRectangle2 As Rectangle
        Dim numberString As String
        height = pboxCanvas.Height
        width = pboxCanvas.Width

        ' Draw rows first
        ' Leave space for the top row of labels
        y = labelActualSize.Height
        number = 1


        ' If in a row entry and no numbers have been pressed yet, highlight all row labels
        ' If in row entry and numbers have been pressed, highlight matching row labels
        Do While (y < height)
            ' Draw top line
            g.DrawLine(thePen, 0, y, width, y)

            ' Draw labels
            numberString = CStr(number)
            theSize = g.MeasureString(numberString, labelActualFont)
            boxRectangle1 = New Rectangle(0, y, theSize.Width, theSize.Height)
            boxRectangle2 = New Rectangle(width - theSize.Width, y, theSize.Width, theSize.Height)

            If (theEntryMode = EntryMode.Row And rowEntry = "") Or
                (theEntryMode = EntryMode.Row And Strings.Left(numberString, Strings.Len(rowEntry)) = rowEntry) Or
                numberString = rowEntry Then
                g.FillRectangle(labelBackHighlightBrush, boxRectangle1)
                g.FillRectangle(labelBackHighlightBrush, boxRectangle2)
            Else
                g.FillRectangle(labelBackNormalBrush, boxRectangle1)
                g.FillRectangle(labelBackNormalBrush, boxRectangle2)
            End If
            g.DrawString(numberString, labelActualFont, textBrush, 0, y)
            g.DrawString(numberString, labelActualFont, textBrush, boxRectangle2.X, y)

            y = y + cmdLineArgs.RowHeight
            number = number + 1
        Loop

        ' Draw columns
        ' Leave space for the left column of labels
        x = labelActualSize.Width
        number = 1

        Do While (x < width)
            ' Draw left line
            g.DrawLine(thePen, x, 0, x, height)

            ' Draw label
            numberString = CStr(number)
            theSize = g.MeasureString(numberString, labelActualFont)
            boxRectangle1 = New Rectangle(x, 0, theSize.Width, theSize.Height)
            boxRectangle2 = New Rectangle(x, height - theSize.Height, theSize.Width, theSize.Height)
            If (theEntryMode = EntryMode.Column And columnEntry = "") Or
                (theEntryMode = EntryMode.Column And Strings.Left(numberString, Strings.Len(columnEntry)) = columnEntry) Or
                numberString = columnEntry Then
                g.FillRectangle(labelBackHighlightBrush, boxRectangle1)
                g.FillRectangle(labelBackHighlightBrush, boxRectangle2)
            Else
                g.FillRectangle(labelBackNormalBrush, boxRectangle1)
                g.FillRectangle(labelBackNormalBrush, boxRectangle2)
            End If

            g.DrawString(numberString, labelActualFont, textBrush, x, 0)
            g.DrawString(numberString, labelActualFont, textBrush, x, boxRectangle2.Y)

            x = x + cmdLineArgs.ColumnWidth
            number = number + 1
        Loop

    End Sub

    Private Sub CalculateLabelFont(g As Graphics)
        ' This should be called After the form has been resized

        ' Label size depends on the column and row width/height,
        ' as well as the number of rows/columns.  The more rows/columns,
        ' the bigger the number and the more space needed for the text.

        Dim cmdLineArgs As CommandLineArgs = My.Application.cmdLineArgs
        Dim numRows As Integer, numCols As Integer
        numRows = CInt(Math.Ceiling(pboxCanvas.Height / cmdLineArgs.RowHeight))
        numCols = CInt(Math.Ceiling(pboxCanvas.Width / cmdLineArgs.ColumnWidth))

        ' God forbid we Should have a grid with more than 999 rows/columns
        If numRows > 99 Or numCols > 99 Then
            ' 3 digits
            labelActualFont = GetAdjustedFont(g, "999", labelReferenceFont, cmdLineArgs.ColumnWidth, labelReferenceFont.Size)
            labelActualSize = g.MeasureString("999", labelActualFont)

        ElseIf numRows > 9 Or numCols > 9 Then
            ' 2 digits
            labelActualFont = GetAdjustedFont(g, "99", labelReferenceFont, cmdLineArgs.ColumnWidth, labelReferenceFont.Size)
            labelActualSize = g.MeasureString("99", labelActualFont)
        Else
            ' 1 digit
            labelActualFont = GetAdjustedFont(g, "9", labelReferenceFont, cmdLineArgs.ColumnWidth, labelReferenceFont.Size)
            labelActualSize = g.MeasureString("9", labelActualFont)
        End If


    End Sub

    Private Function GetAdjustedFont(g As Graphics, text As String, originalFont As Font, containerWidth As Integer,
                                     Optional maxFontSize As Integer = 12, Optional minFontSize As Integer = 1, Optional smallestOnFail As Boolean = False) As Font

        ' Search for smaller font
        For adjustedSize As Integer = maxFontSize To minFontSize Step -1
            Dim testFont As New Font(originalFont.Name, adjustedSize, originalFont.Style)
            Dim adjustedSizeNew As SizeF = g.MeasureString(text, testFont)

            If containerWidth > CInt(adjustedSizeNew.Width) Then
                Return testFont
            End If
        Next

        ' Did not find a font that fits
        If smallestOnFail Then
            Return New Font(originalFont.Name, minFontSize, originalFont.Style)
        Else
            Return originalFont
        End If

    End Function

    Private Sub FrmMain_Closed(sender As Object, e As EventArgs) Handles Me.Closed

    End Sub

    Private Sub FrmMain_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Dim clickPoint As Drawing.Point

        If e.KeyCode = Keys.F1 Then
            If frmHelp.Visible Then
                frmHelp.Focus()
                Exit Sub
            End If
            frmHelp.ShowDialog()

        ElseIf e.KeyCode = Keys.Escape Then
            ClearEntry()

        ElseIf e.KeyCode = Keys.R Then
            rowEntry = ""
            theEntryMode = EntryMode.Row
            RefreshGrid()

        ElseIf e.KeyCode = Keys.C Then
            columnEntry = ""
            theEntryMode = EntryMode.Column
            RefreshGrid()

        ElseIf e.KeyCode = Keys.Enter Then
            ' Finish entry, But leave highlights intact
            FinishEntry()

        ElseIf (e.KeyCode >= Keys.D0 And e.KeyCode <= Keys.D9) Or (e.KeyCode >= Keys.NumPad0 And e.KeyCode <= Keys.NumPad9) Then
            ' Get number and add to current entry
            Dim theChar As String
            If (e.KeyCode >= Keys.D0 And e.KeyCode <= Keys.D9) Then
                theChar = Chr(e.KeyCode)
            Else
                theChar = Chr(e.KeyCode - 48)
            End If

            If theEntryMode = EntryMode.Row Then
                rowEntry = rowEntry & theChar
                RefreshGrid()
            ElseIf theEntryMode = EntryMode.Column Then
                columnEntry = columnEntry & theChar
                RefreshGrid()
            End If

        ElseIf e.KeyCode = Keys.S Then
            ' Single click
            clickPoint = GetClickPoint()
            Me.Hide()
            Thread.Sleep(100)
            If e.KeyData = (Keys.Control + Keys.S) Then
                InputUtils.MouseLeftClick(clickPoint, False, True)
            ElseIf e.KeyData = (Keys.Shift + Keys.S) Then
                InputUtils.MouseLeftClick(clickPoint, True)
            Else
                InputUtils.MouseLeftClick(clickPoint)
            End If
            FinishAction()

        ElseIf e.KeyCode = Keys.D Then
            ' Double click
            clickPoint = GetClickPoint()
            Me.Hide()
            Thread.Sleep(100)
            If e.KeyData = (Keys.Control + Keys.D) Then
                InputUtils.MouseDoubleClick(clickPoint, False, True)
            ElseIf e.KeyData = (Keys.Shift + Keys.D) Then
                InputUtils.MouseDoubleClick(clickPoint, True)
            Else
                InputUtils.MouseDoubleClick(clickPoint)
            End If
            FinishAction()

        ElseIf e.KeyCode = Keys.T Then
            ' Right click
            clickPoint = GetClickPoint()
            Me.Hide()
            Thread.Sleep(100)
            If e.KeyData = (Keys.Control + Keys.T) Then
                InputUtils.MouseRightClick(clickPoint, False, True)
            ElseIf e.KeyData = (Keys.Shift + Keys.T) Then
                InputUtils.MouseRightClick(clickPoint, True)
            Else
                InputUtils.MouseRightClick(clickPoint)
            End If
            FinishAction()

        ElseIf e.KeyCode = Keys.F Then
            ' Prep for drag
            Cursor.Position = GetClickPoint()

        ElseIf e.KeyCode = Keys.M Then
            ' Move the mouse
            Cursor.Position = GetClickPoint()
            FinishAction()

        ElseIf e.KeyCode = Keys.G Then
            ' Drag mouse
            clickPoint = GetClickPoint()
            Me.Hide()
            Thread.Sleep(100)
            If e.KeyData = (Keys.Control + Keys.G) Then
                InputUtils.KeyDown(Keys.Control)
                InputUtils.MouseLeftDown(Cursor.Position)
                Cursor.Position = clickPoint
                Thread.Sleep(100)
                InputUtils.MouseLeftUp(clickPoint)
                InputUtils.KeyUp(Keys.Control)
            ElseIf e.KeyData = (Keys.Shift + Keys.G) Then
                InputUtils.KeyDown(Keys.Shift)
                InputUtils.MouseLeftDown(Cursor.Position)
                Cursor.Position = clickPoint
                Thread.Sleep(100)
                InputUtils.MouseLeftUp(clickPoint)
                InputUtils.KeyUp(Keys.Shift)
            Else
                InputUtils.MouseLeftDown(Cursor.Position)
                Cursor.Position = clickPoint
                Thread.Sleep(100)
                InputUtils.MouseLeftUp(clickPoint)
            End If
            FinishAction()

        ElseIf e.KeyCode = Keys.Y Then
            ' Toggle sticky flag
            isSticky = Not isSticky
            RefreshGrid()

        ElseIf e.KeyCode = keys.Up Then
            HandleIncreaseRowHeight()

        ElseIf e.KeyCode = keys.Down Then
            HandleDecreaseRowHeight()

        ElseIf e.KeyCode = keys.Left Then
            HandleDecreaseColumnWidth()

        ElseIf e.KeyCode = keys.Right Then
            HandleIncreaseColumnWidth()

        ElseIf e.KeyCode = Keys.X Then
            ' Close the application
            Me.Close()

        End If

    End Sub

    Private Sub RefreshGrid()
        pboxCanvas.Invalidate()
        pboxCanvas.Refresh()
    End Sub
    Private Sub ClearEntry()
        theEntryMode = EntryMode.None
        rowEntry = ""
        columnEntry = ""
        pboxCanvas.Invalidate()
        pboxCanvas.Refresh()
    End Sub

    Private Sub FinishEntry()
        theEntryMode = EntryMode.None
        pboxCanvas.Invalidate()
        pboxCanvas.Refresh()
    End Sub

    Private Function GetClickPoint() As Drawing.Point
        ' Get the coordinates of the center of the box at the cross-section of the selected row and column
        ' Click points relative to the picture box control
        ' Return Origin point if there is no selected row or column
        Dim thePoint As New Drawing.Point

        If rowEntry = "" Or columnEntry = "" Then
            thePoint.X = 0
            thePoint.Y = 0
            Return thePoint
        End If

        ' calculate the center of the selected grid box In client coordinates
        Dim rowNumber As Integer, columnNumber As Integer
        rowNumber = CInt(rowEntry)
        columnNumber = CInt(columnEntry)

        thePoint.X = pboxCanvas.Left + labelActualSize.Width + ((columnNumber - 1) * My.Application.cmdLineArgs.ColumnWidth) + (My.Application.cmdLineArgs.ColumnWidth / 2)
        thePoint.Y = pboxCanvas.Top + labelActualSize.Height + ((rowNumber - 1) * My.Application.cmdLineArgs.RowHeight) + (My.Application.cmdLineArgs.RowHeight / 2)

        ' Convert to screen coordinates
        ClientToScreen(Me.Handle, thePoint)

        Return thePoint
    End Function

    Public Sub FinishAction()
        If Not isSticky Then
            Me.Close()
        Else
            ' grab the focus back
            Me.Show()
            BringToForeground()
        End If
    End Sub

    Public Sub BringToForeground()
        SetForegroundWindow(Me.Handle)
    End Sub

    Public Function GetCurrentTargetWindowHandle() As IntPtr
        Return currentTargetWindowHandle
    End Function

    Private Sub frmMain_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

    End Sub

    Private Sub mnuHelp_Click(sender As Object, e As EventArgs) Handles mnuHelp.Click
        frmHelp.Show()
    End Sub

    Private Sub mnuUseScreen_Click(sender As Object, e As EventArgs) Handles mnuUseScreen.Click
        SetRelativeToClient(False)
        UpdateStatusBar()
    End Sub

    Private Sub mnuUseClient_Click(sender As Object, e As EventArgs) Handles mnuUseClient.Click
        SetRelativeToClient(True)
        UpdateStatusBar()
    End Sub

    Private Sub frmMain_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged
        UpdateStatusBar()
    End Sub

    Private Sub frmMain_ClientSizeChanged(sender As Object, e As EventArgs) Handles Me.ClientSizeChanged
        UpdateStatusBar()
        pboxCanvas.Invalidate()
        Me.Refresh()
    End Sub

    Private Sub mnuSticky_Click(sender As Object, e As EventArgs) Handles mnuSticky.Click
        My.Application.cmdLineArgs.IsSticky = Not My.Application.cmdLineArgs.IsSticky
        mnuSticky.Checked = My.Application.cmdLineArgs.IsSticky
        UpdateStatusBar()
    End Sub

    Private Sub mnuOpacity_Click(sender As Object, e As EventArgs) Handles mnuOpacity.Click
        frmOpacity.Init(Me)
        If frmOpacity.ShowDialog() = DialogResult.OK Then
            My.Application.cmdLineArgs.Opacity = frmOpacity.GridOpacity
            Me.Opacity = frmOpacity.GridOpacity
            UpdateStatusBar()
        End If
        frmOpacity.Dispose()
    End Sub

    Private Sub mnuAlwaysOnTop_Click(sender As Object, e As EventArgs) Handles mnuAlwaysOnTop.Click
        My.Application.cmdLineArgs.IsAlwaysOnTop = Not My.Application.cmdLineArgs.IsAlwaysOnTop
        mnuAlwaysOnTop.Checked = My.Application.cmdLineArgs.IsAlwaysOnTop
        UpdateStatusBar()
    End Sub

    Private Sub mnuCopy_Click(sender As Object, e As EventArgs) Handles mnuCopy.Click
        Clipboard.SetText(GetParameterString())
    End Sub

    Private Sub mnuCopyAndExit_Click(sender As Object, e As EventArgs) Handles mnuCopyAndExit.Click
        Clipboard.SetText(GetParameterString())
        Me.Close()
    End Sub

    Private Sub mnuExit_Click(sender As Object, e As EventArgs) Handles mnuExit.Click
        Me.Close()
    End Sub

    Private Sub pboxCanvas_Paint(sender As Object, e As PaintEventArgs) Handles pboxCanvas.Paint
        RenderGrid(e.Graphics)
    End Sub

    Private Sub mnuIncRowHeight_Click(sender As Object, e As EventArgs) Handles mnuIncRowHeight.Click
        HandleIncreaseRowHeight()
    End Sub

    Private Sub mnuDecRowHeight_Click(sender As Object, e As EventArgs) Handles mnuDecRowHeight.Click
        HandleDecreaseRowHeight()
    End Sub

    Private Sub mnuIncColWidth_Click(sender As Object, e As EventArgs) Handles mnuIncColWidth.Click
        HandleIncreaseColumnWidth()
    End Sub

    Private Sub mnuDecColWidth_Click(sender As Object, e As EventArgs) Handles mnuDecColWidth.Click
        HandleDecreaseColumnWidth()
    End Sub

    Private Sub HandleIncreaseRowHeight()
        My.Application.cmdLineArgs.RowHeight = My.Application.cmdLineArgs.RowHeight + 1
        recalculateLabelSize = True
        RefreshGrid()
    End Sub

    Private Sub HandleDecreaseRowHeight()
        My.Application.cmdLineArgs.RowHeight = My.Application.cmdLineArgs.RowHeight - 1
        recalculateLabelSize = True
        RefreshGrid()
    End Sub
    Private Sub HandleDecreaseColumnWidth()
        My.Application.cmdLineArgs.ColumnWidth = My.Application.cmdLineArgs.ColumnWidth - 1
        recalculateLabelSize = True
        RefreshGrid()
    End Sub

    Private Sub HandleIncreaseColumnWidth()
        My.Application.cmdLineArgs.ColumnWidth = My.Application.cmdLineArgs.ColumnWidth + 1
        recalculateLabelSize = True
        RefreshGrid()
    End Sub
End Class



<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.mnuMain = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCopy = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCopyAndExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.EditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuUseScreen = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuUseClient = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuIncRowHeight = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuDecRowHeight = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuIncColWidth = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuDecColWidth = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuSticky = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuOpacity = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuAlwaysOnTop = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutMouseGridToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.statusBarLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.pboxCanvas = New System.Windows.Forms.PictureBox()
        Me.statusBar = New System.Windows.Forms.StatusStrip()
        Me.mnuMain.SuspendLayout()
        CType(Me.pboxCanvas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.statusBar.SuspendLayout()
        Me.SuspendLayout()
        '
        'mnuMain
        '
        Me.mnuMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.EditToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.mnuMain.Location = New System.Drawing.Point(0, 0)
        Me.mnuMain.Name = "mnuMain"
        Me.mnuMain.Size = New System.Drawing.Size(693, 24)
        Me.mnuMain.TabIndex = 6
        Me.mnuMain.Text = "MenuStrip1"
        Me.mnuMain.Visible = False
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCopy, Me.mnuCopyAndExit, Me.mnuExit})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'mnuCopy
        '
        Me.mnuCopy.Name = "mnuCopy"
        Me.mnuCopy.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.mnuCopy.Size = New System.Drawing.Size(208, 22)
        Me.mnuCopy.Text = "&Copy Parameters"
        '
        'mnuCopyAndExit
        '
        Me.mnuCopyAndExit.Name = "mnuCopyAndExit"
        Me.mnuCopyAndExit.Size = New System.Drawing.Size(208, 22)
        Me.mnuCopyAndExit.Text = "Copy &Parameters and Exit"
        '
        'mnuExit
        '
        Me.mnuExit.Name = "mnuExit"
        Me.mnuExit.Size = New System.Drawing.Size(208, 22)
        Me.mnuExit.Text = "E&xit"
        '
        'EditToolStripMenuItem
        '
        Me.EditToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSeparator1, Me.mnuUseScreen, Me.mnuUseClient, Me.ToolStripSeparator2, Me.mnuIncRowHeight, Me.mnuDecRowHeight, Me.mnuIncColWidth, Me.mnuDecColWidth, Me.ToolStripSeparator3, Me.mnuSticky, Me.mnuOpacity, Me.mnuAlwaysOnTop})
        Me.EditToolStripMenuItem.Name = "EditToolStripMenuItem"
        Me.EditToolStripMenuItem.Size = New System.Drawing.Size(61, 20)
        Me.EditToolStripMenuItem.Text = "&Options"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(257, 6)
        '
        'mnuUseScreen
        '
        Me.mnuUseScreen.Name = "mnuUseScreen"
        Me.mnuUseScreen.Size = New System.Drawing.Size(260, 22)
        Me.mnuUseScreen.Text = "Use &Screen Coordinates"
        '
        'mnuUseClient
        '
        Me.mnuUseClient.Name = "mnuUseClient"
        Me.mnuUseClient.Size = New System.Drawing.Size(260, 22)
        Me.mnuUseClient.Text = "Use &Client Coordinates"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(257, 6)
        '
        'mnuIncRowHeight
        '
        Me.mnuIncRowHeight.Name = "mnuIncRowHeight"
        Me.mnuIncRowHeight.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Up), System.Windows.Forms.Keys)
        Me.mnuIncRowHeight.Size = New System.Drawing.Size(260, 22)
        Me.mnuIncRowHeight.Text = "&Increase Row Height"
        '
        'mnuDecRowHeight
        '
        Me.mnuDecRowHeight.Name = "mnuDecRowHeight"
        Me.mnuDecRowHeight.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Down), System.Windows.Forms.Keys)
        Me.mnuDecRowHeight.Size = New System.Drawing.Size(260, 22)
        Me.mnuDecRowHeight.Text = "&Decrease Row Height"
        '
        'mnuIncColWidth
        '
        Me.mnuIncColWidth.Name = "mnuIncColWidth"
        Me.mnuIncColWidth.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Right), System.Windows.Forms.Keys)
        Me.mnuIncColWidth.Size = New System.Drawing.Size(260, 22)
        Me.mnuIncColWidth.Text = "Increase Column &Width"
        '
        'mnuDecColWidth
        '
        Me.mnuDecColWidth.Name = "mnuDecColWidth"
        Me.mnuDecColWidth.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Left), System.Windows.Forms.Keys)
        Me.mnuDecColWidth.Size = New System.Drawing.Size(260, 22)
        Me.mnuDecColWidth.Text = "Decrease Column Wi&dth"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(257, 6)
        '
        'mnuSticky
        '
        Me.mnuSticky.Name = "mnuSticky"
        Me.mnuSticky.Size = New System.Drawing.Size(260, 22)
        Me.mnuSticky.Text = "&Sticky Grid"
        '
        'mnuOpacity
        '
        Me.mnuOpacity.Name = "mnuOpacity"
        Me.mnuOpacity.Size = New System.Drawing.Size(260, 22)
        Me.mnuOpacity.Text = "&Opacity"
        '
        'mnuAlwaysOnTop
        '
        Me.mnuAlwaysOnTop.Name = "mnuAlwaysOnTop"
        Me.mnuAlwaysOnTop.Size = New System.Drawing.Size(260, 22)
        Me.mnuAlwaysOnTop.Text = "&Always On Top"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuHelp, Me.AboutMouseGridToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "&Help"
        '
        'mnuHelp
        '
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Size = New System.Drawing.Size(180, 22)
        Me.mnuHelp.Text = "Mouse Grid &Help"
        '
        'AboutMouseGridToolStripMenuItem
        '
        Me.AboutMouseGridToolStripMenuItem.Name = "AboutMouseGridToolStripMenuItem"
        Me.AboutMouseGridToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.AboutMouseGridToolStripMenuItem.Text = "&About Mouse Grid..."
        '
        'statusBarLabel
        '
        Me.statusBarLabel.Name = "statusBarLabel"
        Me.statusBarLabel.Size = New System.Drawing.Size(103, 17)
        Me.statusBarLabel.Text = "This is a status bar"
        '
        'pboxCanvas
        '
        Me.pboxCanvas.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pboxCanvas.Location = New System.Drawing.Point(0, 0)
        Me.pboxCanvas.Name = "pboxCanvas"
        Me.pboxCanvas.Size = New System.Drawing.Size(693, 410)
        Me.pboxCanvas.TabIndex = 8
        Me.pboxCanvas.TabStop = False
        '
        'statusBar
        '
        Me.statusBar.BackColor = System.Drawing.Color.Gainsboro
        Me.statusBar.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.statusBarLabel})
        Me.statusBar.Location = New System.Drawing.Point(0, 388)
        Me.statusBar.Name = "statusBar"
        Me.statusBar.Size = New System.Drawing.Size(693, 22)
        Me.statusBar.TabIndex = 7
        Me.statusBar.Text = "StatusStrip1"
        Me.statusBar.Visible = False
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(693, 410)
        Me.Controls.Add(Me.pboxCanvas)
        Me.Controls.Add(Me.statusBar)
        Me.Controls.Add(Me.mnuMain)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.mnuMain
        Me.Name = "frmMain"
        Me.Opacity = 0.1R
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Mouse Grid"
        Me.TransparencyKey = System.Drawing.Color.LightGray
        Me.mnuMain.ResumeLayout(False)
        Me.mnuMain.PerformLayout()
        CType(Me.pboxCanvas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.statusBar.ResumeLayout(False)
        Me.statusBar.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblPrompt As Label
    Friend WithEvents mnuMain As MenuStrip
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents mnuExit As ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents mnuHelp As ToolStripMenuItem
    Friend WithEvents mnuCopyAndExit As ToolStripMenuItem
    Friend WithEvents EditToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents mnuUseScreen As ToolStripMenuItem
    Friend WithEvents mnuUseClient As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents mnuCopy As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents mnuIncRowHeight As ToolStripMenuItem
    Friend WithEvents mnuDecRowHeight As ToolStripMenuItem
    Friend WithEvents mnuIncColWidth As ToolStripMenuItem
    Friend WithEvents mnuDecColWidth As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents mnuSticky As ToolStripMenuItem
    Friend WithEvents mnuOpacity As ToolStripMenuItem
    Friend WithEvents mnuAlwaysOnTop As ToolStripMenuItem
    Friend WithEvents statusBarLabel As ToolStripStatusLabel
    Friend WithEvents pboxCanvas As PictureBox
    Friend WithEvents statusBar As StatusStrip
    Friend WithEvents AboutMouseGridToolStripMenuItem As ToolStripMenuItem
End Class

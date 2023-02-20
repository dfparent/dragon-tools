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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.timPrompt = New System.Windows.Forms.Timer(Me.components)
        Me.timSplash = New System.Windows.Forms.Timer(Me.components)
        Me.txtOutput = New System.Windows.Forms.TextBox()
        Me.trayIcon = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.trayIconMenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ExitTrayIconMenuStripItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.timAnimation = New System.Windows.Forms.Timer(Me.components)
        Me.lblPrompt = New System.Windows.Forms.Label()
        Me.trayIconMenuStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'timPrompt
        '
        '
        'timSplash
        '
        '
        'txtOutput
        '
        Me.txtOutput.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.txtOutput.ForeColor = System.Drawing.Color.Red
        Me.txtOutput.Location = New System.Drawing.Point(86, 263)
        Me.txtOutput.Multiline = True
        Me.txtOutput.Name = "txtOutput"
        Me.txtOutput.ReadOnly = True
        Me.txtOutput.Size = New System.Drawing.Size(540, 175)
        Me.txtOutput.TabIndex = 1
        Me.txtOutput.Visible = False
        '
        'trayIcon
        '
        Me.trayIcon.ContextMenuStrip = Me.trayIconMenuStrip
        Me.trayIcon.Icon = CType(resources.GetObject("trayIcon.Icon"), System.Drawing.Icon)
        Me.trayIcon.Text = "Click By Numbers"
        Me.trayIcon.Visible = True
        '
        'trayIconMenuStrip
        '
        Me.trayIconMenuStrip.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.trayIconMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExitTrayIconMenuStripItem})
        Me.trayIconMenuStrip.Name = "trayIconMenuStrip"
        Me.trayIconMenuStrip.Size = New System.Drawing.Size(94, 26)
        '
        'ExitTrayIconMenuStripItem
        '
        Me.ExitTrayIconMenuStripItem.Name = "ExitTrayIconMenuStripItem"
        Me.ExitTrayIconMenuStripItem.Size = New System.Drawing.Size(93, 22)
        Me.ExitTrayIconMenuStripItem.Text = "E&xit"
        '
        'timAnimation
        '
        Me.timAnimation.Interval = 50
        '
        'lblPrompt
        '
        Me.lblPrompt.AutoSize = True
        Me.lblPrompt.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.lblPrompt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPrompt.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.lblPrompt.Font = New System.Drawing.Font("Calibri", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPrompt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPrompt.Location = New System.Drawing.Point(379, 213)
        Me.lblPrompt.Name = "lblPrompt"
        Me.lblPrompt.Size = New System.Drawing.Size(117, 25)
        Me.lblPrompt.TabIndex = 2
        Me.lblPrompt.Text = "Prompt Label"
        Me.lblPrompt.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblPrompt.Visible = False
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.lblPrompt)
        Me.Controls.Add(Me.txtOutput)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMain"
        Me.Opacity = 0.75R
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Click By Numbers"
        Me.TopMost = True
        Me.TransparencyKey = System.Drawing.Color.LightGray
        Me.trayIconMenuStrip.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents timPrompt As Timer
    Friend WithEvents timSplash As Timer
    Friend WithEvents txtOutput As TextBox
    Friend WithEvents trayIcon As NotifyIcon
    Friend WithEvents trayIconMenuStrip As ContextMenuStrip
    Friend WithEvents ExitTrayIconMenuStripItem As ToolStripMenuItem
    Friend WithEvents timAnimation As Timer
    Friend WithEvents lblPrompt As Label
End Class

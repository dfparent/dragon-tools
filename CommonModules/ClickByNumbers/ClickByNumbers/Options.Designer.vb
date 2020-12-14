<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOptions
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOptions))
        Me.btnClose = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboApplyTo = New System.Windows.Forms.ComboBox()
        Me.chkSticky = New System.Windows.Forms.CheckBox()
        Me.chkDisable = New System.Windows.Forms.CheckBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.chkUsesWindowHandles = New System.Windows.Forms.CheckBox()
        Me.chkUsesUIAutomation = New System.Windows.Forms.CheckBox()
        Me.chkUsesMSAA = New System.Windows.Forms.CheckBox()
        Me.btnAddProcess = New System.Windows.Forms.Button()
        Me.btnRemoveProcess = New System.Windows.Forms.Button()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.updnOpacity = New System.Windows.Forms.NumericUpDown()
        Me.GroupBox1.SuspendLayout()
        CType(Me.updnOpacity, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(176, 300)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 0
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(86, 13)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "Options apply to:"
        '
        'cboApplyTo
        '
        Me.cboApplyTo.FormattingEnabled = True
        Me.cboApplyTo.Location = New System.Drawing.Point(104, 6)
        Me.cboApplyTo.Name = "cboApplyTo"
        Me.cboApplyTo.Size = New System.Drawing.Size(176, 21)
        Me.cboApplyTo.TabIndex = 16
        '
        'chkSticky
        '
        Me.chkSticky.AutoSize = True
        Me.chkSticky.Location = New System.Drawing.Point(15, 64)
        Me.chkSticky.Name = "chkSticky"
        Me.chkSticky.Size = New System.Drawing.Size(83, 17)
        Me.chkSticky.TabIndex = 13
        Me.chkSticky.Text = "&Sticky Flags"
        Me.chkSticky.UseVisualStyleBackColor = True
        '
        'chkDisable
        '
        Me.chkDisable.AutoSize = True
        Me.chkDisable.Location = New System.Drawing.Point(15, 87)
        Me.chkDisable.Name = "chkDisable"
        Me.chkDisable.Size = New System.Drawing.Size(110, 17)
        Me.chkDisable.TabIndex = 14
        Me.chkDisable.Text = "&Hide Flags Initially"
        Me.chkDisable.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chkUsesWindowHandles)
        Me.GroupBox1.Controls.Add(Me.chkUsesUIAutomation)
        Me.GroupBox1.Controls.Add(Me.chkUsesMSAA)
        Me.GroupBox1.Location = New System.Drawing.Point(15, 169)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(327, 100)
        Me.GroupBox1.TabIndex = 20
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Discovery Types"
        '
        'chkUsesWindowHandles
        '
        Me.chkUsesWindowHandles.AutoSize = True
        Me.chkUsesWindowHandles.Location = New System.Drawing.Point(6, 30)
        Me.chkUsesWindowHandles.Name = "chkUsesWindowHandles"
        Me.chkUsesWindowHandles.Size = New System.Drawing.Size(294, 17)
        Me.chkUsesWindowHandles.TabIndex = 20
        Me.chkUsesWindowHandles.Text = "Window Handles / EnumWindows (Fastest, fewest finds)"
        Me.chkUsesWindowHandles.UseVisualStyleBackColor = True
        '
        'chkUsesUIAutomation
        '
        Me.chkUsesUIAutomation.AutoSize = True
        Me.chkUsesUIAutomation.Location = New System.Drawing.Point(6, 53)
        Me.chkUsesUIAutomation.Name = "chkUsesUIAutomation"
        Me.chkUsesUIAutomation.Size = New System.Drawing.Size(190, 17)
        Me.chkUsesUIAutomation.TabIndex = 21
        Me.chkUsesUIAutomation.Text = "UI Automation (Slower, many finds)"
        Me.chkUsesUIAutomation.UseVisualStyleBackColor = True
        '
        'chkUsesMSAA
        '
        Me.chkUsesMSAA.AutoSize = True
        Me.chkUsesMSAA.Location = New System.Drawing.Point(6, 76)
        Me.chkUsesMSAA.Name = "chkUsesMSAA"
        Me.chkUsesMSAA.Size = New System.Drawing.Size(300, 17)
        Me.chkUsesMSAA.TabIndex = 22
        Me.chkUsesMSAA.Text = "Microsoft Active Accessibility (MSAA) (Slowest, most finds)"
        Me.chkUsesMSAA.UseVisualStyleBackColor = True
        '
        'btnAddProcess
        '
        Me.btnAddProcess.AutoSize = True
        Me.btnAddProcess.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnAddProcess.Location = New System.Drawing.Point(287, 6)
        Me.btnAddProcess.Name = "btnAddProcess"
        Me.btnAddProcess.Size = New System.Drawing.Size(23, 23)
        Me.btnAddProcess.TabIndex = 21
        Me.btnAddProcess.Text = "+"
        Me.btnAddProcess.UseVisualStyleBackColor = True
        '
        'btnRemoveProcess
        '
        Me.btnRemoveProcess.AutoSize = True
        Me.btnRemoveProcess.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnRemoveProcess.Location = New System.Drawing.Point(316, 6)
        Me.btnRemoveProcess.Name = "btnRemoveProcess"
        Me.btnRemoveProcess.Size = New System.Drawing.Size(20, 23)
        Me.btnRemoveProcess.TabIndex = 22
        Me.btnRemoveProcess.Text = "-"
        Me.btnRemoveProcess.UseVisualStyleBackColor = True
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(92, 300)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(75, 23)
        Me.cmdSave.TabIndex = 23
        Me.cmdSave.Text = "Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 110)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(57, 13)
        Me.Label2.TabIndex = 25
        Me.Label2.Text = "Opacity %:"
        '
        'updnOpacity
        '
        Me.updnOpacity.Location = New System.Drawing.Point(75, 108)
        Me.updnOpacity.Name = "updnOpacity"
        Me.updnOpacity.Size = New System.Drawing.Size(43, 20)
        Me.updnOpacity.TabIndex = 26
        '
        'frmOptions
        '
        Me.AcceptButton = Me.cmdSave
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(355, 342)
        Me.Controls.Add(Me.updnOpacity)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.btnRemoveProcess)
        Me.Controls.Add(Me.btnAddProcess)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboApplyTo)
        Me.Controls.Add(Me.chkSticky)
        Me.Controls.Add(Me.chkDisable)
        Me.Controls.Add(Me.btnClose)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmOptions"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Click By Numbers Options"
        Me.TopMost = True
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.updnOpacity, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnClose As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents cboApplyTo As ComboBox
    Friend WithEvents chkSticky As CheckBox
    Friend WithEvents chkDisable As CheckBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents chkUsesWindowHandles As CheckBox
    Friend WithEvents chkUsesUIAutomation As CheckBox
    Friend WithEvents chkUsesMSAA As CheckBox
    Friend WithEvents btnAddProcess As Button
    Friend WithEvents btnRemoveProcess As Button
    Friend WithEvents cmdSave As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents updnOpacity As NumericUpDown
End Class

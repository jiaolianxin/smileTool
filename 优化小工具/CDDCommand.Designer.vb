<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CDDCommand
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
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

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CDDCommand))
        Me.CellCommand = New System.Windows.Forms.CheckBox()
        Me.MOCommand = New System.Windows.Forms.CheckBox()
        Me.MSCCommand = New System.Windows.Forms.CheckBox()
        Me.Cell_MOCommand = New System.Windows.Forms.CheckBox()
        Me.NDCommand = New System.Windows.Forms.CheckBox()
        Me.start_button = New System.Windows.Forms.Button()
        Me.CommandSave = New System.Windows.Forms.FolderBrowserDialog()
        Me.SuspendLayout()
        '
        'CellCommand
        '
        Me.CellCommand.AutoSize = True
        Me.CellCommand.Location = New System.Drawing.Point(27, 95)
        Me.CellCommand.Name = "CellCommand"
        Me.CellCommand.Size = New System.Drawing.Size(96, 16)
        Me.CellCommand.TabIndex = 0
        Me.CellCommand.Text = "Cell+BSC指令"
        Me.CellCommand.UseVisualStyleBackColor = True
        '
        'MOCommand
        '
        Me.MOCommand.AutoSize = True
        Me.MOCommand.Location = New System.Drawing.Point(158, 95)
        Me.MOCommand.Name = "MOCommand"
        Me.MOCommand.Size = New System.Drawing.Size(60, 16)
        Me.MOCommand.TabIndex = 2
        Me.MOCommand.Text = "MO指令"
        Me.MOCommand.UseVisualStyleBackColor = True
        '
        'MSCCommand
        '
        Me.MSCCommand.AutoSize = True
        Me.MSCCommand.Location = New System.Drawing.Point(27, 146)
        Me.MSCCommand.Name = "MSCCommand"
        Me.MSCCommand.Size = New System.Drawing.Size(66, 16)
        Me.MSCCommand.TabIndex = 3
        Me.MSCCommand.Text = "MSC指令"
        Me.MSCCommand.UseVisualStyleBackColor = True
        '
        'Cell_MOCommand
        '
        Me.Cell_MOCommand.AutoSize = True
        Me.Cell_MOCommand.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Cell_MOCommand.Location = New System.Drawing.Point(27, 46)
        Me.Cell_MOCommand.Name = "Cell_MOCommand"
        Me.Cell_MOCommand.Size = New System.Drawing.Size(114, 16)
        Me.Cell_MOCommand.TabIndex = 4
        Me.Cell_MOCommand.Text = "Cell+MO常用指令"
        Me.Cell_MOCommand.UseVisualStyleBackColor = True
        '
        'NDCommand
        '
        Me.NDCommand.AutoSize = True
        Me.NDCommand.ForeColor = System.Drawing.Color.Red
        Me.NDCommand.Location = New System.Drawing.Point(158, 46)
        Me.NDCommand.Name = "NDCommand"
        Me.NDCommand.Size = New System.Drawing.Size(84, 16)
        Me.NDCommand.TabIndex = 5
        Me.NDCommand.Text = "ND所需指令"
        Me.NDCommand.UseVisualStyleBackColor = True
        '
        'start_button
        '
        Me.start_button.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.start_button.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.start_button.Location = New System.Drawing.Point(95, 213)
        Me.start_button.Name = "start_button"
        Me.start_button.Size = New System.Drawing.Size(75, 23)
        Me.start_button.TabIndex = 6
        Me.start_button.Text = "开始"
        Me.start_button.UseVisualStyleBackColor = True
        '
        'CDDCommand
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(284, 262)
        Me.Controls.Add(Me.start_button)
        Me.Controls.Add(Me.NDCommand)
        Me.Controls.Add(Me.Cell_MOCommand)
        Me.Controls.Add(Me.MSCCommand)
        Me.Controls.Add(Me.MOCommand)
        Me.Controls.Add(Me.CellCommand)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(300, 300)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(300, 300)
        Me.Name = "CDDCommand"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CDDCommand"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CellCommand As System.Windows.Forms.CheckBox
    Friend WithEvents MOCommand As System.Windows.Forms.CheckBox
    Friend WithEvents MSCCommand As System.Windows.Forms.CheckBox
    Friend WithEvents Cell_MOCommand As System.Windows.Forms.CheckBox
    Friend WithEvents NDCommand As System.Windows.Forms.CheckBox
    Friend WithEvents start_button As System.Windows.Forms.Button
    Friend WithEvents CommandSave As System.Windows.Forms.FolderBrowserDialog
End Class

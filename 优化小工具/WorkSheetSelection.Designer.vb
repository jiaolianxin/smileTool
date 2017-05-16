<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WorkSheetsSelection
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.WorkSheetsGroup = New System.Windows.Forms.ComboBox()
        Me.StartImport = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.WorkSheetsGroup)
        Me.GroupBox1.Location = New System.Drawing.Point(30, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(319, 101)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "请选择要操作的Sheet"
        '
        'WorkSheetsGroup
        '
        Me.WorkSheetsGroup.FormattingEnabled = True
        Me.WorkSheetsGroup.Items.AddRange(New Object() {"ssdsds", "sdsdsdff", "fffff", "eee", "", "w", "ww"})
        Me.WorkSheetsGroup.Location = New System.Drawing.Point(27, 52)
        Me.WorkSheetsGroup.Name = "WorkSheetsGroup"
        Me.WorkSheetsGroup.Size = New System.Drawing.Size(274, 20)
        Me.WorkSheetsGroup.TabIndex = 0
        '
        'StartImport
        '
        Me.StartImport.Location = New System.Drawing.Point(144, 130)
        Me.StartImport.Name = "StartImport"
        Me.StartImport.Size = New System.Drawing.Size(75, 23)
        Me.StartImport.TabIndex = 1
        Me.StartImport.Text = "Start"
        Me.StartImport.UseVisualStyleBackColor = True
        '
        'WorkSheetsSelection
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(383, 165)
        Me.Controls.Add(Me.StartImport)
        Me.Controls.Add(Me.GroupBox1)
        Me.MaximumSize = New System.Drawing.Size(399, 203)
        Me.MinimumSize = New System.Drawing.Size(399, 203)
        Me.Name = "WorkSheetsSelection"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "WorkSheet选择"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents WorkSheetsGroup As System.Windows.Forms.ComboBox
    Friend WithEvents StartImport As System.Windows.Forms.Button
End Class

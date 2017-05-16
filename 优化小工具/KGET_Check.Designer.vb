<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class KGET_Check
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
        Me.IsCheckPCI = New System.Windows.Forms.CheckBox()
        Me.IsCheckNeighbour = New System.Windows.Forms.CheckBox()
        Me.Start = New System.Windows.Forms.Button()
        Me.BGWcheckKGET = New System.ComponentModel.BackgroundWorker()
        Me.SuspendLayout()
        '
        'IsCheckPCI
        '
        Me.IsCheckPCI.AutoSize = True
        Me.IsCheckPCI.Location = New System.Drawing.Point(42, 51)
        Me.IsCheckPCI.Name = "IsCheckPCI"
        Me.IsCheckPCI.Size = New System.Drawing.Size(66, 16)
        Me.IsCheckPCI.TabIndex = 0
        Me.IsCheckPCI.Text = "检查PCI"
        Me.IsCheckPCI.ThreeState = True
        Me.IsCheckPCI.UseVisualStyleBackColor = True
        '
        'IsCheckNeighbour
        '
        Me.IsCheckNeighbour.AutoSize = True
        Me.IsCheckNeighbour.Location = New System.Drawing.Point(168, 51)
        Me.IsCheckNeighbour.Name = "IsCheckNeighbour"
        Me.IsCheckNeighbour.Size = New System.Drawing.Size(72, 16)
        Me.IsCheckNeighbour.TabIndex = 1
        Me.IsCheckNeighbour.Text = "检查邻区"
        Me.IsCheckNeighbour.UseVisualStyleBackColor = True
        '
        'Start
        '
        Me.Start.BackColor = System.Drawing.Color.Yellow
        Me.Start.Location = New System.Drawing.Point(95, 109)
        Me.Start.Name = "Start"
        Me.Start.Size = New System.Drawing.Size(75, 23)
        Me.Start.TabIndex = 2
        Me.Start.Text = "Start"
        Me.Start.UseVisualStyleBackColor = False
        '
        'BGWcheckKGET
        '
        Me.BGWcheckKGET.WorkerReportsProgress = True
        Me.BGWcheckKGET.WorkerSupportsCancellation = True
        '
        'KGET_Check
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 142)
        Me.Controls.Add(Me.Start)
        Me.Controls.Add(Me.IsCheckNeighbour)
        Me.Controls.Add(Me.IsCheckPCI)
        Me.MaximumSize = New System.Drawing.Size(300, 180)
        Me.MinimumSize = New System.Drawing.Size(300, 180)
        Me.Name = "KGET_Check"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "KGET_Check"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents IsCheckPCI As System.Windows.Forms.CheckBox
    Friend WithEvents IsCheckNeighbour As System.Windows.Forms.CheckBox
    Friend WithEvents Start As System.Windows.Forms.Button
    Friend WithEvents BGWcheckKGET As System.ComponentModel.BackgroundWorker
End Class

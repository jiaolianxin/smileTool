<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class HW_CDD_PARSE
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
        Me.GcellFilePath = New System.Windows.Forms.TextBox()
        Me.GtrxFilePath = New System.Windows.Forms.TextBox()
        Me.importGcell = New System.Windows.Forms.Button()
        Me.importGTRX = New System.Windows.Forms.Button()
        Me.startParse = New System.Windows.Forms.Button()
        Me.BGWForHWPara = New System.ComponentModel.BackgroundWorker()
        Me.SuspendLayout()
        '
        'GcellFilePath
        '
        Me.GcellFilePath.Location = New System.Drawing.Point(13, 30)
        Me.GcellFilePath.Name = "GcellFilePath"
        Me.GcellFilePath.ReadOnly = True
        Me.GcellFilePath.Size = New System.Drawing.Size(275, 21)
        Me.GcellFilePath.TabIndex = 0
        Me.GcellFilePath.Text = "请选择GCELL文件!"
        '
        'GtrxFilePath
        '
        Me.GtrxFilePath.Location = New System.Drawing.Point(13, 94)
        Me.GtrxFilePath.Name = "GtrxFilePath"
        Me.GtrxFilePath.ReadOnly = True
        Me.GtrxFilePath.Size = New System.Drawing.Size(275, 21)
        Me.GtrxFilePath.TabIndex = 1
        Me.GtrxFilePath.Text = "请选择GTRX文件!"
        '
        'importGcell
        '
        Me.importGcell.Location = New System.Drawing.Point(294, 28)
        Me.importGcell.Name = "importGcell"
        Me.importGcell.Size = New System.Drawing.Size(82, 23)
        Me.importGcell.TabIndex = 2
        Me.importGcell.Text = "导入GCELL"
        Me.importGcell.UseVisualStyleBackColor = True
        '
        'importGTRX
        '
        Me.importGTRX.Location = New System.Drawing.Point(294, 94)
        Me.importGTRX.Name = "importGTRX"
        Me.importGTRX.Size = New System.Drawing.Size(82, 23)
        Me.importGTRX.TabIndex = 3
        Me.importGTRX.Text = "导入GTRX"
        Me.importGTRX.UseVisualStyleBackColor = True
        '
        'startParse
        '
        Me.startParse.Location = New System.Drawing.Point(150, 174)
        Me.startParse.Name = "startParse"
        Me.startParse.Size = New System.Drawing.Size(75, 23)
        Me.startParse.TabIndex = 4
        Me.startParse.Text = "开始转换"
        Me.startParse.UseVisualStyleBackColor = True
        '
        'BGWForHWPara
        '
        Me.BGWForHWPara.WorkerReportsProgress = True
        Me.BGWForHWPara.WorkerSupportsCancellation = True
        '
        'HW_CDD_PARSE
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(388, 209)
        Me.Controls.Add(Me.startParse)
        Me.Controls.Add(Me.importGTRX)
        Me.Controls.Add(Me.importGcell)
        Me.Controls.Add(Me.GtrxFilePath)
        Me.Controls.Add(Me.GcellFilePath)
        Me.MaximumSize = New System.Drawing.Size(404, 247)
        Me.MinimumSize = New System.Drawing.Size(404, 247)
        Me.Name = "HW_CDD_PARSE"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "华为6900配置转换"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GcellFilePath As System.Windows.Forms.TextBox
    Friend WithEvents GtrxFilePath As System.Windows.Forms.TextBox
    Friend WithEvents importGcell As System.Windows.Forms.Button
    Friend WithEvents importGTRX As System.Windows.Forms.Button
    Friend WithEvents startParse As System.Windows.Forms.Button
    Protected Friend WithEvents BGWForHWPara As System.ComponentModel.BackgroundWorker
End Class

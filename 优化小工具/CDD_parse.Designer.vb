<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CDD_parse
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
        Me.components = New System.ComponentModel.Container()
        Me.start = New System.Windows.Forms.Button()
        Me.generateCarrier = New System.Windows.Forms.CheckBox()
        Me.generateCDUtype = New System.Windows.Forms.CheckBox()
        Me.generateCellfile = New System.Windows.Forms.CheckBox()
        Me.BGWParseCDD = New System.ComponentModel.BackgroundWorker()
        Me.TimerForCddParse = New System.Windows.Forms.Timer(Me.components)
        Me.WhatCanParse = New System.Windows.Forms.LinkLabel()
        Me.LoadFaultName = New System.Windows.Forms.CheckBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.generateSeeSiteFile = New System.Windows.Forms.CheckBox()
        Me.generateCellConfig = New System.Windows.Forms.CheckBox()
        Me.CancleAll = New System.Windows.Forms.Button()
        Me.SelectAll = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.useFileName = New System.Windows.Forms.RadioButton()
        Me.useCONNECTED = New System.Windows.Forms.RadioButton()
        Me.useIOEXP = New System.Windows.Forms.RadioButton()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'start
        '
        Me.start.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.start.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.start.Location = New System.Drawing.Point(260, 227)
        Me.start.Name = "start"
        Me.start.Size = New System.Drawing.Size(113, 30)
        Me.start.TabIndex = 0
        Me.start.Text = "Start!"
        Me.start.UseVisualStyleBackColor = False
        '
        'generateCarrier
        '
        Me.generateCarrier.AutoSize = True
        Me.generateCarrier.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.generateCarrier.Location = New System.Drawing.Point(11, 60)
        Me.generateCarrier.Name = "generateCarrier"
        Me.generateCarrier.Size = New System.Drawing.Size(204, 20)
        Me.generateCarrier.TabIndex = 1
        Me.generateCarrier.Text = "生成 McomCarrier 和 Neighbour"
        Me.generateCarrier.UseCompatibleTextRendering = True
        Me.generateCarrier.UseVisualStyleBackColor = True
        '
        'generateCDUtype
        '
        Me.generateCDUtype.AutoSize = True
        Me.generateCDUtype.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.generateCDUtype.Location = New System.Drawing.Point(11, 109)
        Me.generateCDUtype.Name = "generateCDUtype"
        Me.generateCDUtype.Size = New System.Drawing.Size(97, 16)
        Me.generateCDUtype.TabIndex = 2
        Me.generateCDUtype.Text = "CDU类型检查"
        Me.generateCDUtype.UseVisualStyleBackColor = True
        '
        'generateCellfile
        '
        Me.generateCellfile.AutoSize = True
        Me.generateCellfile.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.generateCellfile.Location = New System.Drawing.Point(11, 150)
        Me.generateCellfile.Name = "generateCellfile"
        Me.generateCellfile.Size = New System.Drawing.Size(158, 16)
        Me.generateCellfile.TabIndex = 3
        Me.generateCellfile.Text = "生成 .CELL文件（DT）"
        Me.generateCellfile.UseVisualStyleBackColor = True
        '
        'BGWParseCDD
        '
        Me.BGWParseCDD.WorkerReportsProgress = True
        Me.BGWParseCDD.WorkerSupportsCancellation = True
        '
        'TimerForCddParse
        '
        Me.TimerForCddParse.Enabled = True
        Me.TimerForCddParse.Interval = 1000
        '
        'WhatCanParse
        '
        Me.WhatCanParse.AutoSize = True
        Me.WhatCanParse.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.WhatCanParse.Font = New System.Drawing.Font("宋体", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.WhatCanParse.Location = New System.Drawing.Point(115, 26)
        Me.WhatCanParse.Name = "WhatCanParse"
        Me.WhatCanParse.Size = New System.Drawing.Size(223, 13)
        Me.WhatCanParse.TabIndex = 5
        Me.WhatCanParse.TabStop = True
        Me.WhatCanParse.Text = "PS:查看目前支持的CDD转换指令！"
        '
        'LoadFaultName
        '
        Me.LoadFaultName.AutoSize = True
        Me.LoadFaultName.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LoadFaultName.Location = New System.Drawing.Point(240, 109)
        Me.LoadFaultName.Name = "LoadFaultName"
        Me.LoadFaultName.Size = New System.Drawing.Size(137, 16)
        Me.LoadFaultName.TabIndex = 6
        Me.LoadFaultName.Text = "加载RXMFP告警分析"
        Me.LoadFaultName.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.GroupBox1.Controls.Add(Me.generateSeeSiteFile)
        Me.GroupBox1.Controls.Add(Me.generateCellConfig)
        Me.GroupBox1.Controls.Add(Me.WhatCanParse)
        Me.GroupBox1.Controls.Add(Me.CancleAll)
        Me.GroupBox1.Controls.Add(Me.SelectAll)
        Me.GroupBox1.Controls.Add(Me.generateCarrier)
        Me.GroupBox1.Controls.Add(Me.LoadFaultName)
        Me.GroupBox1.Controls.Add(Me.generateCDUtype)
        Me.GroupBox1.Controls.Add(Me.generateCellfile)
        Me.GroupBox1.Font = New System.Drawing.Font("宋体", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(5, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(417, 209)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "转换CDD的同时，是否进行以下工作："
        '
        'generateSeeSiteFile
        '
        Me.generateSeeSiteFile.AutoSize = True
        Me.generateSeeSiteFile.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.generateSeeSiteFile.Location = New System.Drawing.Point(240, 60)
        Me.generateSeeSiteFile.Name = "generateSeeSiteFile"
        Me.generateSeeSiteFile.Size = New System.Drawing.Size(141, 20)
        Me.generateSeeSiteFile.TabIndex = 10
        Me.generateSeeSiteFile.Text = "生成SeeSite工参文件"
        Me.generateSeeSiteFile.UseCompatibleTextRendering = True
        Me.generateSeeSiteFile.UseVisualStyleBackColor = True
        '
        'generateCellConfig
        '
        Me.generateCellConfig.AutoSize = True
        Me.generateCellConfig.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.generateCellConfig.ForeColor = System.Drawing.Color.Black
        Me.generateCellConfig.Location = New System.Drawing.Point(240, 150)
        Me.generateCellConfig.Name = "generateCellConfig"
        Me.generateCellConfig.Size = New System.Drawing.Size(154, 16)
        Me.generateCellConfig.TabIndex = 9
        Me.generateCellConfig.Text = "生成小区基础配置信息" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.generateCellConfig.UseVisualStyleBackColor = True
        '
        'CancleAll
        '
        Me.CancleAll.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.CancleAll.Location = New System.Drawing.Point(63, 21)
        Me.CancleAll.Name = "CancleAll"
        Me.CancleAll.Size = New System.Drawing.Size(46, 23)
        Me.CancleAll.TabIndex = 8
        Me.CancleAll.Text = "反选"
        Me.CancleAll.UseVisualStyleBackColor = False
        '
        'SelectAll
        '
        Me.SelectAll.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.SelectAll.Location = New System.Drawing.Point(11, 21)
        Me.SelectAll.Name = "SelectAll"
        Me.SelectAll.Size = New System.Drawing.Size(46, 23)
        Me.SelectAll.TabIndex = 7
        Me.SelectAll.Text = "全选"
        Me.SelectAll.UseVisualStyleBackColor = False
        '
        'GroupBox2
        '
        Me.GroupBox2.AutoSize = True
        Me.GroupBox2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.GroupBox2.Controls.Add(Me.useFileName)
        Me.GroupBox2.Controls.Add(Me.useCONNECTED)
        Me.GroupBox2.Controls.Add(Me.useIOEXP)
        Me.GroupBox2.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(428, 12)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(194, 130)
        Me.GroupBox2.TabIndex = 8
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "BSC名称使用设置"
        '
        'useFileName
        '
        Me.useFileName.AutoSize = True
        Me.useFileName.Location = New System.Drawing.Point(7, 94)
        Me.useFileName.Name = "useFileName"
        Me.useFileName.Size = New System.Drawing.Size(122, 16)
        Me.useFileName.TabIndex = 2
        Me.useFileName.TabStop = True
        Me.useFileName.Text = "以log文件名为准"
        Me.useFileName.UseVisualStyleBackColor = True
        '
        'useCONNECTED
        '
        Me.useCONNECTED.AutoSize = True
        Me.useCONNECTED.Location = New System.Drawing.Point(7, 60)
        Me.useCONNECTED.Name = "useCONNECTED"
        Me.useCONNECTED.Size = New System.Drawing.Size(181, 16)
        Me.useCONNECTED.TabIndex = 1
        Me.useCONNECTED.TabStop = True
        Me.useCONNECTED.Text = "以*** CONNECTED TO 为准"
        Me.useCONNECTED.UseVisualStyleBackColor = True
        '
        'useIOEXP
        '
        Me.useIOEXP.AutoSize = True
        Me.useIOEXP.Checked = True
        Me.useIOEXP.Location = New System.Drawing.Point(7, 26)
        Me.useIOEXP.Name = "useIOEXP"
        Me.useIOEXP.Size = New System.Drawing.Size(97, 16)
        Me.useIOEXP.TabIndex = 0
        Me.useIOEXP.TabStop = True
        Me.useIOEXP.Text = "以IOEXP为准"
        Me.useIOEXP.UseVisualStyleBackColor = True
        '
        'CDD_parse
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(623, 265)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.start)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "CDD_parse"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CDD_parse"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents start As System.Windows.Forms.Button
    Friend WithEvents generateCarrier As System.Windows.Forms.CheckBox
    Friend WithEvents generateCDUtype As System.Windows.Forms.CheckBox
    Friend WithEvents generateCellfile As System.Windows.Forms.CheckBox
    Friend WithEvents BGWParseCDD As System.ComponentModel.BackgroundWorker
    Friend WithEvents TimerForCddParse As System.Windows.Forms.Timer
    Friend WithEvents WhatCanParse As System.Windows.Forms.LinkLabel
    Friend WithEvents LoadFaultName As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CancleAll As System.Windows.Forms.Button
    Friend WithEvents SelectAll As System.Windows.Forms.Button
    Friend WithEvents generateCellConfig As System.Windows.Forms.CheckBox
    Friend WithEvents generateSeeSiteFile As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents useFileName As System.Windows.Forms.RadioButton
    Friend WithEvents useCONNECTED As System.Windows.Forms.RadioButton
    Friend WithEvents useIOEXP As System.Windows.Forms.RadioButton
End Class

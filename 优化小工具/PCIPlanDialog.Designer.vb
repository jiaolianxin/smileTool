<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PCIPlanDialog
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
        Me.BSICPlanStatus = New System.Windows.Forms.ToolStripStatusLabel()
        Me.BSICPlanProgressBar = New System.Windows.Forms.ProgressBar()
        Me.StartPlan = New System.Windows.Forms.Button()
        Me.selectPlanCellButton = New System.Windows.Forms.Button()
        Me.planCellFilePath = New System.Windows.Forms.TextBox()
        Me.selectCarrierButton = New System.Windows.Forms.Button()
        Me.selectSiteFileButton = New System.Windows.Forms.Button()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.carrierFilePath = New System.Windows.Forms.TextBox()
        Me.importFileGroup = New System.Windows.Forms.GroupBox()
        Me.siteFilePath = New System.Windows.Forms.TextBox()
        Me.maxDistanse = New System.Windows.Forms.NumericUpDown()
        Me.disLimitLabel = New System.Windows.Forms.Label()
        Me.canNotUseBSIC = New System.Windows.Forms.TextBox()
        Me.canNotUsedLabel = New System.Windows.Forms.Label()
        Me.LimitConditions = New System.Windows.Forms.GroupBox()
        Me.PCIoptions = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.MinPCIgroup = New System.Windows.Forms.NumericUpDown()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.BGW_BSICPlan = New System.ComponentModel.BackgroundWorker()
        Me.BSICPlanTimer = New System.Windows.Forms.Timer(Me.components)
        Me.MaxPCIgroup = New System.Windows.Forms.NumericUpDown()
        Me.StatusStrip1.SuspendLayout()
        Me.importFileGroup.SuspendLayout()
        CType(Me.maxDistanse, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.LimitConditions.SuspendLayout()
        Me.PCIoptions.SuspendLayout()
        CType(Me.MinPCIgroup, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MaxPCIgroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BSICPlanStatus
        '
        Me.BSICPlanStatus.BackColor = System.Drawing.Color.WhiteSmoke
        Me.BSICPlanStatus.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BSICPlanStatus.ForeColor = System.Drawing.Color.Black
        Me.BSICPlanStatus.Name = "BSICPlanStatus"
        Me.BSICPlanStatus.Size = New System.Drawing.Size(70, 17)
        Me.BSICPlanStatus.Text = "欢迎使用！"
        '
        'BSICPlanProgressBar
        '
        Me.BSICPlanProgressBar.Location = New System.Drawing.Point(0, 281)
        Me.BSICPlanProgressBar.Name = "BSICPlanProgressBar"
        Me.BSICPlanProgressBar.Size = New System.Drawing.Size(357, 20)
        Me.BSICPlanProgressBar.TabIndex = 14
        '
        'StartPlan
        '
        Me.StartPlan.BackColor = System.Drawing.Color.Yellow
        Me.StartPlan.Font = New System.Drawing.Font("宋体", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.StartPlan.Location = New System.Drawing.Point(317, 79)
        Me.StartPlan.Name = "StartPlan"
        Me.StartPlan.Size = New System.Drawing.Size(28, 124)
        Me.StartPlan.TabIndex = 11
        Me.StartPlan.Text = "开始规划"
        Me.StartPlan.UseVisualStyleBackColor = False
        '
        'selectPlanCellButton
        '
        Me.selectPlanCellButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.selectPlanCellButton.Font = New System.Drawing.Font("宋体", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.selectPlanCellButton.Location = New System.Drawing.Point(210, 86)
        Me.selectPlanCellButton.Name = "selectPlanCellButton"
        Me.selectPlanCellButton.Size = New System.Drawing.Size(82, 23)
        Me.selectPlanCellButton.TabIndex = 5
        Me.selectPlanCellButton.Text = "导入规划列表"
        Me.selectPlanCellButton.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.selectPlanCellButton.UseVisualStyleBackColor = False
        '
        'planCellFilePath
        '
        Me.planCellFilePath.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.planCellFilePath.Location = New System.Drawing.Point(6, 85)
        Me.planCellFilePath.MaximumSize = New System.Drawing.Size(198, 21)
        Me.planCellFilePath.MinimumSize = New System.Drawing.Size(198, 21)
        Me.planCellFilePath.Name = "planCellFilePath"
        Me.planCellFilePath.ReadOnly = True
        Me.planCellFilePath.Size = New System.Drawing.Size(198, 21)
        Me.planCellFilePath.TabIndex = 1
        Me.planCellFilePath.TabStop = False
        Me.planCellFilePath.Text = "请选择PCI规划小区列表文件!"
        '
        'selectCarrierButton
        '
        Me.selectCarrierButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.selectCarrierButton.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.selectCarrierButton.Location = New System.Drawing.Point(210, 55)
        Me.selectCarrierButton.Name = "selectCarrierButton"
        Me.selectCarrierButton.Size = New System.Drawing.Size(82, 23)
        Me.selectCarrierButton.TabIndex = 3
        Me.selectCarrierButton.Text = "导入Carrier"
        Me.selectCarrierButton.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.selectCarrierButton.UseVisualStyleBackColor = False
        '
        'selectSiteFileButton
        '
        Me.selectSiteFileButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.selectSiteFileButton.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.selectSiteFileButton.Location = New System.Drawing.Point(210, 22)
        Me.selectSiteFileButton.Name = "selectSiteFileButton"
        Me.selectSiteFileButton.Size = New System.Drawing.Size(82, 23)
        Me.selectSiteFileButton.TabIndex = 2
        Me.selectSiteFileButton.Text = "导入Site"
        Me.selectSiteFileButton.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.selectSiteFileButton.UseVisualStyleBackColor = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BSICPlanStatus})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 304)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(357, 22)
        Me.StatusStrip1.TabIndex = 15
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'carrierFilePath
        '
        Me.carrierFilePath.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.carrierFilePath.Location = New System.Drawing.Point(6, 54)
        Me.carrierFilePath.MaximumSize = New System.Drawing.Size(198, 21)
        Me.carrierFilePath.MinimumSize = New System.Drawing.Size(198, 21)
        Me.carrierFilePath.Name = "carrierFilePath"
        Me.carrierFilePath.ReadOnly = True
        Me.carrierFilePath.Size = New System.Drawing.Size(198, 21)
        Me.carrierFilePath.TabIndex = 1
        Me.carrierFilePath.Text = "请选择Mcom Carrier文件!"
        '
        'importFileGroup
        '
        Me.importFileGroup.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.importFileGroup.Controls.Add(Me.selectPlanCellButton)
        Me.importFileGroup.Controls.Add(Me.planCellFilePath)
        Me.importFileGroup.Controls.Add(Me.selectCarrierButton)
        Me.importFileGroup.Controls.Add(Me.selectSiteFileButton)
        Me.importFileGroup.Controls.Add(Me.carrierFilePath)
        Me.importFileGroup.Controls.Add(Me.siteFilePath)
        Me.importFileGroup.Font = New System.Drawing.Font("宋体", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.importFileGroup.ForeColor = System.Drawing.Color.Black
        Me.importFileGroup.Location = New System.Drawing.Point(7, 159)
        Me.importFileGroup.Name = "importFileGroup"
        Me.importFileGroup.Size = New System.Drawing.Size(298, 116)
        Me.importFileGroup.TabIndex = 13
        Me.importFileGroup.TabStop = False
        Me.importFileGroup.Text = "输入文件"
        '
        'siteFilePath
        '
        Me.siteFilePath.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.siteFilePath.Location = New System.Drawing.Point(6, 23)
        Me.siteFilePath.MaximumSize = New System.Drawing.Size(198, 21)
        Me.siteFilePath.MinimumSize = New System.Drawing.Size(198, 21)
        Me.siteFilePath.Name = "siteFilePath"
        Me.siteFilePath.ReadOnly = True
        Me.siteFilePath.Size = New System.Drawing.Size(198, 21)
        Me.siteFilePath.TabIndex = 0
        Me.siteFilePath.Text = "请选择Mcom Site文件!"
        '
        'maxDistanse
        '
        Me.maxDistanse.Location = New System.Drawing.Point(221, 66)
        Me.maxDistanse.Name = "maxDistanse"
        Me.maxDistanse.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.maxDistanse.Size = New System.Drawing.Size(62, 22)
        Me.maxDistanse.TabIndex = 4
        Me.maxDistanse.Value = New Decimal(New Integer() {20, 0, 0, 0})
        '
        'disLimitLabel
        '
        Me.disLimitLabel.AutoSize = True
        Me.disLimitLabel.Font = New System.Drawing.Font("宋体", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.disLimitLabel.ForeColor = System.Drawing.Color.Blue
        Me.disLimitLabel.Location = New System.Drawing.Point(9, 68)
        Me.disLimitLabel.Name = "disLimitLabel"
        Me.disLimitLabel.Size = New System.Drawing.Size(213, 13)
        Me.disLimitLabel.TabIndex = 2
        Me.disLimitLabel.Text = "PCIgroup规划限制距离（KM）："
        '
        'canNotUseBSIC
        '
        Me.canNotUseBSIC.Location = New System.Drawing.Point(9, 41)
        Me.canNotUseBSIC.MaximumSize = New System.Drawing.Size(284, 21)
        Me.canNotUseBSIC.MinimumSize = New System.Drawing.Size(284, 21)
        Me.canNotUseBSIC.Multiline = True
        Me.canNotUseBSIC.Name = "canNotUseBSIC"
        Me.canNotUseBSIC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.canNotUseBSIC.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.canNotUseBSIC.Size = New System.Drawing.Size(284, 21)
        Me.canNotUseBSIC.TabIndex = 1
        Me.canNotUseBSIC.TabStop = False
        '
        'canNotUsedLabel
        '
        Me.canNotUsedLabel.AutoSize = True
        Me.canNotUsedLabel.Font = New System.Drawing.Font("宋体", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.canNotUsedLabel.ForeColor = System.Drawing.Color.Blue
        Me.canNotUsedLabel.Location = New System.Drawing.Point(6, 25)
        Me.canNotUsedLabel.Name = "canNotUsedLabel"
        Me.canNotUsedLabel.Size = New System.Drawing.Size(263, 13)
        Me.canNotUsedLabel.TabIndex = 0
        Me.canNotUsedLabel.Text = "不允许使用的PCIgroup（用"",""分隔）："
        '
        'LimitConditions
        '
        Me.LimitConditions.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.LimitConditions.Controls.Add(Me.maxDistanse)
        Me.LimitConditions.Controls.Add(Me.disLimitLabel)
        Me.LimitConditions.Controls.Add(Me.canNotUseBSIC)
        Me.LimitConditions.Controls.Add(Me.canNotUsedLabel)
        Me.LimitConditions.Font = New System.Drawing.Font("宋体", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LimitConditions.ForeColor = System.Drawing.Color.Black
        Me.LimitConditions.Location = New System.Drawing.Point(7, 54)
        Me.LimitConditions.Name = "LimitConditions"
        Me.LimitConditions.Size = New System.Drawing.Size(299, 99)
        Me.LimitConditions.TabIndex = 12
        Me.LimitConditions.TabStop = False
        Me.LimitConditions.Text = "限制条件"
        '
        'PCIoptions
        '
        Me.PCIoptions.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.PCIoptions.Controls.Add(Me.MaxPCIgroup)
        Me.PCIoptions.Controls.Add(Me.Label2)
        Me.PCIoptions.Controls.Add(Me.MinPCIgroup)
        Me.PCIoptions.Controls.Add(Me.Label1)
        Me.PCIoptions.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PCIoptions.Location = New System.Drawing.Point(7, 1)
        Me.PCIoptions.Name = "PCIoptions"
        Me.PCIoptions.Size = New System.Drawing.Size(299, 47)
        Me.PCIoptions.TabIndex = 10
        Me.PCIoptions.TabStop = False
        Me.PCIoptions.Text = "可用PCIgroup"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("宋体", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.Location = New System.Drawing.Point(135, 20)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(39, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "To :"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'MinPCIgroup
        '
        Me.MinPCIgroup.Location = New System.Drawing.Point(73, 19)
        Me.MinPCIgroup.Name = "MinPCIgroup"
        Me.MinPCIgroup.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.MinPCIgroup.Size = New System.Drawing.Size(53, 21)
        Me.MinPCIgroup.TabIndex = 5
        Me.MinPCIgroup.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("宋体", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(55, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "From :"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'BGW_BSICPlan
        '
        Me.BGW_BSICPlan.WorkerReportsProgress = True
        Me.BGW_BSICPlan.WorkerSupportsCancellation = True
        '
        'BSICPlanTimer
        '
        Me.BSICPlanTimer.Enabled = True
        Me.BSICPlanTimer.Interval = 1000
        '
        'MaxPCIgroup
        '
        Me.MaxPCIgroup.Location = New System.Drawing.Point(180, 19)
        Me.MaxPCIgroup.Maximum = New Decimal(New Integer() {167, 0, 0, 0})
        Me.MaxPCIgroup.Name = "MaxPCIgroup"
        Me.MaxPCIgroup.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.MaxPCIgroup.Size = New System.Drawing.Size(53, 21)
        Me.MaxPCIgroup.TabIndex = 7
        Me.MaxPCIgroup.Value = New Decimal(New Integer() {167, 0, 0, 0})
        '
        'PCIPlanDialog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(357, 326)
        Me.Controls.Add(Me.BSICPlanProgressBar)
        Me.Controls.Add(Me.StartPlan)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.importFileGroup)
        Me.Controls.Add(Me.LimitConditions)
        Me.Controls.Add(Me.PCIoptions)
        Me.MaximumSize = New System.Drawing.Size(373, 364)
        Me.MinimumSize = New System.Drawing.Size(373, 364)
        Me.Name = "PCIPlanDialog"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "PCI规划小工具"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.importFileGroup.ResumeLayout(False)
        Me.importFileGroup.PerformLayout()
        CType(Me.maxDistanse, System.ComponentModel.ISupportInitialize).EndInit()
        Me.LimitConditions.ResumeLayout(False)
        Me.LimitConditions.PerformLayout()
        Me.PCIoptions.ResumeLayout(False)
        Me.PCIoptions.PerformLayout()
        CType(Me.MinPCIgroup, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MaxPCIgroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BSICPlanStatus As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents BSICPlanProgressBar As System.Windows.Forms.ProgressBar
    Friend WithEvents StartPlan As System.Windows.Forms.Button
    Friend WithEvents selectPlanCellButton As System.Windows.Forms.Button
    Friend WithEvents planCellFilePath As System.Windows.Forms.TextBox
    Friend WithEvents selectCarrierButton As System.Windows.Forms.Button
    Friend WithEvents selectSiteFileButton As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents carrierFilePath As System.Windows.Forms.TextBox
    Friend WithEvents importFileGroup As System.Windows.Forms.GroupBox
    Friend WithEvents siteFilePath As System.Windows.Forms.TextBox
    Friend WithEvents maxDistanse As System.Windows.Forms.NumericUpDown
    Friend WithEvents disLimitLabel As System.Windows.Forms.Label
    Friend WithEvents canNotUseBSIC As System.Windows.Forms.TextBox
    Friend WithEvents canNotUsedLabel As System.Windows.Forms.Label
    Friend WithEvents LimitConditions As System.Windows.Forms.GroupBox
    Friend WithEvents PCIoptions As System.Windows.Forms.GroupBox
    Friend WithEvents BGW_BSICPlan As System.ComponentModel.BackgroundWorker
    Friend WithEvents BSICPlanTimer As System.Windows.Forms.Timer
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents MinPCIgroup As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents MaxPCIgroup As System.Windows.Forms.NumericUpDown
End Class

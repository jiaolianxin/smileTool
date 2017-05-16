<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BSICPlanDialog
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
        Me.NCC = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.BSICoptions = New System.Windows.Forms.GroupBox()
        Me.BCC5 = New System.Windows.Forms.CheckBox()
        Me.NCC5 = New System.Windows.Forms.CheckBox()
        Me.BCC4 = New System.Windows.Forms.CheckBox()
        Me.NCC7 = New System.Windows.Forms.CheckBox()
        Me.BCC3 = New System.Windows.Forms.CheckBox()
        Me.NCC4 = New System.Windows.Forms.CheckBox()
        Me.BCC2 = New System.Windows.Forms.CheckBox()
        Me.NCC6 = New System.Windows.Forms.CheckBox()
        Me.BCC1 = New System.Windows.Forms.CheckBox()
        Me.NCC3 = New System.Windows.Forms.CheckBox()
        Me.BCC0 = New System.Windows.Forms.CheckBox()
        Me.BCC7 = New System.Windows.Forms.CheckBox()
        Me.NCC2 = New System.Windows.Forms.CheckBox()
        Me.BCC6 = New System.Windows.Forms.CheckBox()
        Me.NCC1 = New System.Windows.Forms.CheckBox()
        Me.NCC0 = New System.Windows.Forms.CheckBox()
        Me.StartPlan = New System.Windows.Forms.Button()
        Me.LimitConditions = New System.Windows.Forms.GroupBox()
        Me.maxDistanse = New System.Windows.Forms.NumericUpDown()
        Me.disLimitLabel = New System.Windows.Forms.Label()
        Me.canNotUseBSIC = New System.Windows.Forms.TextBox()
        Me.canNotUsedLabel = New System.Windows.Forms.Label()
        Me.importFileGroup = New System.Windows.Forms.GroupBox()
        Me.selectPlanCellButton = New System.Windows.Forms.Button()
        Me.planCellFilePath = New System.Windows.Forms.TextBox()
        Me.selectCarrierButton = New System.Windows.Forms.Button()
        Me.selectSiteFileButton = New System.Windows.Forms.Button()
        Me.carrierFilePath = New System.Windows.Forms.TextBox()
        Me.siteFilePath = New System.Windows.Forms.TextBox()
        Me.BSICPlanProgressBar = New System.Windows.Forms.ProgressBar()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.BSICPlanStatus = New System.Windows.Forms.ToolStripStatusLabel()
        Me.BGW_BSICPlan = New System.ComponentModel.BackgroundWorker()
        Me.BSICPlanTimer = New System.Windows.Forms.Timer(Me.components)
        Me.BSICoptions.SuspendLayout()
        Me.LimitConditions.SuspendLayout()
        CType(Me.maxDistanse, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.importFileGroup.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'NCC
        '
        Me.NCC.AutoSize = True
        Me.NCC.Font = New System.Drawing.Font("宋体", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.NCC.Location = New System.Drawing.Point(9, 17)
        Me.NCC.Name = "NCC"
        Me.NCC.Size = New System.Drawing.Size(39, 13)
        Me.NCC.TabIndex = 2
        Me.NCC.Text = "NCC:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("宋体", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(56, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "BCC:"
        '
        'BSICoptions
        '
        Me.BSICoptions.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.BSICoptions.Controls.Add(Me.BCC5)
        Me.BSICoptions.Controls.Add(Me.NCC5)
        Me.BSICoptions.Controls.Add(Me.BCC4)
        Me.BSICoptions.Controls.Add(Me.NCC7)
        Me.BSICoptions.Controls.Add(Me.BCC3)
        Me.BSICoptions.Controls.Add(Me.NCC4)
        Me.BSICoptions.Controls.Add(Me.BCC2)
        Me.BSICoptions.Controls.Add(Me.NCC6)
        Me.BSICoptions.Controls.Add(Me.BCC1)
        Me.BSICoptions.Controls.Add(Me.NCC3)
        Me.BSICoptions.Controls.Add(Me.BCC0)
        Me.BSICoptions.Controls.Add(Me.BCC7)
        Me.BSICoptions.Controls.Add(Me.NCC2)
        Me.BSICoptions.Controls.Add(Me.BCC6)
        Me.BSICoptions.Controls.Add(Me.NCC1)
        Me.BSICoptions.Controls.Add(Me.NCC0)
        Me.BSICoptions.Controls.Add(Me.Label1)
        Me.BSICoptions.Controls.Add(Me.NCC)
        Me.BSICoptions.Font = New System.Drawing.Font("宋体", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BSICoptions.Location = New System.Drawing.Point(12, 12)
        Me.BSICoptions.Name = "BSICoptions"
        Me.BSICoptions.Size = New System.Drawing.Size(105, 227)
        Me.BSICoptions.TabIndex = 4
        Me.BSICoptions.TabStop = False
        Me.BSICoptions.Text = "可用BSIC"
        '
        'BCC5
        '
        Me.BCC5.AutoSize = True
        Me.BCC5.Checked = True
        Me.BCC5.CheckState = System.Windows.Forms.CheckState.Checked
        Me.BCC5.Location = New System.Drawing.Point(59, 152)
        Me.BCC5.Name = "BCC5"
        Me.BCC5.Size = New System.Drawing.Size(35, 19)
        Me.BCC5.TabIndex = 17
        Me.BCC5.Text = "5"
        Me.BCC5.UseVisualStyleBackColor = True
        '
        'NCC5
        '
        Me.NCC5.AutoSize = True
        Me.NCC5.Checked = True
        Me.NCC5.CheckState = System.Windows.Forms.CheckState.Checked
        Me.NCC5.Location = New System.Drawing.Point(12, 152)
        Me.NCC5.Name = "NCC5"
        Me.NCC5.Size = New System.Drawing.Size(35, 19)
        Me.NCC5.TabIndex = 9
        Me.NCC5.Text = "5"
        Me.NCC5.UseVisualStyleBackColor = True
        '
        'BCC4
        '
        Me.BCC4.AutoSize = True
        Me.BCC4.Checked = True
        Me.BCC4.CheckState = System.Windows.Forms.CheckState.Checked
        Me.BCC4.Location = New System.Drawing.Point(59, 130)
        Me.BCC4.Name = "BCC4"
        Me.BCC4.Size = New System.Drawing.Size(35, 19)
        Me.BCC4.TabIndex = 16
        Me.BCC4.Text = "4"
        Me.BCC4.UseVisualStyleBackColor = True
        '
        'NCC7
        '
        Me.NCC7.AutoSize = True
        Me.NCC7.Checked = True
        Me.NCC7.CheckState = System.Windows.Forms.CheckState.Checked
        Me.NCC7.Location = New System.Drawing.Point(12, 196)
        Me.NCC7.Name = "NCC7"
        Me.NCC7.Size = New System.Drawing.Size(35, 19)
        Me.NCC7.TabIndex = 11
        Me.NCC7.Text = "7"
        Me.NCC7.UseVisualStyleBackColor = True
        '
        'BCC3
        '
        Me.BCC3.AutoSize = True
        Me.BCC3.Checked = True
        Me.BCC3.CheckState = System.Windows.Forms.CheckState.Checked
        Me.BCC3.Location = New System.Drawing.Point(59, 108)
        Me.BCC3.Name = "BCC3"
        Me.BCC3.Size = New System.Drawing.Size(35, 19)
        Me.BCC3.TabIndex = 15
        Me.BCC3.Text = "3"
        Me.BCC3.UseVisualStyleBackColor = True
        '
        'NCC4
        '
        Me.NCC4.AutoSize = True
        Me.NCC4.Checked = True
        Me.NCC4.CheckState = System.Windows.Forms.CheckState.Checked
        Me.NCC4.Location = New System.Drawing.Point(12, 130)
        Me.NCC4.Name = "NCC4"
        Me.NCC4.Size = New System.Drawing.Size(35, 19)
        Me.NCC4.TabIndex = 8
        Me.NCC4.Text = "4"
        Me.NCC4.UseVisualStyleBackColor = True
        '
        'BCC2
        '
        Me.BCC2.AutoSize = True
        Me.BCC2.Checked = True
        Me.BCC2.CheckState = System.Windows.Forms.CheckState.Checked
        Me.BCC2.Location = New System.Drawing.Point(59, 86)
        Me.BCC2.Name = "BCC2"
        Me.BCC2.Size = New System.Drawing.Size(35, 19)
        Me.BCC2.TabIndex = 14
        Me.BCC2.Text = "2"
        Me.BCC2.UseVisualStyleBackColor = True
        '
        'NCC6
        '
        Me.NCC6.AutoSize = True
        Me.NCC6.Checked = True
        Me.NCC6.CheckState = System.Windows.Forms.CheckState.Checked
        Me.NCC6.Location = New System.Drawing.Point(12, 174)
        Me.NCC6.Name = "NCC6"
        Me.NCC6.Size = New System.Drawing.Size(35, 19)
        Me.NCC6.TabIndex = 10
        Me.NCC6.Text = "6"
        Me.NCC6.UseVisualStyleBackColor = True
        '
        'BCC1
        '
        Me.BCC1.AutoSize = True
        Me.BCC1.Checked = True
        Me.BCC1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.BCC1.Location = New System.Drawing.Point(59, 64)
        Me.BCC1.Name = "BCC1"
        Me.BCC1.Size = New System.Drawing.Size(35, 19)
        Me.BCC1.TabIndex = 13
        Me.BCC1.Text = "1"
        Me.BCC1.UseVisualStyleBackColor = True
        '
        'NCC3
        '
        Me.NCC3.AutoSize = True
        Me.NCC3.Checked = True
        Me.NCC3.CheckState = System.Windows.Forms.CheckState.Checked
        Me.NCC3.Location = New System.Drawing.Point(12, 108)
        Me.NCC3.Name = "NCC3"
        Me.NCC3.Size = New System.Drawing.Size(35, 19)
        Me.NCC3.TabIndex = 7
        Me.NCC3.Text = "3"
        Me.NCC3.UseVisualStyleBackColor = True
        '
        'BCC0
        '
        Me.BCC0.AutoSize = True
        Me.BCC0.Checked = True
        Me.BCC0.CheckState = System.Windows.Forms.CheckState.Checked
        Me.BCC0.Location = New System.Drawing.Point(59, 42)
        Me.BCC0.Name = "BCC0"
        Me.BCC0.Size = New System.Drawing.Size(35, 19)
        Me.BCC0.TabIndex = 12
        Me.BCC0.Text = "0"
        Me.BCC0.UseVisualStyleBackColor = True
        '
        'BCC7
        '
        Me.BCC7.AutoSize = True
        Me.BCC7.Checked = True
        Me.BCC7.CheckState = System.Windows.Forms.CheckState.Checked
        Me.BCC7.Location = New System.Drawing.Point(59, 196)
        Me.BCC7.Name = "BCC7"
        Me.BCC7.Size = New System.Drawing.Size(35, 19)
        Me.BCC7.TabIndex = 19
        Me.BCC7.Text = "7"
        Me.BCC7.UseVisualStyleBackColor = True
        '
        'NCC2
        '
        Me.NCC2.AutoSize = True
        Me.NCC2.Checked = True
        Me.NCC2.CheckState = System.Windows.Forms.CheckState.Checked
        Me.NCC2.Location = New System.Drawing.Point(12, 86)
        Me.NCC2.Name = "NCC2"
        Me.NCC2.Size = New System.Drawing.Size(35, 19)
        Me.NCC2.TabIndex = 6
        Me.NCC2.Text = "2"
        Me.NCC2.UseVisualStyleBackColor = True
        '
        'BCC6
        '
        Me.BCC6.AutoSize = True
        Me.BCC6.Checked = True
        Me.BCC6.CheckState = System.Windows.Forms.CheckState.Checked
        Me.BCC6.Location = New System.Drawing.Point(59, 174)
        Me.BCC6.Name = "BCC6"
        Me.BCC6.Size = New System.Drawing.Size(35, 19)
        Me.BCC6.TabIndex = 18
        Me.BCC6.Text = "6"
        Me.BCC6.UseVisualStyleBackColor = True
        '
        'NCC1
        '
        Me.NCC1.AutoSize = True
        Me.NCC1.Checked = True
        Me.NCC1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.NCC1.Location = New System.Drawing.Point(12, 64)
        Me.NCC1.Name = "NCC1"
        Me.NCC1.Size = New System.Drawing.Size(35, 19)
        Me.NCC1.TabIndex = 5
        Me.NCC1.Text = "1"
        Me.NCC1.UseVisualStyleBackColor = True
        '
        'NCC0
        '
        Me.NCC0.AutoSize = True
        Me.NCC0.Checked = True
        Me.NCC0.CheckState = System.Windows.Forms.CheckState.Checked
        Me.NCC0.Location = New System.Drawing.Point(12, 42)
        Me.NCC0.Name = "NCC0"
        Me.NCC0.Size = New System.Drawing.Size(35, 19)
        Me.NCC0.TabIndex = 4
        Me.NCC0.Text = "0"
        Me.NCC0.UseVisualStyleBackColor = True
        '
        'StartPlan
        '
        Me.StartPlan.BackColor = System.Drawing.Color.Yellow
        Me.StartPlan.Font = New System.Drawing.Font("宋体", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.StartPlan.Location = New System.Drawing.Point(181, 245)
        Me.StartPlan.Name = "StartPlan"
        Me.StartPlan.Size = New System.Drawing.Size(75, 23)
        Me.StartPlan.TabIndex = 5
        Me.StartPlan.Text = "开始规划"
        Me.StartPlan.UseVisualStyleBackColor = False
        '
        'LimitConditions
        '
        Me.LimitConditions.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.LimitConditions.Controls.Add(Me.maxDistanse)
        Me.LimitConditions.Controls.Add(Me.disLimitLabel)
        Me.LimitConditions.Controls.Add(Me.canNotUseBSIC)
        Me.LimitConditions.Controls.Add(Me.canNotUsedLabel)
        Me.LimitConditions.Font = New System.Drawing.Font("宋体", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LimitConditions.ForeColor = System.Drawing.Color.Black
        Me.LimitConditions.Location = New System.Drawing.Point(123, 12)
        Me.LimitConditions.Name = "LimitConditions"
        Me.LimitConditions.Size = New System.Drawing.Size(299, 105)
        Me.LimitConditions.TabIndex = 6
        Me.LimitConditions.TabStop = False
        Me.LimitConditions.Text = "限制条件"
        '
        'maxDistanse
        '
        Me.maxDistanse.Location = New System.Drawing.Point(196, 68)
        Me.maxDistanse.Name = "maxDistanse"
        Me.maxDistanse.Size = New System.Drawing.Size(87, 25)
        Me.maxDistanse.TabIndex = 4
        Me.maxDistanse.Value = New Decimal(New Integer() {20, 0, 0, 0})
        '
        'disLimitLabel
        '
        Me.disLimitLabel.AutoSize = True
        Me.disLimitLabel.Font = New System.Drawing.Font("宋体", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.disLimitLabel.ForeColor = System.Drawing.Color.Blue
        Me.disLimitLabel.Location = New System.Drawing.Point(9, 74)
        Me.disLimitLabel.Name = "disLimitLabel"
        Me.disLimitLabel.Size = New System.Drawing.Size(181, 13)
        Me.disLimitLabel.TabIndex = 2
        Me.disLimitLabel.Text = "BSIC规划限制距离（KM）："
        '
        'canNotUseBSIC
        '
        Me.canNotUseBSIC.Location = New System.Drawing.Point(9, 37)
        Me.canNotUseBSIC.MaximumSize = New System.Drawing.Size(284, 25)
        Me.canNotUseBSIC.MinimumSize = New System.Drawing.Size(284, 25)
        Me.canNotUseBSIC.Multiline = True
        Me.canNotUseBSIC.Name = "canNotUseBSIC"
        Me.canNotUseBSIC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.canNotUseBSIC.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.canNotUseBSIC.Size = New System.Drawing.Size(284, 25)
        Me.canNotUseBSIC.TabIndex = 1
        Me.canNotUseBSIC.TabStop = False
        '
        'canNotUsedLabel
        '
        Me.canNotUsedLabel.AutoSize = True
        Me.canNotUsedLabel.Font = New System.Drawing.Font("宋体", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.canNotUsedLabel.ForeColor = System.Drawing.Color.Blue
        Me.canNotUsedLabel.Location = New System.Drawing.Point(6, 21)
        Me.canNotUsedLabel.Name = "canNotUsedLabel"
        Me.canNotUsedLabel.Size = New System.Drawing.Size(231, 13)
        Me.canNotUsedLabel.TabIndex = 0
        Me.canNotUsedLabel.Text = "不允许使用的BSIC（用"",""分隔）："
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
        Me.importFileGroup.Font = New System.Drawing.Font("宋体", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.importFileGroup.ForeColor = System.Drawing.Color.Black
        Me.importFileGroup.Location = New System.Drawing.Point(124, 123)
        Me.importFileGroup.Name = "importFileGroup"
        Me.importFileGroup.Size = New System.Drawing.Size(298, 116)
        Me.importFileGroup.TabIndex = 7
        Me.importFileGroup.TabStop = False
        Me.importFileGroup.Text = "输入文件"
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
        Me.planCellFilePath.Text = "请选择BSIC规划小区列表文件!"
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
        'siteFilePath
        '
        Me.siteFilePath.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.siteFilePath.Location = New System.Drawing.Point(6, 23)
        Me.siteFilePath.MaximumSize = New System.Drawing.Size(198, 25)
        Me.siteFilePath.MinimumSize = New System.Drawing.Size(198, 25)
        Me.siteFilePath.Name = "siteFilePath"
        Me.siteFilePath.ReadOnly = True
        Me.siteFilePath.Size = New System.Drawing.Size(198, 21)
        Me.siteFilePath.TabIndex = 0
        Me.siteFilePath.Text = "请选择Mcom Site文件!"
        '
        'BSICPlanProgressBar
        '
        Me.BSICPlanProgressBar.Location = New System.Drawing.Point(-2, 272)
        Me.BSICPlanProgressBar.MaximumSize = New System.Drawing.Size(436, 20)
        Me.BSICPlanProgressBar.MinimumSize = New System.Drawing.Size(436, 20)
        Me.BSICPlanProgressBar.Name = "BSICPlanProgressBar"
        Me.BSICPlanProgressBar.Size = New System.Drawing.Size(436, 20)
        Me.BSICPlanProgressBar.TabIndex = 8
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BSICPlanStatus})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 290)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(434, 22)
        Me.StatusStrip1.TabIndex = 9
        Me.StatusStrip1.Text = "StatusStrip1"
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
        'BSICPlanDialog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(434, 312)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.BSICPlanProgressBar)
        Me.Controls.Add(Me.importFileGroup)
        Me.Controls.Add(Me.LimitConditions)
        Me.Controls.Add(Me.StartPlan)
        Me.Controls.Add(Me.BSICoptions)
        Me.MaximumSize = New System.Drawing.Size(450, 350)
        Me.MinimumSize = New System.Drawing.Size(450, 350)
        Me.Name = "BSICPlanDialog"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "BSIC规划工具"
        Me.BSICoptions.ResumeLayout(False)
        Me.BSICoptions.PerformLayout()
        Me.LimitConditions.ResumeLayout(False)
        Me.LimitConditions.PerformLayout()
        CType(Me.maxDistanse, System.ComponentModel.ISupportInitialize).EndInit()
        Me.importFileGroup.ResumeLayout(False)
        Me.importFileGroup.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents NCC As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents BSICoptions As System.Windows.Forms.GroupBox
    Friend WithEvents StartPlan As System.Windows.Forms.Button
    Friend WithEvents BCC5 As System.Windows.Forms.CheckBox
    Friend WithEvents BCC4 As System.Windows.Forms.CheckBox
    Friend WithEvents BCC3 As System.Windows.Forms.CheckBox
    Friend WithEvents BCC2 As System.Windows.Forms.CheckBox
    Friend WithEvents BCC1 As System.Windows.Forms.CheckBox
    Friend WithEvents BCC0 As System.Windows.Forms.CheckBox
    Friend WithEvents BCC7 As System.Windows.Forms.CheckBox
    Friend WithEvents BCC6 As System.Windows.Forms.CheckBox
    Friend WithEvents NCC5 As System.Windows.Forms.CheckBox
    Friend WithEvents NCC4 As System.Windows.Forms.CheckBox
    Friend WithEvents NCC3 As System.Windows.Forms.CheckBox
    Friend WithEvents NCC2 As System.Windows.Forms.CheckBox
    Friend WithEvents NCC1 As System.Windows.Forms.CheckBox
    Friend WithEvents NCC0 As System.Windows.Forms.CheckBox
    Friend WithEvents NCC6 As System.Windows.Forms.CheckBox
    Friend WithEvents NCC7 As System.Windows.Forms.CheckBox
    Friend WithEvents LimitConditions As System.Windows.Forms.GroupBox
    Friend WithEvents canNotUsedLabel As System.Windows.Forms.Label
    Friend WithEvents canNotUseBSIC As System.Windows.Forms.TextBox
    Friend WithEvents disLimitLabel As System.Windows.Forms.Label
    Friend WithEvents importFileGroup As System.Windows.Forms.GroupBox
    Friend WithEvents carrierFilePath As System.Windows.Forms.TextBox
    Friend WithEvents siteFilePath As System.Windows.Forms.TextBox
    Friend WithEvents selectSiteFileButton As System.Windows.Forms.Button
    Friend WithEvents selectCarrierButton As System.Windows.Forms.Button
    Friend WithEvents BSICPlanProgressBar As System.Windows.Forms.ProgressBar
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents BSICPlanStatus As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents BGW_BSICPlan As System.ComponentModel.BackgroundWorker
    Friend WithEvents BSICPlanTimer As System.Windows.Forms.Timer
    Friend WithEvents maxDistanse As System.Windows.Forms.NumericUpDown
    Friend WithEvents selectPlanCellButton As System.Windows.Forms.Button
    Friend WithEvents planCellFilePath As System.Windows.Forms.TextBox
End Class

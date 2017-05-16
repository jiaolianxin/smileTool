<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DistanceToolDialog
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DistanceToolDialog))
        Me.BGWForDistance = New System.ComponentModel.BackgroundWorker()
        Me.distanceBar = New System.Windows.Forms.ProgressBar()
        Me.DistanceStatus = New System.Windows.Forms.StatusStrip()
        Me.statusBar = New System.Windows.Forms.ToolStripStatusLabel()
        Me.caculateSelect = New System.Windows.Forms.Button()
        Me.caculatePathLabel = New System.Windows.Forms.Label()
        Me.SiteSelect = New System.Windows.Forms.Button()
        Me.caculateSample = New System.Windows.Forms.LinkLabel()
        Me.siteSample = New System.Windows.Forms.LinkLabel()
        Me.StartForDistance = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.caculateFilePath = New System.Windows.Forms.TextBox()
        Me.siteFilePath = New System.Windows.Forms.TextBox()
        Me.OneToMulti = New System.Windows.Forms.RadioButton()
        Me.OneToOne = New System.Windows.Forms.RadioButton()
        Me.TimerForDistance = New System.Windows.Forms.Timer(Me.components)
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.ToDis = New System.Windows.Forms.NumericUpDown()
        Me.FromDis = New System.Windows.Forms.NumericUpDown()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.isAllowDuplicate = New System.Windows.Forms.CheckBox()
        Me.isShowCoSite = New System.Windows.Forms.CheckBox()
        Me.isCacuMinDis = New System.Windows.Forms.CheckBox()
        Me.DistanceStatus.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.ToDis, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FromDis, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'BGWForDistance
        '
        Me.BGWForDistance.WorkerReportsProgress = True
        Me.BGWForDistance.WorkerSupportsCancellation = True
        '
        'distanceBar
        '
        Me.distanceBar.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.distanceBar.Location = New System.Drawing.Point(-2, 214)
        Me.distanceBar.Name = "distanceBar"
        Me.distanceBar.Size = New System.Drawing.Size(436, 23)
        Me.distanceBar.TabIndex = 11
        '
        'DistanceStatus
        '
        Me.DistanceStatus.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.statusBar})
        Me.DistanceStatus.Location = New System.Drawing.Point(0, 240)
        Me.DistanceStatus.Name = "DistanceStatus"
        Me.DistanceStatus.ShowItemToolTips = True
        Me.DistanceStatus.Size = New System.Drawing.Size(434, 22)
        Me.DistanceStatus.TabIndex = 12
        Me.DistanceStatus.Text = "StatusStrip1"
        '
        'statusBar
        '
        Me.statusBar.AutoToolTip = True
        Me.statusBar.BackColor = System.Drawing.SystemColors.Control
        Me.statusBar.Name = "statusBar"
        Me.statusBar.Size = New System.Drawing.Size(67, 17)
        Me.statusBar.Tag = "statusText"
        Me.statusBar.Text = "欢迎使用！"
        '
        'caculateSelect
        '
        Me.caculateSelect.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.caculateSelect.Location = New System.Drawing.Point(322, 37)
        Me.caculateSelect.Name = "caculateSelect"
        Me.caculateSelect.Size = New System.Drawing.Size(100, 23)
        Me.caculateSelect.TabIndex = 22
        Me.caculateSelect.Tag = ""
        Me.caculateSelect.Text = "导入Caculate"
        Me.caculateSelect.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.caculateSelect.UseVisualStyleBackColor = True
        '
        'caculatePathLabel
        '
        Me.caculatePathLabel.AutoSize = True
        Me.caculatePathLabel.BackColor = System.Drawing.Color.WhiteSmoke
        Me.caculatePathLabel.Font = New System.Drawing.Font("宋体", 8.25!)
        Me.caculatePathLabel.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.caculatePathLabel.Location = New System.Drawing.Point(8, 39)
        Me.caculatePathLabel.Name = "caculatePathLabel"
        Me.caculatePathLabel.Size = New System.Drawing.Size(0, 11)
        Me.caculatePathLabel.TabIndex = 32
        '
        'SiteSelect
        '
        Me.SiteSelect.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.SiteSelect.Location = New System.Drawing.Point(322, 5)
        Me.SiteSelect.Name = "SiteSelect"
        Me.SiteSelect.Size = New System.Drawing.Size(100, 23)
        Me.SiteSelect.TabIndex = 20
        Me.SiteSelect.Text = "导入McomSite"
        Me.SiteSelect.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.SiteSelect.UseVisualStyleBackColor = True
        '
        'caculateSample
        '
        Me.caculateSample.AutoSize = True
        Me.caculateSample.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.caculateSample.Location = New System.Drawing.Point(255, 18)
        Me.caculateSample.Name = "caculateSample"
        Me.caculateSample.Size = New System.Drawing.Size(127, 13)
        Me.caculateSample.TabIndex = 30
        Me.caculateSample.TabStop = True
        Me.caculateSample.Text = "Caculate文件格式"
        '
        'siteSample
        '
        Me.siteSample.AutoSize = True
        Me.siteSample.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.siteSample.Location = New System.Drawing.Point(40, 18)
        Me.siteSample.Name = "siteSample"
        Me.siteSample.Size = New System.Drawing.Size(135, 13)
        Me.siteSample.TabIndex = 29
        Me.siteSample.TabStop = True
        Me.siteSample.Text = "Mcom Site文件格式"
        '
        'StartForDistance
        '
        Me.StartForDistance.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.StartForDistance.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.StartForDistance.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.StartForDistance.Location = New System.Drawing.Point(178, 189)
        Me.StartForDistance.Name = "StartForDistance"
        Me.StartForDistance.Size = New System.Drawing.Size(75, 23)
        Me.StartForDistance.TabIndex = 27
        Me.StartForDistance.Text = "开始计算"
        Me.StartForDistance.UseVisualStyleBackColor = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label2.Location = New System.Drawing.Point(178, 55)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(19, 12)
        Me.Label2.TabIndex = 26
        Me.Label2.Text = "to"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label1.Location = New System.Drawing.Point(4, 55)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(97, 12)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "距离限制(米)："
        '
        'caculateFilePath
        '
        Me.caculateFilePath.Font = New System.Drawing.Font("宋体", 8.25!)
        Me.caculateFilePath.Location = New System.Drawing.Point(10, 37)
        Me.caculateFilePath.Name = "caculateFilePath"
        Me.caculateFilePath.ReadOnly = True
        Me.caculateFilePath.ShortcutsEnabled = False
        Me.caculateFilePath.Size = New System.Drawing.Size(306, 20)
        Me.caculateFilePath.TabIndex = 21
        Me.caculateFilePath.Text = "请选择含有计算内容的文件，务必按正确格式（如下）"
        '
        'siteFilePath
        '
        Me.siteFilePath.Font = New System.Drawing.Font("宋体", 8.25!)
        Me.siteFilePath.Location = New System.Drawing.Point(10, 5)
        Me.siteFilePath.Name = "siteFilePath"
        Me.siteFilePath.ReadOnly = True
        Me.siteFilePath.Size = New System.Drawing.Size(306, 20)
        Me.siteFilePath.TabIndex = 19
        Me.siteFilePath.Text = "请选择McomSite文件，务必按正确格式（如下）"
        '
        'OneToMulti
        '
        Me.OneToMulti.AutoSize = True
        Me.OneToMulti.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.OneToMulti.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.OneToMulti.Location = New System.Drawing.Point(172, 20)
        Me.OneToMulti.Name = "OneToMulti"
        Me.OneToMulti.Size = New System.Drawing.Size(88, 16)
        Me.OneToMulti.TabIndex = 18
        Me.OneToMulti.TabStop = True
        Me.OneToMulti.Text = "一对多计算"
        Me.OneToMulti.UseVisualStyleBackColor = True
        '
        'OneToOne
        '
        Me.OneToOne.AutoSize = True
        Me.OneToOne.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.OneToOne.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.OneToOne.Location = New System.Drawing.Point(6, 20)
        Me.OneToOne.Name = "OneToOne"
        Me.OneToOne.Size = New System.Drawing.Size(122, 20)
        Me.OneToOne.TabIndex = 17
        Me.OneToOne.Text = "一对一计算(邻区)"
        Me.OneToOne.UseCompatibleTextRendering = True
        Me.OneToOne.UseVisualStyleBackColor = True
        '
        'TimerForDistance
        '
        Me.TimerForDistance.Interval = 1000
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.GroupBox1.Controls.Add(Me.siteSample)
        Me.GroupBox1.Controls.Add(Me.caculateSample)
        Me.GroupBox1.Font = New System.Drawing.Font("宋体", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.Black
        Me.GroupBox1.Location = New System.Drawing.Point(6, 146)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(416, 37)
        Me.GroupBox1.TabIndex = 33
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "导入文件需要格式："
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.GroupBox2.Controls.Add(Me.ToDis)
        Me.GroupBox2.Controls.Add(Me.FromDis)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.OneToOne)
        Me.GroupBox2.Controls.Add(Me.OneToMulti)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(6, 63)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(310, 77)
        Me.GroupBox2.TabIndex = 34
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "限制条件"
        '
        'ToDis
        '
        Me.ToDis.Location = New System.Drawing.Point(206, 50)
        Me.ToDis.Maximum = New Decimal(New Integer() {100000000, 0, 0, 0})
        Me.ToDis.Name = "ToDis"
        Me.ToDis.Size = New System.Drawing.Size(86, 21)
        Me.ToDis.TabIndex = 28
        Me.ToDis.Value = New Decimal(New Integer() {100000000, 0, 0, 0})
        '
        'FromDis
        '
        Me.FromDis.Location = New System.Drawing.Point(100, 50)
        Me.FromDis.Maximum = New Decimal(New Integer() {1000000, 0, 0, 0})
        Me.FromDis.Name = "FromDis"
        Me.FromDis.Size = New System.Drawing.Size(72, 21)
        Me.FromDis.TabIndex = 27
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.GroupBox3.Controls.Add(Me.isCacuMinDis)
        Me.GroupBox3.Controls.Add(Me.isAllowDuplicate)
        Me.GroupBox3.Controls.Add(Me.isShowCoSite)
        Me.GroupBox3.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(322, 61)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(100, 79)
        Me.GroupBox3.TabIndex = 35
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "一对多选项"
        '
        'isAllowDuplicate
        '
        Me.isAllowDuplicate.AutoSize = True
        Me.isAllowDuplicate.Location = New System.Drawing.Point(7, 41)
        Me.isAllowDuplicate.Name = "isAllowDuplicate"
        Me.isAllowDuplicate.Size = New System.Drawing.Size(76, 16)
        Me.isAllowDuplicate.TabIndex = 1
        Me.isAllowDuplicate.Text = "双向计算"
        Me.isAllowDuplicate.UseVisualStyleBackColor = True
        '
        'isShowCoSite
        '
        Me.isShowCoSite.AutoSize = True
        Me.isShowCoSite.Location = New System.Drawing.Point(7, 21)
        Me.isShowCoSite.Name = "isShowCoSite"
        Me.isShowCoSite.Size = New System.Drawing.Size(76, 16)
        Me.isShowCoSite.TabIndex = 0
        Me.isShowCoSite.Text = "共站计算"
        Me.isShowCoSite.UseVisualStyleBackColor = True
        '
        'isCacuMinDis
        '
        Me.isCacuMinDis.AutoSize = True
        Me.isCacuMinDis.Location = New System.Drawing.Point(7, 63)
        Me.isCacuMinDis.Name = "isCacuMinDis"
        Me.isCacuMinDis.Size = New System.Drawing.Size(76, 16)
        Me.isCacuMinDis.TabIndex = 2
        Me.isCacuMinDis.Text = "只算最近"
        Me.isCacuMinDis.UseVisualStyleBackColor = True
        '
        'DistanceToolDialog
        '
        Me.AllowDrop = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.AntiqueWhite
        Me.ClientSize = New System.Drawing.Size(434, 262)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.caculateSelect)
        Me.Controls.Add(Me.caculatePathLabel)
        Me.Controls.Add(Me.SiteSelect)
        Me.Controls.Add(Me.StartForDistance)
        Me.Controls.Add(Me.caculateFilePath)
        Me.Controls.Add(Me.siteFilePath)
        Me.Controls.Add(Me.DistanceStatus)
        Me.Controls.Add(Me.distanceBar)
        Me.DoubleBuffered = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(450, 300)
        Me.MinimumSize = New System.Drawing.Size(450, 300)
        Me.Name = "DistanceToolDialog"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "距离计算工具"
        Me.DistanceStatus.ResumeLayout(False)
        Me.DistanceStatus.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.ToDis, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FromDis, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BGWForDistance As System.ComponentModel.BackgroundWorker
    Friend WithEvents distanceBar As System.Windows.Forms.ProgressBar
    Friend WithEvents DistanceStatus As System.Windows.Forms.StatusStrip
    Friend WithEvents statusBar As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents caculateSelect As System.Windows.Forms.Button
    Friend WithEvents caculatePathLabel As System.Windows.Forms.Label
    Friend WithEvents SiteSelect As System.Windows.Forms.Button
    Friend WithEvents caculateSample As System.Windows.Forms.LinkLabel
    Friend WithEvents siteSample As System.Windows.Forms.LinkLabel
    Friend WithEvents StartForDistance As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents caculateFilePath As System.Windows.Forms.TextBox
    Friend WithEvents siteFilePath As System.Windows.Forms.TextBox
    Friend WithEvents OneToMulti As System.Windows.Forms.RadioButton
    Friend WithEvents OneToOne As System.Windows.Forms.RadioButton
    Friend WithEvents TimerForDistance As System.Windows.Forms.Timer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents ToDis As System.Windows.Forms.NumericUpDown
    Friend WithEvents FromDis As System.Windows.Forms.NumericUpDown
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents isShowCoSite As System.Windows.Forms.CheckBox
    Friend WithEvents isAllowDuplicate As System.Windows.Forms.CheckBox
    Friend WithEvents isCacuMinDis As System.Windows.Forms.CheckBox
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ConvertLongitudeAndLatitude
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
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.newLatitude = New System.Windows.Forms.TextBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.newLongitude = New System.Windows.Forms.TextBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.latMiao = New System.Windows.Forms.NumericUpDown()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.latFen = New System.Windows.Forms.NumericUpDown()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.latDu = New System.Windows.Forms.NumericUpDown()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.longMiao = New System.Windows.Forms.NumericUpDown()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.longFen = New System.Windows.Forms.NumericUpDown()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.longDu = New System.Windows.Forms.NumericUpDown()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.latMiao, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.latFen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.latDu, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        CType(Me.longMiao, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.longFen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.longDu, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox6)
        Me.GroupBox1.Controls.Add(Me.GroupBox5)
        Me.GroupBox1.Controls.Add(Me.GroupBox4)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Location = New System.Drawing.Point(13, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(444, 145)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "单一转换"
        Me.GroupBox1.UseCompatibleTextRendering = True
        '
        'GroupBox6
        '
        Me.GroupBox6.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.GroupBox6.Controls.Add(Me.newLatitude)
        Me.GroupBox6.Location = New System.Drawing.Point(298, 78)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(123, 51)
        Me.GroupBox6.TabIndex = 9
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "转换后纬度"
        '
        'newLatitude
        '
        Me.newLatitude.Location = New System.Drawing.Point(9, 18)
        Me.newLatitude.Name = "newLatitude"
        Me.newLatitude.ReadOnly = True
        Me.newLatitude.Size = New System.Drawing.Size(108, 21)
        Me.newLatitude.TabIndex = 1
        '
        'GroupBox5
        '
        Me.GroupBox5.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.GroupBox5.Controls.Add(Me.newLongitude)
        Me.GroupBox5.Location = New System.Drawing.Point(298, 21)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(123, 51)
        Me.GroupBox5.TabIndex = 8
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "转换后经度"
        '
        'newLongitude
        '
        Me.newLongitude.Location = New System.Drawing.Point(9, 18)
        Me.newLongitude.Name = "newLongitude"
        Me.newLongitude.ReadOnly = True
        Me.newLongitude.Size = New System.Drawing.Size(108, 21)
        Me.newLongitude.TabIndex = 0
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Label4)
        Me.GroupBox4.Controls.Add(Me.latMiao)
        Me.GroupBox4.Controls.Add(Me.Label5)
        Me.GroupBox4.Controls.Add(Me.latFen)
        Me.GroupBox4.Controls.Add(Me.Label6)
        Me.GroupBox4.Controls.Add(Me.latDu)
        Me.GroupBox4.Location = New System.Drawing.Point(6, 78)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(286, 51)
        Me.GroupBox4.TabIndex = 7
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "纬度"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("宋体", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label4.Location = New System.Drawing.Point(255, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(22, 15)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "秒"
        '
        'latMiao
        '
        Me.latMiao.DecimalPlaces = 3
        Me.latMiao.Location = New System.Drawing.Point(193, 21)
        Me.latMiao.Maximum = New Decimal(New Integer() {60, 0, 0, 0})
        Me.latMiao.Name = "latMiao"
        Me.latMiao.Size = New System.Drawing.Size(59, 21)
        Me.latMiao.TabIndex = 5
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("宋体", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label5.Location = New System.Drawing.Point(165, 23)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(22, 15)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "分"
        '
        'latFen
        '
        Me.latFen.DecimalPlaces = 5
        Me.latFen.Location = New System.Drawing.Point(88, 21)
        Me.latFen.Maximum = New Decimal(New Integer() {60, 0, 0, 0})
        Me.latFen.Name = "latFen"
        Me.latFen.Size = New System.Drawing.Size(75, 21)
        Me.latFen.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("宋体", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label6.Location = New System.Drawing.Point(60, 24)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(22, 15)
        Me.Label6.TabIndex = 2
        Me.Label6.Text = "度"
        '
        'latDu
        '
        Me.latDu.Location = New System.Drawing.Point(7, 21)
        Me.latDu.Maximum = New Decimal(New Integer() {90, 0, 0, 0})
        Me.latDu.Name = "latDu"
        Me.latDu.Size = New System.Drawing.Size(50, 21)
        Me.latDu.TabIndex = 0
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.longMiao)
        Me.GroupBox3.Controls.Add(Me.Label1)
        Me.GroupBox3.Controls.Add(Me.longFen)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Controls.Add(Me.longDu)
        Me.GroupBox3.Location = New System.Drawing.Point(7, 21)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(285, 51)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "经度"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("宋体", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label3.Location = New System.Drawing.Point(254, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(22, 15)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "秒"
        '
        'longMiao
        '
        Me.longMiao.DecimalPlaces = 3
        Me.longMiao.Location = New System.Drawing.Point(193, 21)
        Me.longMiao.Maximum = New Decimal(New Integer() {60, 0, 0, 0})
        Me.longMiao.Name = "longMiao"
        Me.longMiao.Size = New System.Drawing.Size(58, 21)
        Me.longMiao.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("宋体", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(165, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(22, 15)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "分"
        '
        'longFen
        '
        Me.longFen.DecimalPlaces = 5
        Me.longFen.Location = New System.Drawing.Point(88, 21)
        Me.longFen.Maximum = New Decimal(New Integer() {60, 0, 0, 0})
        Me.longFen.Name = "longFen"
        Me.longFen.Size = New System.Drawing.Size(74, 21)
        Me.longFen.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("宋体", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.Location = New System.Drawing.Point(60, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(22, 15)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "度"
        '
        'longDu
        '
        Me.longDu.Location = New System.Drawing.Point(7, 21)
        Me.longDu.Maximum = New Decimal(New Integer() {180, 0, 0, 0})
        Me.longDu.Name = "longDu"
        Me.longDu.Size = New System.Drawing.Size(50, 21)
        Me.longDu.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Location = New System.Drawing.Point(12, 163)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(444, 103)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "批量转换"
        '
        'ConvertLongitudeAndLatitude
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(469, 275)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "ConvertLongitudeAndLatitude"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "经纬度转换"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        CType(Me.latMiao, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.latFen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.latDu, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        CType(Me.longMiao, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.longFen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.longDu, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents longMiao As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents longFen As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents longDu As System.Windows.Forms.NumericUpDown
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents latMiao As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents latFen As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents latDu As System.Windows.Forms.NumericUpDown
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents newLatitude As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents newLongitude As System.Windows.Forms.TextBox
End Class

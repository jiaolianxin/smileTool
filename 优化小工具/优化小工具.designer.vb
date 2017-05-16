<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.statusBar = New System.Windows.Forms.ToolStripStatusLabel()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.文件ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.数据库设置ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.退出 = New System.Windows.Forms.ToolStripMenuItem()
        Me.GSMmenu = New System.Windows.Forms.ToolStripMenuItem()
        Me.CDD相关ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CDD转换 = New System.Windows.Forms.ToolStripMenuItem()
        Me.生成CDD脚本ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.调整CDDLOG格式ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.华为配置转换 = New System.Windows.Forms.ToolStripMenuItem()
        Me.一致性检查 = New System.Windows.Forms.ToolStripMenuItem()
        Me.CoSiteFreqCheck = New System.Windows.Forms.ToolStripMenuItem()
        Me.NeighborCoAdjFreqCheck = New System.Windows.Forms.ToolStripMenuItem()
        Me.NeighborsCoBCCHCoBSICCheck = New System.Windows.Forms.ToolStripMenuItem()
        Me.BSIC规划ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CEL制作 = New System.Windows.Forms.ToolStripMenuItem()
        Me.同BCCH同BSIC距离检查ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NeighborPlan = New System.Windows.Forms.ToolStripMenuItem()
        Me.LTEMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.CheckKGET = New System.Windows.Forms.ToolStripMenuItem()
        Me.PCI规划ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SiteDatabase转Mcom文件 = New System.Windows.Forms.ToolStripMenuItem()
        Me.McomSite转Planet导入文件 = New System.Windows.Forms.ToolStripMenuItem()
        Me.小工具ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Forte文件制作ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.通过Site或基站信息表制作ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.通过McomSite和Carrier制作ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.文件处理相关ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.增加文件后缀 = New System.Windows.Forms.ToolStripMenuItem()
        Me.增加第二层目录下文件的后缀 = New System.Windows.Forms.ToolStripMenuItem()
        Me.生成目录下所有文件名的list = New System.Windows.Forms.ToolStripMenuItem()
        Me.经纬度转换 = New System.Windows.Forms.ToolStripMenuItem()
        Me.距离计算ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.VBA宏密码破解ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.关于ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutAuthor = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.BackgroundWorkerForFormatCDD = New System.ComponentModel.BackgroundWorker()
        Me.MainProgressBar = New System.Windows.Forms.ProgressBar()
        Me.BGWForteMakerWithSiteAndCarrier = New System.ComponentModel.BackgroundWorker()
        Me.BGWForteMakerWithSiteOrSDB = New System.ComponentModel.BackgroundWorker()
        Me.MyNowTime = New System.Windows.Forms.Timer(Me.components)
        Me.Time_Label = New System.Windows.Forms.Label()
        Me.date_Label = New System.Windows.Forms.Label()
        Me.EvaluationTime = New System.Windows.Forms.Timer(Me.components)
        Me.mainMessageBox = New System.Windows.Forms.RichTextBox()
        Me.BGW_CheckCoBcchBsic = New System.ComponentModel.BackgroundWorker()
        Me.BGW_CreateCelFile = New System.ComponentModel.BackgroundWorker()
        Me.BGW_VBACrack = New System.ComponentModel.BackgroundWorker()
        Me.BGWSiteDBtoMcom = New System.ComponentModel.BackgroundWorker()
        Me.BGWSiteToPlanetImport = New System.ComponentModel.BackgroundWorker()
        Me.NotifyShow = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.QuitWindow = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.退出ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip1.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.QuitWindow.SuspendLayout()
        Me.SuspendLayout()
        '
        'StatusStrip1
        '
        Me.StatusStrip1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.statusBar})
        resources.ApplyResources(Me.StatusStrip1, "StatusStrip1")
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.ShowItemToolTips = True
        '
        'statusBar
        '
        Me.statusBar.AutoToolTip = True
        Me.statusBar.Name = "statusBar"
        resources.ApplyResources(Me.statusBar, "statusBar")
        Me.statusBar.Tag = "statusText"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.文件ToolStripMenuItem, Me.GSMmenu, Me.LTEMenuItem1, Me.小工具ToolStripMenuItem, Me.关于ToolStripMenuItem})
        resources.ApplyResources(Me.MenuStrip1, "MenuStrip1")
        Me.MenuStrip1.Name = "MenuStrip1"
        '
        '文件ToolStripMenuItem
        '
        Me.文件ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.数据库设置ToolStripMenuItem, Me.退出})
        Me.文件ToolStripMenuItem.Name = "文件ToolStripMenuItem"
        resources.ApplyResources(Me.文件ToolStripMenuItem, "文件ToolStripMenuItem")
        '
        '数据库设置ToolStripMenuItem
        '
        resources.ApplyResources(Me.数据库设置ToolStripMenuItem, "数据库设置ToolStripMenuItem")
        Me.数据库设置ToolStripMenuItem.Name = "数据库设置ToolStripMenuItem"
        '
        '退出
        '
        Me.退出.Name = "退出"
        resources.ApplyResources(Me.退出, "退出")
        '
        'GSMmenu
        '
        Me.GSMmenu.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CDD相关ToolStripMenuItem, Me.一致性检查, Me.BSIC规划ToolStripMenuItem, Me.CEL制作, Me.同BCCH同BSIC距离检查ToolStripMenuItem, Me.NeighborPlan})
        Me.GSMmenu.Name = "GSMmenu"
        resources.ApplyResources(Me.GSMmenu, "GSMmenu")
        '
        'CDD相关ToolStripMenuItem
        '
        Me.CDD相关ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CDD转换, Me.生成CDD脚本ToolStripMenuItem, Me.调整CDDLOG格式ToolStripMenuItem, Me.华为配置转换})
        Me.CDD相关ToolStripMenuItem.Name = "CDD相关ToolStripMenuItem"
        resources.ApplyResources(Me.CDD相关ToolStripMenuItem, "CDD相关ToolStripMenuItem")
        '
        'CDD转换
        '
        Me.CDD转换.AutoToolTip = True
        Me.CDD转换.Name = "CDD转换"
        resources.ApplyResources(Me.CDD转换, "CDD转换")
        '
        '生成CDD脚本ToolStripMenuItem
        '
        Me.生成CDD脚本ToolStripMenuItem.Name = "生成CDD脚本ToolStripMenuItem"
        resources.ApplyResources(Me.生成CDD脚本ToolStripMenuItem, "生成CDD脚本ToolStripMenuItem")
        '
        '调整CDDLOG格式ToolStripMenuItem
        '
        Me.调整CDDLOG格式ToolStripMenuItem.Name = "调整CDDLOG格式ToolStripMenuItem"
        resources.ApplyResources(Me.调整CDDLOG格式ToolStripMenuItem, "调整CDDLOG格式ToolStripMenuItem")
        '
        '华为配置转换
        '
        Me.华为配置转换.Name = "华为配置转换"
        resources.ApplyResources(Me.华为配置转换, "华为配置转换")
        '
        '一致性检查
        '
        Me.一致性检查.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CoSiteFreqCheck, Me.NeighborCoAdjFreqCheck, Me.NeighborsCoBCCHCoBSICCheck})
        Me.一致性检查.Name = "一致性检查"
        resources.ApplyResources(Me.一致性检查, "一致性检查")
        '
        'CoSiteFreqCheck
        '
        Me.CoSiteFreqCheck.Name = "CoSiteFreqCheck"
        resources.ApplyResources(Me.CoSiteFreqCheck, "CoSiteFreqCheck")
        '
        'NeighborCoAdjFreqCheck
        '
        Me.NeighborCoAdjFreqCheck.Name = "NeighborCoAdjFreqCheck"
        resources.ApplyResources(Me.NeighborCoAdjFreqCheck, "NeighborCoAdjFreqCheck")
        '
        'NeighborsCoBCCHCoBSICCheck
        '
        resources.ApplyResources(Me.NeighborsCoBCCHCoBSICCheck, "NeighborsCoBCCHCoBSICCheck")
        Me.NeighborsCoBCCHCoBSICCheck.Name = "NeighborsCoBCCHCoBSICCheck"
        '
        'BSIC规划ToolStripMenuItem
        '
        Me.BSIC规划ToolStripMenuItem.Name = "BSIC规划ToolStripMenuItem"
        resources.ApplyResources(Me.BSIC规划ToolStripMenuItem, "BSIC规划ToolStripMenuItem")
        '
        'CEL制作
        '
        Me.CEL制作.Name = "CEL制作"
        resources.ApplyResources(Me.CEL制作, "CEL制作")
        '
        '同BCCH同BSIC距离检查ToolStripMenuItem
        '
        Me.同BCCH同BSIC距离检查ToolStripMenuItem.Name = "同BCCH同BSIC距离检查ToolStripMenuItem"
        resources.ApplyResources(Me.同BCCH同BSIC距离检查ToolStripMenuItem, "同BCCH同BSIC距离检查ToolStripMenuItem")
        '
        'NeighborPlan
        '
        resources.ApplyResources(Me.NeighborPlan, "NeighborPlan")
        Me.NeighborPlan.Name = "NeighborPlan"
        '
        'LTEMenuItem1
        '
        Me.LTEMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CheckKGET, Me.PCI规划ToolStripMenuItem, Me.SiteDatabase转Mcom文件, Me.McomSite转Planet导入文件})
        Me.LTEMenuItem1.Name = "LTEMenuItem1"
        resources.ApplyResources(Me.LTEMenuItem1, "LTEMenuItem1")
        '
        'CheckKGET
        '
        Me.CheckKGET.Name = "CheckKGET"
        resources.ApplyResources(Me.CheckKGET, "CheckKGET")
        '
        'PCI规划ToolStripMenuItem
        '
        Me.PCI规划ToolStripMenuItem.Name = "PCI规划ToolStripMenuItem"
        resources.ApplyResources(Me.PCI规划ToolStripMenuItem, "PCI规划ToolStripMenuItem")
        '
        'SiteDatabase转Mcom文件
        '
        Me.SiteDatabase转Mcom文件.Name = "SiteDatabase转Mcom文件"
        resources.ApplyResources(Me.SiteDatabase转Mcom文件, "SiteDatabase转Mcom文件")
        '
        'McomSite转Planet导入文件
        '
        Me.McomSite转Planet导入文件.Name = "McomSite转Planet导入文件"
        resources.ApplyResources(Me.McomSite转Planet导入文件, "McomSite转Planet导入文件")
        '
        '小工具ToolStripMenuItem
        '
        Me.小工具ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Forte文件制作ToolStripMenuItem, Me.文件处理相关ToolStripMenuItem, Me.经纬度转换, Me.距离计算ToolStripMenuItem, Me.VBA宏密码破解ToolStripMenuItem})
        Me.小工具ToolStripMenuItem.Name = "小工具ToolStripMenuItem"
        resources.ApplyResources(Me.小工具ToolStripMenuItem, "小工具ToolStripMenuItem")
        '
        'Forte文件制作ToolStripMenuItem
        '
        Me.Forte文件制作ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.通过Site或基站信息表制作ToolStripMenuItem, Me.通过McomSite和Carrier制作ToolStripMenuItem})
        Me.Forte文件制作ToolStripMenuItem.Name = "Forte文件制作ToolStripMenuItem"
        resources.ApplyResources(Me.Forte文件制作ToolStripMenuItem, "Forte文件制作ToolStripMenuItem")
        '
        '通过Site或基站信息表制作ToolStripMenuItem
        '
        Me.通过Site或基站信息表制作ToolStripMenuItem.Name = "通过Site或基站信息表制作ToolStripMenuItem"
        resources.ApplyResources(Me.通过Site或基站信息表制作ToolStripMenuItem, "通过Site或基站信息表制作ToolStripMenuItem")
        '
        '通过McomSite和Carrier制作ToolStripMenuItem
        '
        Me.通过McomSite和Carrier制作ToolStripMenuItem.Name = "通过McomSite和Carrier制作ToolStripMenuItem"
        resources.ApplyResources(Me.通过McomSite和Carrier制作ToolStripMenuItem, "通过McomSite和Carrier制作ToolStripMenuItem")
        '
        '文件处理相关ToolStripMenuItem
        '
        Me.文件处理相关ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.增加文件后缀, Me.增加第二层目录下文件的后缀, Me.生成目录下所有文件名的list})
        Me.文件处理相关ToolStripMenuItem.Name = "文件处理相关ToolStripMenuItem"
        resources.ApplyResources(Me.文件处理相关ToolStripMenuItem, "文件处理相关ToolStripMenuItem")
        '
        '增加文件后缀
        '
        Me.增加文件后缀.Name = "增加文件后缀"
        resources.ApplyResources(Me.增加文件后缀, "增加文件后缀")
        '
        '增加第二层目录下文件的后缀
        '
        Me.增加第二层目录下文件的后缀.Name = "增加第二层目录下文件的后缀"
        resources.ApplyResources(Me.增加第二层目录下文件的后缀, "增加第二层目录下文件的后缀")
        '
        '生成目录下所有文件名的list
        '
        Me.生成目录下所有文件名的list.Name = "生成目录下所有文件名的list"
        resources.ApplyResources(Me.生成目录下所有文件名的list, "生成目录下所有文件名的list")
        '
        '经纬度转换
        '
        Me.经纬度转换.Name = "经纬度转换"
        resources.ApplyResources(Me.经纬度转换, "经纬度转换")
        '
        '距离计算ToolStripMenuItem
        '
        Me.距离计算ToolStripMenuItem.Name = "距离计算ToolStripMenuItem"
        resources.ApplyResources(Me.距离计算ToolStripMenuItem, "距离计算ToolStripMenuItem")
        '
        'VBA宏密码破解ToolStripMenuItem
        '
        Me.VBA宏密码破解ToolStripMenuItem.Name = "VBA宏密码破解ToolStripMenuItem"
        resources.ApplyResources(Me.VBA宏密码破解ToolStripMenuItem, "VBA宏密码破解ToolStripMenuItem")
        '
        '关于ToolStripMenuItem
        '
        Me.关于ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutAuthor})
        Me.关于ToolStripMenuItem.Name = "关于ToolStripMenuItem"
        resources.ApplyResources(Me.关于ToolStripMenuItem, "关于ToolStripMenuItem")
        '
        'AboutAuthor
        '
        Me.AboutAuthor.Name = "AboutAuthor"
        resources.ApplyResources(Me.AboutAuthor, "AboutAuthor")
        '
        'OpenFileDialog1
        '
        resources.ApplyResources(Me.OpenFileDialog1, "OpenFileDialog1")
        Me.OpenFileDialog1.Multiselect = True
        Me.OpenFileDialog1.ReadOnlyChecked = True
        Me.OpenFileDialog1.ShowHelp = True
        Me.OpenFileDialog1.ShowReadOnly = True
        Me.OpenFileDialog1.SupportMultiDottedExtensions = True
        '
        'BackgroundWorkerForFormatCDD
        '
        Me.BackgroundWorkerForFormatCDD.WorkerReportsProgress = True
        Me.BackgroundWorkerForFormatCDD.WorkerSupportsCancellation = True
        '
        'MainProgressBar
        '
        resources.ApplyResources(Me.MainProgressBar, "MainProgressBar")
        Me.MainProgressBar.Name = "MainProgressBar"
        '
        'BGWForteMakerWithSiteAndCarrier
        '
        Me.BGWForteMakerWithSiteAndCarrier.WorkerReportsProgress = True
        Me.BGWForteMakerWithSiteAndCarrier.WorkerSupportsCancellation = True
        '
        'BGWForteMakerWithSiteOrSDB
        '
        Me.BGWForteMakerWithSiteOrSDB.WorkerReportsProgress = True
        Me.BGWForteMakerWithSiteOrSDB.WorkerSupportsCancellation = True
        '
        'MyNowTime
        '
        Me.MyNowTime.Interval = 1000
        '
        'Time_Label
        '
        resources.ApplyResources(Me.Time_Label, "Time_Label")
        Me.Time_Label.ForeColor = System.Drawing.Color.Blue
        Me.Time_Label.Name = "Time_Label"
        '
        'date_Label
        '
        resources.ApplyResources(Me.date_Label, "date_Label")
        Me.date_Label.ForeColor = System.Drawing.Color.Blue
        Me.date_Label.Name = "date_Label"
        '
        'EvaluationTime
        '
        Me.EvaluationTime.Interval = 1000
        '
        'mainMessageBox
        '
        resources.ApplyResources(Me.mainMessageBox, "mainMessageBox")
        Me.mainMessageBox.Name = "mainMessageBox"
        Me.mainMessageBox.ReadOnly = True
        '
        'BGW_CheckCoBcchBsic
        '
        Me.BGW_CheckCoBcchBsic.WorkerReportsProgress = True
        Me.BGW_CheckCoBcchBsic.WorkerSupportsCancellation = True
        '
        'BGW_CreateCelFile
        '
        Me.BGW_CreateCelFile.WorkerReportsProgress = True
        Me.BGW_CreateCelFile.WorkerSupportsCancellation = True
        '
        'BGW_VBACrack
        '
        Me.BGW_VBACrack.WorkerReportsProgress = True
        Me.BGW_VBACrack.WorkerSupportsCancellation = True
        '
        'BGWSiteDBtoMcom
        '
        Me.BGWSiteDBtoMcom.WorkerReportsProgress = True
        Me.BGWSiteDBtoMcom.WorkerSupportsCancellation = True
        '
        'BGWSiteToPlanetImport
        '
        Me.BGWSiteToPlanetImport.WorkerReportsProgress = True
        Me.BGWSiteToPlanetImport.WorkerSupportsCancellation = True
        '
        'NotifyShow
        '
        Me.NotifyShow.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info
        resources.ApplyResources(Me.NotifyShow, "NotifyShow")
        Me.NotifyShow.ContextMenuStrip = Me.QuitWindow
        Me.NotifyShow.Tag = ""
        '
        'QuitWindow
        '
        Me.QuitWindow.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.QuitWindow.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.退出ToolStripMenuItem})
        Me.QuitWindow.Name = "QuitWindow"
        resources.ApplyResources(Me.QuitWindow, "QuitWindow")
        Me.QuitWindow.TabStop = True
        '
        '退出ToolStripMenuItem
        '
        Me.退出ToolStripMenuItem.Name = "退出ToolStripMenuItem"
        resources.ApplyResources(Me.退出ToolStripMenuItem, "退出ToolStripMenuItem")
        '
        'MainForm
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        resources.ApplyResources(Me, "$this")
        Me.Controls.Add(Me.mainMessageBox)
        Me.Controls.Add(Me.date_Label)
        Me.Controls.Add(Me.Time_Label)
        Me.Controls.Add(Me.MainProgressBar)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.DoubleBuffered = True
        Me.HelpButton = True
        Me.KeyPreview = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MaximizeBox = False
        Me.Name = "MainForm"
        Me.TransparencyKey = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.QuitWindow.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents 文件ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 退出 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 小工具ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 关于ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutAuthor As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents statusBar As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Forte文件制作ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 通过Site或基站信息表制作ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 通过McomSite和Carrier制作ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 文件处理相关ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 增加文件后缀 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 生成目录下所有文件名的list As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents 增加第二层目录下文件的后缀 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BackgroundWorkerForFormatCDD As System.ComponentModel.BackgroundWorker
    Friend WithEvents MainProgressBar As System.Windows.Forms.ProgressBar
    Friend WithEvents BGWForteMakerWithSiteAndCarrier As System.ComponentModel.BackgroundWorker
    Friend WithEvents BGWForteMakerWithSiteOrSDB As System.ComponentModel.BackgroundWorker
    Friend WithEvents MyNowTime As System.Windows.Forms.Timer
    Friend WithEvents Time_Label As System.Windows.Forms.Label
    Friend WithEvents date_Label As System.Windows.Forms.Label
    Friend WithEvents EvaluationTime As System.Windows.Forms.Timer
    Friend WithEvents 距离计算ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GSMmenu As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CDD相关ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CDD转换 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 生成CDD脚本ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 调整CDDLOG格式ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BSIC规划ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 同BCCH同BSIC距离检查ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LTEMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CheckKGET As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mainMessageBox As System.Windows.Forms.RichTextBox
    Friend WithEvents BGW_CheckCoBcchBsic As System.ComponentModel.BackgroundWorker
    Friend WithEvents CEL制作 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BGW_CreateCelFile As System.ComponentModel.BackgroundWorker
    Friend WithEvents VBA宏密码破解ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BGW_VBACrack As System.ComponentModel.BackgroundWorker
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents PCI规划ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 数据库设置ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SiteDatabase转Mcom文件 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BGWSiteDBtoMcom As System.ComponentModel.BackgroundWorker
    Friend WithEvents 经纬度转换 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents McomSite转Planet导入文件 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BGWSiteToPlanetImport As System.ComponentModel.BackgroundWorker
    Friend WithEvents NeighborPlan As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NotifyShow As System.Windows.Forms.NotifyIcon
    Friend WithEvents QuitWindow As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents 退出ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 华为配置转换 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 一致性检查 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NeighborsCoBCCHCoBSICCheck As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NeighborCoAdjFreqCheck As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CoSiteFreqCheck As System.Windows.Forms.ToolStripMenuItem

End Class

﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DBconDialog
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
        Me.StartCon = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'StartCon
        '
        Me.StartCon.Location = New System.Drawing.Point(131, 174)
        Me.StartCon.Name = "StartCon"
        Me.StartCon.Size = New System.Drawing.Size(75, 23)
        Me.StartCon.TabIndex = 0
        Me.StartCon.Text = "确定"
        Me.StartCon.UseVisualStyleBackColor = True
        '
        'DBconDialog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(345, 209)
        Me.Controls.Add(Me.StartCon)
        Me.Name = "DBconDialog"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "数据库连接设置"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents StartCon As System.Windows.Forms.Button
End Class

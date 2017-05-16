Imports Microsoft.Office.Interop
Imports System.ComponentModel
Imports 优化小工具.CreateCelFile

Public Class MainForm
    Friend totalEvaTime
    Public isSiteDatabaseToMcom As Boolean = False
    Private siteDBFileName As String = ""
    Protected Friend Sub startup() Handles Me.Load
        Me.Text = "优化小工具V" & My.Application.Info.Version.ToString()
        getInitTime(Application.ExecutablePath)
        Time_Label.Text = Now
        If Weekday(Now, FirstDayOfWeek.Monday) = 1 Then
            date_Label.Text = "星期一"
        ElseIf Weekday(Now, FirstDayOfWeek.Monday) = 2 Then
            date_Label.Text = "星期二"
        ElseIf Weekday(Now, FirstDayOfWeek.Monday) = 3 Then
            date_Label.Text = "星期三"
        ElseIf Weekday(Now, FirstDayOfWeek.Monday) = 4 Then
            date_Label.Text = "星期四"
        ElseIf Weekday(Now, FirstDayOfWeek.Monday) = 5 Then
            date_Label.Text = "星期五"
        ElseIf Weekday(Now, FirstDayOfWeek.Monday) = 6 Then
            date_Label.Text = "星期六"
        ElseIf Weekday(Now, FirstDayOfWeek.Monday) = 7 Then
            date_Label.Text = "星期日"
        End If
        If timeRemainForTool Then
            If remainDays <= 3 Then
                aboutMe.ShowDialog()
                MsgBox("此版本的有效期还有" & remainDays & "天！" & vbCrLf & "请尽快到QQ群：202196045下载新版本，或者邮件联系作者，免费获得新版本！", MsgBoxStyle.Information)
            End If
            MyNowTime.Start()
        Else
            'newVersionAddress.Show()
            aboutMe.ShowDialog()
            MsgBox("软件使用到期！仅为了保证您使用的是最新的版本，请联系作者更新：18601105814@163.com！或者加QQ群：202196045下载新版本，或者邮件联系作者，免费获得新版本！", MsgBoxStyle.Critical)
            MyNowTime.Stop()
            Close()
        End If
    End Sub
    Private Sub 退出优化小工具() Handles Me.FormClosed
        MyNowTime.Stop()
        NotifyShow.Dispose()
    End Sub
    Private Sub minimalMainForm() Handles Me.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            Me.NotifyShow.Visible = True
            Me.Hide()
            Me.NotifyShow.ShowBalloonTip(500)
        End If
        Activate()
    End Sub
    Private Sub 通知栏图标控制(sender As Object, e As MouseEventArgs) Handles NotifyShow.DoubleClick
        Me.Show()
        '   Me.WindowState = FormWindowState.Maximized
        Me.WindowState = FormWindowState.Normal
        Me.CenterToScreen()
        Me.NotifyShow.Visible = False
    End Sub
    Private Sub 退出ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 退出ToolStripMenuItem.Click
        MyNowTime.Stop()
        Me.Close()
        Me.Dispose()
        NotifyShow.Dispose()

    End Sub
    Private Sub 退出_Click(sender As Object, e As EventArgs) Handles 退出.Click
        MyNowTime.Stop()
        Me.Close()
        Me.Dispose()
        NotifyShow.Dispose()

    End Sub

    Private Sub showNowTime() Handles MyNowTime.Tick
        Time_Label.Text = Now
        totalEvaTime = totalEvaTime + 1
        If Not timeRemainForTool Then
            MsgBox("软件使用到期！仅为了保证您使用的是最新的版本，请联系作者更新：18601105814@163.com！", MsgBoxStyle.Critical)
            MyNowTime.Stop()
            Close()
        End If
    End Sub

    Private Sub ChangeStatus(status As String)
        statusBar.Text = status
        StatusStrip1.Refresh()
    End Sub
    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutAuthor.Click
        aboutMe.Show()
    End Sub

    Private Sub CDD解析_Click(sender As Object, e As EventArgs) Handles CDD转换.Click
        CDD_parse.Show()
    End Sub

    Private Sub 调整CDDLOG格式ToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles 调整CDDLOG格式ToolStripMenuItem.Click
        If BackgroundWorkerForFormatCDD.IsBusy() Then
            MsgBox("正在进行的CDDlog调整任务尚未完成，稍候再试吧！", MsgBoxStyle.Critical)
        Else
            If FormatCDDlog_backgroundworker.getFiles().length = 0 Then
                Exit Sub
            Else
                BackgroundWorkerForFormatCDD.RunWorkerAsync()
            End If
        End If
    End Sub
    Private Sub BackgroundWorkerForFormatCDD_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorkerForFormatCDD.DoWork
        FormatCDDlog_backgroundworker.doFormatWork(sender, e)
    End Sub
    Private Sub BackgroundWorkerForFormatCDD_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BackgroundWorkerForFormatCDD.ProgressChanged
        MainProgressBar.Value = CInt(e.ProgressPercentage)
        statusBar.Text = CStr(e.UserState.ToString)
    End Sub
    Private Sub BackgroundWorkerForFormatCDD_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorkerForFormatCDD.RunWorkerCompleted
        If e.Cancelled Then
            '            MsgBox("执行完毕！")
            ChangeStatus("欢迎使用！")
            MainProgressBar.Value = 0

        Else
            BackgroundWorkerForFormatCDD.CancelAsync()
            ChangeStatus("欢迎使用！")
            MainProgressBar.Value = 0

        End If
    End Sub

    Private Sub 通过McomSite和Carrier制作ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 通过McomSite和Carrier制作ToolStripMenuItem.Click
        If BGWForteMakerWithSiteAndCarrier.IsBusy() Then
            MsgBox("正在进行的Forte文件制作任务尚未完成，稍候再试吧！", MsgBoxStyle.Information)
        Else
            If makeForteFilewithSite_Carrier.getFiles().length < 2 Then
                MsgBox("请同时选择MCOM Site 和 Carrier 两个文件！", MsgBoxStyle.Information)
                Exit Sub
            Else
                BGWForteMakerWithSiteAndCarrier.RunWorkerAsync()
            End If
        End If
    End Sub
    Private Sub BKWForteMakerWithSiteAndCarrier_DoWork(sender As Object, e As DoWorkEventArgs) Handles BGWForteMakerWithSiteAndCarrier.DoWork
        Try
            makeForteFilewithSite_Carrier.makeForteFilewithSiteCarrer(sender, e)
        Catch ex As Exception
            BGWForteMakerWithSiteAndCarrier.CancelAsync()
            MsgBox("请先检查，导入文件是否存在格式错误？或数据格式错误？如：经度或纬度的前后存在换行符？" & vbCrLf & "如不是上述问题，请联系作者：jiaolianxin227@163.com !" & vbCrLf & ex.ToString)
        End Try
    End Sub
    Private Sub BKWForteMakerWithSiteAndCarrier_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BGWForteMakerWithSiteAndCarrier.ProgressChanged
        MainProgressBar.Value = CInt(e.ProgressPercentage)
        statusBar.Text = CStr(e.UserState.ToString)
    End Sub
    Private Sub BKWForteMakerWithSiteAndCarrier_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BGWForteMakerWithSiteAndCarrier.RunWorkerCompleted
        If e.Cancelled Then
            ChangeStatus("欢迎使用！")
            MainProgressBar.Value = 0
        Else
            BGWForteMakerWithSiteAndCarrier.CancelAsync()
            ChangeStatus("欢迎使用！")
            MainProgressBar.Value = 0
        End If
    End Sub

    Private Sub 通过Site或基站信息表制作ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 通过Site或基站信息表制作ToolStripMenuItem.Click
        If BGWForteMakerWithSiteOrSDB.IsBusy Then
            MsgBox("正在进行的Forte文件制作任务尚未完成，稍候再试吧！", MsgBoxStyle.Critical)
        Else
            If MakeForteFileWithSiteOrSiteDatabase.getFiles().length = 0 Then
            Else
                BGWForteMakerWithSiteOrSDB.RunWorkerAsync()
            End If
        End If
    End Sub
    Private Sub BGWForteMakerWithSiteOrSDB_DoWork(sender As Object, e As DoWorkEventArgs) Handles BGWForteMakerWithSiteOrSDB.DoWork
        Try
            MakeForteFileWithSiteOrSiteDatabase.doMaker(sender, e)
        Catch ex As Exception
            BGWForteMakerWithSiteOrSDB.CancelAsync()
            MsgBox("请先检查，导入文件是否存在格式错误？或数据格式错误？如：经度或纬度的前后存在换行符？" & vbCrLf & "如不是上述问题，请联系作者：jiaolianxin227@163.com !" & vbCrLf & ex.ToString)
        End Try
    End Sub
    Private Sub BGWForteMakerWithSiteOrSDB_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BGWForteMakerWithSiteOrSDB.ProgressChanged
        MainProgressBar.Value = e.ProgressPercentage
        statusBar.Text = e.UserState.ToString
    End Sub
    Private Sub BGWForteMakerWithSiteOrSDB_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BGWForteMakerWithSiteOrSDB.RunWorkerCompleted
        If e.Cancelled Then
            ChangeStatus("欢迎使用！")
            MainProgressBar.Value = 0
        Else
            BGWForteMakerWithSiteOrSDB.CancelAsync()
            ChangeStatus("欢迎使用！")
            MainProgressBar.Value = 0
        End If
    End Sub

    Private Sub 生成目录下所有文件名的listToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 生成目录下所有文件名的list.Click
        Dim myPath
        Dim fso, filename, outputFile, fileShortName, fileExtName

        MsgBox("请选择想要生成文件名list的文件夹！")
        FolderBrowserDialog1.Description = "请选择需要生成文件列表的文件夹！"
        FolderBrowserDialog1.ShowDialog()
        myPath = FolderBrowserDialog1.SelectedPath
        FolderBrowserDialog1.Reset()
        If myPath = "" Then
            Exit Sub
        End If
        fso = CreateObject("Scripting.FileSystemObject")
        filename = Dir(myPath & "\*.*")
        outputFile = fso.openTextFile(myPath & "\fileNameList.csv", 2, 1, 0)
        outputFile.writeline("文件名,扩展名")
        Do While filename <> ""
            fileShortName = Strings.Split(Strings.Right(filename, Len(filename) - InStrRev(filename, "\")), ".")(0)
            If UBound(Strings.Split(filename, ".")) = 0 Then
                fileExtName = ""
            Else
                fileExtName = Strings.Split(filename, ".")(1)
            End If
            outputFile.writeline(fileShortName & "," & fileExtName)
            filename = Dir()
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
        fso = Nothing
        ChangeStatus("生成完毕！")
        MsgBox("生成完毕！")
        Shell("explorer.exe " & myPath, 1)
    End Sub

    Private Sub 增加文件后缀_Click(sender As Object, e As EventArgs) Handles 增加文件后缀.Click
        Dim myPath, fname, extname, oldname, newname
        Dim fso As Object
        Dim folder
        Dim a(1000) As String
        Dim b

        MsgBox("请选择带有子目录的文件夹!")
        FolderBrowserDialog1.Description = "请选择增加后缀文件所在文件夹的上级文件夹！"
        FolderBrowserDialog1.ShowDialog()
        myPath = FolderBrowserDialog1.SelectedPath
        FolderBrowserDialog1.Reset()
        If myPath = "" Then
            Exit Sub
        End If
        fso = CreateObject("Scripting.FileSystemObject")
        folder = fso.getfolder(myPath)
        b = 1
        For Each thing In folder.subfolders
            a(b) = Strings.Right(thing.path, Len(thing.path) - InStrRev(thing.path, "\"))
            Dim filename As String
            filename = Dir(thing.path & "\*.*")
            Do While filename <> ""
                ChangeStatus("正在整理文件夹：" & a(b) & "下的文件!")
                fname = Split(filename, ".")(0)
                If UBound(Split(filename, ".")) < 1 Then
                    extname = ""
                Else
                    extname = Split(filename, ".")(1)
                End If
                oldname = thing.path & "\" & filename
                newname = myPath & "\" & fname & "_" & a(b) & "." & extname
                FileSystem.FileCopy(oldname, newname)
                filename = Dir()
            Loop
            'If thing.type = "File folder" And thing.size = 0 And thing.subfolders.count = 0 Then
            '    thing.delete()
            'End If
            b = b + 1
        Next

        folder = Nothing
        fso = Nothing
        ChangeStatus("转换完毕！")
        MsgBox("转换完毕！谢谢使用，再见！")
        Shell("explorer.exe " & myPath, 1)
        myPath = Nothing
    End Sub

    Private Sub 生成CDD脚本ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 生成CDD脚本ToolStripMenuItem.Click
        CDDCommand.Show()
    End Sub

    Private Sub 增加第二层目录下文件的后缀_Click(sender As Object, e As EventArgs) Handles 增加第二层目录下文件的后缀.Click
        Dim myPath, fname, extname, oldname, newname
        Dim fso As Object
        Dim folder1, folder2
        Dim a(1000) As String
        Dim b

        MsgBox("请选择带有子目录的文件夹!")
        FolderBrowserDialog1.Description = "请选择增加后缀文件所在文件夹的上级文件夹！"
        FolderBrowserDialog1.ShowDialog()
        myPath = FolderBrowserDialog1.SelectedPath
        FolderBrowserDialog1.Reset()
        If myPath = "" Then
            Exit Sub
        End If
        fso = CreateObject("Scripting.FileSystemObject")
        folder1 = fso.getfolder(myPath)
        b = 1
        For Each thing1 In folder1.subfolders
            folder2 = fso.getFolder(thing1.path)
            For Each thing In folder2.subFolders
                a(b) = Strings.Replace(Strings.Right(thing.path, Len(thing.path) - InStrRev(thing.path, "\")), "-", "_")
                a(b) = Strings.Replace(Strings.Right(thing1.path, Len(thing1.path) - InStrRev(thing1.path, "\")), "-", "_") & "_" & a(b)
                Dim filename As String
                filename = Dir(thing.path & "\*.*")
                Do While filename <> ""
                    ChangeStatus("正在整理文件夹：" & a(b) & "下的文件!")
                    fname = Split(filename, ".")(0)
                    If UBound(Split(filename, ".")) < 1 Then
                        extname = ""
                    Else
                        extname = Split(filename, ".")(1)
                    End If
                    oldname = thing.path & "\" & filename
                    newname = myPath & "\" & fname & "_" & a(b) & "." & extname
                    FileSystem.FileCopy(oldname, newname)
                    filename = Dir()
                Loop
                'If thing.type = "File folder" And thing.size = 0 And thing.subfolders.count = 0 Then
                '    thing.delete()
                'End If
                b = b + 1
            Next

        Next

        folder1 = Nothing
        folder2 = Nothing
        fso = Nothing
        ChangeStatus("转换完毕！")
        MsgBox("转换完毕！谢谢使用，再见！")
        Shell("explorer.exe " & myPath, 1)
        myPath = Nothing
    End Sub

    'Private Sub 删除重复对(sender As Object, e As EventArgs)
    '    Dim fso As Object
    '    Dim outputfile As Object
    '    Dim templine As String
    '    Dim templine1 As String
    '    Dim aaa As ArrayList
    '    Dim temp_arr As Array
    '    aaa = New ArrayList
    '    fso = CreateObject("Scripting.FileSystemObject")
    '    outputfile = fso.OpenTextFile("d:\b.csv", 2, 1, 0)
    '    FileClose(1)
    '    FileOpen(1, "d:\distance.txt", OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
    '    Do While Not EOF(1)
    '        templine = LineInput(1)
    '        temp_arr = Split(templine, "	")
    '        templine = temp_arr(0) & "#" & temp_arr(1)
    '        templine1 = temp_arr(1) & "#" & temp_arr(0)
    '        If Not aaa.Contains(templine) And Not aaa.Contains(templine1) Then
    '            aaa.Add(templine)
    '            outputfile.write(temp_arr(0))
    '            For x = 1 To UBound(temp_arr)
    '                outputfile.write("," & temp_arr(x))
    '            Next
    '            '            outputfile.writeLine(temp_arr(0) & "," & temp_arr(1) & "," & temp_arr(2) & "," & temp_arr(3) & "," & temp_arr(4) & "," & temp_arr(5) & "," & temp_arr(6))
    '            outputfile.writeLine()
    '        End If
    '    Loop

    '    FileClose(1)
    '    outputfile.close()
    '    outputfile = Nothing
    '    fso = Nothing
    '    MsgBox("完成！")

    'End Sub

    Private Sub 距离计算ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 距离计算ToolStripMenuItem.Click
        DistanceToolDialog.Show()
    End Sub

    Private Sub KGET检查_Click(sender As Object, e As EventArgs) Handles CheckKGET.Click
        KGET_Check.Show()
    End Sub

    Private Sub 同BCCH同BSIC距离检查ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 同BCCH同BSIC距离检查ToolStripMenuItem.Click
        Dim siteFileName As String, carrierFileName As String
        MsgBox("请务必保证Site 和 Carrier文件的小区信息对应！", MsgBoxStyle.Information + MsgBoxStyle.DefaultButton1, "提示")
        MsgBox("请选择Mcom Site文件!", MsgBoxStyle.Information)
        siteFileName = selectFile("选择Mcom Site文件", "Mcom Site文件(*.txt)|*.txt", False, True, True)
        If siteFileName = "" Then
            MsgBox("请选择有效的Mcom site文件！", MsgBoxStyle.Critical)
            Exit Sub
        Else
            MsgBox("请选择Mcom Carrier文件!", MsgBoxStyle.Information)
            carrierFileName = selectFile("选择Mcom Carrier文件", "Mcom Carrier文件(*.txt)|*.txt", False, True, True)
            If carrierFileName = "" Then
                MsgBox("请选择有效的Mcom Carrier文件！", MsgBoxStyle.Critical)
                Exit Sub
            Else
                totalEvaTime = 0
                mainMessageBox.Clear()
                BGW_CheckCoBcchBsic.RunWorkerAsync(siteFileName & "#" & carrierFileName)
            End If
        End If
    End Sub
    Private Sub checkCoBcchBsic_doWork(sender As Object, e As DoWorkEventArgs) Handles BGW_CheckCoBcchBsic.DoWork
        CheckCoBCCH_BSIC.Check_CoBcchBsic(sender, e, Me)
    End Sub
    Private Sub checkCoBcchBsic_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BGW_CheckCoBcchBsic.ProgressChanged
        MainProgressBar.Value = IIf(CInt(e.ProgressPercentage) > 100, 100, CInt(e.ProgressPercentage))
        If InStr(CStr(e.UserState.ToString), "@") Then
            mainMessageBox.AppendText(Split(CStr(e.UserState.ToString), "@")(1) & vbCrLf)
            statusBar.Text = Split(CStr(e.UserState.ToString), "@")(0)
        Else
            statusBar.Text = CStr(e.UserState.ToString)
        End If
    End Sub
    Private Sub checkCoBcchBsic_runCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BGW_CheckCoBcchBsic.RunWorkerCompleted
        If e.Cancelled Then
            ChangeStatus("欢迎使用！")
            MainProgressBar.Value = 0
        Else
            BGW_CheckCoBcchBsic.CancelAsync()
            ChangeStatus("欢迎使用！")
            MainProgressBar.Value = 0
        End If
    End Sub

    Private Sub BSIC规划ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BSIC规划ToolStripMenuItem.Click
        If timeRemainForTool Then
            BSICPlanDialog.Show()
        Else
            MsgBox("BSIC规划功能已过期，请联系作者！", MsgBoxStyle.Critical)
        End If
    End Sub

    Private Sub CEL制作_Click(sender As Object, e As EventArgs) Handles CEL制作.Click
        Dim response As Object
        Dim siteFileName As String, carrierFileName As String, neighborFileName As String
        MsgBox("本模块主要通过Mcom Site，Mcom Carrier, Mcom Neighbor(可选)三个文件制作.CEL文件！", MsgBoxStyle.Information, "友情提示")

        MsgBox("请选择Mcom Site文件！如果需要显示中文名，请在siteName列加入中文名！", MsgBoxStyle.Information)
        siteFileName = selectFile("选择需要MCOM Site文件(必须包含:Cell,Longitude,Latitude,Dir,SiteName<可选>）", "Mcom Site文件(*.txt)|*.txt", False, True, True)
        If siteFileName <> "" Then
            MsgBox("请选择Mcom Carrier文件！(必须包含CELL,LAI,CI,BCCH,BSIC)", MsgBoxStyle.Information)
            carrierFileName = selectFile("选择需要MCOM Carrier文件", "Mcom Carrier文件(*.txt)|*.txt", False, True, True)
            If carrierFileName <> "" Then
                MsgBox("请选择Mcom Neighbor文件！", MsgBoxStyle.Information)
                neighborFileName = selectFile("选择需要MCOM Neighbor文件", "Mcom Neighbor文件(*.txt)|*.txt", False, True, True)
                If neighborFileName = "" Then
                    response = MsgBox("如果不选择Neighbor文件，生成的.CEL文件中将不会出现邻区的LAC,CI.是否继续？", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "友情提示")
                    If response = vbYes Then
                        If Not BGW_CreateCelFile.IsBusy Then
                            BGW_CreateCelFile.RunWorkerAsync(siteFileName & "#" & carrierFileName)
                        End If
                    Else
                        Exit Sub
                    End If
                Else
                    If Not BGW_CreateCelFile.IsBusy Then
                        BGW_CreateCelFile.RunWorkerAsync(siteFileName & "#" & carrierFileName & "#" & neighborFileName)
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub createCelFile_dowork(sender As Object, e As DoWorkEventArgs) Handles BGW_CreateCelFile.DoWork
        CreateCelFile.startCreateCelFile(sender, e)
    End Sub
    Private Sub createCelFile_changedProgress(sender As Object, e As ProgressChangedEventArgs) Handles BGW_CreateCelFile.ProgressChanged
        MainProgressBar.Value = IIf(CInt(e.ProgressPercentage) > 100, 100, CInt(e.ProgressPercentage))
        If InStr(CStr(e.UserState.ToString), "@") Then
            mainMessageBox.AppendText(Split(CStr(e.UserState.ToString), "@")(1) & vbCrLf)
            statusBar.Text = Split(CStr(e.UserState.ToString), "@")(0)
        Else
            statusBar.Text = CStr(e.UserState.ToString)
        End If
    End Sub
    Private Sub createCelFile_runWorkCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BGW_CreateCelFile.RunWorkerCompleted
        If e.Cancelled Then
            '            MsgBox("执行完毕！")
            ChangeStatus("欢迎使用！")
            MainProgressBar.Value = 0

        Else
            BGW_CreateCelFile.CancelAsync()
            ChangeStatus("欢迎使用！")
            MainProgressBar.Value = 0

        End If
    End Sub

    Private Sub VBA宏密码破解ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VBA宏密码破解ToolStripMenuItem.Click
        If timeRemainForTool Then
            '你要解保护的Excel文件路径
            Dim fileName As String
            MsgBox("请选择要破解VBA宏的EXCEL文件(仅限EXCEL97-2003）!", MsgBoxStyle.Information)

            fileName = selectFile("选择要破解VBA宏的EXCEL文件", "Excel文件(*.xls;*.xla;*.xlt;*.xlsm)|*.xls;*.xla;*.xlt;*.xlsm", _
                                  False, True, True)
            If Dir(fileName) = "" Or fileName = "" Then
                MsgBox("没找到相关文件,请重新设置。", MsgBoxStyle.Critical)
                Exit Sub
            Else
                FileCopy(fileName, fileName & ".bak") '备份文件。
                BGW_VBACrack.RunWorkerAsync(fileName)
            End If
        Else
            MsgBox("VBA宏密码破解功能已过期，请联系作者！", MsgBoxStyle.Critical)
        End If
    End Sub
    Private Sub VBAmacro_dowork(sender As Object, e As DoWorkEventArgs) Handles BGW_VBACrack.DoWork
        VBAMacroCrack.crack97_2003VBAmacro(e.Argument, sender, e)
    End Sub
    Private Sub VBAmacro_progressChaneged(sender As Object, e As ProgressChangedEventArgs) Handles BGW_VBACrack.ProgressChanged
        MainProgressBar.Value = IIf(CInt(e.ProgressPercentage) > 100, 100, CInt(e.ProgressPercentage))
        If InStr(CStr(e.UserState.ToString), "@") Then
            mainMessageBox.AppendText(Split(CStr(e.UserState.ToString), "@")(1) & vbCrLf)
            statusBar.Text = Split(CStr(e.UserState.ToString), "@")(0)
        Else
            statusBar.Text = CStr(e.UserState.ToString)
        End If
    End Sub
    Private Sub VBAmacro_runWorkCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BGW_VBACrack.RunWorkerCompleted
        If e.Cancelled Then
            '            MsgBox("执行完毕！")
            ChangeStatus("欢迎使用！")
            MainProgressBar.Value = 0
            '    fileName = ""
        Else
            BGW_VBACrack.CancelAsync()
            ChangeStatus("欢迎使用！")
            MainProgressBar.Value = 0
            '      fileName = ""
        End If
    End Sub

    Private Sub PCI规划ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PCI规划ToolStripMenuItem.Click
        If timeRemainForTool Then
            PCIPlanDialog.Show()
        Else
            MsgBox("PCI规划功能已过期，请联系作者！", MsgBoxStyle.Critical)
        End If

    End Sub

    Private Sub 数据库设置ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 数据库设置ToolStripMenuItem.Click
        DBconDialog.Show()
    End Sub

    Private Sub SiteDatabase转Mcom文件_Click(sender As Object, e As EventArgs) Handles SiteDatabase转Mcom文件.Click
        isSiteDatabaseToMcom = True

        MsgBox("请选择SiteDatabase文件！", MsgBoxStyle.Information)
        siteDBFileName = selectFile("选择SiteDatabase文件", "SiteDatabase文件(*.xls;*.xlsx)|*.xls;*.xlsx", False, True, True)
        If siteDBFileName <> "" Then
            selectWorkSheet(siteDBFileName)
        End If
    End Sub
    Private Sub siteDBtoMcom_doWork(sender As Object, e As DoWorkEventArgs) Handles BGWSiteDBtoMcom.DoWork
        totalEvaTime = 0
        SiteDatabaseToMcomFile.startConvertSiteDbToMcom(siteDBFileName, e.Argument, sender, e)
    End Sub
    Private Sub siteDBtoMcom_progressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BGWSiteDBtoMcom.ProgressChanged
        MainProgressBar.Value = IIf(CInt(e.ProgressPercentage) > 100, 100, CInt(e.ProgressPercentage))
        If InStr(CStr(e.UserState.ToString), "@") Then
            mainMessageBox.AppendText(Split(CStr(e.UserState.ToString), "@")(1) & vbCrLf)
            statusBar.Text = Split(CStr(e.UserState.ToString), "@")(0)
        Else
            statusBar.Text = CStr(e.UserState.ToString)
        End If
    End Sub
    Private Sub siteDBtoMcom_runWorkCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BGWSiteDBtoMcom.RunWorkerCompleted
        If e.Cancelled Then
            '            MsgBox("执行完毕！")
            ChangeStatus("欢迎使用！")
            MainProgressBar.Value = 0
            '    fileName = ""
        Else
            BGW_VBACrack.CancelAsync()
            ChangeStatus("欢迎使用！")
            MainProgressBar.Value = 0
            '      fileName = ""
        End If
    End Sub

    Private Sub 经纬度转换_Click(sender As Object, e As EventArgs) Handles 经纬度转换.Click
        ConvertLongitudeAndLatitude.Show()
    End Sub

    Private Sub McomSite转Planet导入文件_Click(sender As Object, e As EventArgs) Handles McomSite转Planet导入文件.Click
        Dim SiteFileName As String

        MsgBox("请选择McomSite文件！", MsgBoxStyle.Information)
        SiteFileName = selectFile("请选择McomSite文件", "McomSite文件(*.txt)|*.txt", False, True, True)
        If SiteFileName <> "" Then
            BGWSiteToPlanetImport.RunWorkerAsync(SiteFileName)
        End If
    End Sub
    Private Sub mcomSiteToPlnaetImport_dowork(sender As Object, e As DoWorkEventArgs) Handles BGWSiteToPlanetImport.DoWork
        SiteToPlanetFile.startConvertSiteToPlanetFile(sender, e, e.Argument)
    End Sub
    Private Sub mcomSiteToPlnaetImport_progressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BGWSiteToPlanetImport.ProgressChanged
        MainProgressBar.Value = IIf(CInt(e.ProgressPercentage) > 100, 100, CInt(e.ProgressPercentage))
        If InStr(CStr(e.UserState.ToString), "@") Then
            mainMessageBox.AppendText(Split(CStr(e.UserState.ToString), "@")(1) & vbCrLf)
            statusBar.Text = Split(CStr(e.UserState.ToString), "@")(0)
        Else
            statusBar.Text = CStr(e.UserState.ToString)
        End If
    End Sub
    Private Sub mcomSiteToPlnaetImport_runWorkCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BGWSiteToPlanetImport.RunWorkerCompleted
        If e.Cancelled Then
            '            MsgBox("执行完毕！")
            ChangeStatus("欢迎使用！")
            MainProgressBar.Value = 0
            '    fileName = ""
        Else
            BGW_VBACrack.CancelAsync()
            ChangeStatus("欢迎使用！")
            MainProgressBar.Value = 0
            '      fileName = ""
        End If
    End Sub

    Private Sub NeighborPlan_Click(sender As Object, e As EventArgs) Handles NeighborPlan.Click
        'Dim cellInfo As ArrayList
        'cellInfo = loadCarrierFile()
        'MsgBox(cellInfo(0).BCCH)
        'MsgBox(importSiteFileMethod("d:\ab.txt").Count)
        'MsgBox(isAgainstCells(121.14645, 41.12741, 10, 121.15514, 41.13013, 230, 120))
    End Sub

    Private Sub 华为配置转换_Click(sender As Object, e As EventArgs) Handles 华为配置转换.Click
        HW_CDD_PARSE.Show()
    End Sub

    Private Sub CoSiteFreqCheck_Click(sender As Object, e As EventArgs) Handles CoSiteFreqCheck.Click
        '   MsgBox("本功能需要本工具通过CDD转换的Mcom Carrier 和 Mcom Neighbor两个文件，请先准备好！", MsgBoxStyle.Information)
        Dim fileName As String
        Dim outFolder As String
        MsgBox("本功能需要本工具通过CDD转换的Mcom Carrier文件，请先准备好！", MsgBoxStyle.Information)
        fileName = selectFile("选择需要MCOM Carrier文件", "Mcom Carrier文件(*.txt)|*.txt", False, True, True)
        If fileName <> "" Then
            ConsistencyCheck.loadCarrierData(fileName)
            outFolder = Strings.Left(fileName, InStrRev(fileName, "\")) & "FreqCheck\"
            ConsistencyCheck.coSiteFreqCheck(outFolder)
        End If
    End Sub
    Private Sub NeighborCoAdjFreqCheck_Click(sender As Object, e As EventArgs) Handles NeighborCoAdjFreqCheck.Click
        '   MsgBox("本功能需要本工具通过CDD转换的Mcom Carrier 和 Mcom Neighbor两个文件，请先准备好！", MsgBoxStyle.Information)
        Dim carrierFileName As String
        Dim neighborFileName As String
        Dim outFolder As String
        MsgBox("本功能需要本工具通过CDD转换的Mcom Carrier文件和Mcom Neighbor文件，请先准备好！", MsgBoxStyle.Information)
        MsgBox("请先导入Mcom Carrier文件！", MsgBoxStyle.Information)
        carrierFileName = selectFile("选择需要MCOM Carrier文件", "Mcom Carrier文件(*.txt)|*.txt", False, True, True)
        MsgBox("请导入Mcom Neighbor文件！", MsgBoxStyle.Information)
        neighborFileName = selectFile("选择需要MCOM Neighbor文件", "Mcom Neighbor文件(*.txt)|*.txt", False, True, True)
        If carrierFileName <> "" And neighborFileName <> "" Then
            ConsistencyCheck.loadCarrierData(carrierFileName)
            ConsistencyCheck.loadNeighborData(neighborFileName)
            outFolder = Strings.Left(carrierFileName, InStrRev(carrierFileName, "\")) & "FreqCheck\"
            ConsistencyCheck.neighborCellCoBCCHCheck(outFolder)
        End If
    End Sub

    Private Sub mainMessageBox_TextChanged(sender As Object, e As EventArgs) Handles mainMessageBox.TextChanged

    End Sub
End Class

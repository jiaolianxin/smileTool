Imports System.ComponentModel

Public Class DistanceToolDialog
    Private Dis_TotalCacuTime As Integer
    Private site_dic As New Dictionary(Of String, String), HOcount_dic As New Dictionary(Of String, String), SourceCell_arr As ArrayList, Ncell_arr As ArrayList, dual_dic As New Dictionary(Of String, String)
    Private siteFileName As String, CaculFileName As String, minDis As String, maxDis As String
    Private isSiteFileOK As Boolean, isCaculateFileOK As Boolean

    Private Sub countTotalTime() Handles TimerForDistance.Tick
        Dis_TotalCacuTime = Dis_TotalCacuTime + 1
    End Sub
    Private Sub DistanceToolDialog_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MainForm.Visible = False
    End Sub
    Private Sub quit优化小工具() Handles Me.FormClosed
        MainForm.Visible = True
        If BGWForDistance.IsBusy Then
            BGWForDistance.CancelAsync()
            BGWForDistance.Dispose()
        End If
    End Sub

    Private Sub StartForDistance_Click(sender As Object, e As EventArgs) Handles StartForDistance.Click
        'If FromDis.Value = "" Then
        '    minDis = 0
        '    FromDis.Value = 0
        'Else
        '    minDis = FromDis.Value
        'End If
        minDis = FromDis.Value
        maxDis = ToDis.Value

        'If ToDis.Value = "" Then
        '    maxDis = 3000000
        '    ToDis.Value = 3000000
        'Else
        '    maxDis = ToDis.Value
        'End If
        If siteFilePath.Text = "请选择site文件，务必按正确格式（如下）" Then
            MsgBox("请选择完Mcom Site文件之后再开始计算！", MsgBoxStyle.Critical)
            Exit Sub
        Else
            siteFileName = siteFilePath.Text
        End If
        If caculateFilePath.Text = "请选择含有计算内容的文件，务必按正确格式（如下）" Then
            MsgBox("请选择完Caculate文件之后再开始计算！", MsgBoxStyle.Critical)
            Exit Sub
        Else
            CaculFileName = caculateFilePath.Text
        End If
        If Not OneToOne.Checked And Not OneToMulti.Checked Then
            MsgBox("请选择完计算方式之后再开始计算！", MsgBoxStyle.Critical)
            Exit Sub
        End If
        If siteFilePath.Text <> "" And caculateSelect.Text <> "" Then
            If BGWForDistance.IsBusy Then
                MsgBox("当前正在进行距离计算，请稍等！", MsgBoxStyle.Critical)
            Else
                TimerForDistance.Start()
                BGWForDistance.RunWorkerAsync()
            End If
        Else

        End If
    End Sub
    Private Sub Distance_doWork(sender As Object, e As DoWorkEventArgs) Handles BGWForDistance.DoWork
        statusBar.Text = "开始导入数据..."
        importSiteFile(sender, e)
        importCaculateFile(sender, e)

        If isSiteFileOK And isCaculateFileOK Then
            If OneToOne.Checked Then
                OneToOne_Caculate(sender, e, Me)
            Else
                oneToMulti_Caculate(sender, e, Me)
            End If
        Else
            Exit Sub
        End If
        TimerForDistance.Stop()
    End Sub
    Private Sub Distance_changed(sender As Object, e As ProgressChangedEventArgs) Handles BGWForDistance.ProgressChanged
        distanceBar.Value = e.ProgressPercentage
        statusBar.Text = e.UserState.ToString
    End Sub
    Private Sub distance_runWorkCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BGWForDistance.RunWorkerCompleted
        If e.Cancelled Then
            statusBar.Text = "欢迎使用！"
            distanceBar.Value = 0
        Else
            BGWForDistance.CancelAsync()
            statusBar.Text = "欢迎使用！"
            distanceBar.Value = 0
        End If
    End Sub
    Private Sub SiteSelect_Click(sender As Object, e As EventArgs) Handles SiteSelect.Click
        MsgBox("请选择Mcom Site文件!", MsgBoxStyle.Information)
        siteFileName = selectFile("选择Mcom Site文件", "Mcom Site文件(*.txt)|*.txt", False, True, True)
        If siteFileName <> "" Then
            siteFilePath.Text = siteFileName
        End If
    End Sub
    Private Sub caculateSelect_Click(sender As Object, e As EventArgs) Handles caculateSelect.Click
        MsgBox("请选择Caculate文件!", MsgBoxStyle.Information)
        CaculFileName = selectFile("选择Caculate文件", "Caculate文件(*.csv)|*.csv", False, True, True)
        If CaculFileName <> "" Then
            caculateFilePath.Text = CaculFileName
        End If
    End Sub
    Private Sub siteSample_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles siteSample.LinkClicked
        createSiteFile()
    End Sub
    Private Sub caculateSample_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles caculateSample.LinkClicked
        createCaculateFile()
    End Sub

    Private Sub createSiteFile()
        Dim fso As Object
        Dim outputFile As Object
        Dim nowPath As String
        outputFile = ""
        nowPath = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\Mcom Site.txt"
        MsgBox("将在目录" & System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\Mcom Site.txt" & "下生成Mcom Site模板文件！", MsgBoxStyle.Information)

        fso = CreateObject("Scripting.FileSystemObject")
        Try
            outputFile = fso.OpenTextFile(nowPath, 2, 1, 0)
            outputFile.writeLine("Cell" & "	" & "Site" & "	" & "Longitude" & "	" & "Latitude" & "	" & "Dir" & "	" & "Height" & "	" & "Tilt" & "	" & "Ground_Height" & "	" & "Cell_Type" & "	" & "Ant_Type" & "	" & "Ant_BW" & "	" & "Note" & "	" & "SiteName" & "	" & "Flag1" & "	" & "Flag2" & "	" & "Flag3" & "	" & "Flag4" & "	" & "Flag5" & "	" & "CI" & "	" & "BSC" & "	" & "Ant_size")
            outputFile.writeLine("D060555" & "	" & "D06055" & "	" & "119.108529" & "	" & "36.729514" & "	" & "120")
        Catch ex As Exception
            MsgBox("写入文件" & System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\Mcom Site.txt" & vbCrLf & "处于打开状态，请关闭后重试！" & ex.ToString, MsgBoxStyle.Critical)
            outputFile.close()
        Finally
            outputFile.close()
        End Try

        outputFile = Nothing
        fso = Nothing
        Shell("explorer.exe " & nowPath, AppWinStyle.NormalFocus)
    End Sub
    Private Sub createCaculateFile()
        Dim fso As Object
        Dim outputFile As Object
        Dim nowPath As String
        outputFile = ""
        nowPath = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\CaculateFile.csv"
        MsgBox("将在目录" & nowPath & "下生成CaculateFile模板文件！", MsgBoxStyle.Information)

        fso = CreateObject("Scripting.FileSystemObject")
        Try
            outputFile = fso.OpenTextFile(nowPath, 2, 1, 0)
            outputFile.writeLine("Cell" & "," & "Ncell" & "," & "切换申请数（可选）" & "," & "切换成功数（可选）")
            outputFile.writeLine("D060555" & "," & "D060541" & "," & "" & "," & "")
            outputFile.writeLine("D060555" & "," & "D060543" & "," & "" & "," & "")
            outputFile.writeLine("D060555" & "," & "D060542" & "," & "" & "," & "")
        Catch ex As Exception
            MsgBox("写入文件" & nowPath & vbCrLf & "处于打开状态，请关闭后重试！" & ex.ToString, MsgBoxStyle.Critical)
            outputFile.close()
        Finally
            outputFile.close()
        End Try
        outputFile = Nothing
        fso = Nothing
        Shell("explorer.exe " & nowPath, 1)
    End Sub
    Private Sub importSiteFile(sender As Object, e As DoWorkEventArgs)
        Dim site_tempLine As String, site_temp_arr As Object
        Dim d As Integer = 0
        Dim isNoticeDuplicate As Boolean = True
        isSiteFileOK = True
        site_dic.Clear()
        FileClose(1)
        FileOpen(1, siteFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        sender.reportProgress(5, "开始site文件导入")
        Do While Not EOF(1)
            d = d + 1
            site_tempLine = LineInput(1)
            If InStr(site_tempLine, "	") > 0 Then
                site_temp_arr = Split(site_tempLine, "	")
                Try
                    If site_dic.ContainsKey(site_temp_arr(0)) Then
                        If isNoticeDuplicate Then
                            If MsgBox(site_temp_arr(0) & "的经纬度信息重复，将不考虑" & d & "行的经纬度！是否继续提示？", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Warning") = vbNo Then
                                isNoticeDuplicate = False
                            End If
                        End If
                    Else
                        If site_temp_arr(2) <> "" And site_temp_arr(3) <> "" Then
                            site_dic.Add(site_temp_arr(0), site_temp_arr(2) & "#" & site_temp_arr(3))
                        ElseIf site_temp_arr(1) = "" And site_temp_arr(2) <> "" Then
                            site_dic.Add(site_temp_arr(0), 0 & "#" & site_temp_arr(3))
                        ElseIf site_temp_arr(1) <> "" And site_temp_arr(2) = "" Then
                            site_dic.Add(site_temp_arr(0), site_temp_arr(2) & "#" & 0)
                        End If
                    End If
                Catch ex As Exception
                    FileClose(1)
                    MsgBox("Mcom Site文件格式不正确，请严格按模板填写！请检查" & d & "行的数据！", MsgBoxStyle.Critical)
                    isSiteFileOK = False
                    site_dic.Clear()
                    Exit Sub
                End Try
            Else
                If MsgBox(site_tempLine & "的格式与模板不匹配或无经纬度信息，是否继续？", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Warning") = vbYes Then
                    site_dic.Add(site_tempLine, 0 & "#" & 0)
                Else
                    isSiteFileOK = False
                    FileClose(1)
                    site_dic.Clear()
                    Exit Sub
                End If
            End If
        Loop
        FileClose(1)
    End Sub
    Private Sub importCaculateFile(sender As Object, e As DoWorkEventArgs)
        Dim caculate_tempLine As String, caculate_temp_arr As Object
        Dim nowLineNo As Integer = 0
        SourceCell_arr = New ArrayList
        Ncell_arr = New ArrayList
        isCaculateFileOK = True
        HOcount_dic.Clear()
        FileClose(1)
        FileOpen(1, CaculFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        sender.reportProgress(10, "开始Caculate文件导入")
        Do While Not EOF(1)
            nowLineNo = nowLineNo + 1
            caculate_tempLine = LineInput(1)
            If InStr(caculate_tempLine, ",") Then
                Try
                    caculate_temp_arr = Split(caculate_tempLine, ",")
                    SourceCell_arr.Add(caculate_temp_arr(0))
                    Ncell_arr.Add(caculate_temp_arr(1))
                    If Not HOcount_dic.ContainsKey(caculate_temp_arr(0) & "#" & caculate_temp_arr(1)) Then
                        If UBound(caculate_temp_arr) = 1 Then
                            HOcount_dic.Add(caculate_temp_arr(0) & "#" & caculate_temp_arr(1), "" & "," & "")
                        ElseIf UBound(caculate_temp_arr) = 2 Then
                            HOcount_dic.Add(caculate_temp_arr(0) & "#" & caculate_temp_arr(1), caculate_temp_arr(2) & "," & "")
                        ElseIf UBound(caculate_temp_arr) = 3 Then
                            HOcount_dic.Add(caculate_temp_arr(0) & "#" & caculate_temp_arr(1), caculate_temp_arr(2) & "," & caculate_temp_arr(3))
                        End If
                    End If
                Catch ex As Exception
                    FileClose(1)
                    MsgBox("Caculate文件中第" & nowLineNo & "行的格式不正确，请严格按模板填写！" & ex.ToString, MsgBoxStyle.Critical)
                    SourceCell_arr = Nothing
                    Ncell_arr = Nothing
                    isCaculateFileOK = False
                    Exit Sub
                End Try
            Else
                If MsgBox(caculate_tempLine & "是否为正确的源小区?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Warning") = vbYes Then
                    SourceCell_arr.Add(caculate_tempLine)
                    Ncell_arr.Add("NONE")
                Else
                    isCaculateFileOK = False
                    FileClose(1)
                    SourceCell_arr = Nothing
                    Ncell_arr = Nothing
                    Exit Sub
                End If
            End If
        Loop
        FileClose(1)
        sender.reportProgress(100, "导入完成！")

    End Sub
    Private Sub OneToOne_Caculate(sender As Object, e As DoWorkEventArgs, distanceDialog As DistanceToolDialog)
        Dim fso As Object
        Dim outPutFile As Object
        Dim outFilePath As String
        Dim distanceResult As Double, lon1 As Double, lat1 As Double, lon2 As Double, lat2 As Double
        Dim falseCell As String, NoticeStatue As Boolean
        Dim y As Integer
        falseCell = ""
        NoticeStatue = False
        fso = CreateObject("Scripting.FileSystemObject")
        sender.reportProgress(15, "开始计算距离...")
        outFilePath = Strings.Left(siteFileName, InStrRev(siteFileName, "\")) & "DistanceResults" & Format(Now, "yyMMdd") & ".csv"
        Try
            outPutFile = fso.OpenTextFile(outFilePath, 2, 1, 0)
        Catch ex As Exception
            MsgBox("导出的文件处于打开状态，请先关闭该文件后，重新执行距离计算！", MsgBoxStyle.Critical)
            SourceCell_arr.Clear()
            Ncell_arr.Clear()
            site_dic.Clear()
            Exit Sub
        End Try
        outPutFile.writeLine("S_cell" & "," & "N_cell" & "," & "S_lon" & "," & "S_lat" & "," & "N_lon" & "," & "N_lat" & "," & "切换申请数" & "," & "切换成功数" & "," & "Distance(M)")
        Try
            For y = 1 To SourceCell_arr.Count - 1
                If site_dic.ContainsKey(SourceCell_arr(y)) Then
                    lon1 = Split(site_dic(SourceCell_arr(y)), "#")(0)
                    lat1 = Split(site_dic(SourceCell_arr(y)), "#")(1)
                    If math.abs(lat1) > 90 Then
                        MsgBox("源小区：" & SourceCell_arr(y) & "的纬度异常！", MsgBoxStyle.Critical)
                        outPutFile.close()
                        site_dic.Clear()
                        Exit Sub
                    End If
                    If site_dic.ContainsKey(Ncell_arr(y)) Then
                        lon2 = Split(site_dic(Ncell_arr(y)), "#")(0)
                        lat2 = Split(site_dic(Ncell_arr(y)), "#")(1)
                        sender.reportProgress(Mid(15 + y / (SourceCell_arr.Count - 1) * 85, 1, 2), "正在计算距离..." & Mid(15 + y / (SourceCell_arr.Count - 1) * 85, 1, 2) & "%")
                        If math.abs(lat2) > 90 Then
                            MsgBox("目标小区：" & Ncell_arr(y) & "的纬度异常！", MsgBoxStyle.Critical)
                            outPutFile.close()
                            site_dic.Clear()
                            Exit Sub
                        End If
                        If lon1.ToString = "" Or lat1.ToString = "" Or lon2.ToString = "" Or lat2.ToString = "" Then
                            outPutFile.writeLine(SourceCell_arr(y) & "," & Ncell_arr(y) & "," & lon1 & "," & lat1 & "," & lon2 & "," & lat2 & "," & HOcount_dic(SourceCell_arr(y) & "#" & Ncell_arr(y)) & "," & "经纬度信息不全")
                        Else
                            If Math.abs(lon1) > 0 And Math.abs(lat1) > 0 Then
                                If Math.abs(lon2) > 0 And Math.abs(lat2) > 0 Then
                                    If lon1 = lon2 And lat1 = lat2 Then
                                        If SourceCell_arr(y) <> Ncell_arr(y) Then
                                            If distanceResult >= CInt(minDis) And distanceResult <= CInt(maxDis) Then
                                                outPutFile.writeLine(SourceCell_arr(y) & "," & Ncell_arr(y) & "," & lon1 & "," & lat1 & "," & lon2 & "," & lat2 & "," & HOcount_dic(SourceCell_arr(y) & "#" & Ncell_arr(y)) & "," & 0)
                                            End If
                                        End If
                                    Else
                                        distanceResult = 1000 * 111.12 * Math.Acos((Math.Sin(lat1 * 3.14159265358979 / 180) * Math.Sin(lat2 * 3.14159265358979 / 180) + Math.Cos(lat1 * 3.14159265358979 / 180) * Math.Cos(lat2 * 3.14159265358979 / 180) * Math.Cos((lon2 - lon1) * 3.14159265358979 / 180))) * 180 / 3.14159265358979
                                        If distanceResult >= CInt(minDis) And distanceResult <= CInt(maxDis) Then
                                            outPutFile.writeLine(SourceCell_arr(y) & "," & Ncell_arr(y) & "," & lon1 & "," & lat1 & "," & lon2 & "," & lat2 & "," & HOcount_dic(SourceCell_arr(y) & "#" & Ncell_arr(y)) & "," & distanceResult)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        outPutFile.writeLine(SourceCell_arr(y) & "," & Ncell_arr(y) & "," & lon1 & "," & lat1 & "," & 0 & "," & 0 & "," & HOcount_dic(SourceCell_arr(y) & "#" & Ncell_arr(y)) & "," & "目标小区无经纬度信息！")
                    End If
                Else
                    outPutFile.writeLine(SourceCell_arr(y) & "," & Ncell_arr(y) & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & HOcount_dic(SourceCell_arr(y) & "#" & Ncell_arr(y)) & "," & "源小区无经纬度信息！")
                End If
            Next
        Catch ex As Exception
            MsgBox(SourceCell_arr(y) & "或" & Ncell_arr(y) & "经纬度异常，请检查！" & ex.ToString(), MsgBoxStyle.Critical)
            Exit Sub
        Finally
            FileClose(1)
            outPutFile.close()
            fso = Nothing
            HOcount_dic.Clear()
            site_dic.Clear()
            SourceCell_arr.Clear()
            Ncell_arr.Clear()
            dual_dic.Clear()
        End Try
        sender.reportProgress(100, "距离计算完毕！")
        MsgBox("计算完毕！共耗时：" & distanceDialog.Dis_TotalCacuTime & "秒！", MsgBoxStyle.Information)
        Shell("explorer.exe " & outFilePath, 1)
    End Sub
    Private Sub oneToMulti_Caculate(sender As Object, e As DoWorkEventArgs, distanceDialog As DistanceToolDialog)
        Dim fso As Object
        Dim outPutFile As Object
        Dim outFilePath As String
        Dim distanceResult As Double, lon1 As Double, lat1 As Double, lon2 As Double, lat2 As Double
        Dim falseNCell As String, NoticeNcellStatue As Boolean, Response As Integer, falseSCell As String, NoticeScellStatue As Boolean
        Dim noticeNcellStatue_arr As New ArrayList
        Dim noticeScellStatue_arr As New ArrayList
        Dim y As Integer, j As Integer
        Dim nowCellMinDis As Integer = 999999999
        Dim minDisLine As String = ""
        falseNCell = ""
        falseSCell = ""
        NoticeScellStatue = False
        NoticeNcellStatue = False

        fso = CreateObject("Scripting.FileSystemObject")
        sender.reportProgress(15, "开始计算距离...")
        outFilePath = Strings.Left(siteFileName, InStrRev(siteFileName, "\")) & "DistanceResults" & Format(Now, "yyMMdd") & ".csv"
        Try
            outPutFile = fso.OpenTextFile(outFilePath, 2, 1, 0)
        Catch ex As Exception
            MsgBox("导出的文件处于打开状态，请先关闭该文件后，重新执行距离计算！", MsgBoxStyle.Critical)
            site_dic.Clear()
            SourceCell_arr.Clear()
            Ncell_arr.Clear()
            Exit Sub
        End Try
        outPutFile.writeLine("S_cell" & "," & "N_cell" & "," & "S_lon" & "," & "S_lat" & "," & "N_lon" & "," & "N_lat" & "," & "Distance(M)")
        Try
            For y = 1 To SourceCell_arr.Count - 1
                sender.reportProgress(Mid(15 + y / (SourceCell_arr.Count - 1) * 85, 1, 2), "正在计算距离..." & Mid(15 + y / (SourceCell_arr.Count - 1) * 85, 1, 2) & "%")
                If site_dic.ContainsKey(SourceCell_arr(y)) Then
                    If y > 2 And distanceDialog.isCacuMinDis.Checked Then
                        outPutFile.writeLine(minDisLine)
                    End If
                    nowCellMinDis = 999999999
                    minDisLine = ""
                    lon1 = CDbl(Split(site_dic(SourceCell_arr(y)), "#")(0))
                    lat1 = CDbl(Split(site_dic(SourceCell_arr(y)), "#")(1))
                    If math.abs(lat1) > 90 Then
                        MsgBox("源小区：" & SourceCell_arr(y) & "的纬度异常！", MsgBoxStyle.Critical)
                        Exit Sub
                    End If
                    For j = 1 To Ncell_arr.Count - 1
                        '                   sender.reportProgress((15 + ((y * j) / ((SourceCell_arr.Count - 1) * (Ncell_arr.Count - 1))) * 85), "正在计算" & SourceCell_arr(y) & "-" & Ncell_arr(j) & "的距离！")
                        If site_dic.ContainsKey(Ncell_arr(j)) Then
                            lon2 = CDbl(Split(site_dic(Ncell_arr(j)), "#")(0))
                            lat2 = CDbl(Split(site_dic(Ncell_arr(j)), "#")(1))
                            If math.abs(lat2) > 90 Then
                                MsgBox("目标小区：" & Ncell_arr(j) & "的纬度异常！", MsgBoxStyle.Critical)
                                Exit Sub
                            End If
                            If Math.abs(lon1) > 0 And Math.abs(lat1) > 0 Then
                                If Math.abs(lon2) > 0 And Math.abs(lat2) > 0 Then
                                    If lon1 = lon2 And lat1 = lat2 Then
                                        If distanceDialog.isCacuMinDis.Checked Then
                                            If SourceCell_arr(y) <> Ncell_arr(j) And Strings.Left(SourceCell_arr(y), Len(SourceCell_arr(y)) - 1) <> Strings.Left(Ncell_arr(j), Len(Ncell_arr(j)) - 1) Then
                                                nowCellMinDis = 0
                                                minDisLine = SourceCell_arr(y) & "," & Ncell_arr(j) & "," & lon1 & "," & lat1 & "," & lon2 & "," & lat2 & "," & 0
                                                dual_dic.Add(SourceCell_arr(y) & "#" & Ncell_arr(j), "OK")
                                            End If
                                        Else
                                            If Not dual_dic.ContainsKey(SourceCell_arr(y) & "#" & Ncell_arr(j)) And Not dual_dic.ContainsKey(Ncell_arr(j) & "#" & SourceCell_arr(y)) Then
                                                If SourceCell_arr(y) <> Ncell_arr(j) Then
                                                    If distanceDialog.isAllowDuplicate.Checked Then
                                                        outPutFile.writeLine(SourceCell_arr(y) & "," & Ncell_arr(j) & "," & lon1 & "," & lat1 & "," & lon2 & "," & lat2 & "," & 0)
                                                    Else
                                                        If distanceDialog.isShowCoSite.Checked And Strings.Left(SourceCell_arr(y), Len(SourceCell_arr(y)) - 1) = Strings.Left(Ncell_arr(j), Len(Ncell_arr(j)) - 1) Then
                                                            outPutFile.writeLine(SourceCell_arr(y) & "," & Ncell_arr(j) & "," & lon1 & "," & lat1 & "," & lon2 & "," & lat2 & "," & 0)
                                                        Else
                                                            outPutFile.writeLine(SourceCell_arr(y) & "," & Ncell_arr(j) & "," & lon1 & "," & lat1 & "," & lon2 & "," & lat2 & "," & 0)
                                                            dual_dic.Add(SourceCell_arr(y) & "#" & Ncell_arr(j), "OK")
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If Not dual_dic.ContainsKey(SourceCell_arr(y) & "#" & Ncell_arr(j)) And Not dual_dic.ContainsKey(Ncell_arr(j) & "#" & SourceCell_arr(y)) Then
                                            distanceResult = 1000 * 111.12 * Math.Acos((Math.Sin(lat1 * 3.14159265358979 / 180) * Math.Sin(lat2 * 3.14159265358979 / 180) + Math.Cos(lat1 * 3.14159265358979 / 180) * Math.Cos(lat2 * 3.14159265358979 / 180) * Math.Cos((lon2 - lon1) * 3.14159265358979 / 180))) * 180 / 3.14159265358979
                                            If distanceDialog.isCacuMinDis.Checked Then
                                                If distanceResult >= CInt(minDis) And distanceResult <= CInt(maxDis) Then
                                                    If distanceResult < nowCellMinDis Then
                                                        nowCellMinDis = distanceResult
                                                        minDisLine = SourceCell_arr(y) & "," & Ncell_arr(j) & "," & lon1 & "," & lat1 & "," & lon2 & "," & lat2 & "," & nowCellMinDis
                                                        dual_dic.Add(SourceCell_arr(y) & "#" & Ncell_arr(j), "OK")
                                                    End If
                                                End If
                                            Else
                                                If distanceResult >= CInt(minDis) And distanceResult <= CInt(maxDis) Then
                                                    If distanceDialog.isAllowDuplicate.Checked Then
                                                        outPutFile.writeLine(SourceCell_arr(y) & "," & Ncell_arr(j) & "," & lon1 & "," & lat1 & "," & lon2 & "," & lat2 & "," & distanceResult)
                                                    Else
                                                        If distanceDialog.isShowCoSite.Checked And Strings.Left(SourceCell_arr(y), Len(SourceCell_arr(y)) - 1) = Strings.Left(Ncell_arr(j), Len(Ncell_arr(j)) - 1) Then
                                                            outPutFile.writeLine(SourceCell_arr(y) & "," & Ncell_arr(j) & "," & lon1 & "," & lat1 & "," & lon2 & "," & lat2 & "," & distanceResult)
                                                        Else
                                                            outPutFile.writeLine(SourceCell_arr(y) & "," & Ncell_arr(j) & "," & lon1 & "," & lat1 & "," & lon2 & "," & lat2 & "," & distanceResult)
                                                            dual_dic.Add(SourceCell_arr(y) & "#" & Ncell_arr(j), "OK")
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            If falseNCell = Ncell_arr(j) Then
                            Else
                                If NoticeNcellStatue And noticeNcellStatue_arr.Contains(Ncell_arr(j)) Then
                                Else
                                    Response = MsgBox("目标小区:" & Ncell_arr(j) & "无经纬度信息，是否继续?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Critical, "问题提示")
                                    If Response = vbNo Then
                                        site_dic.Clear()
                                        outPutFile.close()
                                        Exit Sub
                                    Else
                                        falseNCell = Ncell_arr(j)
                                        NoticeNcellStatue = True
                                        noticeNcellStatue_arr.Add(Ncell_arr(j))
                                    End If
                                End If
                            End If
                        End If
                    Next
                Else
                    If falseSCell = SourceCell_arr(y) Then
                    Else
                        If NoticeScellStatue And noticeScellStatue_arr.Contains(SourceCell_arr(y)) Then

                        Else
                            Response = MsgBox("源小区:" & SourceCell_arr(y) & "无经纬度信息，是否继续?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Critical, "问题提示")
                            If Response = vbNo Then
                                site_dic.Clear()
                                outPutFile.close()
                                Exit Sub
                            Else
                                falseSCell = SourceCell_arr(y)
                                NoticeScellStatue = True
                                noticeScellStatue_arr.Add(SourceCell_arr(y))
                            End If
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(SourceCell_arr(y) & "或" & Ncell_arr(j) & "经纬度异常，请检查！" & ex.ToString, MsgBoxStyle.Critical)
        Finally
            FileClose(1)
            outPutFile.close()
            HOcount_dic.Clear()
            site_dic.Clear()
            SourceCell_arr.Clear()
            Ncell_arr.Clear()
            dual_dic.Clear()
            noticeNcellStatue_arr = Nothing
            noticeScellStatue_arr = Nothing
            fso = Nothing
        End Try
        sender.reportProgress(100, "距离计算完毕！")
        MsgBox("计算完毕！共耗时：" & distanceDialog.Dis_TotalCacuTime & "秒！", MsgBoxStyle.Information)
        Shell("explorer.exe " & outFilePath, 1)
    End Sub

End Class
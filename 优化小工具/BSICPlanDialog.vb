Imports System.Reflection
Imports System.ComponentModel

Public Class BSICPlanDialog
    Private site_position_arr As New Dictionary(Of String, String)  '(cell,lon#lat#dir)
    Private carrier_cell_arr As New ArrayList, BCCH_arr As New ArrayList, BSIC_arr As New ArrayList, TCH_arr As New ArrayList
    Private BSC_arr As New ArrayList, LAI_arr As New ArrayList, CI_arr As New ArrayList, HOP_arr As New ArrayList, HSN_arr As New ArrayList, FONT_SIZE_arr As New ArrayList
    Private OLD_BSIC_arr As New ArrayList, NEW_BSIC_arr As New ArrayList
    Private planCell_arr As New ArrayList
    Private allBSIC_ARR As New ArrayList
    Private total_calcu_time As Integer
    Private Sub doAtClose() Handles Me.FormClosing
        MainForm.Visible = True
        site_position_arr.Clear()
        carrier_cell_arr.Clear()
        BSC_arr.Clear()
        LAI_arr.Clear()
        CI_arr.Clear()
        BCCH_arr.Clear()
        BSIC_arr.Clear()
        TCH_arr.Clear()
        HOP_arr.Clear()
        HSN_arr.Clear()
        FONT_SIZE_arr.Clear()
        OLD_BSIC_arr.Clear()
        NEW_BSIC_arr.Clear()
        allBSIC_ARR.Clear()
        planCell_arr.Clear()
    End Sub
    Private Sub recordTime() Handles BSICPlanTimer.Tick
        total_calcu_time = total_calcu_time + 1
    End Sub
    Private Sub doAtLoad() Handles MyBase.Load
        MainForm.Visible = False
    End Sub
    Private Sub selectSiteFileButton_Click(sender As Object, e As EventArgs) Handles selectSiteFileButton.Click
        Dim siteFileName As String
        MsgBox("请选择Mcom Site文件!", MsgBoxStyle.Information)
        siteFileName = selectFile("选择Mcom Site文件", "Mcom Site文件(*.txt)|*.txt", False, True, True)
        If siteFileName <> "" Then
            siteFilePath.Text = siteFileName
            importSiteData()
        End If
    End Sub
    Private Sub importSiteData()
        Dim fileNo As Integer
        Dim temp_arr As Object
        Dim tempLine As String
        Dim lineNo As Integer = 0
        site_position_arr.Clear()
        fileNo = FreeFile()
        Try
            FileOpen(fileNo, siteFilePath.Text, OpenMode.Input, OpenAccess.Read, OpenShare.Default)
        Catch ex As Exception
            MsgBox("Mcom site文件导入失败！请检查导入文件是否处于打开状态！", MsgBoxStyle.Critical)
            Exit Sub
        End Try
        Try
            Do While Not EOF(fileNo)
                tempLine = LineInput(fileNo)
                lineNo = lineNo + 1
                If InStr(tempLine, "	") <> 0 Then
                    temp_arr = Split(tempLine, "	")
                    If lineNo = 1 Then
                        If UCase(temp_arr(0)) <> "CELL" Or Strings.Left(UCase(temp_arr(2)), 3) <> "LON" Or Strings.Left(UCase(temp_arr(3)), 3) <> "LAT" Or UCase(temp_arr(4)) <> "DIR" Then
                            MsgBox("Mcom site文件格式有误！请检查第一行数据，表头是否为标准的Site文件表头！", MsgBoxStyle.Critical)
                            siteFilePath.Text = "请选择Mcom Site文件!"
                            Exit Sub
                        End If
                    Else
                        Try
                            If Not site_position_arr.ContainsKey(temp_arr(0)) Then
                                site_position_arr.Add(temp_arr(0), temp_arr(2) & "#" & temp_arr(3) & "#" & temp_arr(4))
                            Else
                                MsgBox("Mcom Site文件中小区：" & temp_arr(0) & "存在重复信息！第" & lineNo & "行的数据，将被忽略！", MsgBoxStyle.Critical)
                            End If
                        Catch ex As Exception
                            MsgBox("Mcom site文件格式有误！请检查第" & lineNo & "行的数据！", MsgBoxStyle.Critical)
                            siteFilePath.Text = "请选择Mcom Site文件!"
                            Exit Sub
                        End Try
                    End If
                Else
                    MsgBox("Mcom site文件格式有误！请检查第" & lineNo & "行的数据！", MsgBoxStyle.Critical)
                    siteFilePath.Text = "请选择Mcom Site文件!"
                    Exit Sub
                End If
            Loop
        Catch ex As Exception
            MsgBox("Mcom site文件导入失败！请检查！", MsgBoxStyle.Critical)
            site_position_arr.Clear()
            siteFilePath.Text = "请选择Mcom Site文件!"
            Exit Sub
        Finally
            FileClose(fileNo)
        End Try
        MsgBox("Mcom Site文件读取完毕！", MsgBoxStyle.Information)
    End Sub
    Private Sub selectCarrierButton_Click(sender As Object, e As EventArgs) Handles selectCarrierButton.Click
        Dim carrierFileName As String
        MsgBox("请选择Mcom Carrier文件!", MsgBoxStyle.Information)
        carrierFileName = selectFile("选择Mcom Carrier文件", "Mcom Carrier文件(*.txt)|*.txt", False, True, True)
        If carrierFileName <> "" Then
            carrierFilePath.Text = carrierFileName
            importCarrierData()
        End If
    End Sub
    Private Sub importCarrierData()
        Dim fileNo As Integer
        Dim temp_arr As Object
        Dim tempLine As String
        Dim lineNo As Integer = 0
        Dim duplicate_record As String = ""
        Dim total_duplicate_no As Integer
        carrier_cell_arr.Clear()
        BSC_arr.Clear()
        LAI_arr.Clear()
        CI_arr.Clear()
        BCCH_arr.Clear()
        BSIC_arr.Clear()
        TCH_arr.Clear()
        HOP_arr.Clear()
        HSN_arr.Clear()
        FONT_SIZE_arr.Clear()
        OLD_BSIC_arr.Clear()
        NEW_BSIC_arr.Clear()
        fileNo = FreeFile()
        Try
            FileOpen(fileNo, carrierFilePath.Text, OpenMode.Input, OpenAccess.Read, OpenShare.Default)
        Catch ex As Exception
            MsgBox("Mcom Carrier文件导入失败！请检查导入文件是否处于打开状态！", MsgBoxStyle.Critical)
            Exit Sub
        End Try
        Try
            Do While Not EOF(fileNo)
                tempLine = LineInput(fileNo)
                lineNo = lineNo + 1
                If InStr(tempLine, "	") <> 0 Then
                    temp_arr = Split(tempLine, "	")
                    If lineNo = 1 Then
                        If UCase(temp_arr(0)) <> "CELL" Or UCase(temp_arr(4)) <> "BCCH" Or UCase(temp_arr(5)) <> "BSIC" Or UCase(temp_arr(6)) <> "TCH" Then
                            MsgBox("Mcom Carrier文件格式有误！请检查！", MsgBoxStyle.Critical)
                            carrierFilePath.Text = "请选择Mcom Carrier文件!"
                            Exit Sub
                        Else
                            carrier_cell_arr.Add(temp_arr(0))
                            BSC_arr.Add(temp_arr(1))
                            LAI_arr.Add(temp_arr(2))
                            CI_arr.Add(temp_arr(3))
                            BCCH_arr.Add(temp_arr(4))
                            If Len(Trim(temp_arr(5))) <= 1 Then
                                BSIC_arr.Add("0" & Trim(temp_arr(5)))
                            Else
                                BSIC_arr.Add(temp_arr(5))
                            End If
                            TCH_arr.Add(temp_arr(6))
                            HOP_arr.Add(temp_arr(7))
                            HSN_arr.Add(temp_arr(8))
                            FONT_SIZE_arr.Add(temp_arr(9))
                            OLD_BSIC_arr.Add("Original_BSIC")
                            '                         NEW_BSIC_arr.Add("Planed_BSIC")
                        End If
                    Else
                        Try
                            If Not carrier_cell_arr.Contains(temp_arr(0)) Then
                                carrier_cell_arr.Add(temp_arr(0))
                                BSC_arr.Add(temp_arr(1))
                                LAI_arr.Add(temp_arr(2))
                                CI_arr.Add(temp_arr(3))
                                BCCH_arr.Add(temp_arr(4))
                                If Len(Trim(temp_arr(5))) <= 1 Then
                                    BSIC_arr.Add("0" & Trim(temp_arr(5)))
                                Else
                                    BSIC_arr.Add(temp_arr(5))
                                End If
                                TCH_arr.Add(temp_arr(6))
                                HOP_arr.Add(temp_arr(7))
                                HSN_arr.Add(temp_arr(8))
                                FONT_SIZE_arr.Add(temp_arr(9))
                                If UBound(temp_arr) >= 10 Then
                                    If temp_arr(10) <> "" Then
                                        OLD_BSIC_arr.Add(temp_arr(10))
                                    Else
                                        OLD_BSIC_arr.Add("")
                                    End If
                                Else
                                    OLD_BSIC_arr.Add("")
                                End If
                                '    NEW_BSIC_arr.Add("")
                            Else
                                duplicate_record = Trim(duplicate_record & "Mcom Carrier文件中小区：" & temp_arr(0) & "存在重复信息！第" & lineNo & "行的数据，已被忽略！" & vbCrLf)
                                ' MsgBox("Mcom Carrier文件中小区：" & temp_arr(0) & "存在重复信息！第" & lineNo & "行的数据，将被忽略！", MsgBoxStyle.Critical)
                            End If
                        Catch ex As Exception
                            MsgBox("Mcom Carrier文件格式有误！请检查第" & lineNo & "行的数据！", MsgBoxStyle.Critical)
                            carrierFilePath.Text = "请选择Mcom Carrier文件!"
                            Exit Sub
                        End Try
                    End If
                Else
                    MsgBox("Mcom Carrier文件导入失败！请检查！", MsgBoxStyle.Critical)
                    carrierFilePath.Text = "请选择Mcom Carrier文件!"
                    Exit Sub
                End If
            Loop
        Catch ex As Exception
            MsgBox("Mcom Carrier文件导入失败！请检查！", MsgBoxStyle.Critical)
            carrierFilePath.Text = "请选择Mcom Carrier文件!"
            carrier_cell_arr.Clear()
            BSC_arr.Clear()
            LAI_arr.Clear()
            CI_arr.Clear()
            BCCH_arr.Clear()
            BSIC_arr.Clear()
            TCH_arr.Clear()
            HOP_arr.Clear()
            HSN_arr.Clear()
            FONT_SIZE_arr.Clear()
            OLD_BSIC_arr.Clear()
            NEW_BSIC_arr.Clear()
            Exit Sub
        Finally
            FileClose(fileNo)
        End Try
        If duplicate_record <> "" Then
            total_duplicate_no = UBound(Split(duplicate_record, vbCrLf)) + 1
            If UBound(Split(duplicate_record, vbCrLf)) > 8 Then
                duplicate_record = Split(duplicate_record, vbCrLf)(0) & vbCrLf & Split(duplicate_record, vbCrLf)(1) & vbCrLf & Split(duplicate_record, vbCrLf)(2) & vbCrLf & Split(duplicate_record, vbCrLf)(3) & vbCrLf & Split(duplicate_record, vbCrLf)(4) & vbCrLf & Split(duplicate_record, vbCrLf)(5) & vbCrLf & Split(duplicate_record, vbCrLf)(6) & vbCrLf & Split(duplicate_record, vbCrLf)(7) & vbCrLf & Split(duplicate_record, vbCrLf)(8) & vbCrLf
                MsgBox("重复项仅使用第一个出现的值！" & vbCrLf & duplicate_record & vbCrLf & "...还有" & total_duplicate_no - 9 & "行的重复项，已被忽略！", MsgBoxStyle.Critical)
            Else
                MsgBox("重复项仅使用第一个出现的值！" & vbCrLf & duplicate_record, MsgBoxStyle.Critical)
            End If
        End If
        MsgBox("Mcom Carrier文件读取完毕！", MsgBoxStyle.Information)
    End Sub
    Private Sub selectPlanCellButton_Click(sender As Object, e As EventArgs) Handles selectPlanCellButton.Click
        Dim planCellFileName As String
        MsgBox("请选择BSIC规划小区列表文件!", MsgBoxStyle.Information)
        planCellFileName = selectFile("选择BSIC规划小区列表文件", "BSIC规划小区列表文件(*.txt)|*.txt", False, True, True)
        If planCellFileName <> "" Then
            planCellFilePath.Text = planCellFileName
            importPlanCellData()
        End If
    End Sub
    Private Sub importPlanCellData()
        Dim fileNo As Integer
        Dim tempLine As String
        Dim lineNo As Integer = 0
        planCell_arr.Clear()
        fileNo = FreeFile()
        Try
            FileOpen(fileNo, planCellFilePath.Text, OpenMode.Input, OpenAccess.Read, OpenShare.Default)
        Catch ex As Exception
            MsgBox("BSIC规划小区列表文件导入失败！请检查导入文件是否处于打开状态！", MsgBoxStyle.Critical)
            Exit Sub
        End Try
        Try
            Do While Not EOF(fileNo)
                tempLine = LineInput(fileNo)
                lineNo = lineNo + 1
                Try
                    If InStr(Trim(tempLine), "	") > 0 Or InStr(Trim(tempLine), ",") > 0 Or InStr(Trim(tempLine), "，") > 0 Then
                        MsgBox("BSIC规划小区列表文件中第" & lineNo & "行数据格式有误,请按如下格式：" & vbCrLf & "CELL" & vbCrLf & "G05221A" & vbCrLf & "G05221B", MsgBoxStyle.Critical)
                        planCellFilePath.Text = "请选择BSIC规划小区列表文件!"
                        Exit Sub
                    Else
                        If Not planCell_arr.Contains(Trim(tempLine)) Then
                            planCell_arr.Add(Trim(tempLine))
                        Else
                            MsgBox("BSIC规划小区列表文件中小区：" & Trim(tempLine) & "存在重复信息！第" & lineNo & "行的数据，将被忽略！", MsgBoxStyle.Critical)
                        End If
                    End If
                Catch ex As Exception
                    MsgBox("BSIC规划小区列表文件格式有误！请检查！", MsgBoxStyle.Critical)
                    planCellFilePath.Text = "请选择BSIC规划小区列表文件!"
                    Exit Sub
                End Try
            Loop
        Catch ex As Exception
            MsgBox("BSIC规划小区列表文件导入失败！请检查！", MsgBoxStyle.Critical)
            planCellFilePath.Text = "请选择BSIC规划小区列表文件!"
            planCell_arr.clear()
            Exit Sub
        Finally
            FileClose(fileNo)
        End Try
        MsgBox("BSIC规划小区列表读取完毕！", MsgBoxStyle.Information)
    End Sub
    Private Sub StartPlan_Click(sender As Object, e As EventArgs) Handles StartPlan.Click
        Dim t As Type = Me.GetType
        allBSIC_ARR.Clear()
        For x = 0 To 7
            Dim NCCbox As FieldInfo = t.GetField("_" & "NCC" & x, BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.Public)
            Dim NCC As CheckBox = CType(NCCbox.GetValue(Me), CheckBox)
            If NCC.Checked Then
                For y = 0 To 7
                    Dim BCCbox As FieldInfo = t.GetField("_" & "BCC" & y, BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.Public)
                    Dim BCC As CheckBox = CType(BCCbox.GetValue(Me), CheckBox)
                    If BCC.Checked Then
                        If Not allBSIC_ARR.Contains(x & y) Then
                            allBSIC_ARR.Add(x & y)
                        End If
                    End If
                Next
            End If
        Next
        'For p = 31 To 167
        '    allBSIC_ARR.Add(CStr(p))
        'Next

        If allBSIC_ARR.Count = 0 Then
            MsgBox("无可用的BSIC，请选择规划时允许使用的NCC,BCC！", MsgBoxStyle.Critical)
            Exit Sub
        End If
        If siteFilePath.Text = "请选择Mcom Site文件!" Then
            MsgBox("请选择Mcom Site文件!", MsgBoxStyle.Critical)
            Exit Sub
        ElseIf site_position_arr.Count = 0 And siteFilePath.Text <> "请选择Mcom Site文件!" Then
            importSiteData()
        End If
        If carrierFilePath.Text = "请选择Mcom Carrier文件!" Then
            MsgBox("请选择Mcom Carrier文件!", MsgBoxStyle.Critical)
            Exit Sub
        ElseIf carrier_cell_arr.Count = 0 And carrierFilePath.Text <> "请选择Mcom Carrier文件!" Then
            importCarrierData()
        End If
        If planCellFilePath.Text = "请选择BSIC规划小区列表文件!" Then
            MsgBox("请选择BSIC规划小区列表文件!", MsgBoxStyle.Critical)
            Exit Sub
        ElseIf planCell_arr.Count = 0 And planCellFilePath.Text <> "请选择BSIC规划小区列表文件!" Then
            importPlanCellData()
        End If
        If Not BGW_BSICPlan.IsBusy() Then
            BSICPlanTimer.Start()
            total_calcu_time = 1
            BGW_BSICPlan.RunWorkerAsync()
        Else
            MsgBox("正在进行BSIC规划，请稍候再试!", MsgBoxStyle.Critical)
        End If
    End Sub
    Private Sub bsicPlan_dowork(sender As Object, e As DoWorkEventArgs) Handles BGW_BSICPlan.DoWork
        Try
            BSICPlan(sender, e)
        Catch ex As Exception
            MsgBox("BSIC规划未成功，请自已检查数据或反馈作者！" & vbCrLf & ex.ToString, MsgBoxStyle.Critical)
        Finally
            BSICPlanTimer.Stop()
        End Try
    End Sub
    Private Sub bsicPlan_progressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BGW_BSICPlan.ProgressChanged
        BSICPlanProgressBar.Value = e.ProgressPercentage
        BSICPlanStatus.Text = e.UserState.ToString
    End Sub
    Private Sub bsicPlan_runWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BGW_BSICPlan.RunWorkerCompleted
        If e.Cancelled Then
            BSICPlanStatus.Text = "欢迎使用！"
            BSICPlanProgressBar.Value = 0
        Else
            BGW_BSICPlan.CancelAsync()
            BSICPlanStatus.Text = "欢迎使用！"
            BSICPlanProgressBar.Value = 0
        End If
    End Sub
    Private Sub BSICPlan(sender As Object, e As DoWorkEventArgs)
        Dim temp_IgnoreBSIC_arr As Object
        Dim temp_BSIC_Results_arr As New ArrayList, temp_arr2 As Object, tempStr As String, coBCCHBSIC_minDis As Double, cell_minDis As String
        Dim canUseBSIC_maxDis As Double, canUseBSIC_maxDis_position As Integer
        Dim BSIC_result_arr As New ArrayList, Final_BSIC_Results_arr As New ArrayList
        Dim baseLineUsedBSIC_arr As ArrayList
        Dim lon1 As Double, lat1 As Double, lon2 As Double, lat2 As Double
        Dim notUsedBSIC As String
        Dim outPath As String
        Dim fso As Object, outPutResultFile As Object, outNewCarrier As Object
        Dim site_BSIC_arr As New ArrayList, site_bsic_temp As String = "", site_temp As String = ""
        Dim x As Integer
        outPath = Strings.Left(planCellFilePath.Text, InStrRev(planCellFilePath.Text, "\"))
        fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.folderExists(outPath & "Bsic Plan Result\") Then
            fso.createFolder(outPath & "Bsic Plan Result\")
        End If
        Try
            outPutResultFile = fso.OpenTextFile(outPath & "Bsic Plan Result\BSIC_Plan_Results_" & Format(Now, "yyyyMMdd") & ".csv", 2, 1, 0)
            outNewCarrier = fso.OpenTextFile(outPath & "Bsic Plan Result\Mcom Carrier with NEW BSIC_" & Format(Now, "yyyyMMdd") & ".txt", 2, 1, 0)
            outPutResultFile.writeLine("源小区,源小区BCCH,规划BSIC,与其最近的同BCCH同BSIC小区,与最近的同BCCH同BSIC小区的距离,是否满足限制条件？")
        Catch ex As Exception
            MsgBox("输出文件处于打开状态，请关闭后，重新开始！", MsgBoxStyle.Critical)
            Exit Sub
        End Try
        '''''''''去除不可用的BSIC
        sender.reportProgress(0, "开始整理可用的BSIC...")
        Try
            If canNotUseBSIC.Text <> "" Then
                If InStr(canNotUseBSIC.Text, ",") > 0 Then
                    temp_IgnoreBSIC_arr = Split(canNotUseBSIC.Text, ",")
                    For x = 0 To UBound(temp_IgnoreBSIC_arr)
                        If Len(temp_IgnoreBSIC_arr(x)) = 1 Then
                            temp_IgnoreBSIC_arr(x) = 0 & temp_IgnoreBSIC_arr(x)
                        End If
                        If allBSIC_ARR.Contains(temp_IgnoreBSIC_arr(x)) Then
                            allBSIC_ARR.Remove(temp_IgnoreBSIC_arr(x))
                        End If
                    Next
                ElseIf InStr(canNotUseBSIC.Text, "，") > 0 Then
                    temp_IgnoreBSIC_arr = Split(canNotUseBSIC.Text, "，")
                    For x = 0 To UBound(temp_IgnoreBSIC_arr)
                        If Len(temp_IgnoreBSIC_arr(x)) = 1 Then
                            temp_IgnoreBSIC_arr(x) = 0 & temp_IgnoreBSIC_arr(x)
                        End If
                        If allBSIC_ARR.Contains(temp_IgnoreBSIC_arr(x)) Then
                            allBSIC_ARR.Remove(temp_IgnoreBSIC_arr(x))
                        End If
                    Next
                ElseIf InStr(canNotUseBSIC.Text, "，") > 0 And InStr(canNotUseBSIC.Text, ",") > 0 Then
                    MsgBox("请正确填写限制使用的BSIC，请用英文','作为间隔！", MsgBoxStyle.Critical)
                    Exit Sub
                Else
                    If Len(canNotUseBSIC.Text) = 1 Then
                        If allBSIC_ARR.Contains(0 & canNotUseBSIC.Text) Then
                            allBSIC_ARR.Remove(0 & canNotUseBSIC.Text)
                        End If
                    Else
                        If allBSIC_ARR.Contains(canNotUseBSIC.Text) Then
                            allBSIC_ARR.Remove(canNotUseBSIC.Text)
                        End If
                    End If
                End If
            End If
            ' ''''''检查全部可用BSIC
            'Console.WriteLine("=============可用BSIC==Start========")
            'For y = 0 To allBSIC_ARR.Count - 1
            '    Console.WriteLine(allBSIC_ARR(y))
            'Next
            'Console.WriteLine("=============可用BSIC==END==========")
            sender.reportProgress(5, "正在整理可用的BSIC...")
            For x = 0 To planCell_arr.Count - 1
                sender.reportProgress(5 + ((x + 1) / planCell_arr.Count) * 90, "正在规划小区：" & planCell_arr(x) & "的BSIC...")
                If UCase(planCell_arr(x)) <> "CELL" Then
                    If Not carrier_cell_arr.Contains(planCell_arr(x)) Then
                        MsgBox("Mcom Carrier表中不存在小区：" & planCell_arr(x) & " !请检查后，重新规划！", MsgBoxStyle.Critical)
                        Exit Sub
                    End If
                    '''''''''''''找出所有于planCell_arr（x）同频小区的距离
                    baseLineUsedBSIC_arr = New ArrayList
                    notUsedBSIC = ""
                    BSIC_result_arr = New ArrayList
                    baseLineUsedBSIC_arr.Clear()
                    BSIC_result_arr.Clear()
                    For y = 0 To carrier_cell_arr.Count - 1
                        If BCCH_arr(y) = BCCH_arr(carrier_cell_arr.IndexOf(planCell_arr(x))) And carrier_cell_arr(y) <> planCell_arr(x) Then
                            If site_position_arr.ContainsKey(planCell_arr(x)) And site_position_arr.ContainsKey(carrier_cell_arr(y)) Then
                                If Split(site_position_arr(planCell_arr(x)), "#")(0) <> "" And Split(site_position_arr(planCell_arr(x)), "#")(1) <> "" And Split(site_position_arr(carrier_cell_arr(y)), "#")(0) <> "" And Split(site_position_arr(carrier_cell_arr(y)), "#")(1) <> "" Then
                                    lon1 = CDbl(Split(site_position_arr(planCell_arr(x)), "#")(0))
                                    lat1 = CDbl(Split(site_position_arr(planCell_arr(x)), "#")(1))
                                    lon2 = CDbl(Split(site_position_arr(carrier_cell_arr(y)), "#")(0))
                                    lat2 = CDbl(Split(site_position_arr(carrier_cell_arr(y)), "#")(1))
                                    If lat1 < 90 And lat2 < 90 Then
                                        BSIC_result_arr.Add(planCell_arr(x) & "," & carrier_cell_arr(y) & "," & BCCH_arr(y) & "#" & BSIC_arr(y) & "," & getDistance(lon1, lat1, lon2, lat2))
                                        If Not baseLineUsedBSIC_arr.Contains(Trim(BSIC_arr(y))) And Trim(BSIC_arr(y)) <> "" Then
                                            baseLineUsedBSIC_arr.Add(Trim(BSIC_arr(y)))
                                        End If
                                    Else
                                        MsgBox("Mcom Carrier表中的小区：" & carrier_cell_arr(y) & "的经纬度有误，请检查！", MsgBoxStyle.Critical)
                                        Exit Sub
                                    End If
                                Else
                                    BSIC_result_arr.Add(planCell_arr(x) & "," & carrier_cell_arr(y) & "," & BCCH_arr(y) & "#" & BSIC_arr(y) & "," & "a")
                                    If Not baseLineUsedBSIC_arr.Contains(Trim(BSIC_arr(y))) And Trim(BSIC_arr(y)) <> "" Then
                                             baseLineUsedBSIC_arr.Add(Trim(BSIC_arr(y)))
                                    End If
                                End If
                            Else
                                BSIC_result_arr.Add(planCell_arr(x) & "," & carrier_cell_arr(y) & "," & BCCH_arr(y) & "#" & BSIC_arr(y) & "," & "a")
                                If Not baseLineUsedBSIC_arr.Contains(Trim(BSIC_arr(y))) And Trim(BSIC_arr(y)) <> "" Then
                                            baseLineUsedBSIC_arr.Add(Trim(BSIC_arr(y)))
                                End If
                            End If
                        End If
                    Next
                    site_temp = Trim(Strings.Left(planCell_arr(x), Len(planCell_arr(x)) - 1))
                    For m = 0 To allBSIC_ARR.Count - 1
                        site_bsic_temp = site_temp & "#" & Trim(allBSIC_ARR(m))
                        If Not baseLineUsedBSIC_arr.Contains(Trim(allBSIC_ARR(m))) Then
                            If site_BSIC_arr.Contains(site_bsic_temp) Then
                                'allBSIC_ARR.Reverse()
                            Else
                                site_BSIC_arr.Add(site_bsic_temp)
                                '                       Console.WriteLine("===========================查找现网未使用的BCCH#BSIC=====START==============")
                                notUsedBSIC = Trim(allBSIC_ARR(m))
                                '                      Console.WriteLine("===========================查找现网未使用的BCCH#BSIC=====END==============")
                                Exit For
                                site_bsic_temp = ""
                                site_temp = ""
                            End If
                        End If
                    Next
                    If notUsedBSIC <> "" Then
                        outPutResultFile.writeLine(planCell_arr(x) & "," & BCCH_arr(carrier_cell_arr.IndexOf(planCell_arr(x))) & "," & notUsedBSIC & "," & "无" & "," & "现网中第一次使用" & "," & "满足")
                        If OLD_BSIC_arr(carrier_cell_arr.IndexOf(planCell_arr(x))) = "" Then
                            OLD_BSIC_arr.Item(carrier_cell_arr.IndexOf(planCell_arr(x))) = BSIC_arr(carrier_cell_arr.IndexOf(planCell_arr(x)))
                        End If
                        BSIC_arr(carrier_cell_arr.IndexOf(planCell_arr(x))) = notUsedBSIC
                    Else
                        '''''每个BSIC仅保留一个值，且为同BCCH同BSIC中与源小区距离最近的值
                        '     temp_arr2 = ""
                        tempStr = ""
                        temp_BSIC_Results_arr.Clear()
                        Final_BSIC_Results_arr.Clear()
                        canUseBSIC_maxDis_position = 0
                        For n = 0 To baseLineUsedBSIC_arr.Count - 1
                            '   site_bsic_temp = site_temp & "#" & Trim(baseLineUsedBSIC_arr(n))
                            If allBSIC_ARR.Contains(baseLineUsedBSIC_arr(n)) Then
                                tempStr = BCCH_arr(carrier_cell_arr.IndexOf(planCell_arr(x))) & "#" & baseLineUsedBSIC_arr(n)
                                coBCCHBSIC_minDis = 10000
                                cell_minDis = ""
                                For i = 0 To BSIC_result_arr.Count - 1
                                    temp_arr2 = ""
                                    temp_arr2 = Split(BSIC_result_arr(i), ",")
                                    If planCell_arr(x) <> temp_arr2(1) And tempStr = temp_arr2(2) Then
                                        If temp_arr2(3) = "a" Then
                                            coBCCHBSIC_minDis = 0
                                            cell_minDis = temp_arr2(1)
                                        Else
                                            If coBCCHBSIC_minDis = 10000 Then
                                                coBCCHBSIC_minDis = CDbl(temp_arr2(3))
                                                cell_minDis = temp_arr2(1)
                                            End If
                                            If coBCCHBSIC_minDis >= CDbl(temp_arr2(3)) Then
                                                coBCCHBSIC_minDis = CDbl(temp_arr2(3))
                                                cell_minDis = temp_arr2(1)
                                            End If
                                        End If
                                    End If
                                Next
                                If coBCCHBSIC_minDis >= maxDistanse.Value Then
                                    temp_BSIC_Results_arr.Add(coBCCHBSIC_minDis)
                                    Final_BSIC_Results_arr.Add(planCell_arr(x) & "," & BCCH_arr(carrier_cell_arr.IndexOf(planCell_arr(x))) & "," & baseLineUsedBSIC_arr(n) & "," & cell_minDis & "," & coBCCHBSIC_minDis & "," & "满足")
                                Else
                                    temp_BSIC_Results_arr.Add(coBCCHBSIC_minDis)
                                    Final_BSIC_Results_arr.Add(planCell_arr(x) & "," & BCCH_arr(carrier_cell_arr.IndexOf(planCell_arr(x))) & "," & baseLineUsedBSIC_arr(n) & "," & cell_minDis & "," & coBCCHBSIC_minDis & "," & "不满足")
                                End If
                                canUseBSIC_maxDis = 0
                                canUseBSIC_maxDis_position = 0
                                For q = 0 To temp_BSIC_Results_arr.Count - 1
                                    If q = 0 Then
                                        canUseBSIC_maxDis = temp_BSIC_Results_arr(0)
                                        canUseBSIC_maxDis_position = 0
                                    End If
                                    If canUseBSIC_maxDis < temp_BSIC_Results_arr(q) Then
                                        canUseBSIC_maxDis = temp_BSIC_Results_arr(q)
                                        canUseBSIC_maxDis_position = q
                                    End If
                                Next
                                'site_BSIC_arr.Add(site_bsic_temp)
                                'site_bsic_temp = ""
                                'site_temp = ""
                            End If
                        Next
                        outPutResultFile.writeLine(Final_BSIC_Results_arr(canUseBSIC_maxDis_position))
                        If OLD_BSIC_arr(carrier_cell_arr.IndexOf(planCell_arr(x))) = "" Then
                            OLD_BSIC_arr.Item(carrier_cell_arr.IndexOf(planCell_arr(x))) = BSIC_arr(carrier_cell_arr.IndexOf(planCell_arr(x)))
                        End If
                        BSIC_arr(carrier_cell_arr.IndexOf(planCell_arr(x))) = Split(Final_BSIC_Results_arr(canUseBSIC_maxDis_position), ",")(2)

                        'Console.WriteLine("===========================去重复后的BCCH#BSIC=====START==============")
                        'For a = 0 To Final_BSIC_Results_arr.Count - 1
                        '    Console.WriteLine(Final_BSIC_Results_arr(a))
                        'Next
                        'Console.WriteLine("===========================去重复后的BCCH#BSIC=====END==============")

                    End If
                End If
            Next
            sender.reportProgress(95, "开始生成带有规划后BSIC的Mcom Carrier文件...")

            For o = 0 To carrier_cell_arr.Count - 1
                outNewCarrier.writeLine(carrier_cell_arr(o) & "	" & BSC_arr(o) & "	" & LAI_arr(o) & "	" & CI_arr(o) & "	" & BCCH_arr(o) & "	" & BSIC_arr(o) & "	" & TCH_arr(o) & "	" & HOP_arr(o) & "	" & HSN_arr(o) & "	" & FONT_SIZE_arr(o) & "	" & OLD_BSIC_arr(o))
            Next
            sender.reportProgress(100, "BSIC规划完毕！")
        Catch ex As Exception
            MsgBox("BSIC规划失败！请检查小区：" & planCell_arr(x) & "的数据！" & vbCrLf & ex.ToString, MsgBoxStyle.Critical)
            site_BSIC_arr.Clear()
            Exit Sub
        Finally
            outPutResultFile.close()
            outNewCarrier.close()
            fso = Nothing
        End Try
        MsgBox("PCI group规划完毕！累死我了！用时：" & total_calcu_time & "秒！" & vbCrLf & "请到目录：" & vbCrLf & outPath & "BSIC Plan Result\" & vbCrLf & "查看规划结果！", MsgBoxStyle.Information)
        Shell("explorer.exe " & outPath, 1)
    End Sub


End Class
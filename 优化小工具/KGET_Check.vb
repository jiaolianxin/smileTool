Imports Microsoft.Office.Interop
Imports System.ComponentModel
Imports Microsoft.Office.Interop.Excel

Public Class KGET_Check
    Private sitedatabaseFile As String, cellTDD As String, cellRelation As String
    Private enbcell As ArrayList, eNodeBname As ArrayList, cluster As ArrayList, eNodeBID As ArrayList, cellName As ArrayList, cellID As ArrayList, PCIgroup As ArrayList, subCellId As ArrayList
    Private logicName As ArrayList, myCellName As ArrayList, nodeBid_cellID As ArrayList, S_cell_arr As ArrayList
    Private KgetMo_Cell As ArrayList, BaselinePCI As ArrayList, baseLinePCIgroup As ArrayList, baseLineSubCellId As ArrayList, baseLineState As ArrayList
    Private progressPercentage As Integer
    Private EUtranCellTDDName1 As String
    Private S_cellForCount As New ArrayList, NcellCount_arr As New ArrayList

    Private Sub Start_Click(sender As Object, e As EventArgs) Handles Start.Click
        If Not IsCheckPCI.Checked And Not IsCheckNeighbour.Checked Then
            MsgBox("请至少选择一个检查内容！")
            Exit Sub
        End If
        MsgBox("请选择最新的Site database文件！", MsgBoxStyle.Information)
        sitedatabaseFile = selectFile("选择最新的Site database文件", "Site database(*.xls)|*.xls", _
                              False, True, True)
        If sitedatabaseFile <> "" Then
            MsgBox("请选择EUtranCellTDD文件！", MsgBoxStyle.Information)
            cellTDD = selectFile("选择EUtranCellTDD文件", "EUtranCellTDD文件(*.txt)|*.txt", _
                                  False, True, True)
            If cellTDD <> "" Then
                If IsCheckNeighbour.Checked Then
                    MsgBox("请选择EUtranCellRelatin文件！", MsgBoxStyle.Information)
                    cellRelation = selectFile("选择EUtranCellRelatin文件", "EUtranCellRelatin文件(*.txt)|*.txt", _
                                          False, True, True)
                End If
                Close()
                BGWcheckKGET.RunWorkerAsync()
            End If
        End If
    End Sub

    Private Sub checkPCI_dowork(sender As Object, e As DoWorkEventArgs) Handles BGWcheckKGET.DoWork
        If sitedatabaseFile <> "" Then
            getSiteDatabase(sitedatabaseFile, sender, e)
            If IsCheckNeighbour.Checked Then
                If cellRelation <> "" Then
                    getEUtranCellRelation(cellRelation)
                    checkNeighbour(sender, e)
                End If
            End If
            If IsCheckPCI.Checked Then
                If cellTDD <> "" Then
                    getEUtranCellTDD(cellTDD, sender, e)
                    CheckPCI(sender, e)
                End If
            End If
            MsgBox("KGET 检查完毕!")
        End If
        sitedatabaseFile = ""
        cellTDD = ""
        cellRelation = ""
    End Sub
    Private Sub checkPCI_changedProgress(sender As Object, e As ProgressChangedEventArgs) Handles BGWcheckKGET.ProgressChanged
        MainForm.MainProgressBar.Value = CInt(e.ProgressPercentage)
        MainForm.statusBar.Text = CStr(e.UserState.ToString)
    End Sub
    Private Sub checkPCI_runWorkCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BGWcheckKGET.RunWorkerCompleted
        If e.Cancelled Then
            MainForm.statusBar.Text = "欢迎使用！"
            MainForm.MainProgressBar.Value = 0

        Else
            BGWcheckKGET.CancelAsync()
            MainForm.statusBar.Text = "欢迎使用！"
            MainForm.MainProgressBar.Value = 0
        End If
    End Sub

    Private Sub getSiteDatabase(siteDatabaseName As String, sender As Object, e As DoWorkEventArgs)
        Dim excel_app As New Excel.Application
        Dim excel_book As Object
        Dim sheetsCount As Integer
        Dim total_cell As Integer
        Dim ENBCELL_pos As Integer, LOGIC_NAME_pos As Integer, ENODEB_NAME_pos As Integer, CLUSTER_pos As Integer, eNB_id_pos As Integer, CELL_NAME_pos As Integer, CELL_ID_pos As Integer, PCIgroup_pos As Integer, subCellId_pos As Integer
        enbcell = New ArrayList
        eNodeBname = New ArrayList
        cluster = New ArrayList
        eNodeBID = New ArrayList
        cellName = New ArrayList
        cellID = New ArrayList
        PCIgroup = New ArrayList
        subCellId = New ArrayList
        logicName = New ArrayList
        myCellName = New ArrayList
        nodeBid_cellID = New ArrayList
        Dim progressPercentage As Integer
        Dim isFoundSite As Boolean, isFoundSiteIndoor As Boolean
        Dim x As Integer

        isFoundSite = False
        excel_app = GetObject(, "Excel.Application")
        excel_book = excel_app.Workbooks.Open(siteDatabaseName)
        Try
            '      excel_book = GetObject(siteDatabaseName)
            sheetsCount = excel_book.Sheets.Count
            progressPercentage = 0
            sender.reportProgress(progressPercentage, "读取Site database数据......")
            For x = 1 To sheetsCount
                If excel_book.sheets(x).name = "Site database" Then
                    total_cell = excel_app.ActiveWorkbook.Sheets("Site database").UsedRange.rows.count
                    isFoundSite = True
                    For y = 1 To 50
                        If UCase(Trim(excel_book.sheets("Site database").cells(1, y).value)) = "ENBCELL" Or UCase(Trim(excel_book.sheets("Site database").cells(2, y).value)) = "ENBCELL" Then
                            ENBCELL_pos = y
                        ElseIf UCase(Trim(excel_book.sheets("Site database").cells(1, y).value)) = "LOGIC NAME" Or UCase(Trim(excel_book.sheets("Site database").cells(2, y).value)) = "LOGIC NAME" Then
                            LOGIC_NAME_pos = y
                        ElseIf UCase(Trim(excel_book.sheets("Site database").cells(1, y).value)) = "ENODEB NAME" Or UCase(Trim(excel_book.sheets("Site database").cells(2, y).value)) = "ENODEB NAME" Then
                            ENODEB_NAME_pos = y
                        ElseIf UCase(Trim(excel_book.sheets("Site database").cells(1, y).value)) = "CLUSTER" Or UCase(Trim(excel_book.sheets("Site database").cells(2, y).value)) = "CLUSTER" Then
                            CLUSTER_pos = y
                        ElseIf Trim(excel_book.sheets("Site database").cells(1, y).value) = "新eNB id(十进制)" Or Trim(excel_book.sheets("Site database").cells(2, y).value) = "新eNB id(十进制)" Then
                            eNB_id_pos = y
                        ElseIf UCase(Trim(excel_book.sheets("Site database").cells(1, y).value)) = "CELL NAME" Or UCase(Trim(excel_book.sheets("Site database").cells(2, y).value)) = "CELL NAME" Then
                            CELL_NAME_pos = y
                        ElseIf UCase(Trim(excel_book.sheets("Site database").cells(1, y).value)) = "CELL ID" Or UCase(Trim(excel_book.sheets("Site database").cells(2, y).value)) = "CELL ID" Then
                            CELL_ID_pos = y
                        ElseIf InStr(Trim(excel_book.sheets("Site database").cells(1, y).value), "CellIdGroup") Or InStr(Trim(excel_book.sheets("Site database").cells(2, y).value), "CellIdGroup") Then
                            PCIgroup_pos = y
                        ElseIf InStr(Trim(excel_book.sheets("Site database").cells(1, y).value), "SubCellId") Or InStr(Trim(excel_book.sheets("Site database").cells(2, y).value), "SubCellId") Then
                            subCellId_pos = y
                        End If
                    Next
                    sender.reportProgress(10, "读取Site database数据......")
                    For m = 1 To total_cell
                        enbcell.Add(excel_book.sheets("Site database").cells(m, ENBCELL_pos).value)
                        logicName.Add(excel_book.sheets("Site database").cells(m, LOGIC_NAME_pos).value)
                        eNodeBname.Add(excel_book.sheets("Site database").cells(m, ENODEB_NAME_pos).value)
                        cluster.Add(excel_book.sheets("Site database").cells(m, CLUSTER_pos).value)
                        eNodeBID.Add(excel_book.sheets("Site database").cells(m, eNB_id_pos).value)
                        cellName.Add(excel_book.sheets("Site database").cells(m, CELL_NAME_pos).value)
                        cellID.Add(excel_book.sheets("Site database").cells(m, CELL_ID_pos).value)
                        PCIgroup.Add(excel_book.sheets("Site database").cells(m, PCIgroup_pos).value)
                        subCellId.Add(excel_book.sheets("Site database").cells(m, subCellId_pos).value)
                        nodeBid_cellID.Add(excel_book.sheets("Site database").cells(m, eNB_id_pos).value & "-" & excel_book.sheets("Site database").cells(m, CELL_ID_pos).value)
                        sender.reportProgress(10 + (m / total_cell) * 20, "读取Site database数据......")
                        myCellName.Add(Strings.Split(excel_book.sheets("Site database").cells(m, LOGIC_NAME_pos).value, "-")(0) & "_" & Strings.Right(excel_book.sheets("Site database").cells(m, CELL_ID_pos).value, 1))
                    Next
                ElseIf excel_book.sheets(x).name = "IBS" Then
                    total_cell = excel_app.Sheets("IBS").UsedRange.rows.count
                    isFoundSiteIndoor = True
                    sender.reportProgress(30, "读取IBS(室分)数据......")
                    For y = 1 To 50
                        If UCase(Trim(excel_book.sheets("IBS").cells(1, y).value)) = "ENBCELL" Or UCase(Trim(excel_book.sheets("IBS").cells(2, y).value)) = "ENBCELL" Then
                            ENBCELL_pos = y
                        ElseIf UCase(Trim(excel_book.sheets("IBS").cells(1, y).value)) = "LOGIC NAME" Or UCase(Trim(excel_book.sheets("IBS").cells(2, y).value)) = "LOGIC NAME" Then
                            LOGIC_NAME_pos = y
                        ElseIf UCase(Trim(excel_book.sheets("IBS").cells(1, y).value)) = "ENODEB NAME" Or UCase(Trim(excel_book.sheets("IBS").cells(2, y).value)) = "ENODEB NAME" Then
                            ENODEB_NAME_pos = y
                        ElseIf UCase(Trim(excel_book.sheets("IBS").cells(1, y).value)) = "CLUSTER" Or UCase(Trim(excel_book.sheets("IBS").cells(2, y).value)) = "CLUSTER" Then
                            CLUSTER_pos = y
                        ElseIf Trim(excel_book.sheets("IBS").cells(1, y).value) = "新eNodeB Id(十进制)" Or Trim(excel_book.sheets("IBS").cells(2, y).value) = "新eNodeB Id(十进制)" Then
                            eNB_id_pos = y
                        ElseIf UCase(Trim(excel_book.sheets("IBS").cells(1, y).value)) = "CELL NAME" Or UCase(Trim(excel_book.sheets("IBS").cells(2, y).value)) = "CELL NAME" Then
                            CELL_NAME_pos = y
                        ElseIf UCase(Trim(excel_book.sheets("IBS").cells(1, y).value)) = "CELL ID" Or UCase(Trim(excel_book.sheets("IBS").cells(2, y).value)) = "CELL ID" Then
                            CELL_ID_pos = y
                        ElseIf InStr(Trim(excel_book.sheets("IBS").cells(1, y).value), "CellIdGroup") Or InStr(Trim(excel_book.sheets("IBS").cells(2, y).value), "CellIdGroup") Then
                            PCIgroup_pos = y
                        ElseIf InStr(Trim(excel_book.sheets("IBS").cells(1, y).value), "SubCellId") Or InStr(Trim(excel_book.sheets("IBS").cells(2, y).value), "SubCellId") Then
                            subCellId_pos = y
                        End If
                    Next
                    For m = 1 To total_cell
                        enbcell.Add(excel_book.sheets("IBS").cells(m, ENBCELL_pos).value)
                        logicName.Add(excel_book.sheets("IBS").cells(m, LOGIC_NAME_pos).value)
                        eNodeBname.Add(excel_book.sheets("IBS").cells(m, ENODEB_NAME_pos).value)
                        cluster.Add(excel_book.sheets("IBS").cells(m, CLUSTER_pos).value)
                        eNodeBID.Add(excel_book.sheets("IBS").cells(m, eNB_id_pos).value)
                        cellName.Add(excel_book.sheets("IBS").cells(m, CELL_NAME_pos).value)
                        cellID.Add(excel_book.sheets("IBS").cells(m, CELL_ID_pos).value)
                        PCIgroup.Add(excel_book.sheets("IBS").cells(m, PCIgroup_pos).value)
                        subCellId.Add(excel_book.sheets("IBS").cells(m, subCellId_pos).value)
                        nodeBid_cellID.Add(excel_book.sheets("IBS").cells(m, eNB_id_pos).value & "-" & excel_book.sheets("IBS").cells(m, CELL_ID_pos).value)
                        myCellName.Add(Strings.Split(excel_book.sheets("IBS").cells(m, LOGIC_NAME_pos).value, "-")(0) & "_" & Strings.Right(excel_book.sheets("IBS").cells(m, CELL_ID_pos).value, 1))
                        sender.reportProgress(30 + (m / total_cell) * 20, "读取IBS(室分)数据......")
                    Next
                End If
            Next
            If Not isFoundSite Then
                MsgBox("请确认打开的文件是SiteDatabase文件！(Sheet名为Site database）", MsgBoxStyle.Critical)
                isFoundSite = False
                isFoundSiteIndoor = False
                excel_app.DisplayAlerts = False
                excel_book.close()
                excel_book = Nothing
                excel_app.Quit()
                Exit Sub
            End If
            'For x = 0 To logicName.Count - 1
            'Next
        Catch ex As Exception
            excel_app.DisplayAlerts = False
            excel_book.close()
            excel_app.Quit()
            MsgBox("发生错误！请检查输入文件是否正确！", MsgBoxStyle.Critical)
            Exit Sub
        Finally
            excel_app.DisplayAlerts = False
            excel_book.close()
            excel_book = Nothing
            excel_app.Quit()
            excel_app = Nothing
        End Try
    End Sub
    Private Sub getEUtranCellTDD(EUtranCellTDDName As String, sender As Object, e As DoWorkEventArgs)
        Dim nowCellTDDLine As String
        Dim temp3 As Array
        KgetMo_Cell = New ArrayList
        BaselinePCI = New ArrayList
        baseLinePCIgroup = New ArrayList
        baseLineSubCellId = New ArrayList
        baseLineState = New ArrayList
        sender.reportProgress(55, "导入EUtranCellTDD文件...")
        EUtranCellTDDName1 = EUtranCellTDDName
        FileOpen(1, EUtranCellTDDName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        Try
            Do While Not EOF(1)
                nowCellTDDLine = LineInput(1)
                temp3 = Strings.Split(nowCellTDDLine, "	")
                'SubNetwork=QDLTE,MeContext=LDE02A005H,ManagedElement=1,ENodeBFunction=1,EUtranCellTDD=3
                If InStr(temp3(1), "=") > 0 Then
                    KgetMo_Cell.Add(Strings.Split(Strings.Split(temp3(1), ",")(1).ToString, "=")(1).ToString & "_" & temp3(2).ToString)
                    BaselinePCI.Add(CInt(temp3(109)) * 3 + CInt(temp3(110)))
                Else
                    KgetMo_Cell.Add(temp3(1).ToString)
                    BaselinePCI.Add("PCIgroup")
                End If
                baseLinePCIgroup.Add(temp3(109).ToString)
                baseLineSubCellId.Add(temp3(110).ToString)
                baseLineState.Add(temp3(26).ToString)
            Loop
            If InStr(baseLineState(2), "LOCK") = 0 Then
                FileClose(1)
                MsgBox("请检查输入文件EUtranCellTDD格式是否正确！", MsgBoxStyle.Critical)
                Exit Sub
            End If
        Catch ex As Exception
            FileClose(1)
            MsgBox("未知错误，请检查输入文件EUtranCellTDD格式是否正确！", MsgBoxStyle.Critical)
            Exit Sub
        Finally
            FileClose(1)
            sender.reportProgress(70, "EUtranCellTDD文件导入完毕！")
        End Try
    End Sub
    Private Sub getEUtranCellRelation(NeighbourName As String)
        Dim templine As String
        Dim temp_arr As Array
        Dim i As Integer
        Dim fso As Object
        Dim outCellRelationFile As Object
        Dim S_CELL As String, N_CELL As String, S_CELL_NAME As String, N_CELL_NAME As String, S_nodeBID As String, N_nodeBID As String
        S_cell_arr = New ArrayList
        fso = CreateObject("Scripting.FileSystemObject")
        FileClose(2)
        outCellRelationFile = fso.OpenTextFile(Strings.Left(NeighbourName, InStrRev(NeighbourName, "\")) & "\EUtranCellRelation_CellName" & Format(FileDateTime(NeighbourName), "MMdd") & ".xls", 2, 1, 0)
        FileOpen(2, NeighbourName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        S_CELL_NAME = ""
        S_nodeBID = ""
        N_nodeBID = ""
        N_CELL_NAME = ""
        S_cell_arr.Add("cell")
        Try
            Do While Not EOF(2)
                i = i + 1
                templine = LineInput(2)
                temp_arr = Strings.Split(templine, "	")
                If i = 1 Then
                    If temp_arr(2) <> "EUtranCellRelationId" Then
                        FileClose(2)
                        MsgBox("请检查输入文件EUtranCellRelation格式是否正确！", MsgBoxStyle.Critical)
                        Exit Sub
                    End If
                    outCellRelationFile.writeLine(temp_arr(0) & "	" & temp_arr(1) & "	" & "S_ENBCELL" & "	" & "S_CELL NAME" & "	" & "S_eNodBID-CellId" & "	" & "T_CELL NAME" & "	" & "T_eNodBID-CellId" & "	" & temp_arr(2) & "	" & temp_arr(3) & "	" & temp_arr(4) & "	" & temp_arr(5) & "	" & temp_arr(6) & "	" & temp_arr(7) & "	" & temp_arr(8) & "	" & temp_arr(9) & "	" & temp_arr(10) & "	" & temp_arr(11) & "	" & temp_arr(12) & "	" & temp_arr(13) & "	" & temp_arr(14) & "	" & temp_arr(15) & "	" & temp_arr(16) & "	" & temp_arr(17) & "	" & temp_arr(18) & "	" & temp_arr(19) & "	" & temp_arr(20) & "	" & temp_arr(21))
                Else
                    Try
                        S_CELL = Split(Split(temp_arr(1), ",")(1), "=")(1) & "_" & Split(Split(temp_arr(1), ",")(4), "=")(1)
                    Catch ex As Exception
                        S_CELL = Trim(temp_arr(1))
                    End Try
                    S_cell_arr.Add(S_CELL)
                    If myCellName.Contains(S_CELL) Then
                        S_CELL_NAME = Trim(cellName(myCellName.IndexOf(S_CELL)))
                        S_nodeBID = Trim(nodeBid_cellID(myCellName.IndexOf(S_CELL)))
                    Else
                        S_CELL_NAME = ""
                        S_nodeBID = ""
                    End If
                    If InStr(Trim(temp_arr(2)), "-") > 0 Then
                        If UBound(Split(temp_arr(2), "-")) = 1 Then
                            N_CELL = Strings.Left(Split(temp_arr(2), "-")(1), Len(Split(temp_arr(2), "-")(1)) - 1) & "_" & Strings.Right(Split(temp_arr(2), "-")(1), 1)
                            If enbcell.Contains(N_CELL) Then
                                N_CELL_NAME = Trim(cellName(enbcell.IndexOf(N_CELL)))
                            Else
                                N_CELL_NAME = ""
                            End If
                        Else
                            N_CELL = Split(Trim(temp_arr(2)), "-")(1) & "-" & Split(Trim(temp_arr(2)), "-")(2)
                            If nodeBid_cellID.Contains(N_CELL) Then
                                N_CELL_NAME = Trim(cellName(nodeBid_cellID.IndexOf(N_CELL)))
                            Else
                                N_CELL_NAME = ""
                            End If

                        End If
                    Else
                        N_CELL = Strings.Left(S_CELL, Len(S_CELL) - 1) & Trim(temp_arr(2))
                        If myCellName.Contains(N_CELL) Then
                            N_CELL_NAME = Trim(cellName(myCellName.IndexOf(N_CELL)))
                            N_CELL = Trim(nodeBid_cellID(myCellName.IndexOf(N_CELL)))
                        Else
                            N_CELL_NAME = ""
                            N_CELL = ""
                        End If
                    End If
                    outCellRelationFile.writeLine(temp_arr(0) & "	" & temp_arr(1) & "	" & S_CELL & "	" & S_CELL_NAME & "	" & S_nodeBID & "	" & N_CELL_NAME & "	" & N_CELL & "	" & temp_arr(2) & "	" & temp_arr(3) & "	" & temp_arr(4) & "	" & temp_arr(5) & "	" & temp_arr(6) & "	" & temp_arr(7) & "	" & temp_arr(8) & "	" & temp_arr(9) & "	" & temp_arr(10) & "	" & temp_arr(11) & "	" & temp_arr(12) & "	" & temp_arr(13) & "	" & temp_arr(14) & "	" & temp_arr(15) & "	" & temp_arr(16) & "	" & temp_arr(17) & "	" & temp_arr(18) & "	" & temp_arr(19) & "	" & temp_arr(20) & "	" & temp_arr(21))
                End If
            Loop
        Catch ex As Exception
            MsgBox("未知错误，请检查输入文件EUtranCellRelation格式是否正确！", MsgBoxStyle.Critical)
            Exit Sub
        Finally
            FileClose(2)
            outCellRelationFile.close()
            outCellRelationFile = Nothing
            fso = Nothing
        End Try
       

    End Sub
    Private Sub CheckPCI(sender As Object, e As DoWorkEventArgs)
        Dim outPutResultFileName As String
        Dim fso As Object, outPutFile As Object
        Dim sitePCI As String, sitePCIgroup As String, siteSubCellId As String, ncell_total_count As String
        outPutResultFileName = Strings.Left(EUtranCellTDDName1, InStrRev(EUtranCellTDDName1, "\")) & "PCI_check_" & Format(FileDateTime(EUtranCellTDDName1), "yyyyMMdd") & ".csv"
        sender.reportProgress(72, "PCI_Check Progress......")
        FileClose(1)
        fso = CreateObject("Scripting.FileSystemObject")
        outPutFile = fso.OpenTextFile(outPutResultFileName, 2, 1, 0)
        Try
            outPutFile.WriteLine("CELL NAME" & "," & "ENBCELL(KGET)" & "," & "administrativeState（KGET)" & "," & baseLinePCIgroup(0) & "(KGET)," & baseLineSubCellId(0) & "(KGET)," & "PCI(KGET)" & "," & "PCIgroup(siteDatabase)" & "," & "subCellId(siteDatabase)" & "," & "PCI(siteDatabase)" & "," & "PCI是否相同" & "," & "邻区个数")
            For x = 1 To KgetMo_Cell.Count - 1
                If S_cellForCount.Contains((KgetMo_Cell(x))) Then
                    ncell_total_count = NcellCount_arr(S_cellForCount.IndexOf(KgetMo_Cell(x)))
                Else
                    ncell_total_count = "0"
                End If
                If myCellName.Contains(KgetMo_Cell(x).ToString) Then
                    sitePCIgroup = PCIgroup(myCellName.IndexOf(KgetMo_Cell(x)))
                    siteSubCellId = subCellId(myCellName.IndexOf(KgetMo_Cell(x)))
                    If sitePCIgroup <> "" And siteSubCellId <> "" Then
                        sitePCI = PCIgroup(myCellName.IndexOf(KgetMo_Cell(x))) * 3 + subCellId(myCellName.IndexOf(KgetMo_Cell(x)))
                    Else
                        sitePCI = ""
                    End If
                    outPutFile.WriteLine(cellName(myCellName.IndexOf(KgetMo_Cell(x).ToString)) & "," & KgetMo_Cell(x) & "," & baseLineState(x) & "," & baseLinePCIgroup(x) & "," & baseLineSubCellId(x) & "," & BaselinePCI(x) & "," & sitePCIgroup & "," & siteSubCellId & "," & sitePCI & "," & IIf(BaselinePCI(x) = sitePCI, "是", "否") & "," & ncell_total_count)
                Else
                    sitePCIgroup = ""
                    sitePCI = ""
                    siteSubCellId = ""
                    outPutFile.WriteLine("" & "," & KgetMo_Cell(x) & "," & baseLineState(x) & "," & baseLinePCIgroup(x) & "," & baseLineSubCellId(x) & "," & BaselinePCI(x) & "," & "" & "," & "" & "," & "" & "," & "否" & "," & ncell_total_count)
                End If
            Next
            sender.reportProgress(100, "PCI checking Complete!")
        Catch ex As Exception
            FileClose(1)
            outPutFile.close()
            MsgBox("检查KGET时发生问题，请检查输入文件是否正确！" & ex.ToString, MsgBoxStyle.Critical)
            Exit Sub
        Finally
            FileClose(1)
            outPutFile.close()
            outPutFile = Nothing
        End Try
    End Sub
    Private Sub checkNeighbour(sender As Object, e As DoWorkEventArgs)
        Dim NcellCount As Integer
        Dim fso As Object
        fso = CreateObject("Scripting.FileSystemObject")
        S_cell_arr.Sort()
        S_cellForCount = New ArrayList
        NcellCount_arr = New ArrayList
        S_cellForCount.Add("CELL")
        NcellCount_arr.Add("邻区个数")

        For x = 1 To S_cell_arr.Count - 1
            If x + 1 = S_cell_arr.Count Then
                If S_cell_arr(x) = S_cell_arr(x - 1) Then
                    NcellCount = NcellCount + 1
                    S_cellForCount.Add(S_cell_arr(x))
                    NcellCount_arr.Add(NcellCount)
                Else
                    S_cellForCount.Add(S_cell_arr(x))
                    NcellCount_arr.Add("1")
                End If
            Else
                If S_cell_arr(x) = S_cell_arr(x - 1) And S_cell_arr(x) = S_cell_arr(x + 1) Then
                    NcellCount = NcellCount + 1
                ElseIf S_cell_arr(x) = S_cell_arr(x - 1) And S_cell_arr(x) <> S_cell_arr(x + 1) Then
                    NcellCount = NcellCount + 1
                    S_cellForCount.Add(S_cell_arr(x))
                    NcellCount_arr.Add(NcellCount)
                ElseIf S_cell_arr(x) <> S_cell_arr(x - 1) And S_cell_arr(x) = S_cell_arr(x + 1) Then
                    NcellCount = 1
                ElseIf S_cell_arr(x) <> S_cell_arr(x - 1) And S_cell_arr(x) <> S_cell_arr(x + 1) Then
                    S_cellForCount.Add(S_cell_arr(x))
                    NcellCount_arr.Add("1")
                End If
            End If
        Next
    End Sub

End Class
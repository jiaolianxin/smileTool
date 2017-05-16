Imports System.ComponentModel

Public Class HW_CDD_PARSE
    Private BSC_INDEX_CELLNAME_ARR As ArrayList
    Private CSYSTYPE_ARR As ArrayList
    Private LAC_ARR As ArrayList, CI_ARR As ArrayList
    Private BSIC_ARR As ArrayList
    Private BSC_INDEX_CELLNAME_BCCH_DIC As Dictionary(Of String, String)
    Private BSC_INDEX_CELLNAME_TCH_DIC As Dictionary(Of String, String)

    Private Sub startParse_Click(sender As Object, e As EventArgs) Handles startParse.Click
        If GcellFilePath.Text = "请选择GCELL文件!" Then
            MsgBox("请先选择GCELL文件!", MsgBoxStyle.Critical)
            Exit Sub
        End If
        If GtrxFilePath.Text = "请选择GTRX文件!" Then
            MsgBox("请先选择GTRX文件!", MsgBoxStyle.Critical)
            Exit Sub
        End If
        Hide()
        BGWForHWPara.RunWorkerAsync()
    End Sub
    Private Sub parse_parameter(sender As Object, e As DoWorkEventArgs)
        Dim outFileName As String  '全路径
        Dim tempStr As String
        Dim x As Integer
        tempStr = ""
        outFileName = Strings.Left(GcellFilePath.Text, InStrRev(GcellFilePath.Text, "\")) & "Results_" & Format(FileDateTime(GtrxFilePath.Text), "yyyyMMdd") & ".csv"
        tempStr = "CELLID" & "," & "CELLNAME" & "," & "BSC" & "," & "频段" & "," & "LAC" & "," & "CI" & "," & "NCC" & "," & "BCC" & "," & "BSIC" & "," & "BCCH" & "," & "TCH1" & "," & "TCH2" & "," & "TCH3" & "," & "TCH4" & "," & "TCH5" & "," & "TCH6" & "," & "TCH7" & "," & "TCH8" & "," & "TCH9" & "," & "TCH10" & "," & "TCH11" & "," & "TCH12"
        myWriteLine(outFileName, tempStr)
        '     tempStr = "小区索引" & "," & "小区名称" & "," & "网元名称" & "," & "频段" & "," & "小区LAC" & "," & "小区CI" & "," & "NCC" & "," & "BCC" & "," & "BSIC" & "," & "BCCH" & "," & "TCH1" & "," & "TCH2" & "," & "TCH3" & "," & "TCH4" & "," & "TCH5" & "," & "TCH6" & "," & "TCH7" & "," & "TCH8" & "," & "TCH9" & "," & "TCH10" & "," & "TCH11" & "," & "TCH12"
        '    myWriteLine(outFileName, tempStr)
        Try
            For x = 0 To BSC_INDEX_CELLNAME_ARR.Count - 1
                tempStr = Split(BSC_INDEX_CELLNAME_ARR(x), "_")(1) & "," & Split(BSC_INDEX_CELLNAME_ARR(x), "_")(2) & "," & Split(BSC_INDEX_CELLNAME_ARR(x), "_")(0) & "," & CSYSTYPE_ARR(x) & "," & LAC_ARR(x) & "," & CI_ARR(x) & "," & Strings.Left(BSIC_ARR(x), 1) & "," & Strings.Right(BSIC_ARR(x), 1) & "," & BSIC_ARR(x) & "," & BSC_INDEX_CELLNAME_BCCH_DIC(BSC_INDEX_CELLNAME_ARR(x)) & "," & BSC_INDEX_CELLNAME_TCH_DIC(BSC_INDEX_CELLNAME_ARR(x))
                myWriteLine(outFileName, tempStr)
            Next
            MsgBox("转换完成！", MsgBoxStyle.Information)
            Shell("explorer.exe " & outFileName, 1)
        Catch ex As Exception
            MsgBox("请检查小区" & BSC_INDEX_CELLNAME_ARR(x) & "是否存在问题！？", MsgBoxStyle.RetryCancel)
            Print(BSC_INDEX_CELLNAME_BCCH_DIC(BSC_INDEX_CELLNAME_ARR(x)))
            Exit Sub
        Finally
            BSC_INDEX_CELLNAME_ARR.Clear()
            BSC_INDEX_CELLNAME_BCCH_DIC.Clear()
            BSC_INDEX_CELLNAME_TCH_DIC.Clear()
            LAC_ARR.Clear()
            CI_ARR.Clear()
            BSIC_ARR.Clear()
        End Try
    End Sub
    Private Sub DoWork(sender As Object, e As DoWorkEventArgs) Handles BGWForHWPara.DoWork
        parse_parameter(sender, e)
    End Sub
    Private Sub ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BGWForHWPara.ProgressChanged
        MainForm.MainProgressBar.Value = IIf(CInt(e.ProgressPercentage) > 100, 100, CInt(e.ProgressPercentage))
        If InStr(CStr(e.UserState.ToString), "@") Then
            MainForm.mainMessageBox.AppendText(Split(CStr(e.UserState.ToString), "@")(1) & vbCrLf)
            MainForm.statusBar.Text = Split(CStr(e.UserState.ToString), "@")(0)
        Else
            MainForm.statusBar.Text = CStr(e.UserState.ToString)
        End If
    End Sub
    Private Sub RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BGWForHWPara.RunWorkerCompleted
        If e.Cancelled Then
            MainForm.statusBar.Text = "欢迎使用！"
            MainForm.MainProgressBar.Value = 0
        Else
            BGWForHWPara.CancelAsync()
            MainForm.statusBar.Text = "欢迎使用！"
            MainForm.MainProgressBar.Value = 0
        End If
    End Sub
    Private Sub importGcell_Click(sender As Object, e As EventArgs) Handles importGcell.Click
        Dim filePath As String
        filePath = selectFile("选择GCELL文件", "GCELL文件(*.txt)|*.txt", False, True, True)
        If filePath <> "" Then
            GcellFilePath.Text = filePath
            importGCELLFile(filePath)
        End If
    End Sub
    Private Sub importGTRX_Click(sender As Object, e As EventArgs) Handles importGTRX.Click
        Dim filePath As String
        filePath = selectFile("选择GTRX文件", "GTRX文件(*.txt)|*.txt", False, True, True)
        If filePath <> "" Then
            GtrxFilePath.Text = filePath
            importGTRXFile(filePath)
        End If
    End Sub
    Private Sub importGCELLFile(filePath As String)
        Dim fileNo As Integer
        Dim tempLine As String
        Dim temp_arr As Object
        Dim nowLineNo As Integer = 1
        BSC_INDEX_CELLNAME_ARR = New ArrayList
        CSYSTYPE_ARR = New ArrayList
        LAC_ARR = New ArrayList
        CI_ARR = New ArrayList
        BSIC_ARR = New ArrayList
        fileNo = FreeFile()
        FileOpen(fileNo, filePath, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        Try
            Do While Not EOF(fileNo)
                tempLine = Replace(LineInput(fileNo), """", "")
                temp_arr = Split(UCase(tempLine), ",")
                If nowLineNo = 1 Then
                    If temp_arr(0) <> "BSCNAME" Or temp_arr(2) <> "CELLID" Or temp_arr(3) <> "CELLNAME" Or temp_arr(4) <> "TYPE" Then
                        MsgBox("GCELL文件格式不正确，请重新选择！", MsgBoxStyle.Critical)
                        GcellFilePath.Text = "请选择GCELL文件!"
                        Exit Sub
                    End If
                ElseIf nowLineNo = 2 Then
                    If temp_arr(0) <> "BSC NAME" Then
                        BSC_INDEX_CELLNAME_ARR.Add(temp_arr(0) & "_" & temp_arr(2) & "_" & temp_arr(3))
                        CSYSTYPE_ARR.Add(temp_arr(4))
                        LAC_ARR.Add(temp_arr(7))
                        CI_ARR.Add(temp_arr(8))
                        BSIC_ARR.Add(temp_arr(9) & temp_arr(10))
                    End If
                Else
                    BSC_INDEX_CELLNAME_ARR.Add(temp_arr(0) & "_" & temp_arr(2) & "_" & temp_arr(3))
                    CSYSTYPE_ARR.Add(temp_arr(4))
                    LAC_ARR.Add(temp_arr(7))
                    CI_ARR.Add(temp_arr(8))
                    BSIC_ARR.Add(temp_arr(9) & temp_arr(10))
                End If
                nowLineNo = nowLineNo + 1
            Loop
        Catch ex As Exception
            MsgBox("GCELL文件导入失败，请检查！", MsgBoxStyle.Critical)
            Exit Sub
        Finally
            FileClose(fileNo)
        End Try

    End Sub
    Private Sub importGTRXFile(filePath As String)
        Dim fileNo As Integer
        Dim tempLine As String
        Dim temp_arr As Object
        Dim nowLineNo As Integer = 1
        Dim lastIndex As String
        Dim nowIndex As String
        Dim TCH_temp As String
        Dim freqNo As String
        Dim bcchMark As String
        Dim verFlag As Integer

        BSC_INDEX_CELLNAME_BCCH_DIC = New Dictionary(Of String, String)
        BSC_INDEX_CELLNAME_TCH_DIC = New Dictionary(Of String, String)
        lastIndex = ""
        nowIndex = ""
        TCH_temp = ""
        freqNo = ""
        bcchMark = ""
        verFlag = 1
        fileNo = FreeFile()
        FileOpen(fileNo, filePath, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        Try
            Do While Not EOF(fileNo)
                tempLine = Replace(LineInput(fileNo), """", "")
                temp_arr = Split(UCase(tempLine), ",")
                If verFlag = 1 Then
                    freqNo = temp_arr(4)
                    bcchMark = temp_arr(5)
                ElseIf verFlag = 2 Then
                    freqNo = temp_arr(5)
                    bcchMark = temp_arr(6)
                End If
                If nowLineNo = 1 Then
                    If temp_arr(0) <> "BSCNAME" Or temp_arr(2) <> "CELLID" Or temp_arr(1) <> "CELLNAME" Then
                        MsgBox("GTRX文件格式不正确，请重新选择！", MsgBoxStyle.Critical)
                        GtrxFilePath.Text = "请选择GTRX文件!"
                        Exit Sub
                    Else
                        If temp_arr(4) <> "FREQ" And temp_arr(5) <> "FREQ" Then
                            MsgBox("GTRX文件格式不正确，请重新选择！", MsgBoxStyle.Critical)
                            GtrxFilePath.Text = "请选择GTRX文件!"
                            Exit Sub
                        Else
                            If temp_arr(5) <> "ISMAINBCCH" And temp_arr(6) <> "ISMAINBCCH" Then
                                MsgBox("GTRX文件格式不正确，请重新选择！", MsgBoxStyle.Critical)
                                GtrxFilePath.Text = "请选择GTRX文件!"
                                Exit Sub
                            End If
                        End If
                        If temp_arr(4) = "FREQ" Then
                            verFlag = 1
                        ElseIf temp_arr(5) = "FREQ" Then
                            verFlag = 2
                        End If
                    End If
                ElseIf nowLineNo = 2 Then
                    If temp_arr(0) <> "BSC NAME" Then
                        nowIndex = temp_arr(0) & "_" & temp_arr(2) & "_" & temp_arr(1)
                        If bcchMark = "YES" Then
                            If BSC_INDEX_CELLNAME_BCCH_DIC.ContainsKey(nowIndex) Then
                                MsgBox("小区 " & nowIndex & " 存在两个BCCH，请核查是否正确！", MsgBoxStyle.Critical)
                            Else
                                BSC_INDEX_CELLNAME_BCCH_DIC.Add(nowIndex, freqNo)
                            End If
                        Else
                            TCH_temp = freqNo
                        End If
                        lastIndex = nowIndex
                    End If
                Else
                    nowIndex = temp_arr(0) & "_" & temp_arr(2) & "_" & temp_arr(1)
                    If lastIndex = nowIndex Then
                        If bcchMark = "YES" Then
                            If BSC_INDEX_CELLNAME_BCCH_DIC.ContainsKey(nowIndex) Then
                                MsgBox("小区 " & nowIndex & " 存在两个BCCH，请核查是否正确！", MsgBoxStyle.Critical)
                            Else
                                BSC_INDEX_CELLNAME_BCCH_DIC.Add(nowIndex, freqNo)
                            End If
                        Else
                            If TCH_temp = "" Then
                                TCH_temp = freqNo
                            Else
                                TCH_temp = TCH_temp & "," & freqNo
                            End If
                        End If
                        'If TCH_temp = "" Then
                        '    TCH_temp = freqNo
                        'Else
                        '    TCH_temp = TCH_temp & "," & freqNo
                        'End If
                    Else
                        BSC_INDEX_CELLNAME_TCH_DIC.Add(lastIndex, TCH_temp)
                        TCH_temp = ""
                        If bcchMark = "YES" Then
                            If BSC_INDEX_CELLNAME_BCCH_DIC.ContainsKey(nowIndex) Then
                                MsgBox("小区 " & nowIndex & " 存在两个BCCH，请核查是否正确！", MsgBoxStyle.Critical)
                            Else
                                BSC_INDEX_CELLNAME_BCCH_DIC.Add(nowIndex, freqNo)
                            End If
                        Else
                            If TCH_temp = "" Then
                                TCH_temp = freqNo
                            Else
                                TCH_temp = TCH_temp & "," & freqNo
                            End If
                        End If
                    End If
                    lastIndex = nowIndex
                    End If
                nowLineNo = nowLineNo + 1
            Loop
            BSC_INDEX_CELLNAME_TCH_DIC.Add(lastIndex, TCH_temp)
            TCH_temp = ""
        Catch ex As Exception
            MsgBox("GTRX文件导入失败，请检查！", MsgBoxStyle.Critical)
            Exit Sub
        Finally
            FileClose(fileNo)
        End Try
    End Sub
End Class
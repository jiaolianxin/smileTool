Imports System.ComponentModel
Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports 优化小工具
Imports Microsoft.Office.Interop

Public Class CDD_parse
    Private totalTime As Integer
    Private Filenames As Object
    Private fso As Object
    Private BSCNAME As String
    Private CDDTime As String
    Private outPath As String
    Private unParseCommand As String
    Private unParseCommand_arr As ArrayList
    '''''''制作Mcom Carrier
    Private DEP_Cell_arr As ArrayList
    Private BSC_arr As ArrayList
    Private DEP_BSC_CELL As ArrayList
    Private CGI_arr As ArrayList
    Private LAI_arr As ArrayList
    Private CI_arr As ArrayList
    Private BCCH_arr As ArrayList
    Private BSIC_arr As ArrayList
    Private CFP_Cell_arr As ArrayList
    Private CFP_BSC_CELL As ArrayList
    Private TCH_arr As ArrayList
    Private HOP_arr As ArrayList
    Private HSN_arr As ArrayList
    Private CHGR_ARR As ArrayList
    Private MBC_CELL_ARR As ArrayList
    Private TRX_ARR As ArrayList
    Private CSYSTYPE_ARR As ArrayList
    Private MaxTCHNo As Integer = 0
    '''''''''''''''''''制作Mcom Neighbour
    Private S_cell_Neighbour As ArrayList
    Private N_cell_Neighbour As ArrayList
    Private MaxNeighborNo As Integer = 0

    '''''''制作CDU检查文件
    Private TCP_CELL_MO_DIC As New Dictionary(Of String, String)
    Private MFP_MO_CDUtype_DIC As New Dictionary(Of String, String)
    Private I As Long = 0
    Private CDU_DIC As New Dictionary(Of String, String)
    Private Final_EQP_DIC As New Dictionary(Of String, String)

    ''''''制作CELL文件
    Private siteFileName As String
    Private cellOfSite_arr As ArrayList
    Private Lat_arr As ArrayList
    Private Lon_arr As ArrayList
    Private Dir_arr As ArrayList
    Private Height_arr As ArrayList
    Private Dtile_arr As ArrayList
    Private SiteName_arr As ArrayList

    ''''''RXMFP告警信息
    Private faultyCode_dic As New Dictionary(Of String, String)

    ''''''RLSTP小区状态
    Private STP_BSC_CELL_STATE As New Dictionary(Of String, String)

    '''''''小区TRXS数（RXMOP:MOTY=RXOTRX;)
    Private TRX_BSC_MO_CELL_DIC As New Dictionary(Of String, String) 'BSC#CELL,BSC#MO

    ''''''小区TRX类型
    Private TRX_TYPE_DIC As New Dictionary(Of String, String)

    ''''''小区DIP信息
    Dim dipDic As New Dictionary(Of String, String)
    Dim dipStateDic As New Dictionary(Of String, String)

    ''''记录temp2执行成功的；
    Dim temp2succArr As New ArrayList



    Private Sub countTotalTime() Handles TimerForCddParse.Tick
        totalTime = totalTime + 1
    End Sub
    Private Sub SelectAll_Click(sender As Object, e As EventArgs) Handles SelectAll.Click
        generateCarrier.Checked = True
        generateCDUtype.Checked = True
        generateCellfile.Checked = True
        generateCellConfig.Checked = True
        generateSeeSiteFile.Checked = True
        LoadFaultName.Checked = True
    End Sub
    Private Sub CancleAll_Click(sender As Object, e As EventArgs) Handles CancleAll.Click
        generateCarrier.Checked = False
        generateCDUtype.Checked = False
        generateCellfile.Checked = False
        generateCellConfig.Checked = False
        generateSeeSiteFile.Checked = False
        LoadFaultName.Checked = False
    End Sub
    Private Sub WhatCanParse_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles WhatCanParse.LinkClicked
        Dim fso As Object
        Dim outPutFile As Object
        Dim resultFile As String
        resultFile = System.Environment.CurrentDirectory & "\目前支持转换的CDD指令.txt"
        MsgBox("将在目录：" & resultFile & "下生成文件，生成后如不需要可以删除！", MsgBoxStyle.Information)
        fso = CreateObject("Scripting.FilesystemObject")
        outPutFile = fso.OpenTextFile(resultFile, 2, 1, 0)
        outPutFile.writeline("IOEXP;")
        outPutFile.writeline("ALLIP;")
        outPutFile.writeline("CACLP;")
        outPutFile.writeline("DBTSP:TAB=AXEPARS;")
        outPutFile.writeline("DTSTP:DIP=ALL;")
        outPutFile.writeLine("DTQUP:DIP=ALL;")
        outPutFile.writeline("EXRPP:RP=ALL;")
        outPutFile.writeline("MGCEP:CELL=ALL;")
        outPutFile.writeline("MGEPP:ID=ALL;")
        outPutFile.writeline("MGIDP;")
        outPutFile.writeline("MGOCP:CELL=ALL;")
        outPutFile.writeline("PLLDP;")
        outPutFile.writeline("RAEPP:ID=ALL;")
        outPutFile.writeline("RFCAP:CELL=ALL;")
        outPutFile.writeline("RLACP:CELL=ALL;")
        outPutFile.writeline("RLADP:SET=ALL;")
        outPutFile.writeline("RLAPP:CELL=ALL;")
        outPutFile.writeline("RLBCP:CELL=ALL;")
        outPutFile.writeline("RLBDP:CELL=ALL;")
        outPutFile.writeline("RLCAP;")
        outPutFile.writeline("RLCCP:CELL=ALL;")
        outPutFile.writeline("RLCDP:CELL=ALL;")
        outPutFile.writeline("RLCFP:CELL=ALL;")
        outPutFile.writeline("RLCHP:CELL=ALL;")
        outPutFile.writeline("RLCLP:CELL=ALL;")
        outPutFile.writeline("RLCPP:CELL=ALL;")
        outPutFile.writeline("RLCPP:CELL=ALL,EXT;")
        outPutFile.writeline("RLCRP:CELL=ALL;")
        outPutFile.writeline("RLCXP:CELL=ALL;")
        outPutFile.writeline("RLDCP;")
        outPutFile.writeline("RLDEP:CELL=ALL;")
        outPutFile.writeline("RLDEP:CELL=ALL,EXT;")
        outPutFile.writeline("RLDGP:CELL=ALL;")
        outPutFile.writeline("RLDHP:CELL=ALL;")
        outPutFile.writeline("RLDMP:CELL=ALL;")
        outPutFile.writeline("RLDRP:CELL=ALL;")
        outPutFile.writeline("RLDTP:CELL=ALL;")
        outPutFile.writeline("RLDUP:CELL=ALL;")
        outPutFile.writeline("RLGAP:CELL=ALL;")
        outPutFile.writeline("RLGRP:CELL=ALL;")
        outPutFile.writeline("RLGSP:CELL=ALL;")
        outPutFile.writeline("RLHBP;")
        outPutFile.writeline("RLHPP:CELL=ALL;")
        outPutFile.writeline("RLIHP:CELL=ALL;")
        outPutFile.writeline("RLIMP:CELL=ALL;")
        outPutFile.writeline("RLLBP;")
        outPutFile.writeline("RLLCP:CELL=ALL;")
        outPutFile.writeline("RLLDP:CELL=ALL;")
        outPutFile.writeline("RLLFP:CELL=ALL;")
        outPutFile.writeline("RLLHP:CELL=ALL;")
        outPutFile.writeline("RLLHP:CELL=ALL,EXT;")
        outPutFile.writeline("RLLLP:CELL=ALL;")
        outPutFile.writeline("RLLOP:CELL=ALL;")
        outPutFile.writeline("RLLOP:CELL=ALL,EXT;")
        outPutFile.writeline("RLLPP:CELL=ALL;")
        outPutFile.writeline("RLLSP;")
        outPutFile.writeline("RLLUP:CELL=ALL;")
        outPutFile.writeline("RLMFP:CELL=ALL;")
        outPutFile.writeline("RLNRP:CELL=ALL;")
        outPutFile.writeline("RLOLP:CELL=ALL;")
        outPutFile.writeline("RLOMP;")
        outPutFile.writeline("RLPBP:CELL=ALL;")
        outPutFile.writeline("RLPCP:CELL=ALL;")
        outPutFile.writeline("RLPRP:CELL=ALL;")
        outPutFile.writeline("RLSBP:CELL=ALL;")
        outPutFile.writeline("RLSLP:CELL=ALL;")
        outPutFile.writeline("RLSMP:CELL=ALL;")
        outPutFile.writeline("RLSSP:CELL=ALL;")
        outPutFile.writeline("RLSTP:CELL=ALL;")
        outPutFile.writeline("RLSUP:CELL=ALL;")
        outPutFile.writeline("RLTDP:MSC=ALL;")
        outPutFile.writeline("RLTYP;")
        outPutFile.writeline("RLUMP:CELL=ALL;")
        outPutFile.writeline("RLVLP;")
        outPutFile.writeline("RRMBP:MSC=ALL;")
        outPutFile.writeline("RRTPP:TRAPOOL=ALL;")
        outPutFile.writeline("RXAPP:MOTY=RXOTG;")
        outPutFile.writeline("RXASP:MOTY=RXOTG;")
        outPutFile.writeline("RXELP:MOTY=RXOTG;")
        outPutFile.writeline("RXMDP:MOTY=RXOTRX;")
        outPutFile.writeline("RXMFP:MOTY=RXOMCTR;")
        outPutFile.writeline("RXMFP:MOTY=RXOMCTR,FAULTY;")
        outPutFile.writeline("RXMFP:MOTY=RXOCF;")
        outPutFile.writeline("RXMFP:MOTY=RXOCF,FAULTY;")
        outPutFile.writeline("RXMFP:MOTY=RXOCON;")
        outPutFile.writeline("RXMFP:MOTY=RXOCON,FAULTY;")
        outPutFile.writeline("RXMFP:MOTY=RXORX;")
        outPutFile.writeline("RXMFP:MOTY=RXORX,FAULTY;")
        outPutFile.writeline("RXMFP:MOTY=RXOTF;")
        outPutFile.writeline("RXMFP:MOTY=RXOTF,FAULTY;")
        outPutFile.writeline("RXMFP:MOTY=RXOTRX;")
        outPutFile.writeline("RXMFP:MOTY=RXOTRX,FAULTY;")
        outPutFile.writeline("RXMFP:MOTY=RXOTS;")
        outPutFile.writeline("RXMFP:MOTY=RXOTS,FAULTY;")
        outPutFile.writeline("RXMFP:MOTY=RXOTX;")
        outPutFile.writeline("RXMFP:MOTY=RXOTX,FAULTY;")
        outPutFile.writeline("RXMOP:MOTY=RXOCF;")
        outPutFile.writeline("RXMOP:MOTY=RXOMCTR;")
        outPutFile.writeline("RXMOP:MOTY=RXORX;")
        outPutFile.writeline("RXMOP:MOTY=RXOTF;")
        outPutFile.writeline("RXMOP:MOTY=RXOTG;")
        outPutFile.writeline("RXMOP:MOTY=RXOTRX;")
        outPutFile.writeline("RXMOP:MOTY=RXOTX;")
        outPutFile.writeline("RXMSP:MOTY=RXOTRX;")
        outPutFile.writeline("RXTCP:MOTY=RXOTG;")
        outPutFile.writeline("SAAEP:SAE=ALL;")

        outPutFile.close()
        outPutFile = Nothing
        fso = Nothing
        Shell("explorer.exe " & resultFile, 1)
    End Sub
    Private Sub start_Click(sender As Object, e As EventArgs) Handles start.Click
        Dim siteFilePath As String
        Dim response As Object

        siteFilePath = ""
        siteFileName = ""
        If BGWParseCDD.IsBusy() Then
            MsgBox("正在进行的CDD解析任务尚未完成，稍候再试吧！", MsgBoxStyle.Critical)
        Else
            MsgBox("请选择需要转换的CDDlog文件！", MsgBoxStyle.Information)
            Filenames = selectFile("选择需要转换的CDDlog文件！", "CDDlog文件(*.log)|*.log|CDD文本文件(*.txt)|*.txt|CDD文本文件(*.*)|*.*", True, True, True)
            If Filenames.length = 0 Then
                Exit Sub
            Else
                If generateCellfile.CheckState = CheckState.Checked Then
                    MsgBox("请选择Mcom Site文件！如果需要显示中文名，请在siteName列加入中文名！", MsgBoxStyle.Information)
                    siteFileName = selectFile("选择需要MCOM Site文件(必须包含:Cell,Longitude,Latitude,Dir,SiteName<可选>）", "Mcom Site文件(*.txt)|*.txt", False, True, True)
                    '     siteFilePath = getSiteFile()
                    If siteFileName.Length = 0 Then
                        MsgBox("如果需要生成CELL文件，必须导入Mcom Site文件！", MsgBoxStyle.Critical)
                        Exit Sub
                    End If
                End If
                If generateCellConfig.CheckState = CheckState.Checked Or generateSeeSiteFile.CheckState = CheckState.Checked Then
                    If siteFileName.Length = 0 Then
                        MsgBox("请选择Mcom Site文件！", MsgBoxStyle.Information)
                        siteFileName = selectFile("选择需要MCOM Site文件", "Mcom Site文件(*.txt)|*.txt", False, True, True)
                        If siteFileName.Length = 0 Then
                            response = MsgBox("如果不导入MCOM site文件，将会影响【小区基础配置信息】与【SeeSite工参文件】中的工参信息！是否继续？", MsgBoxStyle.YesNo + MsgBoxStyle.Critical)
                            If response = vbNo Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                If generateCellConfig.CheckState = CheckState.Checked And generateCDUtype.CheckState = CheckState.Unchecked Then
                    response = MsgBox("不勾选 CDU类型检查，生成的小区基础配置信息表中无CDU类型相关信息！是否继续？", MsgBoxStyle.YesNo + MsgBoxStyle.Critical)
                    If response = vbNo Then
                        Exit Sub
                    End If
                End If
                Hide()
                totalTime = 0
                TimerForCddParse.Start()
                MainForm.mainMessageBox.Clear()
                BGWParseCDD.RunWorkerAsync()
            End If
        End If
    End Sub
    Private Sub BGWParseCDD_DoWork(sender As Object, e As DoWorkEventArgs) Handles BGWParseCDD.DoWork
        parseCDD(sender, e)
        TimerForCddParse.Stop()
    End Sub
    Private Sub BGWParseCDD_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BGWParseCDD.ProgressChanged
        MainForm.MainProgressBar.Value = IIf(CInt(e.ProgressPercentage) > 100, 100, CInt(e.ProgressPercentage))
        If InStr(CStr(e.UserState.ToString), "@") Then
            MainForm.mainMessageBox.AppendText(Split(CStr(e.UserState.ToString), "@")(1) & vbCrLf)
            MainForm.statusBar.Text = Split(CStr(e.UserState.ToString), "@")(0)
        Else
            MainForm.statusBar.Text = CStr(e.UserState.ToString)
        End If
    End Sub
    Private Sub BGWParseCDD_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BGWParseCDD.RunWorkerCompleted
        If e.Cancelled Then
            MainForm.statusBar.Text = "欢迎使用！"
            MainForm.MainProgressBar.Value = 0
        Else
            BGWParseCDD.CancelAsync()
            MainForm.statusBar.Text = "欢迎使用！"
            MainForm.MainProgressBar.Value = 0
        End If
    End Sub
    Private Function getNowLine()
        If EOF(1) Then
            getNowLine = ""
        Else
            getNowLine = LineInput(1)
        End If
    End Function
    Private Function isEnd(str As String)
        If Strings.Left(Trim(str), 3) = "END" Or Strings.Left(Trim(str), 1) = "<" Or Trim(str) = "NOT ACCEPTED" Or _
      Trim(str) = "FAULT INTERRUPT" Or InStr(Trim(str), "FUNCTION BUSY") Then
            isEnd = True
        Else
            isEnd = False
        End If
    End Function
    Private Sub getBSC(fileName As String)
        Dim nowLineContent As String
        'Dim existsMark As Boolean
        Dim x As Integer
        Dim CddTimeExistsMark As Boolean = False
        Dim IOEXP_BSCNAME As String
        Dim TOP_BSCNAME As String
        Dim FILENAME_BSCNAME As String
        Dim EAW_BSCNAME As String
        Dim tempLine As String

        IOEXP_BSCNAME = ""
        TOP_BSCNAME = ""
        FILENAME_BSCNAME = ""
        EAW_BSCNAME = ""
        CDDTime = ""
        BSCNAME = ""

        FileOpen(2, fileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        Try
            Do While Not EOF(2) And x < 60
                nowLineContent = LineInput(2)
                If UCase(Strings.Left(Trim(nowLineContent), 7)) = "<IOEXP;" Or UCase(Strings.Right(Trim(nowLineContent), 6)) = "IOEXP;" Then
                    If InStr(nowLineContent, "/") Then
                        IOEXP_BSCNAME = Strings.Left(nowLineContent, InStr(nowLineContent, "/") - 1)
                        'If Strings.right(BSCNAME, 3) = "R12" Then
                        '    BSCNAME = Strings.left(BSCNAME, Len(BSCNAME) - 3)
                        'End If
                    ElseIf InStr(nowLineContent, "_") Then
                        IOEXP_BSCNAME = Strings.Left(nowLineContent, InStr(nowLineContent, "_") - 1)
                    End If
                ElseIf UCase(Strings.Left(Trim(nowLineContent), 7)) = "<CACLP;" Or UCase(Strings.Right(Trim(nowLineContent), 6)) = "CACLP;" Then
                    CddTimeExistsMark = True
                ElseIf UCase(Strings.Left(Trim(nowLineContent), 16)) = "*** CONNECTED TO" Then
                    '   TOP_BSCNAME = Trim(Mid(Trim(nowLineContent), 17, 9))
                    tempLine = UCase(Trim(nowLineContent))
                    TOP_BSCNAME = Trim(Mid(tempLine, InStr(tempLine, " TO") + 4, InStr(tempLine, " ***") - InStr(tempLine, " TO") - 4))
                End If
                If CddTimeExistsMark Then
                    If Trim(nowLineContent) = "DATE     TIME     SUMMERTIME     DAY      DCAT" Then
                        nowLineContent = LineInput(2)
                        CDDTime = Trim(Mid(nowLineContent, 1, 9))
                    End If
                End If
                x = x + 1
            Loop
            FILENAME_BSCNAME = UCase(Strings.Right(Trim(fileName), Len(Trim(fileName)) - InStrRev(Trim(fileName), "\")))
            If InStr(FILENAME_BSCNAME, ".") Then
                FILENAME_BSCNAME = Strings.Split(FILENAME_BSCNAME, ".")(0)
            End If
            If CDDTime = "" Then
                CDDTime = Format(FileDateTime(fileName), "yyMMdd")
            End If
            If useCONNECTED.Checked And TOP_BSCNAME <> "" Then
                BSCNAME = TOP_BSCNAME
            ElseIf useFileName.Checked And FILENAME_BSCNAME <> "" Then
                BSCNAME = FILENAME_BSCNAME
            ElseIf useIOEXP.Checked And IOEXP_BSCNAME <> "" Then
                BSCNAME = IOEXP_BSCNAME
            Else
                BSCNAME = IOEXP_BSCNAME
            End If
            If BSCNAME = "" Then
                BSCNAME = FILENAME_BSCNAME
            End If
        Catch ex As Exception
            MsgBox("获取BSC名称及CDDlog生成时间时出错！" & ex.ToString, MsgBoxStyle.Critical)
            Exit Sub
        Finally
            FileClose(2)
        End Try

    End Sub
    'Private Sub getBSC(fileName As String)                '旧版本的
    '    Dim nowLineContent As String
    '    Dim existsMark As Boolean
    '    Dim x As Integer
    '    Dim CddTimeExistsMark As Boolean = False
    '    Dim IOEXP_BSCNAME As String
    '    Dim TOP_BSCNAME As String
    '    Dim FILENAME_BSCNAME As String

    '    IOEXP_BSCNAME = ""
    '    TOP_BSCNAME = ""
    '    FILENAME_BSCNAME = ""
    '    CDDTime = ""
    '    BSCNAME = ""

    '    FileOpen(2, fileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
    '    Try
    '        Do While Not EOF(2) And x < 60
    '            nowLineContent = LineInput(2)
    '            If UCase(Strings.Left(Trim(nowLineContent), 7)) = "<IOEXP;" Or UCase(Strings.Right(Trim(nowLineContent), 6)) = "IOEXP;" Then
    '                existsMark = True
    '            ElseIf UCase(Strings.Left(Trim(nowLineContent), 7)) = "<CACLP;" Or UCase(Strings.Right(Trim(nowLineContent), 6)) = "CACLP;" Then
    '                CddTimeExistsMark = True
    '            ElseIf UCase(Strings.Left(Trim(nowLineContent), 16)) = "*** CONNECTED TO" Then
    '                BSCNAME = Trim(Mid(Trim(nowLineContent), 17, 9))
    '            End If
    '            If existsMark Then
    '                If InStr(nowLineContent, "/") Then
    '                    BSCNAME = Strings.Left(nowLineContent, InStr(nowLineContent, "/") - 1)
    '                    'If Strings.right(BSCNAME, 3) = "R12" Then
    '                    '    BSCNAME = Strings.left(BSCNAME, Len(BSCNAME) - 3)
    '                    'End If
    '                ElseIf InStr(nowLineContent, "_") Then
    '                    BSCNAME = Strings.Left(nowLineContent, InStr(nowLineContent, "_") - 1)
    '                End If
    '            End If
    '            If CddTimeExistsMark Then
    '                If Trim(nowLineContent) = "DATE     TIME     SUMMERTIME     DAY      DCAT" Then
    '                    nowLineContent = LineInput(2)
    '                    CDDTime = Trim(Mid(nowLineContent, 1, 9))
    '                End If
    '            End If
    '            x = x + 1
    '        Loop
    '        If BSCNAME = "" Then
    '            BSCNAME = UCase(Strings.Right(Trim(fileName), Len(Trim(fileName)) - InStrRev(Trim(fileName), "\")))
    '            If InStr(BSCNAME, ".") Then
    '                BSCNAME = Strings.Split(BSCNAME, ".")(0)
    '            End If
    '        End If
    '        If CDDTime = "" Then
    '            CDDTime = Format(FileDateTime(fileName), "yyMMdd")
    '        End If
    '    Catch ex As Exception
    '        MsgBox("获取BSC名称及CDDlog生成时间时出错！" & ex.ToString, MsgBoxStyle.Critical)
    '        Exit Sub
    '    Finally
    '        FileClose(2)
    '    End Try

    'End Sub
    Private Sub makeFaultyCodeDic()
        If faultyCode_dic.Count = 0 Then
            faultyCode_dic.Add("CF1A0", "重启.自动恢复")
            faultyCode_dic.Add("CF1A1", "重启.加电")
            faultyCode_dic.Add("CF1A10", "室内温度超出安全范围")
            faultyCode_dic.Add("CF1A12", "直流电压超出范围")
            faultyCode_dic.Add("CF1A14", "总线故障")
            faultyCode_dic.Add("CF1A15", "RBS数据库或IDB损坏")
            faultyCode_dic.Add("CF1A16", "RU数据库遭到破坏")
            faultyCode_dic.Add("CF1A17", "硬件和IDB不相容")
            faultyCode_dic.Add("CF1A18", "内部配置失败")
            faultyCode_dic.Add("CF1A19", "HW and SW Inconsistent")
            faultyCode_dic.Add("CF1A2", "重启.切换")
            faultyCode_dic.Add("CF1A21", "硬件故障")
            faultyCode_dic.Add("CF1A22", "无法计算开始时间")
            faultyCode_dic.Add("CF1A23", "时钟分配故障")
            faultyCode_dic.Add("CF1A24", "Temperature Close to Destructive Limit")
            faultyCode_dic.Add("CF1A3", "重启.看门狗")
            faultyCode_dic.Add("CF1A4", "软件故障重启")
            faultyCode_dic.Add("CF1A5", "RAM故障重启")
            faultyCode_dic.Add("CF1A6", "重启.内部功能修改")
            faultyCode_dic.Add("CF1A7", "X-Bus总线故障")
            faultyCode_dic.Add("CF1A8", "定时装置VCO故障")
            faultyCode_dic.Add("CF1A9", "时钟分配或定时总线故障")
            faultyCode_dic.Add("CF2A0", "重启.重启动尝试失败")
            faultyCode_dic.Add("CF2A1", "重启.加电")
            faultyCode_dic.Add("CF2A10", "DXU optional EEPROM校验故障、DXU容错校验错误")
            faultyCode_dic.Add("CF2A11", "ALNA/TMA故障")
            faultyCode_dic.Add("CF2A12", "RX最大/最小增益超标")
            faultyCode_dic.Add("CF2A13", "定时装置VCO老化")
            faultyCode_dic.Add("CF2A14", "CDU无法进行监管/通信")
            faultyCode_dic.Add("CF2A15", "无法监测驻波比/输出功率")
            faultyCode_dic.Add("CF2A16", "超出正常的室内温度范围")
            faultyCode_dic.Add("CF2A17", "超出正常的室内湿度范围")
            faultyCode_dic.Add("CF2A18", "直流电压超出正常范围")
            faultyCode_dic.Add("CF2A19", "电源和环境系统处于独立模式")
            faultyCode_dic.Add("CF2A2", "重启.切换")
            faultyCode_dic.Add("CF2A20", "外部电源故障")
            faultyCode_dic.Add("CF2A21", "内部电源容量下降")
            faultyCode_dic.Add("CF2A22", "备用电池的电量下降")
            faultyCode_dic.Add("CF2A23", "环境温度适应能力下降")
            faultyCode_dic.Add("CF2A24", "硬件故障")
            faultyCode_dic.Add("CF2A25", "DXU或ECU负载文件丢失")
            faultyCode_dic.Add("CF2A26", "温度传感器故障")
            faultyCode_dic.Add("CF2A27", "系统电压传感器故障")
            faultyCode_dic.Add("CF2A28", "A/D转换器故障")
            faultyCode_dic.Add("CF2A29", "变阻器错误")
            faultyCode_dic.Add("CF2A3", "重启.看门狗")
            faultyCode_dic.Add("CF2A30", "BUS总线故障")
            faultyCode_dic.Add("CF2A31", "频繁出现的软件故障")
            faultyCode_dic.Add("CF2A32", "内存损坏")
            faultyCode_dic.Add("CF2A33", "RX分级接受丢失")
            faultyCode_dic.Add("CF2A34", "输出电压故障")
            faultyCode_dic.Add("CF2A35", "Optional synchronization source、可选同步源")
            faultyCode_dic.Add("CF2A36", "RU数据库遭到破坏")
            faultyCode_dic.Add("CF2A37", "断路器或保险丝断开")
            faultyCode_dic.Add("CF2A38", "使用默认值")
            faultyCode_dic.Add("CF2A39", "RX线缆连接中断")
            faultyCode_dic.Add("CF2A4", "重启.软件故障")
            faultyCode_dic.Add("CF2A40", "重启.DXU链路丢失")
            faultyCode_dic.Add("CF2A41", "与TRU失去通信")
            faultyCode_dic.Add("CF2A42", "与ECU失去通信")
            faultyCode_dic.Add("CF2A43", "内部配置失败")
            faultyCode_dic.Add("CF2A44", "ESB分配失败")
            faultyCode_dic.Add("CF2A45", "温度过高")
            faultyCode_dic.Add("CF2A46", "DB参数故障")
            faultyCode_dic.Add("CF2A47", "天线跳频失败")
            faultyCode_dic.Add("CF2A48", "GPS 同步故障")
            faultyCode_dic.Add("CF2A49", "电池备用时间比预定时间短")
            faultyCode_dic.Add("CF2A5", "重启.内存故障")
            faultyCode_dic.Add("CF2A50", "RBS 在电池上运行")
            faultyCode_dic.Add("CF2A51", "TMA无法进行监管/通信")
            faultyCode_dic.Add("CF2A52", "CXU监测/通信丢失")
            faultyCode_dic.Add("CF2A53", "硬件和IDB不兼容")
            faultyCode_dic.Add("CF2A54", "定时总线故障、时钟总线错误")
            faultyCode_dic.Add("CF2A55", "X-Bus总线故障")
            faultyCode_dic.Add("CF2A57", "RX分路接收不平衡")
            faultyCode_dic.Add("CF2A58", "Disconnected")
            faultyCode_dic.Add("CF2A59", "Operating Temperature Too High、 Main Load Disconnected ")
            faultyCode_dic.Add("CF2A6", "重启.内部功能改变")
            faultyCode_dic.Add("CF2A60", "Operating Temperature Too High、 Battery Disconnected")
            faultyCode_dic.Add("CF2A61", "Operating Temperature Too High、 Capacity Reduced Related Rus")
            faultyCode_dic.Add("CF2A62", "Operating Temperature Too Low、 Capacity Reduced Related Rus")
            faultyCode_dic.Add("CF2A63", "Operating Temperature Too High、 No Service Related Rus")
            faultyCode_dic.Add("CF2A64", "Operating Temperature Too Low、 Communication Lost")
            faultyCode_dic.Add("CF2A65", "Battery Voltage Too Low、 Main Load Disconnected")
            faultyCode_dic.Add("CF2A66", "Battery Voltage Too Low、 Prio Load Disconnected")
            faultyCode_dic.Add("CF2A67", "System Undervoltage")
            faultyCode_dic.Add("CF2A68", "System Overvoltage")
            faultyCode_dic.Add("CF2A69", "Cabinet Product Data Mismatch")
            faultyCode_dic.Add("CF2A7", "RX内部放大器错误")
            faultyCode_dic.Add("CF2A70", "Battery Missing")
            faultyCode_dic.Add("CF2A71", "Low Battery Capacity")
            faultyCode_dic.Add("CF2A72", "Software Load of RUS Failed")
            faultyCode_dic.Add("CF2A73", "Degraded or Lost Communication to Radio Unit")
            faultyCode_dic.Add("CF2A79", "Configuration Fault of CPRI System")
            faultyCode_dic.Add("CF2A8", "驻波比超出限制")
            faultyCode_dic.Add("CF2A80", "Antenna System DC Power Supply Overloaded")
            faultyCode_dic.Add("CF2A81", "Primary Node Disconnected")
            faultyCode_dic.Add("CF2A82", "Radio Unit Incompatible")
            faultyCode_dic.Add("CF2A9", "功率超出限制、超过电源限制")
            faultyCode_dic.Add("CFE1B2", "L/R SWI(本地模式BTS)、基站本地断开链接")
            faultyCode_dic.Add("CFE1B4", "L/R SWI(本地模式BTS)")
            faultyCode_dic.Add("CFE1B5", "L/R TI(链路断开时.本地转换成远程模式)、DXU向远端转换时通讯丢失")
            faultyCode_dic.Add("CFE1B9", "Smoke Alarm")
            faultyCode_dic.Add("CFE2B10", "主电源故障(外部电源故障)")
            faultyCode_dic.Add("CFE2B11", "ALNA/TMA 故障")
            faultyCode_dic.Add("CFE2B12", "ALNA/TMA性能降低")
            faultyCode_dic.Add("CFE2B13", "辅助设备故障")
            faultyCode_dic.Add("CFE2B14", "备用电池的外部保险丝故障")
            faultyCode_dic.Add("CFE2B3", "Smoke Alarm Faulty")
            faultyCode_dic.Add("CFE2B6", "操作链路中断")
            faultyCode_dic.Add("CFE2B9", "RBS door(机柜门打开)")
            faultyCode_dic.Add("CFRU0", "DXU(MU.IXU)")
            faultyCode_dic.Add("CFRU1", "ECU")
            faultyCode_dic.Add("CFRU10", "ACCU")
            faultyCode_dic.Add("CFRU11", "Active cooler")
            faultyCode_dic.Add("CFRU12", "ALNA/TMA A")
            faultyCode_dic.Add("CFRU13", "ALNA/TMA B")
            faultyCode_dic.Add("CFRU14", "Battery")
            faultyCode_dic.Add("CFRU15", "Fan")
            faultyCode_dic.Add("CFRU16", "Heater")
            faultyCode_dic.Add("CFRU17", "Heat exchanger external fan")
            faultyCode_dic.Add("CFRU18", "Heat exchanger internal fan")
            faultyCode_dic.Add("CFRU19", "Humidity sensor")
            faultyCode_dic.Add("CFRU20", "TMA CM")
            faultyCode_dic.Add("CFRU21", "Temperature sensor")
            faultyCode_dic.Add("CFRU22", "CDU HLOUT HLIN cable")
            faultyCode_dic.Add("CFRU23", "CDU RX cable")
            faultyCode_dic.Add("CFRU24", "CU")
            faultyCode_dic.Add("CFRU25", "DU")
            faultyCode_dic.Add("CFRU26", "FU")
            faultyCode_dic.Add("CFRU27", "FU CU/CDU PFWD cable")
            faultyCode_dic.Add("CFRU28", "FU CU/CDU PREFL cable")
            faultyCode_dic.Add("CFRU29", "HLIN cable")
            faultyCode_dic.Add("CFRU3", "Y-link")
            faultyCode_dic.Add("CFRU30", "CDU bus/IOM bus")
            faultyCode_dic.Add("CFRU31", "Environment")
            faultyCode_dic.Add("CFRU32", "Local bus")
            faultyCode_dic.Add("CFRU33", "EPC bus/Power communication loop")
            faultyCode_dic.Add("CFRU34", "RBS DB/IDB")
            faultyCode_dic.Add("CFRU36", "Timing bus")
            faultyCode_dic.Add("CFRU37", "CDU CXU RXA cable")
            faultyCode_dic.Add("CFRU38", "CDU CXU RXB cable")
            faultyCode_dic.Add("CFRU39", "X-Bus")
            faultyCode_dic.Add("CFRU40", "天线")
            faultyCode_dic.Add("CFRU41", "PSU DC cable")
            faultyCode_dic.Add("CFRU42", "CXU")
            faultyCode_dic.Add("CFRU43", "Flash card")
            faultyCode_dic.Add("CFRU44", "OVP")
            faultyCode_dic.Add("CFRU45", "电池温度传感器")
            faultyCode_dic.Add("CFRU46", "FCU")
            faultyCode_dic.Add("CFRU47", "TMA-CM cable")
            faultyCode_dic.Add("CFRU48", "GPS接收器")
            faultyCode_dic.Add("CFRU49", "GPS-DXU连接线缆")
            faultyCode_dic.Add("CFRU5", "CDU")
            faultyCode_dic.Add("CFRU50", "Active cooler fan")
            faultyCode_dic.Add("CFRU51", "BFU fuse or circuit breaker")
            faultyCode_dic.Add("CFRU52", "CDU PFWD cable")
            faultyCode_dic.Add("CFRU53", "CDU PREFL cable")
            faultyCode_dic.Add("CFRU54", "IOM bus")
            faultyCode_dic.Add("CFRU55", "ASU RXA units or cable")
            faultyCode_dic.Add("CFRU56", "ASU RXA units or cable")
            faultyCode_dic.Add("CFRU57", "ASU CDU RXA cable")
            faultyCode_dic.Add("CFRU58", "ASU CDU RXB cable")
            faultyCode_dic.Add("CFRU59", "MCPA")
            faultyCode_dic.Add("CFRU6", "CCU")
            faultyCode_dic.Add("CFRU7", "PSU")
            faultyCode_dic.Add("CFRU8", "BFU")
            faultyCode_dic.Add("CFRU9", "BDM")
            faultyCode_dic.Add("CONE1B8", "LAPD Q CG(LAPD 排队拥塞)")
            faultyCode_dic.Add("CONE2B8", "LAPD Q CG(LAPD 排队拥塞)")
            faultyCode_dic.Add("RX1B0", "RXDA/LNA放大器实时故障、RX内部放大器故障")
            faultyCode_dic.Add("RX1B1", "ALNA/TMA故障")
            faultyCode_dic.Add("RX1B10", "RX初始化故障")
            faultyCode_dic.Add("RX1B11", "CDU输出电压故障")
            faultyCode_dic.Add("RX1B12", "TMA CM输出电压故障")
            faultyCode_dic.Add("RX1B13", "RX接收支路故障")
            faultyCode_dic.Add("RX1B14", "CDU无法进行监测/通信")
            faultyCode_dic.Add("RX1B15", "RX温度过高")
            faultyCode_dic.Add("RX1B17", "TMA监测故障")
            faultyCode_dic.Add("RX1B18", "TMA电源分配故障")
            faultyCode_dic.Add("RX1B19", "滤波器负载文件校验故障")
            faultyCode_dic.Add("RX1B20", "RX监测丢失")
            faultyCode_dic.Add("RX1B21", "Traffic Lost Uplink")
            faultyCode_dic.Add("RX1B22", "Antenna System DC Power Supply Overloaded")
            faultyCode_dic.Add("RX1B23", "Radio Unit Antenna System Output Voltage Fault")
            faultyCode_dic.Add("RX1B3", "RX EEPROM校验故障")
            faultyCode_dic.Add("RX1B4", "RX配置表校验错误")
            faultyCode_dic.Add("RX1B47", "RX辅助设备故障 ")
            faultyCode_dic.Add("RX1B5", "RX合成器A/B未锁定")
            faultyCode_dic.Add("RX1B6", "RX合成器C未锁定")
            faultyCode_dic.Add("RX1B7", "RX通信故障")
            faultyCode_dic.Add("RX1B8", "RX内部电压故障")
            faultyCode_dic.Add("RX1B9", "RX线缆连接中断")
            faultyCode_dic.Add("RX2A0", "CXU无法进行监测/通信")
            faultyCode_dic.Add("RX2A1", "无法找到A接收器端的RX路径")
            faultyCode_dic.Add("RX2A2", "无法找到B接收器端的RX路径")
            faultyCode_dic.Add("RX2A3", "RXC路接收丢失")
            faultyCode_dic.Add("RX2A4", "RXD路接收丢失")
            faultyCode_dic.Add("RX2A5", "A路接收不平衡")
            faultyCode_dic.Add("RX2A6", "B路接收不平衡")
            faultyCode_dic.Add("RX2A7", "C路接收不平衡")
            faultyCode_dic.Add("RX2A8", "D路接收不平衡")
            faultyCode_dic.Add("TF1A0", "Temperature Below Operational Limit")
            faultyCode_dic.Add("TF1A1", "Temperature Above Operational Limit")
            faultyCode_dic.Add("TF1B0", "丢失可选同步源")
            faultyCode_dic.Add("TF1B1", "DXU容错校验错误")
            faultyCode_dic.Add("TF1B2", "GPS 同步故障")
            faultyCode_dic.Add("TF2A0", "帧启动偏移故障")
            faultyCode_dic.Add("TFE1B0", "EXT 同步(无可用的外部引用)")
            faultyCode_dic.Add("TFE1B1", "PCM 同步(无可用的PCM引用)")
            faultyCode_dic.Add("TFE1B6", "多时钟主管理")
            faultyCode_dic.Add("TFE1B7", "ESB测量失败")
            faultyCode_dic.Add("TFE2B0", "EXT 同步(无可用的外部引用)")
            faultyCode_dic.Add("TFE2B1", "PCM 同步(无可用的PCM引用)")
            faultyCode_dic.Add("TFE2B7", "ESB测量失败")
            faultyCode_dic.Add("TRX1A0", "重启、自动恢复")
            faultyCode_dic.Add("TRX1A1", "重启.加电")
            faultyCode_dic.Add("TRX1A10", "RX通讯错误")
            faultyCode_dic.Add("TRX1A11", "DSP-CPU通信故障")
            faultyCode_dic.Add("TRX1A12", "话务信道故障")
            faultyCode_dic.Add("TRX1A13", "RF环路测试故障")
            faultyCode_dic.Add("TRX1A14", "RU数据库遭到破坏")
            faultyCode_dic.Add("TRX1A15", "X-bus通信故障、交叉总线通信故障")
            faultyCode_dic.Add("TRX1A16", "启动故障")
            faultyCode_dic.Add("TRX1A17", "X-接口故障")
            faultyCode_dic.Add("TRX1A18", "DSP故障")
            faultyCode_dic.Add("TRX1A19", "重启.DXU 链路丢失")
            faultyCode_dic.Add("TRX1A2", "重启.切换")
            faultyCode_dic.Add("TRX1A20", "硬件和IDB不相容")
            faultyCode_dic.Add("TRX1A21", "内部配置失败")
            faultyCode_dic.Add("TRX1A22", "电压供应故障")
            faultyCode_dic.Add("TRX1A23", "空中接口计时器丢失")
            faultyCode_dic.Add("TRX1A24", "温度过高")
            faultyCode_dic.Add("TRX1A25", "TX/RX通信故障")
            faultyCode_dic.Add("TRX1A26", "无线控制系统负载")
            faultyCode_dic.Add("TRX1A27", "下行链路丢失")
            faultyCode_dic.Add("TRX1A28", "上行链路丢失")
            faultyCode_dic.Add("TRX1A29", "Y-Link通信故障")
            faultyCode_dic.Add("TRX1A3", "重启.看门狗")
            faultyCode_dic.Add("TRX1A30", "DSP RAM软件错误")
            faultyCode_dic.Add("TRX1A31", "内存故障")
            faultyCode_dic.Add("TRX1A32", "偶合开关丢失")
            faultyCode_dic.Add("TRX1A33", "Low Temperature")
            faultyCode_dic.Add("TRX1A34", "Radio Unit HW Fault")
            faultyCode_dic.Add("TRX1A35", "Radio Unit Fault")
            faultyCode_dic.Add("TRX1A36", "Lost Communication to Radio Unit")
            faultyCode_dic.Add("TRX1A4", "软件故障重启")
            faultyCode_dic.Add("TRX1A5", "RAM故障重启")
            faultyCode_dic.Add("TRX1A6", "内部功能修改重启")
            faultyCode_dic.Add("TRX1A8", "时钟接收故障")
            faultyCode_dic.Add("TRX1A9", "信令处理故障")
            faultyCode_dic.Add("TRX1B0", "CDU/合路器不能使用")
            faultyCode_dic.Add("TRX1B1", "室内温度超过安全范围")
            faultyCode_dic.Add("TRX1B10", "时钟接收故障")
            faultyCode_dic.Add("TRX1B11", "X-Bus总线通信故障")
            faultyCode_dic.Add("TRX1B3", "直流电压超过安全范围")
            faultyCode_dic.Add("TRX1B7", "TX地址冲突")
            faultyCode_dic.Add("TRX1B8", "Y-link通信故障")
            faultyCode_dic.Add("TRX1B9", "Y-link通信丢失")
            faultyCode_dic.Add("TRX2A0", "RX线缆连接中断")
            faultyCode_dic.Add("TRX2A1", "RX EEPROM校验故障")
            faultyCode_dic.Add("TRX2A10", "TX内部电压故障")
            faultyCode_dic.Add("TRX2A11", "TX温度过高")
            faultyCode_dic.Add("TRX2A12", "TX输出功率超限")
            faultyCode_dic.Add("TRX2A13", "TX饱和(saturation)")
            faultyCode_dic.Add("TRX2A14", "电压供应故障")
            faultyCode_dic.Add("TRX2A15", "驻波比/输出功率监测丢失")
            faultyCode_dic.Add("TRX2A16", "内存遭到破坏")
            faultyCode_dic.Add("TRX2A17", "TRU负载文件丢失")
            faultyCode_dic.Add("TRX2A18", "DSP故障")
            faultyCode_dic.Add("TRX2A19", "频繁出现的软件故障")
            faultyCode_dic.Add("TRX2A2", "RX配置表校验故障")
            faultyCode_dic.Add("TRX2A20", "RX初始化失败")
            faultyCode_dic.Add("TRX2A21", "TX初始化失败")
            faultyCode_dic.Add("TRX2A22", "CDU总线通信故障")
            faultyCode_dic.Add("TRX2A23", "使用默认值")
            faultyCode_dic.Add("TRX2A24", "Radio Unit Antenna System Output Voltage Fault")
            faultyCode_dic.Add("TRX2A25", "TX最高功率限制")
            faultyCode_dic.Add("TRX2A26", "DB参数错误")
            faultyCode_dic.Add("TRX2A27", "RX接收之路故障")
            faultyCode_dic.Add("TRX2A29", "功率放大器故障")
            faultyCode_dic.Add("TRX2A3", "RX合成器未锁定")
            faultyCode_dic.Add("TRX2A30", "CPU温度过高")
            faultyCode_dic.Add("TRX2A31", "TX/RX通信故障")
            faultyCode_dic.Add("TRX2A32", "RX温度过高")
            faultyCode_dic.Add("TRX2A33", "TRX内部通信故障")
            faultyCode_dic.Add("TRX2A34", "驻波比超出限制")
            faultyCode_dic.Add("TRX2A35", "话务突发抑制")
            faultyCode_dic.Add("TRX2A36", "RX滤波器负载文件校验故障")
            faultyCode_dic.Add("TRX2A37", "RX内部放大器故障")
            faultyCode_dic.Add("TRX2A38", "环境适应能力下降")
            faultyCode_dic.Add("TRX2A39", "RF环路测试故障.RX降级")
            faultyCode_dic.Add("TRX2A4", "RX内部电压故障")
            faultyCode_dic.Add("TRX2A40", "内存故障")
            faultyCode_dic.Add("TRX2A41", "IR内存未开启")
            faultyCode_dic.Add("TRX2A42", "偶合开关与IDB不匹配")
            faultyCode_dic.Add("TRX2A43", "内部偶合开关损坏、更换载频")
            faultyCode_dic.Add("TRX2A44", "TX Low Temperature")
            faultyCode_dic.Add("TRX2A45", "Radio Unit HW Fault")
            faultyCode_dic.Add("TRX2A46", "Traffic Performance Uplink")
            faultyCode_dic.Add("TRX2A47", "Internal Configuration Failed")
            faultyCode_dic.Add("TRX2A5", "RX通信故障")
            faultyCode_dic.Add("TRX2A6", "TX通信故障")
            faultyCode_dic.Add("TRX2A7", "TX EEPROM校验故障")
            faultyCode_dic.Add("TRX2A8", "TX配置表校验故障")
            faultyCode_dic.Add("TRX2A9", "TX合成器未锁定")
            faultyCode_dic.Add("TRXE1B4", "L/R SWI(TRU 处于本地模式)")
            faultyCode_dic.Add("TRXE1B5", "L/R TI(链路断开时.本地转换成远程模式)")
            faultyCode_dic.Add("TRXE2B16", "TS0 TRA Lost")
            faultyCode_dic.Add("TRXE2B17", "TS0 TRA Lost")
            faultyCode_dic.Add("TRXE2B18", "TS0 PCU Lost")
            faultyCode_dic.Add("TRXE2B20", "TS1 TRA Lost")
            faultyCode_dic.Add("TRXE2B21", "TS1 TRA Lost")
            faultyCode_dic.Add("TRXE2B22", "TS1 PCU Lost")
            faultyCode_dic.Add("TRXE2B24", "TS2 TRA Lost")
            faultyCode_dic.Add("TRXE2B25", "TS2 TRA Lost")
            faultyCode_dic.Add("TRXE2B26", "TS2 PCU Lost")
            faultyCode_dic.Add("TRXE2B28", "TS3 TRA Lost")
            faultyCode_dic.Add("TRXE2B29", "TS3 TRA Lost")
            faultyCode_dic.Add("TRXE2B30", "TS3 PCU Lost")
            faultyCode_dic.Add("TRXE2B32", "TS4 TRA Lost")
            faultyCode_dic.Add("TRXE2B33", "TS4 TRA Lost")
            faultyCode_dic.Add("TRXE2B34", "TS4 PCU Lost")
            faultyCode_dic.Add("TRXE2B36", "TS5 TRA Lost")
            faultyCode_dic.Add("TRXE2B37", "TS5 TRA Lost")
            faultyCode_dic.Add("TRXE2B38", "TS5 PCU Lost")
            faultyCode_dic.Add("TRXE2B40", "TS6 TRA Lost")
            faultyCode_dic.Add("TRXE2B41", "TS6 TRA Lost")
            faultyCode_dic.Add("TRXE2B42", "TS6 PCU Lost")
            faultyCode_dic.Add("TRXE2B44", "TS7 TRA Lost")
            faultyCode_dic.Add("TRXE2B45", "TS7 TRA Lost")
            faultyCode_dic.Add("TRXE2B46", "TS7 PCU Lost")
            faultyCode_dic.Add("TRXRU0", "TRU.dTRU or ATRU")
            faultyCode_dic.Add("TRXRU10", "CDU to TRU PFWD cable")
            faultyCode_dic.Add("TRXRU11", "CDU to TRU PREFL cable")
            faultyCode_dic.Add("TRXRU12", "CDU to TRU RXA cable")
            faultyCode_dic.Add("TRXRU13", "CDU to TRU RXB cable")
            faultyCode_dic.Add("TRXRU14", "CDU-Splitter cable RXA cable")
            faultyCode_dic.Add("TRXRU15", "CDU-Splitter cable RXB cable")
            faultyCode_dic.Add("TRXRU16", "CDU to TRU TX cable")
            faultyCode_dic.Add("TRXRU17", "CXU-Splitter cable RXA cable")
            faultyCode_dic.Add("TRXRU18", "CXU-Splitter cable RXB cable")
            faultyCode_dic.Add("TRXRU19", "Splitter-DRU cable RXA cable")
            faultyCode_dic.Add("TRXRU20", "Splitter-DRU cable RXB cable")
            faultyCode_dic.Add("TRXRU21", "DRU-DRU RXA cable")
            faultyCode_dic.Add("TRXRU22", "DRU-DRU RXB cable")
            faultyCode_dic.Add("TRXRU23", "HCU TRU TX cable or HCU or CDU HCU TX cable")
            faultyCode_dic.Add("TRXRU3", "CXU-TRU RXA cable")
            faultyCode_dic.Add("TRXRU4", "CXU-TRU RXB cable")
            faultyCode_dic.Add("TSE1B3", "TS退服或闭塞、TRA/PCU（无法进行远程代码转换机/PCU通信）")
            faultyCode_dic.Add("TX1A0", "TX 冲突")
            faultyCode_dic.Add("TX1A1", "内部偶合开关损坏、更换载频")
            faultyCode_dic.Add("TX1A2", "偶合开关与IDB不匹配")
            faultyCode_dic.Add("TX1B0", "CU/CDU不可用")
            faultyCode_dic.Add("TX1B1", "CDU/合路器驻波比超限")
            faultyCode_dic.Add("TX1B10", "TX通信故障")
            faultyCode_dic.Add("TX1B11", "TX内部电压故障")
            faultyCode_dic.Add("TX1B12", "TX温度过高")
            faultyCode_dic.Add("TX1B13", "TX输出功率超限")
            faultyCode_dic.Add("TX1B14", "TX 饱和(saturation)")
            faultyCode_dic.Add("TX1B15", "电源供应故障")
            faultyCode_dic.Add("TX1B16", "电源模块未启用")
            faultyCode_dic.Add("TX1B17", "TX初始化故障")
            faultyCode_dic.Add("TX1B18", "CU/CDU硬件故障")
            faultyCode_dic.Add("TX1B19", "CU软件load/start故障")
            faultyCode_dic.Add("TX1B2", "CDU输出功率超限")
            faultyCode_dic.Add("TX1B20", "CU输入功率故障")
            faultyCode_dic.Add("TX1B21", "CU驻留故障")
            faultyCode_dic.Add("TX1B22", "驻波比/输出功率监测丢失")
            faultyCode_dic.Add("TX1B23", "CU重启.加电")
            faultyCode_dic.Add("TX1B24", "CU 重启.通信故障")
            faultyCode_dic.Add("TX1B25", "CU 重启.看门狗")
            faultyCode_dic.Add("TX1B26", "CU/CDU微调故障")
            faultyCode_dic.Add("TX1B27", "TX最高功率限制")
            faultyCode_dic.Add("TX1B28", "CDU温度过高")
            faultyCode_dic.Add("TX1B30", "TX CDU功率控制故障")
            faultyCode_dic.Add("TX1B31", "功率放大器故障")
            faultyCode_dic.Add("TX1B32", "TX低温")
            faultyCode_dic.Add("TX1B33", "CDU-总线通信故障")
            faultyCode_dic.Add("TX1B34", "Y-link与X-Bus冲突")
            faultyCode_dic.Add("TX1B35", "RX接收通路不平衡")
            faultyCode_dic.Add("TX1B36", "Radio Unit HW Fault")
            faultyCode_dic.Add("TX1B4", "TX天线驻波比超限")
            faultyCode_dic.Add("TX1B47", "TX辅助设备故障")
            faultyCode_dic.Add("TX1B6", "TX EEPROM校验故障")
            faultyCode_dic.Add("TX1B7", "TX配置表校验故障")
            faultyCode_dic.Add("TX1B8", "TX合成器A/B未锁定")
            faultyCode_dic.Add("TX1B9", "TX合成器C未锁定")
            faultyCode_dic.Add("TX2A0", "发射分极故障")
            faultyCode_dic.Add("TX2A1", "天线快速跳频失败")
            faultyCode_dic.Add("MCTR1A0", "Radio Unit Fault")
            faultyCode_dic.Add("MCTR1A1", "HW Fault")
            faultyCode_dic.Add("MCTR1A2", "Radio Unit Incompatible")
            faultyCode_dic.Add("MCTR1A3", "HW and IDB Inconsistent")
            faultyCode_dic.Add("MCTR1A4", "Radio Unit in Full Maintenance Mode")
            faultyCode_dic.Add("MCTR1B0", "Radio Unit Connection Fault")
            faultyCode_dic.Add("MCTR1B1", "Temperature Exceptional")
            faultyCode_dic.Add("MCTR1B2", "Lost Communication to Radio Unit")
            faultyCode_dic.Add("MCTR1B3", "Traffic Lost")
            faultyCode_dic.Add("MCTR2A0", "HW Fault")
            faultyCode_dic.Add("MCTR2A1", "RX Cable Disconnected")
            faultyCode_dic.Add("MCTR2A2", "VSWR Over Threshold")
            faultyCode_dic.Add("MCTR2A4", "Temperature Abnormal")
            faultyCode_dic.Add("MCTR2A5", "RX Maxgain Violated")
            faultyCode_dic.Add("MCTR2A6", "Current to high")
            faultyCode_dic.Add("MCTR2A7", "High Frequency of Software Fault")
            faultyCode_dic.Add("MCTR2A8", "Traffic Lost")
            faultyCode_dic.Add("MCTR2A9", "ALNA/TMA Fault")
            faultyCode_dic.Add("MCTR2A10", "Auxiliary Equipment Fault")
            faultyCode_dic.Add("MCTR2A11", "CPRI Delay Too Long")
            faultyCode_dic.Add("MCTR2A12", "Communication Disturbance Between Radio")
            faultyCode_dic.Add("MCTR2A13", "Communication Failure Between Radio Units")
            faultyCode_dic.Add("MCTR2A14", "Communication Equipment Fault in Cascade")
            faultyCode_dic.Add("MCTR2A15", "Lost Communication to Radio Unit in Cascade")
            faultyCode_dic.Add("MCTR2A16", "RX Path Lost on A Receiver Side")
            faultyCode_dic.Add("MCTR2A17", "RX Path Lost on B Receiver Side")
            faultyCode_dic.Add("MCTR2A18", "RX Path Lost on C Receiver Side")
            faultyCode_dic.Add("MCTR2A19", "RX Path Lost on D Receiver Side")
            faultyCode_dic.Add("MCTR2A20", "RX Path A Imbalance")
            faultyCode_dic.Add("MCTR2A21", "RX Path B Imbalance")
            faultyCode_dic.Add("MCTR2A22", "RX Path C Imbalance")
            faultyCode_dic.Add("MCTR2A23", "RX Path D Imbalance")
            faultyCode_dic.Add("MCTREC110", "CC CONF (Inconsistent Combined Cell Configuration)")
        End If
    End Sub
    Private Sub mergeCDDResults(sender As Object, e As DoWorkEventArgs, fileFolder As String)
        Dim excel_app As Excel.Application
        Dim excel_workbook As Excel.Workbook
        Dim nowCSVfile As Excel.Workbook
        Dim fileName As String
        Dim nowFileNo As Integer = 1
        Dim fileCounts As Integer
        excel_app = New Excel.Application
        excel_app.DisplayAlerts = False
        fileName = Dir(fileFolder & "*.csv")
        fileCounts = Dir(fileFolder & "*.csv").Count
        If fileName <> "" Then
            excel_workbook = excel_app.Workbooks.Add
        Else
            Exit Sub
        End If
        sender.reportProgress(0, "开始合并CSV文件...")
        Try
            Do While fileName <> ""
                sender.reportProgress(nowFileNo / fileCounts, "正在合并CSV文件" & fileName & "...")
                nowCSVfile = excel_app.Workbooks.Open(fileFolder & fileName) : excel_workbook.Sheets.Add()
                excel_workbook.ActiveSheet.Name = Strings.Left(Split(fileName, ".")(0), Len(Split(fileName, ".")(0)) - 7)
                nowCSVfile.ActiveSheet.UsedRange.Copy(excel_workbook.ActiveSheet.range("A1"))
                nowCSVfile.Close(False)
                fileName = Dir()
                nowFileNo = nowFileNo + 1
            Loop
            excel_workbook.SaveAs(fileFolder & "CDD_results_" & CDDTime & ".xlsx")
            excel_app.ScreenUpdating = True
            sender.reportProgress(100, "合并CSV文件成功...@=>合并CSV文件成功!文件名为：" & fileFolder & "CDD_results_" & CDDTime & ".xlsx")
        Catch ex As Exception
            MsgBox("合并CSV文件出错！" & ex.ToString, MsgBoxStyle.Critical)
            sender.reportProgress(0, "合并CSV文件失败！")
            Exit Sub
        Finally
            excel_workbook.Close()
            excel_app.Quit()
        End Try
    End Sub
    Private Function getFaultyCodeName(faultyCode As String)
        If faultyCode_dic.ContainsKey(faultyCode) Then
            getFaultyCodeName = faultyCode_dic(faultyCode)
        Else
            getFaultyCodeName = ""
        End If
    End Function
    Private Sub parseCDD(sender As Object, e As DoWorkEventArgs)
        Dim nowFileNo As Integer
        Dim strLineContent As String
        Dim nowProgress As Integer
        Dim nowFileName As String
        DEP_Cell_arr = New ArrayList
        BSC_arr = New ArrayList
        DEP_BSC_CELL = New ArrayList
        CGI_arr = New ArrayList
        LAI_arr = New ArrayList
        CI_arr = New ArrayList
        BCCH_arr = New ArrayList
        BSIC_arr = New ArrayList
        Height_arr = New ArrayList
        Dtile_arr = New ArrayList
        CFP_Cell_arr = New ArrayList
        CFP_BSC_CELL = New ArrayList
        TCH_arr = New ArrayList
        HOP_arr = New ArrayList
        HSN_arr = New ArrayList
        CHGR_ARR = New ArrayList
        MBC_CELL_ARR = New ArrayList
        TRX_ARR = New ArrayList
        S_cell_Neighbour = New ArrayList
        N_cell_Neighbour = New ArrayList
        CSYSTYPE_ARR = New ArrayList
        unParseCommand_arr = New ArrayList
        unParseCommand = ""
        fso = CreateObject("Scripting.FileSystemObject")
        nowProgress = 0
        If LoadFaultName.Checked Then
            Try
                makeFaultyCodeDic()
            Catch ex As Exception
                sender.reportProgress(nowProgress, "加载告警信息失败!@=>加载告警信息失败，请反馈作者！")
            End Try
            sender.reportProgress(nowProgress, "Load RXMFP告警名称!...@=>Load RXMFP告警名称完成！")
        End If
        sender.reportProgress(nowProgress, "Start Parsing!...@=>Start Parsing CDD log！")
        Try
            For nowFileNo = 0 To UBound(Filenames)
                nowFileName = UCase(Strings.Right(Trim(Filenames(nowFileNo)), Len(Trim(Filenames(nowFileNo))) - InStrRev(Trim(Filenames(nowFileNo)), "\")))
                sender.reportProgress(nowProgress, nowFileName & "...parsing!")
                getBSC(Filenames(nowFileNo))
                outPath = Strings.Left(Filenames(nowFileNo), InStrRev(Filenames(nowFileNo), "\"))
                FileOpen(1, Filenames(nowFileNo), OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
                Do While Not EOF(1)
                    strLineContent = LineInput(1)
                    If UCase(Trim(Strings.Left(strLineContent, 1))) = "<" Then
                        doWhichMethod(UCase(Trim(strLineContent)), sender)
                    End If
                Loop
                FileClose(1)
                nowProgress = (nowFileNo + 1) / (UBound(Filenames) + 1) * 100
                sender.reportProgress(nowProgress, nowFileName & "...complete!@=>" & nowFileName & "...complete!")
            Next
            sender.reportProgress(100, "ALL CDD Parse Completed!@=>ALL CDD Parse Completed!")
            If generateCarrier.CheckState = CheckState.Checked Then
                doMcomCarrier(outPath, CDDTime, sender, e)
                doMcomNeighbour(outPath, CDDTime, sender, e)
            End If
            If generateCDUtype.CheckState = CheckState.Checked Then
                doCheckCDU(outPath, CDDTime, sender, e)
            End If
            If generateCellfile.CheckState = CheckState.Checked Then
                cellOfSite_arr = New ArrayList
                Lat_arr = New ArrayList
                Lon_arr = New ArrayList
                Dir_arr = New ArrayList
                SiteName_arr = New ArrayList
                doCellFile(outPath, CDDTime, sender, e)
            End If
            If generateCellConfig.CheckState = CheckState.Checked Then
                If timeRemainForTool Then
                    cellOfSite_arr = New ArrayList
                    Lat_arr = New ArrayList
                    Lon_arr = New ArrayList
                    Dir_arr = New ArrayList
                    SiteName_arr = New ArrayList
                    doCellConfig(outPath, CDDTime, sender, e)
                Else
                    MsgBox("小区基础配置信息功能已过期，请联系作者！或者加QQ群：202196045", MsgBoxStyle.Critical)
                End If
            End If
            If generateSeeSiteFile.CheckState = CheckState.Checked Then
                If timeRemainForTool Then
                    cellOfSite_arr = New ArrayList
                    Lat_arr = New ArrayList
                    Lon_arr = New ArrayList
                    Dir_arr = New ArrayList
                    Height_arr = New ArrayList
                    Dtile_arr = New ArrayList
                    SiteName_arr = New ArrayList
                    doSeeSiteFile(outPath, CDDTime, sender, e)
                Else
                    MsgBox("SeeSite工参文件生成功能已过期，请联系作者！或者加QQ群：202196045", MsgBoxStyle.Critical)
                End If
                sender.reportProgress(0, "欢迎使用！")
                '    MsgBox("SeeSite工参文件用EXCEL打开后，手动另存为.xls文件即可！", MsgBoxStyle.Information)
            End If
            If MsgBox("是否合并CSV文件到一个文件？" & vbCrLf & "提示：此过程可能会占用您较多的时间！", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "提示") = vbYes Then
                mergeCDDResults(sender, e, outPath)
            End If
            sender.reportProgress(0, "欢迎使用！@" & unParseCommand)
        Catch ex As Exception
            MsgBox("大多数情况，是由于当前需生成的结果文件处于打开状态造成的！" & vbCrLf & "如非上述问题！请反馈作者！jiaolianxin227@163.com或者加QQ群：202196045" & vbCrLf & ex.ToString)
        Finally
            DEP_Cell_arr = Nothing
            BSC_arr = Nothing
            DEP_BSC_CELL = Nothing
            CGI_arr = Nothing
            LAI_arr = Nothing
            CI_arr = Nothing
            BCCH_arr = Nothing
            BSIC_arr = Nothing
            CFP_Cell_arr = Nothing
            CFP_BSC_CELL = Nothing
            TCH_arr = Nothing
            HOP_arr = Nothing
            HSN_arr = Nothing
            CHGR_ARR = Nothing
            MBC_CELL_ARR = Nothing
            TRX_ARR = Nothing
            fso = Nothing
            BSCNAME = Nothing
            cellOfSite_arr = Nothing
            Lat_arr = Nothing
            Lon_arr = Nothing
            Dir_arr = Nothing
            SiteName_arr = Nothing
            S_cell_Neighbour = Nothing
            N_cell_Neighbour = Nothing
            CSYSTYPE_ARR = Nothing
            Height_arr.Clear()
            Dtile_arr.Clear()
            STP_BSC_CELL_STATE.Clear()
            TCP_CELL_MO_DIC.Clear()
            MFP_MO_CDUtype_DIC.Clear()
            '    TRX_BSC_CELL_DIC.Clear()
            CDU_DIC.Clear()
            Final_EQP_DIC.Clear()
            TRX_TYPE_DIC.Clear()
            TRX_BSC_MO_CELL_DIC.Clear()
            '     TRX_SITE_TRXNO_DIC.Clear()
            I = 0
            FileClose(1)
            unParseCommand_arr.Clear()
            unParseCommand = Nothing
        End Try
        MsgBox("CDD转换完毕！共用时：" & totalTime & "秒！", MsgBoxStyle.Information)
        Shell("explorer.exe " & outPath, 1)
    End Sub
    
    Private Sub doWhichMethod(str As String, sender As Object)
        Dim temp As String
        Dim temp2 As String
        temp = Trim(Replace(str, "<", ""))
        temp2 = UCase(Strings.Left(temp, InStrRev(temp, "=")))
        If Strings.Left(temp, 17) = "DBTSP:TAB=AXEPARS" Then
            temp = "DBTSP:TAB=AXEPARS;"
        End If
        If Strings.Left(temp, 12) = "PCORP:BLOCK=" Then
            temp = "PCORP:BLOCK=ALL;"
        End If
        Select Case temp2
            Case "RLNCP:CELL="
                Try
                    doRLNCP(sender)
                Catch e As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                End Try
            Case "RLGSP:CELL="
                Try
                    temp2succArr.Add(temp)
                    doRLGSP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case Else

        End Select
        Select Case temp
            Case "IOEXP;"
                Try
                    doIOEXP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "ALLIP;"
                Try
                    doALLIP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "CACLP;"
                Try
                    doCACLP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "DBTSP:TAB=AXEPARS;"
                Try
                    doDBTSP_AXEPARS(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "DTSTP:DIP=ALL;"
                Try
                    doDTSTP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "DTQUP:DIP=ALL;"
                Try
                    doDTQUP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "EXRPP:RP=ALL;"
                Try
                    doEXRPP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "EXEMP:EM=ALL,RP=ALL;"
                Try
                    doEXEMP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "MGCEP:CELL=ALL;"
                Try
                    doMGCEP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "MGEPP:ID=ALL;"
                Try
                    doMGEPP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "MGBRP:RR=ALL;"
                Try
                    doMGBRP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "MGLAP;"
                Try
                    doMGLAP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "MGIDP;"
                Try
                    doMGIDP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "MGOCP:CELL=ALL;"
                Try
                    doMGOCP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "PCORP:BLOCK=ALL;"
                Try
                    doPCORP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "PLLDP;"
                Try
                    doPLLDP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RACLP;"
                Try
                    doRACLP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RAEPP:ID=ALL;"
                Try
                    doRAEPP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RAHHP:SNT=ALL;"
                Try
                    doRAHHP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RASAP;"
                Try
                    doRASAP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RFCAP:CELL=ALL;"
                Try
                    doRFCAP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RFPCP:CELL=ALL;"
                Try
                    doRFPCP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLACP:CELL=ALL;"
                Try
                    doRLACP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLADP:SET=ALL;"
                Try
                    doRLADP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLAPP:CELL=ALL;"
                Try
                    doRLAPP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLBAP;"
                Try
                    doRLBAP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLBCP:CELL=ALL;"
                Try
                    doRLBCP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLBDP:CELL=ALL;"
                Try
                    doRLBDP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLCAP;"
                Try
                    doRLCAP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLCCP:CELL=ALL;"
                Try
                    doRLCCP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLCDP:CELL=ALL;"
                Try
                    doRLCDP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLCFP:CELL=ALL;"
                Try
                    doRLCFP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLCHP:CELL=ALL;"
                Try
                    doRLCHP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLCLP:CELL=ALL;"
                Try
                    doRLCLP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLCPP:CELL=ALL;"
                Try
                    doRLCPP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLCPP:CELL=ALL,EXT;"
                Try
                    doRLCPP_EXT(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLCRP:CELL=ALL;"
                Try
                    doRLCRP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLCXP:CELL=ALL;"
                Try
                    doRLCXP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLDCP;"
                Try
                    doRLDCP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLDEP:CELL=ALL;"
                Try
                    doRLDEP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLDEP:CELL=ALL,EXT;"
                Try
                    doRLDEP_ext(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLDEP:CELL=ALL,EXT,UTRAN;"
                Try
                    doRLDEP_ext_UTRAN(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLDGP:CELL=ALL;"
                Try
                    doRLDGP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLDHP:CELL=ALL;"
                Try
                    doRLDHP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLDMP:CELL=ALL;"
                Try
                    doRLDMP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLDRP:CELL=ALL;"
                Try
                    doRLDRP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLDTP:CELL=ALL;"
                Try
                    doRLDTP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLEFP:CELL=ALL;"
                Try
                    doRLEFP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLDUP:CELL=ALL;"
                Try
                    doRLDUP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLGAP:CELL=ALL;"
                Try
                    doRLGAP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLGRP:CELL=ALL;"
                Try
                    doRLGRP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
                'Case "RLGSP:CELL=ALL;"
                '    Try
                '        doRLGSP(sender)
                '    Catch ex As Exception
                '        unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                '    Finally
                '    End Try
            Case "RLHBP;"
                Try
                    doRLHBP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLHPP:CELL=ALL;"
                Try
                    doRLHPP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLIHP:CELL=ALL;"
                Try
                    doRLIHP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLIMP:CELL=ALL;"
                Try
                    doRLIMP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLLBP;"
                Try
                    doRLLBP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLLCP:CELL=ALL;"
                Try
                    doRLLCP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLLDP:CELL=ALL;"
                Try
                    doRLLDP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLLFP:CELL=ALL;"
                Try
                    doRLLFP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLLHP:CELL=ALL;"
                Try
                    doRLLHP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLLHP:CELL=ALL,EXT;"
                Try
                    doRLLHP_EXT(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLLLP:CELL=ALL;"
                Try
                    doRLLLP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLLOP:CELL=ALL;"
                Try
                    doRLLOP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLLOP:CELL=ALL,EXT;"
                Try
                    doRLLOP_EXT(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLLPP:CELL=ALL;"
                Try
                    doRLLPP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLLSP;"
                Try
                    doRLLSP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLLUP:CELL=ALL;"
                Try
                    doRLLUP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLMAP;"
                Try
                    doRLMAP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLMFP:CELL=ALL;"
                Try
                    doRLMFP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLNMP:CELL=ALL;"
                Try
                    doRLNMP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLNPP:CELL=ALL;"
                Try
                    doRLNPP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLNRP:CELL=ALL;"
                Try
                    doRLNRP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLNRP:CELL=ALL,NODATA;"
                Try
                    doRLNRP_NODATA(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLNRP:CELL=ALL,UTRAN;"
                Try
                    doRLNRP_UTRAN(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLOLP:CELL=ALL;"
                Try
                    doRLOLP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLOMP;"
                Try
                    doRLOMP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLPBP:CELL=ALL;"
                Try
                    doRLPBP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLPCP:CELL=ALL;"
                Try
                    doRLPCP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLPDP:CELL=ALL;"
                Try
                    doRLPDP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLPRP:CELL=ALL;"
                Try
                    doRLPRP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLRLP:CELL=ALL;"
                Try
                    doRLRLP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLRLP:CELL=ALL,COVERAGEE=ALL;"
                Try
                    doRLRLP_COVERAGEE(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLSBP:CELL=ALL;"
                Try
                    doRLSBP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLSCP;"
                Try
                    doRLSCP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLSLP:CELL=ALL;"
                Try
                    doRLSLP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLSMP:CELL=ALL;"
                Try
                    doRLSMP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLSQP;"
                Try
                    doRLSQP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLSRP:CELL=ALL;"
                Try
                    doRLSRP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLSSP:CELL=ALL;"
                Try
                    doRLSSP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLSTP:CELL=ALL;"
                Try
                    doRLSTP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLSUP:CELL=ALL;"
                Try
                    doRLSUP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLSVP:CELL=ALL;"
                Try
                    doRLSVP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLTDP:MSC=ALL;"
                Try
                    doRLTDP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLTYP;"
                Try
                    doRLTYP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLUMP:CELL=ALL;"
                Try
                    doRLUMP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLVAP;"
                Try
                    doRLVAP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RLVLP;"
                Try
                    doRLVLP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RRMBP:MSC=ALL;"
                Try
                    doRRMBP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RRSCP:SCGR=ALL;"
                Try
                    doRRSCP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RRTPP:TRAPOOL=ALL;"
                Try
                    doRRTTP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXAPP:MOTY=RXOTG;"
                Try
                    doRXAPP_RXOTG(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXASP:MOTY=RXOTG;"
                Try
                    doRXASP_RXOTG(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXBFP:MOTY=RXOTG;"
                Try
                    doRXBFP_RXOTG(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXELP:MOTY=RXOTG;"
                Try
                    doRXELP_RXOTG(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMDP:MOTY=RXOCF;"
                Try
                    doRXMDP_RXOCF(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMDP:MOTY=RXOTRX;"
                Try
                    doRXMDP_RXOTRX(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMFP:MOTY=RXOMCTR;"
                Try
                    doRXMFP_RXOMCTR(sender, "MCTR")
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMFP:MOTY=RXOMCTR,FAULTY;"
                Try
                    doRXMFP_FAULTY(sender, "MCTR")
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMFP:MOTY=RXOCF;"
                Try
                    doRXMFP_RXOCF(sender, "CF")
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMFP:MOTY=RXOCF,FAULTY;"
                Try
                    doRXMFP_FAULTY(sender, "CF")
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMFP:MOTY=RXOCON;"
                Try
                    doRXMFP_RXOCON(sender, "CON")
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMFP:MOTY=RXOCON,FAULTY;"
                Try
                    doRXMFP_FAULTY(sender, "CON")
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMFP:MOTY=RXORX;"
                Try
                    doRXMFP_RXORX(sender, "RX")
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMFP:MOTY=RXORX,FAULTY;"
                Try
                    doRXMFP_FAULTY(sender, "RX")
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMFP:MOTY=RXOTF;"
                Try
                    doRXMFP_RXOTF(sender, "TF")
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMFP:MOTY=RXOTF,FAULTY;"
                Try
                    doRXMFP_FAULTY(sender, "TF")
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMFP:MOTY=RXOTRX;"
                Try
                    doRXMFP_RXOTRX(sender, "TRX")
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMFP:MOTY=RXOTRX,FAULTY;"
                Try
                    doRXMFP_FAULTY(sender, "TRX")
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMFP:MOTY=RXOTS;"
                Try
                    doRXMFP_RXOTS(sender, "TS")
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMFP:MOTY=RXOTS,FAULTY;"
                Try
                    doRXMFP_FAULTY(sender, "TS")
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMFP:MOTY=RXOTX;"
                Try
                    doRXMFP_RXOTX(sender, "TX")
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMFP:MOTY=RXOTX,FAULTY;"
                Try
                    doRXMFP_FAULTY(sender, "TX")
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMOP:MOTY=RXOCF;"
                Try
                    doRXMOP_RXOCF(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMOP:MOTY=RXOMCTR;"
                Try
                    doRXMOP_RXOMCTR(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMOP:MOTY=RXORX;"
                Try
                    doRXMOP_RXORX(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMOP:MOTY=RXOTF;"
                Try
                    doRXMOP_RXOTF(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMOP:MOTY=RXOTG;"
                Try
                    doRXMOP_RXOTG(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMOP:MOTY=RXOTRX;"
                Try
                    doRXMOP_RXOTRX(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMOP:MOTY=RXOTX;"
                Try
                    doRXMOP_RXOTX(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMSP:MOTY=RXOTRX;"
                Try
                    doRXMSP_RXOTRX(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMSP:MOTY=RXOCF;"
                Try
                    doRXMSP_RXOCF(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXMSP:MOTY=RXOTF;"
                Try
                    doRXMSP_RXOTF(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "RXTCP:MOTY=RXOTG;"
                Try
                    doRXTCP_RXOTG(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "SAAEP:SAE=ALL;"
                Try
                    doSAAEP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case "TPCOP:SDIP=ALL;"
                Try
                    doTPCOP(sender)
                Catch ex As Exception
                    unParseCommand = unParseCommand & "=>" & BSCNAME & "的" & temp & "指令未正常转换，请稍候检查！" & vbCrLf
                Finally
                End Try
            Case Else
                If Not unParseCommand_arr.Contains("=>" & temp & "指令暂不支持转换，如需要，请反馈！") And Not temp2succArr.Contains(temp) Then
                    unParseCommand_arr.Add("=>" & temp & "指令暂不支持转换，如需要，请反馈！")
                    unParseCommand = unParseCommand & "=>" & temp & "指令暂不支持转换，如需要，请反馈！" & vbCrLf
                End If
        End Select
    End Sub
    Private Sub doMcomCarrier(outFolder As String, fileCreateTime As String, sender As Object, e As DoWorkEventArgs)
        Dim outFile As Object
        Dim outFilewithTRXS As Object
        Dim outFileDEP_CFP As Object
        Dim outMBCFile As Object
        Dim CELL As String, BSC As String, TCH As String, HOP As String, HSN As String, TRX_Count As Integer
        Dim TCH_GROUP As String
        Dim TCH_HEADLINE As String
        Dim minSep As Integer, maxSep As Integer
        Dim tempSep As Integer
        TCH_GROUP = ""
        sender.reportProgress(0, "制作MCOM carrier文件..." & "0%")
        outFile = fso.OpenTextFile(outFolder & "McomCarrier" & fileCreateTime & ".txt", 2, 1, 0)
        outFilewithTRXS = fso.OpenTextFile(outFolder & "McomCarrier_高级_" & fileCreateTime & ".txt", 2, 1, 0)
        outFileDEP_CFP = fso.OpenTextFile(outFolder & "RLDEP_RLCFP_" & fileCreateTime & ".csv", 2, 1, 0)
        'CELL	BSC	LAI	CI	BCCH	BSIC	TCH	HOP	HSN	FONT_SIZE
        outFile.writeLine("CELL" & "	" & "BSC" & "	" & "LAI" & "	" & "CI" & "	" & "BCCH" & "	" & "BSIC" & "	" & "TCH" & "	" & "HOP" & "	" & "HSN" & "	" & "FONT_SIZE")
        outFilewithTRXS.writeLine("CELL" & "	" & "BSC" & "	" & "LAI" & "	" & "CI" & "	" & "BCCH" & "	" & "BSIC" & "	" & "TCH" & "	" & "HOP" & "	" & "HSN" & "	" & "FONT_SIZE" & "	" & "载频数" & "	" & "频点间隔(最小)" & "	" & "频点间隔(最大)")
        TCH_HEADLINE = "TCH1"
        For a = 2 To MaxTCHNo
            TCH_HEADLINE = TCH_HEADLINE & "," & "TCH" & a
        Next
        If MaxTCHNo < 12 Then
            For b = MaxTCHNo + 1 To 12
                TCH_HEADLINE = TCH_HEADLINE & "," & "TCH" & b
            Next
        End If
        outFileDEP_CFP.writeLine("CELL" & "," & "BSC" & "," & "CSYSTYPE" & "," & "LAC" & "," & "CI" & "," & "NCC" & "," & "BCC" & "," & "BSIC" & "," & "BCCHNO" & "," & TCH_HEADLINE)
        If MBC_CELL_ARR.Count > 0 Then
            outMBCFile = fso.OpenTextFile(outFolder & "MBC_McomCarrier" & fileCreateTime & ".txt", 2, 1, 0)
            outMBCFile.writeLine("CELL" & "	" & "BSC" & "	" & "LAI" & "	" & "CI" & "	" & "BCCH" & "	" & "BSIC" & "	" & "TCH" & "	" & "HOP" & "	" & "HSN" & "	" & "FONT_SIZE")
        Else
            outMBCFile = ""
        End If

        For X = 0 To DEP_BSC_CELL.Count - 1
            CELL = Strings.Split(DEP_BSC_CELL(X), "#")(1)
            BSC = Strings.Split(DEP_BSC_CELL(X), "#")(0)

            If CFP_BSC_CELL.Contains(DEP_BSC_CELL(X)) Then
                TCH = Trim(TCH_arr(CFP_BSC_CELL.IndexOf(DEP_BSC_CELL(X)))) & " "
                TCH = Trim(Replace(TCH, " " & Trim(BCCH_arr(X)) & " ", " "))
                If InStr(TCH, "] [") > 0 Then
                    TCH = Trim(Replace(TCH, Mid(TCH, InStr(TCH, "] [") - 2, 5), "["))
                End If
                If Strings.Right(TCH, 1) = "]" Then
                    TCH = Trim(Strings.Left(TCH, Len(TCH) - 3))
                End If
                If Strings.Right(TCH, 1) = "]" Then
                    TCH = Trim(Strings.Left(TCH, Len(TCH) - 3))
                End If

                HOP = Trim(HOP_arr(CFP_BSC_CELL.IndexOf(DEP_BSC_CELL(X))))
                HSN = Trim(HSN_arr(CFP_BSC_CELL.IndexOf(DEP_BSC_CELL(X))))
                TRX_Count = UBound(Split(Trim(TRX_ARR(CFP_BSC_CELL.IndexOf(DEP_BSC_CELL(X)))), " ")) + 1
                outFile.writeLine(CELL & "	" & BSC & "	" & LAI_arr(X) & "	" & CI_arr(X) & "	" & BCCH_arr(X) & "	" & BSIC_arr(X) & "	" & TCH & "	" & HOP & "	" & HSN & "	" & "5")
                minSep = 9999
                maxSep = -1
                For k = 0 To TRX_Count - 1
                    For o = k + 1 To TRX_Count - 1
                        tempSep = Math.Abs(Split(Trim(TRX_ARR(CFP_BSC_CELL.IndexOf(DEP_BSC_CELL(X)))), " ")(k) - Split(Trim(TRX_ARR(CFP_BSC_CELL.IndexOf(DEP_BSC_CELL(X)))), " ")(o))
                        If k = 0 And o = 1 Then
                            minSep = tempSep
                            maxSep = tempSep
                        Else
                            If tempSep < minSep Then
                                minSep = tempSep
                            End If
                            If tempSep > maxSep Then
                                maxSep = tempSep
                            End If
                        End If
                    Next
                Next
                outFilewithTRXS.writeLine(CELL & "	" & BSC & "	" & LAI_arr(X) & "	" & CI_arr(X) & "	" & BCCH_arr(X) & "	" & BSIC_arr(X) & "	" & TCH & "	" & HOP & "	" & HSN & "	" & "5" & "	" & TRX_Count & "	" & minSep & "	" & maxSep)
                TCH_GROUP = Replace(Trim(Replace(" " & TRX_ARR(CFP_BSC_CELL.IndexOf(DEP_BSC_CELL(X))) & " ", " " & BCCH_arr(X) & " ", " ")), " ", ",")
                outFileDEP_CFP.writeLine(CELL & "," & BSC & "," & CSYSTYPE_ARR(X) & "," & LAI_arr(X) & "," & CI_arr(X) & "," & Strings.Left(BSIC_arr(X), 1) & "," & Strings.Right(BSIC_arr(X), 1) & "," & BSIC_arr(X) & "," & BCCH_arr(X) & "," & TCH_GROUP)
                If MBC_CELL_ARR.Contains(CELL) And TCH <> "" Then
                    outMBCFile.writeLine(CELL & "	" & BSC & "	" & LAI_arr(X) & "	" & CI_arr(X) & "	" & BCCH_arr(X) & "	" & BSIC_arr(X) & "	" & TCH & "	" & HOP & "	" & HSN & "	" & "5")
                End If
            Else
                If Trim(BCCH_arr(X)) <> "" Then
                    TRX_Count = 1
                Else
                    TRX_Count = 0
                End If
                outFile.writeLine(CELL & "	" & BSC & "	" & LAI_arr(X) & "	" & CI_arr(X) & "	" & BCCH_arr(X) & "	" & BSIC_arr(X) & "	" & "" & "	" & "" & "	" & "" & "	" & "5")
                outFilewithTRXS.writeLine(CELL & "	" & BSC & "	" & LAI_arr(X) & "	" & CI_arr(X) & "	" & BCCH_arr(X) & "	" & BSIC_arr(X) & "	" & "" & "	" & "" & "	" & "" & "	" & "5" & "	" & TRX_Count)
            End If
            sender.reportProgress(X / (DEP_BSC_CELL.Count - 1) * 100, "制作MCOM carrier文件..." & Mid(X / (DEP_BSC_CELL.Count - 1) * 100, 1, 2) & "%")
        Next
        sender.reportProgress(100, "制作MCOM carrier文件..100% @=>MCOM carrier文件制作完成..100% ")
        outFile.close()
        outFilewithTRXS.CLOSE()
        outFileDEP_CFP.close()
        If MBC_CELL_ARR.Count > 0 Then
            outMBCFile.CLOSE()
            outMBCFile = Nothing
        End If
        outFile = Nothing
        outFilewithTRXS = Nothing
        outFileDEP_CFP = Nothing
    End Sub
    Private Sub doMcomNeighbour(outFolder As String, fileCreateTime As String, sender As Object, e As DoWorkEventArgs)
        Dim outFile As Object
        sender.reportProgress(0, "制作MCOM Neighbour文件..." & "0%")
        outFile = fso.OpenTextFile(outFolder & "McomNeighbour" & fileCreateTime & ".txt", 2, 1, 0)
        outFile.writeLine("CELL" & "	" & "NCELL")

        For X = 0 To S_cell_Neighbour.Count - 1
            outFile.writeLine(S_cell_Neighbour(X) & "	" & N_cell_Neighbour(X))

            sender.reportProgress(X / (DEP_BSC_CELL.Count - 1) * 100, "制作MCOM Neighbour文件..." & Mid(X / (DEP_BSC_CELL.Count - 1) * 100, 1, 2) & "%")
        Next
        sender.reportProgress(100, "制作MCOM Neighbour文件..100% @=>制作MCOM Neighbour文件...100% ")
        outFile.close()
        outFile = Nothing
    End Sub
    Private Sub doCheckCDU(outFolder As String, fileCreateTime As String, sender As Object, e As DoWorkEventArgs)
        Dim outCDUfile As Object
        Dim CDU_type As String, TCP_CELL_TEMP As String, TCP_MO_TEMP As String
        Dim EQP_DIC As New Dictionary(Of String, String)
        Dim second_EQP_DIC As New Dictionary(Of String, String)
        Dim temp As Object, cdu_type_temp As String
        Dim eqp_temp As String, second_eqp_temp As String, cdu_temp As String
        Dim x As Integer
        eqp_temp = ""
        second_eqp_temp = ""
        cdu_temp = ""
        CDU_type = ""
        sender.reportProgress(0, "CDU类型检查...0%")
        outCDUfile = fso.OpenTextFile(outFolder & "CDU类型" & fileCreateTime & ".xls", 2, 1, 0)
        outCDUfile.writeLine("BSC" & "	" & "Cell" & "	" & "BSC#MO" & "	" & "CDU_TYPE" & "	" & "机框类型(优先现网定义)" & "	" & "机框类型(现网定义)" & "	" & "机框类型(由CDU判断)")
        For x = 1 To MFP_MO_CDUtype_DIC.Count
            temp = Split(Trim(Replace(MFP_MO_CDUtype_DIC(x), "RXOCF", "RXOTG")), "@")

            If Trim(temp(1)) <> "" Then
                If Len(temp(1)) < 3 Then
                    cdu_type_temp = temp(1)
                Else
                    cdu_type_temp = Trim(Strings.Right(temp(1), Len(temp(1)) - 3))
                End If
                If Strings.Left(cdu_type_temp, 3) = "DUG" And cdu_type_temp <> "" Then
                    CDU_type = "DUG"
                ElseIf Len(cdu_type_temp) >= 3 And cdu_type_temp <> "" Then
                    CDU_type = IIf(InStr(cdu_type_temp, " ") > 0, Trim(Strings.Left(cdu_type_temp, Len(cdu_type_temp) - 1)), cdu_type_temp)
                ElseIf Len(cdu_type_temp) < 3 And cdu_type_temp <> "" Then
                    CDU_type = cdu_type_temp
                End If
            Else
                CDU_type = ""
            End If
            If Not second_EQP_DIC.ContainsKey(temp(0)) Then
                If InStr(CDU_type, "DUG") > 0 Then
                    second_EQP_DIC.Add(temp(0), "CABI RBS 6000+")
                ElseIf InStr(CDU_type, "MU") > 0 Then
                    second_EQP_DIC.Add(temp(0), "CABI RBS 2111")
                ElseIf InStr(CDU_type, "CDU_A") > 0 Or InStr(CDU_type, "CDUA") > 0 _
                    Or InStr(CDU_type, "CDU_C") > 0 Or InStr(CDU_type, "CDUC") > 0 _
                    Or InStr(CDU_type, "CDU_D") > 0 Or InStr(CDU_type, "CDUD") > 0 _
                    Or InStr(CDU_type, "CU_9") > 0 Or InStr(CDU_type, "FU_9") > 0 Then
                    second_EQP_DIC.Add(temp(0), "CABI RBS 2202")
                ElseIf InStr(CDU_type, "CDU_F") > 0 Or InStr(CDU_type, "CDUF") > 0 _
                    Or InStr(CDU_type, "CDU_K") > 0 Or InStr(CDU_type, "CDUK") > 0 Then
                    second_EQP_DIC.Add(temp(0), "CABI RBS 2206")
                ElseIf InStr(CDU_type, "CDU_G") > 0 Or InStr(CDU_type, "CDUG") > 0 Then
                    second_EQP_DIC.Add(temp(0), "CABI RBS 2106/2206")
                ElseIf InStr(CDU_type, "CDU_J") > 0 Or InStr(CDU_type, "CDUJ") > 0 Then
                    second_EQP_DIC.Add(temp(0), "CABI RBS 2107/2207")
                ElseIf InStr(CDU_type, "DXU-23") > 0 Or InStr(CDU_type, "DXU_23") > 0 _
                    Or InStr(CDU_type, "DXU-31") > 0 Or InStr(CDU_type, "DXU_31") > 0 Then
                    second_EQP_DIC.Add(temp(0), "CABI RBS 2216")
                ElseIf InStr(CDU_type, "CDU_M") > 0 Or InStr(CDU_type, "CDUM") > 0 Then
                    second_EQP_DIC.Add(temp(0), "CABI RBS 2302")
                ElseIf InStr(CDU_type, "IXU-21") > 0 Or InStr(CDU_type, "IXU_21") > 0 Then
                    second_EQP_DIC.Add(temp(0), "CABI RBS 2308/2309")
                End If
            End If
            If InStr(temp(1), "CDU") > 0 Or InStr(temp(1), "CU_9") > 0 Or InStr(temp(1), "DU_9") > 0 Or InStr(temp(1), "FU_9") > 0 Or Strings.Left(temp(1), 2) = "DX" Or Strings.Left(temp(1), 7) = "SRU CDU" Then
                If Not CDU_DIC.ContainsKey(temp(0)) Then
                    CDU_DIC.Add(temp(0), CDU_type)
                Else
                    If InStr(CDU_DIC(temp(0)), "DXU") > 0 And InStr(CDU_DIC(temp(0)), CDU_type) = 0 Then
                        CDU_DIC(temp(0)) = CDU_type
                        '                       CDU_DIC(temp(0)) = CDU_DIC(temp(0)) & "," & CDU_type
                    ElseIf InStr(CDU_DIC(temp(0)), "DXU") = 0 And InStr(CDU_DIC(temp(0)), CDU_type) = 0 And InStr(temp(1), "DXU") = 0 Then
                        CDU_DIC(temp(0)) = CDU_type & "," & CDU_DIC(temp(0))
                    End If
                End If
            ElseIf Strings.Left(temp(1), 4) = "CABI" Then
                If Not EQP_DIC.ContainsKey(temp(0)) Then
                    EQP_DIC.Add(temp(0), temp(1))
                End If
            End If
            sender.reportProgress(x / MFP_MO_CDUtype_DIC.Count * 50, "CDU类型检查..." & Mid(x / (MFP_MO_CDUtype_DIC.Count) * 50, 1, 2) & "%")
        Next
        For m = 1 To TCP_CELL_MO_DIC.Count - 1
            TCP_CELL_TEMP = TCP_CELL_MO_DIC.Keys(m)
            TCP_MO_TEMP = TCP_CELL_MO_DIC.Values(m)
            If InStr(TCP_MO_TEMP, "|") > 0 Then
                outCDUfile.write(Split(TCP_CELL_TEMP, "#")(0) & "	" & Split(TCP_CELL_TEMP, "#")(1) & "	" & TCP_MO_TEMP)
                For n = 0 To UBound(Split(TCP_MO_TEMP, "|"))
                    If n = 0 Then
                        If CDU_DIC.ContainsKey(Split(TCP_MO_TEMP, "|")(n)) Then
                            cdu_temp = CDU_DIC(Split(TCP_MO_TEMP, "|")(n))
                        Else
                            cdu_temp = "无CDU类型"
                        End If
                    Else
                        If CDU_DIC.ContainsKey(Split(TCP_MO_TEMP, "|")(n)) Then
                            If cdu_temp <> CDU_DIC(Split(TCP_MO_TEMP, "|")(n)) Then
                                cdu_temp = CDU_DIC(Split(TCP_MO_TEMP, "|")(n)) & "," & cdu_temp
                            End If
                        Else
                            ' outCDUfile.write("," & "无CDU类型")
                        End If
                    End If
                Next
                outCDUfile.write("	" & cdu_temp)
                For n = 0 To UBound(Split(TCP_MO_TEMP, "|"))
                    If n = 0 Then
                        If EQP_DIC.ContainsKey(Split(TCP_MO_TEMP, "|")(n)) Then
                            eqp_temp = EQP_DIC(Split(TCP_MO_TEMP, "|")(n))
                        Else
                            eqp_temp = "现网未定义"
                        End If
                        If second_EQP_DIC.ContainsKey(Split(TCP_MO_TEMP, "|")(n)) Then
                            second_eqp_temp = second_EQP_DIC(Split(TCP_MO_TEMP, "|")(n))
                        Else
                            second_eqp_temp = "无法准确判断"
                        End If
                    Else
                        If EQP_DIC.ContainsKey(Split(TCP_MO_TEMP, "|")(n)) Then
                            If eqp_temp <> EQP_DIC(Split(TCP_MO_TEMP, "|")(n)) Then
                                eqp_temp = EQP_DIC(Split(TCP_MO_TEMP, "|")(n)) & "," & eqp_temp
                            End If
                        Else
                            'eqp_temp = eqp_temp & "," & "现网未定义"
                        End If
                        If second_EQP_DIC.ContainsKey(Split(TCP_MO_TEMP, "|")(n)) Then
                            If second_eqp_temp <> second_EQP_DIC(Split(TCP_MO_TEMP, "|")(n)) Then
                                second_eqp_temp = second_EQP_DIC(Split(TCP_MO_TEMP, "|")(n)) & "," & second_eqp_temp
                            End If
                        Else
                            ' second_eqp_temp = second_eqp_temp & "," & "无法准确判断"
                        End If
                    End If
                Next
                If eqp_temp <> "现网未定义" Then
                    outCDUfile.writeLine("	" & eqp_temp & "	" & eqp_temp & "	" & second_eqp_temp)
                Else
                    outCDUfile.writeLine("	" & second_eqp_temp & "	" & eqp_temp & "	" & second_eqp_temp)
                End If

            Else
                If EQP_DIC.ContainsKey(TCP_MO_TEMP) And CDU_DIC.ContainsKey(TCP_MO_TEMP) And second_EQP_DIC.ContainsKey(TCP_MO_TEMP) Then
                    If Not Final_EQP_DIC.ContainsKey(TCP_MO_TEMP) Then
                        Final_EQP_DIC.Add(TCP_MO_TEMP, EQP_DIC(TCP_MO_TEMP))
                    End If
                    outCDUfile.writeLine(Split(TCP_CELL_TEMP, "#")(0) & "	" & Split(TCP_CELL_TEMP, "#")(1) & "	" & TCP_MO_TEMP & "	" & CDU_DIC(TCP_MO_TEMP) & "	" & EQP_DIC(TCP_MO_TEMP) & "	" & EQP_DIC(TCP_MO_TEMP) & "	" & second_EQP_DIC(TCP_MO_TEMP))
                ElseIf EQP_DIC.ContainsKey(TCP_MO_TEMP) And CDU_DIC.ContainsKey(TCP_MO_TEMP) And Not second_EQP_DIC.ContainsKey(TCP_MO_TEMP) Then
                    If Not Final_EQP_DIC.ContainsKey(TCP_MO_TEMP) Then
                        Final_EQP_DIC.Add(TCP_MO_TEMP, EQP_DIC(TCP_MO_TEMP))
                    End If
                    outCDUfile.writeLine(Split(TCP_CELL_TEMP, "#")(0) & "	" & Split(TCP_CELL_TEMP, "#")(1) & "	" & TCP_MO_TEMP & "	" & CDU_DIC(TCP_MO_TEMP) & "	" & EQP_DIC(TCP_MO_TEMP) & "	" & EQP_DIC(TCP_MO_TEMP) & "	" & "无法准确判断")
                ElseIf EQP_DIC.ContainsKey(TCP_MO_TEMP) And Not CDU_DIC.ContainsKey(TCP_MO_TEMP) And second_EQP_DIC.ContainsKey(TCP_MO_TEMP) Then
                    If Not Final_EQP_DIC.ContainsKey(TCP_MO_TEMP) Then
                        Final_EQP_DIC.Add(TCP_MO_TEMP, EQP_DIC(TCP_MO_TEMP))
                    End If
                    outCDUfile.writeLine(Split(TCP_CELL_TEMP, "#")(0) & "	" & Split(TCP_CELL_TEMP, "#")(1) & "	" & TCP_MO_TEMP & "	" & "无CDU类型" & "	" & EQP_DIC(TCP_MO_TEMP) & "	" & EQP_DIC(TCP_MO_TEMP) & "	" & second_EQP_DIC(TCP_MO_TEMP))
                ElseIf EQP_DIC.ContainsKey(TCP_MO_TEMP) And Not CDU_DIC.ContainsKey(TCP_MO_TEMP) And Not second_EQP_DIC.ContainsKey(TCP_MO_TEMP) Then
                    If Not Final_EQP_DIC.ContainsKey(TCP_MO_TEMP) Then
                        Final_EQP_DIC.Add(TCP_MO_TEMP, EQP_DIC(TCP_MO_TEMP))
                    End If
                    outCDUfile.writeLine(Split(TCP_CELL_TEMP, "#")(0) & "	" & Split(TCP_CELL_TEMP, "#")(1) & "	" & TCP_MO_TEMP & "	" & "无CDU类型" & "	" & EQP_DIC(TCP_MO_TEMP) & "	" & EQP_DIC(TCP_MO_TEMP) & "	" & "无法准确判断")
                ElseIf Not EQP_DIC.ContainsKey(TCP_MO_TEMP) And CDU_DIC.ContainsKey(TCP_MO_TEMP) And second_EQP_DIC.ContainsKey(TCP_MO_TEMP) Then
                    If Not Final_EQP_DIC.ContainsKey(TCP_MO_TEMP) Then
                        Final_EQP_DIC.Add(TCP_MO_TEMP, second_EQP_DIC(TCP_MO_TEMP))
                    End If
                    outCDUfile.writeLine(Split(TCP_CELL_TEMP, "#")(0) & "	" & Split(TCP_CELL_TEMP, "#")(1) & "	" & TCP_MO_TEMP & "	" & CDU_DIC(TCP_MO_TEMP) & "	" & second_EQP_DIC(TCP_MO_TEMP) & "	" & "现网未定义" & "	" & second_EQP_DIC(TCP_MO_TEMP))
                ElseIf Not EQP_DIC.ContainsKey(TCP_MO_TEMP) And CDU_DIC.ContainsKey(TCP_MO_TEMP) And Not second_EQP_DIC.ContainsKey(TCP_MO_TEMP) Then
                    If Not Final_EQP_DIC.ContainsKey(TCP_MO_TEMP) Then
                        Final_EQP_DIC.Add(TCP_MO_TEMP, "现网未定义")
                    End If
                    outCDUfile.writeLine(Split(TCP_CELL_TEMP, "#")(0) & "	" & Split(TCP_CELL_TEMP, "#")(1) & "	" & TCP_MO_TEMP & "	" & CDU_DIC(TCP_MO_TEMP) & "	" & "现网未定义" & "	" & "现网未定义" & "	" & "无法准确判断")
                ElseIf Not EQP_DIC.ContainsKey(TCP_MO_TEMP) And Not CDU_DIC.ContainsKey(TCP_MO_TEMP) And second_EQP_DIC.ContainsKey(TCP_MO_TEMP) Then
                    If Not Final_EQP_DIC.ContainsKey(TCP_MO_TEMP) Then
                        Final_EQP_DIC.Add(TCP_MO_TEMP, "现网未定义")
                    End If
                    outCDUfile.writeLine(Split(TCP_CELL_TEMP, "#")(0) & "	" & Split(TCP_CELL_TEMP, "#")(1) & "	" & TCP_MO_TEMP & "	" & CDU_DIC(TCP_MO_TEMP) & "	" & "现网未定义" & "	" & "现网未定义" & "	" & second_EQP_DIC(TCP_MO_TEMP))
                ElseIf Not EQP_DIC.ContainsKey(TCP_MO_TEMP) And Not CDU_DIC.ContainsKey(TCP_MO_TEMP) And Not second_EQP_DIC.ContainsKey(TCP_MO_TEMP) Then
                    If Not Final_EQP_DIC.ContainsKey(TCP_MO_TEMP) Then
                        Final_EQP_DIC.Add(TCP_MO_TEMP, "现网未定义")
                    End If
                    outCDUfile.writeLine(Split(TCP_CELL_TEMP, "#")(0) & "	" & Split(TCP_CELL_TEMP, "#")(1) & "	" & TCP_MO_TEMP & "	" & "无CDU类型" & "	" & "现网未定义" & "	" & "现网未定义" & "	" & "无法准确判断")
                End If
            End If
            sender.reportProgress(50 + m / TCP_CELL_MO_DIC.Count * 50, "CDU类型检查..." & Mid(50 + m / (TCP_CELL_MO_DIC.Count) * 50, 1, 2) & "%")
        Next
        sender.reportProgress(100, "CDU类型检查...100%@=>CDU类型检查完成...100%")

        outCDUfile.close()
        outCDUfile = Nothing
        csv2xlsx(outFolder & "CDU类型" & fileCreateTime & ".xls")
    End Sub
    Private Sub doCellFile(outFolder As String, fileCreateTime As String, sender As Object, e As DoWorkEventArgs)
        Dim outFile As Object
        Dim outFilewithTRXS As Object
        Dim templine As String, temp_arr As Object, m As Integer = 0, i As Integer = 0
        Dim Ncell_TEMP As String, Ncell_LAC As String, Ncell_CI As String, Ncell_lac_ci_temp As String, TCH_TEMP As String, TRX_TEMP As String
        Dim Ncell_arr As Object
        Dim CELL As String, ARFCN As String, BSIC As String, Lat As String, Lon As String, LAC As String, CI As String, ANT_DIRECTION As String, ANT_BEAM_WIDTH As String
        Dim MCC As String, MNC As String, NoSiteCell As String
        CELL = ""
        ARFCN = ""
        BSIC = ""
        Lat = ""
        Lon = ""
        LAC = ""
        CI = ""
        MCC = ""
        MNC = ""
        NoSiteCell = ""
        ANT_DIRECTION = ""
        ANT_BEAM_WIDTH = ""
        TCH_TEMP = ""
        TRX_TEMP = ""
        Ncell_TEMP = ""
        Ncell_LAC = ""
        Ncell_CI = ""
        Ncell_lac_ci_temp = ""
        sender.reportProgress(0, "制作.CELL文件..." & "0%")
        If Not fso.folderExists(outFolder & "TEMS cell\") Then
            fso.createFolder(outFolder & "TEMS cell\")
        End If
        outFile = fso.OpenTextFile(outFolder & "TEMS cell\TEMS_" & fileCreateTime & ".cel", 2, 1, 0)
        '2 TEMS_-_Cell_names
        'CELL	ARFCN	BSIC	Lat	Lon	MCC	MNC	LAC	CI	ANT_DIRECTION	ANT_BEAM_WIDTH
        For Y = 1 To MaxTCHNo
            If Y = 1 Then
                TCH_TEMP = "TCH_ARFCN_" & Y
            Else
                TCH_TEMP = Trim(TCH_TEMP & "	" & "TCH_ARFCN_" & Y)
            End If
        Next
        For x = 1 To MaxNeighborNo
            If x = 1 Then
                Ncell_TEMP = "LAC_N_" & x & "	" & "CI_N_" & x
            Else
                Ncell_TEMP = Trim(Ncell_TEMP & "	" & "LAC_N_" & x & "	" & "CI_N_" & x)
            End If
        Next
        outFile.writeLine("2 TEMS_-_Cell_names")
        outFile.writeLine("CELL" & "	" & "ARFCN" & "	" & "BSIC" & "	" & "Lat" & "	" & "Lon" & "	" & "MCC" & "	" & "MNC" & "	" & "LAC" & "	" & "CI" & "	" & "ANT_DIRECTION" & "	" & "ANT_BEAM_WIDTH" & "	" & TCH_TEMP & "	" & Ncell_TEMP)
        TCH_TEMP = ""
        Ncell_TEMP = ""
        FileClose(3)
        FileOpen(3, siteFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        Try
            Do While Not EOF(3)
                m = m + 1
                templine = LineInput(3)
                temp_arr = Split(templine, "	")
                If m = 1 Then
                    If UCase(temp_arr(0)) <> "CELL" Or Strings.Left(UCase(temp_arr(2)), 3) <> "LON" Or Strings.Left(UCase(temp_arr(3)), 3) <> "LAT" Or UCase(temp_arr(4)) <> "DIR" Then
                        MsgBox("Mcom Site文件表头不正确，请重新导入！" & vbCrLf & templine, MsgBoxStyle.Critical)
                        m = 0
                        FileClose(3)
                        Exit Sub
                    End If
                Else
                    cellOfSite_arr.Add(temp_arr(0))
                    Lat_arr.Add(temp_arr(3))
                    Lon_arr.Add(temp_arr(2))
                    Dir_arr.Add(temp_arr(4))
                    SiteName_arr.Add(temp_arr(12))
                End If
            Loop
        Catch ex As Exception
            MsgBox("Mcom Site文件不正确，请重新导入！", MsgBoxStyle.Critical)
            Exit Sub
        Finally
            m = 0
            FileClose(3)
        End Try
        'CELL	ARFCN	BSIC	Lat	Lon	MCC	MNC	LAC	CI	ANT_DIRECTION	ANT_BEAM_WIDTH
        Try
            For X = 0 To DEP_BSC_CELL.Count - 1
                CELL = Strings.Split(DEP_BSC_CELL(X), "#")(1)
                ARFCN = BCCH_arr(X)
                BSIC = BSIC_arr(X)
                LAC = LAI_arr(X)
                CI = CI_arr(X)
                If S_cell_Neighbour.Contains(CELL) Then
                    If InStr(N_cell_Neighbour(S_cell_Neighbour.IndexOf(CELL)), " ") Then
                        Ncell_arr = Split(N_cell_Neighbour(S_cell_Neighbour.IndexOf(CELL)), " ")
                        For i = 0 To UBound(Ncell_arr)
                            If i = 0 Then
                                If DEP_Cell_arr.Contains(Ncell_arr(i)) Then
                                    Ncell_lac_ci_temp = LAI_arr(DEP_Cell_arr.IndexOf(Ncell_arr(i))) & "	" & CI_arr(DEP_Cell_arr.IndexOf(Ncell_arr(i)))
                                Else
                                    Ncell_lac_ci_temp = Ncell_arr(i) & "	" & Ncell_arr(i)
                                End If
                            Else
                                If DEP_Cell_arr.Contains(Ncell_arr(i)) Then
                                    Ncell_lac_ci_temp = Ncell_lac_ci_temp & "	" & LAI_arr(DEP_Cell_arr.IndexOf(Ncell_arr(i))) & "	" & CI_arr(DEP_Cell_arr.IndexOf(Ncell_arr(i)))
                                Else
                                    Ncell_lac_ci_temp = Ncell_lac_ci_temp & "	" & Ncell_arr(i) & "	" & Ncell_arr(i)
                                End If
                            End If
                        Next
                    Else
                        If DEP_Cell_arr.Contains(N_cell_Neighbour(S_cell_Neighbour.IndexOf(CELL))) Then
                            Ncell_lac_ci_temp = LAI_arr(DEP_Cell_arr.IndexOf(N_cell_Neighbour(S_cell_Neighbour.IndexOf(CELL)))) & "	" & CI_arr(DEP_Cell_arr.IndexOf(N_cell_Neighbour(S_cell_Neighbour.IndexOf(CELL))))
                        Else
                            Ncell_lac_ci_temp = N_cell_Neighbour(S_cell_Neighbour.IndexOf(CELL)) & "	" & N_cell_Neighbour(S_cell_Neighbour.IndexOf(CELL))
                        End If
                    End If
                Else
                    Ncell_lac_ci_temp = ""
                End If
                If CFP_BSC_CELL.Contains(DEP_BSC_CELL(X)) Then
                    TRX_TEMP = Trim(TRX_ARR(CFP_BSC_CELL.IndexOf(DEP_BSC_CELL(X))))
                    If InStr(TRX_TEMP, " ") Then
                        For p = 0 To MaxTCHNo - 1
                            If p = 0 Then
                                If Split(TRX_TEMP, " ")(p) <> BCCH_arr(X) Then
                                    TCH_TEMP = Split(TRX_TEMP, " ")(p)
                                End If
                            ElseIf p <> 0 And p <= UBound(Split(TRX_TEMP, " ")) Then
                                If Split(TRX_TEMP, " ")(p) <> BCCH_arr(X) Then
                                    If TCH_TEMP = "" Then
                                        TCH_TEMP = Split(TRX_TEMP, " ")(p)
                                    Else
                                        TCH_TEMP = TCH_TEMP & "	" & Split(TRX_TEMP, " ")(p)
                                    End If
                                End If
                            ElseIf p <> 0 And p > UBound(Split(TRX_TEMP, " ")) Then
                                TCH_TEMP = TCH_TEMP & "	" & ""
                            End If
                        Next
                        If UBound(Split(TCH_TEMP, "	")) < MaxTCHNo - 1 Then
                            TCH_TEMP = TCH_TEMP & "	" & ""
                        End If
                    Else
                        For Y = 1 To MaxTCHNo
                            If Y = 1 Then
                                If TRX_TEMP <> BCCH_arr(X) Then
                                    TCH_TEMP = TRX_TEMP
                                End If
                            Else
                                TCH_TEMP = TCH_TEMP & "	" & ""
                            End If
                        Next
                    End If
                Else
                    For Y = 1 To MaxTCHNo
                        If Y = 1 Then
                            TCH_TEMP = ""
                        Else
                            TCH_TEMP = TCH_TEMP & "	" & ""
                        End If
                    Next
                End If
                If InStr(CGI_arr(X), "-") > 0 Then
                    MCC = Split(CGI_arr(X), "-")(0)
                    MNC = Split(CGI_arr(X), "-")(1)
                Else
                    MCC = ""
                    MNC = ""
                End If

                If cellOfSite_arr.Contains(CELL) Then
                    Lat = Lat_arr(cellOfSite_arr.IndexOf(CELL))
                    Lon = Lon_arr(cellOfSite_arr.IndexOf(CELL))
                    ANT_DIRECTION = Dir_arr(cellOfSite_arr.IndexOf(CELL))
                    If ANT_DIRECTION = "360" Then
                        ANT_BEAM_WIDTH = "360"
                    Else
                        ANT_BEAM_WIDTH = "65"
                    End If
                    CELL = CELL & "_" & SiteName_arr(cellOfSite_arr.IndexOf(CELL))
                Else
                    Lat = ""
                    Lon = ""
                    ANT_DIRECTION = ""
                    ANT_BEAM_WIDTH = ""
                    NoSiteCell = Trim("=>" & CELL & "在Mcom site中无数据！" & vbCrLf & NoSiteCell)
                End If
                outFile.writeLine(CELL & "	" & ARFCN & "	" & BSIC & "	" & Lat & "	" & Lon & "	" & MCC & "	" & MNC & "	" & LAC & "	" & CI & "	" & ANT_DIRECTION & "	" & ANT_BEAM_WIDTH & "	" & TCH_TEMP & "	" & Ncell_lac_ci_temp)

                TCH_TEMP = ""
                Ncell_lac_ci_temp = ""
                sender.reportProgress(X / (DEP_BSC_CELL.Count - 1) * 100, "制作.CELL文件..." & Mid(X / (DEP_BSC_CELL.Count - 1) * 100, 1, 2) & "%")
            Next
            sender.reportProgress(100, "制作.CELL文件完成..100% @=>.CELL文件制作完成..100%@" & NoSiteCell)
        Catch ex As Exception
            MsgBox(".cel文件制作失败！请检查数据！" & ex.ToString, MsgBoxStyle.Critical)
        Finally
            outFile.close()
            outFile = Nothing
            outFilewithTRXS = Nothing
            TCH_TEMP = ""
            Ncell_lac_ci_temp = ""
        End Try
    End Sub
    Private Sub doCellConfig(outFolder As String, fileCreateTime As String, sender As Object, e As DoWorkEventArgs)
        Dim outFile As Object
        Dim templine As String, temp_arr As Object, m As Integer = 0
        Dim BSC As String, CELL As String, SITE As String, SITENAME As String, LONGITUDE As String, LATITUDE As String, CHGR As String, CHGRSTATE As String
        Dim CSYSTYPE As String, CGI As String, LAI As String, LAC_CI As String, CI As String, BSIC As String, BCCHNO As String, DCHNO As String
        Dim 频点数 As String, TRX数 As String, 频点与TRX差 As String, FAULT_TRX As String, DIP As String, DIPNO As String, DIPSTATE As String, MO_TG As String, CDU类型 As String, 设备类型 As String, 载频类型 As String, 基站TRX数 As String, 基站小区数 As String, 基站配置 As String
        Dim cfpIndepNo As Integer
        Dim 基站配置_TEMP As String, 载频类型_temp As String
        Dim site_Config_dic As New Dictionary(Of String, String), temp_siteName As String, site_TRX_no_temp As Integer = 0, site_config_cellTrx As String, site_cells_no As Integer = 0
        Dim aa As Integer
        Dim TRX_BSC_MO_CELL_DIC_KEY_LIST As List(Of String)
        'BSC	CELL	SITE	SITENAME	LONGITUDE	LATITUDE	CHGR	CHGRSTATE	CSYSTYPE	CGI	LAI	LACCI	CI	BSIC	BCCHNO	DCHNO	频点数	TRX数	频点与TRX差	FAULT_TRX	DIP	DIPNO	DIPSTATE	MO_TG	CDU类型	设备类型	载频类型	基站TRX数	基站小区数	基站配置
        cfpIndepNo = 0
        基站配置_TEMP = ""
        载频类型_temp = ""
        temp_siteName = ""
        site_config_cellTrx = ""
        aa = 0
        BSC = ""
        CELL = ""
        SITE = ""
        SITENAME = ""
        LONGITUDE = ""
        LATITUDE = ""
        CHGR = ""
        CHGRSTATE = ""
        CSYSTYPE = ""
        CGI = ""
        LAI = ""
        LAC_CI = ""
        CI = ""
        BSIC = ""
        BCCHNO = ""
        DCHNO = ""
        频点数 = ""
        TRX数 = ""
        频点与TRX差 = ""
        FAULT_TRX = ""
        DIP = ""
        DIPNO = ""
        DIPSTATE = ""
        MO_TG = ""
        CDU类型 = ""
        设备类型 = ""
        载频类型 = ""
        基站TRX数 = ""
        基站小区数 = ""
        基站配置 = ""


        fso = CreateObject("Scripting.FileSystemObject")
        sender.reportProgress(0, "制作小区基础配置信息..." & "0%")
        outFile = fso.OpenTextFile(outFolder & "小区基础配置信息_" & fileCreateTime & ".csv", 2, 1, 0)
        outFile.writeLine("BSC" & "," & "CELL" & "," & "SITE" & "," & "SITENAME" & "," & "LONGITUDE" & "," & "LATITUDE" & "," & "CHGR" & "," & "CHGRSTATE" & "," & "CSYSTYPE" & "," & "CGI" & "," & "LAI" & "," & "LAC_CI" & "," & "CI" & "," & "BSIC" & "," & "BCCHNO" & "," & "DCHNO" & "," & "频点数" & "," & "TRX数" & "," & "频点与TRX差" & "," & "FAULT_TRX" & "," & "DIP" & "," & "DIPNO" & "," & "DIPSTATE" & "," & "MO_TG" & "," & "CDU类型" & "," & "设备类型" & "," & "载频类型" & "," & "基站TRX数" & "," & "基站小区数" & "," & "基站配置")
        '========================读Site文件
        If siteFileName <> "" Then
            FileClose(3)
            FileOpen(3, siteFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
            Try
                Do While Not EOF(3)
                    m = m + 1
                    templine = LineInput(3)
                    temp_arr = Split(templine, "	")
                    If m = 1 Then
                        If UCase(temp_arr(0)) <> "CELL" Or Strings.Left(UCase(temp_arr(2)), 3) <> "LON" Or Strings.Left(UCase(temp_arr(3)), 3) <> "LAT" Or UCase(temp_arr(4)) <> "DIR" Then
                            MsgBox("Mcom Site文件不正确，请重新导入！", MsgBoxStyle.Critical)
                            m = 0
                            FileClose(3)
                            Exit Sub
                        End If
                    Else
                        cellOfSite_arr.Add(temp_arr(0))
                        Lat_arr.Add(temp_arr(3))
                        Lon_arr.Add(temp_arr(2))
                        Dir_arr.Add(temp_arr(4))
                        SiteName_arr.Add(temp_arr(12))
                    End If
                Loop
            Catch ex As Exception
                MsgBox("Mcom Site文件不正确，请重新导入！", MsgBoxStyle.Critical)
            Finally
                m = 0
                FileClose(3)
            End Try
        End If
        '======================================================================
        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@整理载频配置@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        Try
            site_Config_dic.Clear()
            TRX_BSC_MO_CELL_DIC_KEY_LIST = TRX_BSC_MO_CELL_DIC.Keys.ToList
            TRX_BSC_MO_CELL_DIC_KEY_LIST.Sort()
            For m = 0 To TRX_BSC_MO_CELL_DIC_KEY_LIST.Count - 1
                temp_siteName = Strings.Left(TRX_BSC_MO_CELL_DIC_KEY_LIST(m), Len(TRX_BSC_MO_CELL_DIC_KEY_LIST(m)) - 1)
                If m = 0 Then
                    If InStr(TRX_BSC_MO_CELL_DIC(TRX_BSC_MO_CELL_DIC_KEY_LIST(m)), "@") > 0 Then
                        site_TRX_no_temp = UBound(Split(TRX_BSC_MO_CELL_DIC(TRX_BSC_MO_CELL_DIC_KEY_LIST(m)), "@")) + 1
                        site_cells_no = 1
                        site_config_cellTrx = "S" & CStr(UBound(Split(TRX_BSC_MO_CELL_DIC(TRX_BSC_MO_CELL_DIC_KEY_LIST(m)), "@")) + 1)
                        site_Config_dic.Add(temp_siteName, site_TRX_no_temp & "$" & site_cells_no & "$" & site_config_cellTrx)
                    Else
                        site_TRX_no_temp = 1
                        site_cells_no = 1
                        site_config_cellTrx = "S1"
                        site_Config_dic.Add(temp_siteName, site_TRX_no_temp & "$" & site_cells_no & "$" & site_config_cellTrx)
                    End If
                Else
                    If site_Config_dic.ContainsKey(temp_siteName) Then
                        If InStr(TRX_BSC_MO_CELL_DIC(TRX_BSC_MO_CELL_DIC_KEY_LIST(m)), "@") > 0 Then
                            site_TRX_no_temp = Split(site_Config_dic(temp_siteName), "$")(0) + UBound(Split(TRX_BSC_MO_CELL_DIC(TRX_BSC_MO_CELL_DIC_KEY_LIST(m)), "@")) + 1
                            site_cells_no = Split(site_Config_dic(temp_siteName), "$")(1) + 1
                            site_config_cellTrx = Split(site_Config_dic(temp_siteName), "$")(2) & "/" & CStr(UBound(Split(TRX_BSC_MO_CELL_DIC(TRX_BSC_MO_CELL_DIC_KEY_LIST(m)), "@")) + 1)
                            site_Config_dic(temp_siteName) = site_TRX_no_temp & "$" & site_cells_no & "$" & site_config_cellTrx
                        Else
                            site_TRX_no_temp = Split(site_Config_dic(temp_siteName), "$")(0) + 1
                            site_cells_no = Split(site_Config_dic(temp_siteName), "$")(1) + 1
                            site_config_cellTrx = Split(site_Config_dic(temp_siteName), "$")(2) & "/1"
                            site_Config_dic(temp_siteName) = site_TRX_no_temp & "$" & site_cells_no & "$" & site_config_cellTrx
                        End If
                    Else
                        site_TRX_no_temp = 0
                        site_cells_no = 0
                        site_config_cellTrx = ""
                        If InStr(TRX_BSC_MO_CELL_DIC(TRX_BSC_MO_CELL_DIC_KEY_LIST(m)), "@") > 0 Then
                            site_TRX_no_temp = UBound(Split(TRX_BSC_MO_CELL_DIC(TRX_BSC_MO_CELL_DIC_KEY_LIST(m)), "@")) + 1
                            site_cells_no = 1
                            site_config_cellTrx = "S" & CStr(UBound(Split(TRX_BSC_MO_CELL_DIC(TRX_BSC_MO_CELL_DIC_KEY_LIST(m)), "@")) + 1)
                            site_Config_dic.Add(temp_siteName, site_TRX_no_temp & "$" & site_cells_no & "$" & site_config_cellTrx)
                        Else
                            site_TRX_no_temp = 1
                            site_cells_no = 1
                            site_config_cellTrx = "S1"
                            site_Config_dic.Add(temp_siteName, site_TRX_no_temp & "$" & site_cells_no & "$" & site_config_cellTrx)
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox("MO数据整理故障!" & vbCrLf & ex.ToString, MsgBoxStyle.Critical)
            outFile.close()
            outFile = Nothing
            site_Config_dic.Clear()
            fso = Nothing
            Exit Sub
        End Try

        getDIP(outFolder & "RXAPP_RXOTG_" & fileCreateTime & ".csv")
        getDIPState(outFolder & "DTSTP_" & fileCreateTime & ".csv")


        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        Try
            For x = 0 To CFP_BSC_CELL.Count - 1
                BSC = Split(CFP_BSC_CELL(x), "#")(0)
                CELL = Split(CFP_BSC_CELL(x), "#")(1)
                SITE = Strings.Left(CELL, Len(CELL) - 1)
                If cellOfSite_arr.Contains(CELL) Then
                    SITENAME = SiteName_arr(cellOfSite_arr.IndexOf(CELL))
                    LONGITUDE = Lon_arr(cellOfSite_arr.IndexOf(CELL))
                    LATITUDE = Lat_arr(cellOfSite_arr.IndexOf(CELL))
                Else
                    SITENAME = ""
                    LONGITUDE = ""
                    LATITUDE = ""
                End If
                CHGR = CHGR_ARR(x)
                If STP_BSC_CELL_STATE.ContainsKey(CFP_BSC_CELL(x)) Then
                    CHGRSTATE = STP_BSC_CELL_STATE(CFP_BSC_CELL(x))
                Else
                    CHGRSTATE = ""
                End If
                If DEP_BSC_CELL.Contains(CFP_BSC_CELL(x)) Then
                    cfpIndepNo = DEP_BSC_CELL.IndexOf(CFP_BSC_CELL(x))
                    CSYSTYPE = CSYSTYPE_ARR(cfpIndepNo)
                    CGI = CGI_arr(cfpIndepNo)
                    If CGI.Length <> 0 Then
                        LAI = Strings.Left(CGI, InStrRev(CGI, "-") - 1)
                        CI = Strings.Right(CGI, Len(CGI) - InStrRev(CGI, "-"))
                        LAC_CI = Strings.Right(LAI, Len(LAI) - InStrRev(LAI, "-")) & "_" & CI
                    Else
                        LAI = ""
                        LAC_CI = ""
                        CI = ""
                    End If
                    BSIC = BSIC_arr(cfpIndepNo)
                    BCCHNO = BCCH_arr(cfpIndepNo)
                Else
                    CSYSTYPE = ""
                    CGI = ""
                    LAI = ""
                    LAC_CI = ""
                    CI = ""
                    BSIC = ""
                    BCCHNO = ""
                End If
                DCHNO = Trim(TCH_arr(x)) & " "
                DCHNO = Trim(Replace(DCHNO, " " & Trim(BCCHNO) & " ", " "))
                If InStr(DCHNO, "] [") > 0 Then
                    DCHNO = Trim(Replace(DCHNO, Mid(DCHNO, InStr(DCHNO, "] [") - 2, 5), "["))
                End If
                If Strings.Right(DCHNO, 1) = "]" Then
                    DCHNO = Trim(Strings.Left(DCHNO, Len(DCHNO) - 3))
                End If
                If Strings.Right(DCHNO, 1) = "]" Then
                    DCHNO = Trim(Strings.Left(DCHNO, Len(DCHNO) - 3))
                End If
                频点数 = UBound(Split(TRX_ARR(x), " ")) + 1
                If TRX_BSC_MO_CELL_DIC.ContainsKey(CFP_BSC_CELL(x)) Then
                    If InStr(TRX_BSC_MO_CELL_DIC(CFP_BSC_CELL(x)), "@") > 0 Then
                        TRX数 = UBound(Split(TRX_BSC_MO_CELL_DIC(CFP_BSC_CELL(x)), "@")) + 1
                    Else
                        TRX数 = 1
                    End If
                Else
                    TRX数 = 0
                End If
                频点与TRX差 = CStr(CInt(频点数) - CInt(TRX数))

                FAULT_TRX = ""

                '#########################################################
                If TCP_CELL_MO_DIC.ContainsKey(CFP_BSC_CELL(x)) Then
                    If InStr(TCP_CELL_MO_DIC(CFP_BSC_CELL(x)), "|") > 0 Then
                        For q = 0 To UBound(Split(TCP_CELL_MO_DIC(CFP_BSC_CELL(x)), "|"))
                            If q = 0 Then
                                MO_TG = Split(Split(TCP_CELL_MO_DIC(CFP_BSC_CELL(x)), "|")(q), "#")(1)
                                If CDU_DIC.ContainsKey(Split(TCP_CELL_MO_DIC(CFP_BSC_CELL(x)), "|")(q)) Then
                                    CDU类型 = CDU_DIC(Split(TCP_CELL_MO_DIC(CFP_BSC_CELL(x)), "|")(q))
                                    If InStr(CDU类型, ",") > 0 Then
                                        CDU类型 = Replace(CDU类型, ",", "|")
                                    End If
                                Else
                                    CDU类型 = "无CDU类型"
                                End If
                                If Final_EQP_DIC.ContainsKey(Split(TCP_CELL_MO_DIC(CFP_BSC_CELL(x)), "|")(q)) Then
                                    设备类型 = Final_EQP_DIC(Split(TCP_CELL_MO_DIC(CFP_BSC_CELL(x)), "|")(q))
                                Else
                                    设备类型 = "现网未定义"
                                End If
                            Else
                                MO_TG = MO_TG & "/" & Split(Split(TCP_CELL_MO_DIC(CFP_BSC_CELL(x)), "|")(q), "#")(1)
                                If CDU_DIC.ContainsKey(Split(TCP_CELL_MO_DIC(CFP_BSC_CELL(x)), "|")(q)) Then
                                    If CDU类型 <> CDU_DIC(Split(TCP_CELL_MO_DIC(CFP_BSC_CELL(x)), "|")(q)) Or InStr(CDU类型, CDU_DIC(Split(TCP_CELL_MO_DIC(CFP_BSC_CELL(x)), "|")(q))) > 0 Then
                                        CDU类型 = CDU类型 & "/" & CDU_DIC(Split(TCP_CELL_MO_DIC(CFP_BSC_CELL(x)), "|")(q))
                                    End If
                                    If InStr(CDU类型, ",") > 0 Then
                                        CDU类型 = Replace(CDU类型, ",", "|")
                                    End If
                                Else
                                    CDU类型 = CDU类型 & "/" & "无CDU类型"
                                End If
                                If Final_EQP_DIC.ContainsKey(Split(TCP_CELL_MO_DIC(CFP_BSC_CELL(x)), "|")(q)) Then
                                    设备类型 = 设备类型 & "/" & Final_EQP_DIC(Split(TCP_CELL_MO_DIC(CFP_BSC_CELL(x)), "|")(q))
                                Else
                                    设备类型 = 设备类型 & "/" & "现网未定义"
                                End If
                            End If
                        Next
                    Else
                        MO_TG = Split(TCP_CELL_MO_DIC(CFP_BSC_CELL(x)), "#")(1)
                        If CDU_DIC.ContainsKey(TCP_CELL_MO_DIC(CFP_BSC_CELL(x))) Then
                            CDU类型 = CDU_DIC(TCP_CELL_MO_DIC(CFP_BSC_CELL(x)))
                            If InStr(CDU类型, ",") > 0 Then
                                CDU类型 = Replace(CDU类型, ",", "|")
                            End If
                        Else
                            CDU类型 = "无CDU类型"
                        End If
                        If Final_EQP_DIC.ContainsKey(TCP_CELL_MO_DIC(CFP_BSC_CELL(x))) Then
                            设备类型 = Final_EQP_DIC(TCP_CELL_MO_DIC(CFP_BSC_CELL(x)))
                        Else
                            设备类型 = "现网未定义"
                        End If
                    End If
                Else
                    MO_TG = ""
                    CDU类型 = ""
                    设备类型 = ""
                End If

                '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                If InStr(MO_TG, "/") > 0 Then
                    For j = 0 To UBound(Split(MO_TG, "/"))
                        If j = 0 Then
                            If dipDic.ContainsKey(BSC & "#" & Split(MO_TG, "/")(j)) Then
                                DIP = Replace(dipDic(BSC & "#" & Split(MO_TG, "/")(j)), "@", "")
                            End If
                        Else
                            If dipDic.ContainsKey(BSC & "#" & Split(MO_TG, "/")(j)) Then
                                DIP = DIP & "//" & Replace(dipDic(BSC & "#" & Split(MO_TG, "/")(j)), "@", "")
                            End If
                        End If
                    Next
                Else
                    If dipDic.ContainsKey(BSC & "#" & MO_TG) Then
                        DIP = Replace(dipDic(BSC & "#" & MO_TG), "@", "")
                    End If
                End If
                If InStr(DIP, "//") > 0 Then
                    For m = 0 To UBound(Split(DIP, "//"))
                        If m = 0 Then
                            If dipStateDic.ContainsKey(BSC & "#" & Split(DIP, "//")(m)) Then
                                DIPSTATE = dipStateDic(BSC & "#" & Split(DIP, "//")(m))
                            Else
                                DIPSTATE = ""
                            End If
                        Else
                            If dipStateDic.ContainsKey(BSC & "#" & Split(DIP, "//")(m)) Then
                                DIPSTATE = DIPSTATE & "//" & dipStateDic(BSC & "#" & Split(DIP, "//")(m))
                            Else
                                DIPSTATE = ""
                            End If
                        End If
                    Next
                Else
                    If dipStateDic.ContainsKey(BSC & "#" & DIP) Then
                        DIPSTATE = dipStateDic(BSC & "#" & DIP)
                    Else
                        DIPSTATE = ""
                    End If
                End If
                If InStr(DIP, "RAL2") > 0 Then
                    DIPNO = Replace(DIP, "RAL2", "")
                ElseIf InStr(DIP, "RBL2") > 0 Then
                    DIPNO = Replace(DIP, "RBL2", "")
                ElseIf InStr(DIP, "RALT") > 0 Then
                    DIPNO = Replace(DIP, "RALT", "")
                ElseIf InStr(DIP, "RBLT") > 0 Then
                    DIPNO = Replace(DIP, "RBLT", "")
                ElseIf InStr(DIP, "RA2") > 0 Then
                    DIPNO = Replace(DIP, "RA2", "")
                ElseIf InStr(DIP, "RB2") > 0 Then
                    DIPNO = Replace(DIP, "RB2", "")
                ElseIf InStr(DIP, "RAL") > 0 Then
                    DIPNO = Replace(DIP, "RAL", "")
                ElseIf InStr(DIP, "RBL") > 0 Then
                    DIPNO = Replace(DIP, "RBL", "")
                End If
                '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                If TRX_BSC_MO_CELL_DIC.ContainsKey(CFP_BSC_CELL(x)) Then
                    If InStr(TRX_BSC_MO_CELL_DIC(CFP_BSC_CELL(x)), "@") > 0 Then
                        For bb = 0 To UBound(Split(TRX_BSC_MO_CELL_DIC(CFP_BSC_CELL(x)), "@"))
                            If TRX_TYPE_DIC.ContainsKey(Split(TRX_BSC_MO_CELL_DIC(CFP_BSC_CELL(x)), "@")(bb)) Then
                                载频类型_temp = TRX_TYPE_DIC(Split(TRX_BSC_MO_CELL_DIC(CFP_BSC_CELL(x)), "@")(bb))
                            Else
                                载频类型_temp = ""
                            End If
                            If bb = 0 Then
                                If InStr(载频类型_temp, "-") > 0 Then
                                    载频类型 = Split(载频类型_temp, "-")(0)
                                Else
                                    载频类型 = 载频类型_temp
                                End If
                            Else
                                If InStr(载频类型_temp, "-") > 0 Then
                                    载频类型_temp = Split(载频类型_temp, "-")(0)
                                    If InStr(载频类型, 载频类型_temp) = 0 And 载频类型 <> 载频类型_temp Then
                                        载频类型 = 载频类型 & "|" & 载频类型_temp
                                    End If
                                Else
                                    If InStr(载频类型, 载频类型_temp) = 0 And 载频类型 <> 载频类型_temp Then
                                        载频类型 = 载频类型 & "|" & 载频类型_temp
                                    End If
                                End If
                            End If
                        Next
                    Else
                        If TRX_TYPE_DIC.ContainsKey(TRX_BSC_MO_CELL_DIC(CFP_BSC_CELL(x))) Then
                            载频类型 = TRX_TYPE_DIC(TRX_BSC_MO_CELL_DIC(CFP_BSC_CELL(x)))
                        End If
                    End If
                Else
                    载频类型 = ""
                End If
                If site_Config_dic.ContainsKey(BSC & "#" & SITE) Then
                    基站TRX数 = Split(site_Config_dic(BSC & "#" & SITE), "$")(0)
                    基站小区数 = Split(site_Config_dic(BSC & "#" & SITE), "$")(1)
                    基站配置 = Split(site_Config_dic(BSC & "#" & SITE), "$")(2)
                Else
                    基站TRX数 = ""
                    基站小区数 = ""
                    基站配置 = ""
                End If

                outFile.writeLine(BSC & "," & CELL & "," & SITE & "," & SITENAME & "," & LONGITUDE & "," & LATITUDE & "," & CHGR & "," & CHGRSTATE & "," & CSYSTYPE & "," & CGI & "," & LAI & "," & LAC_CI & "," & CI & "," & BSIC & "," & BCCHNO & "," & DCHNO & "," & 频点数 & "," & TRX数 & "," & 频点与TRX差 & "," & FAULT_TRX & "," & DIP & "," & DIPNO & "," & DIPSTATE & "," & MO_TG & "," & CDU类型 & "," & 设备类型 & "," & 载频类型 & "," & 基站TRX数 & "," & 基站小区数 & "," & 基站配置)
                sender.reportProgress(x / (CFP_BSC_CELL.Count - 1) * 100, "制作小区基础配置信息..." & Mid(x / (CFP_BSC_CELL.Count - 1) * 100, 1, 2) & "%")
            Next
            sender.reportProgress(100, "制作小区基础配置信息..100% @=>制作小区基础配置信息制作完成..100% ")
        Catch ex As Exception
            MsgBox("制作小区基础配置信息文件时出错！" & ex.ToString, MsgBoxStyle.Critical)
            Exit Sub
        Finally
            outFile.close()
            outFile = Nothing
            dipDic.Clear()
            site_Config_dic.Clear()
            dipStateDic.Clear()
            dipStateDic.Clear()
            fso = Nothing
        End Try
        ''''''''''''''''''''''转存为EXCEL文件
        csv2xlsx(outFolder & "小区基础配置信息_" & fileCreateTime & ".csv")

    End Sub
    Private Sub getDIP(filepath As String)
        Dim fileNo As Integer
        Dim lineNo As Integer = 0
        Dim lineStr As String
        Dim dipNo As Integer
        Dim dipTypeStr As String
        Dim dipStr As String
        Dim tempARR As Object
        Dim tempARR2 As Object
        Dim tempSTR As Object
        Dim isExsits As Boolean
        If fso.fileExists(filepath) Then
            fileNo = FreeFile()
            FileOpen(fileNo, filepath, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
            Try
                Do While Not EOF(fileNo)
                    lineStr = LineInput(fileNo)
                    lineNo = lineNo + 1
                    If lineNo <> 1 Then
                        tempARR = Split(lineStr, ",")
                        dipNo = CInt((Split(tempARR(2), "-")(1) - 1) \ 32)
                        dipTypeStr = Split(tempARR(2), "-")(0)
                        If Strings.Right(dipTypeStr, 1) <> "T" Then
                            dipTypeStr = Replace(UCase(dipTypeStr), "T", "")
                        End If
                        dipStr = dipNo & "@" & dipTypeStr
                        If Len(dipStr) > 8 Then
                            If Strings.Right(dipStr, 1) = "2" Then
                                dipStr = Replace(dipStr, "L", "")
                                'Else
                                '    dipStr = Replace(dipStr, "L", "")
                            End If
                        End If
                        If dipDic.ContainsKey(tempARR(0) & "#" & tempARR(1)) Then
                            If dipStr <> dipDic(tempARR(0) & "#" & tempARR(1)) Then
                                If InStr(dipDic(tempARR(0) & "#" & tempARR(1)), "//") > 0 Then
                                    isExsits = False
                                    tempARR2 = Split(dipDic(tempARR(0) & "#" & tempARR(1)), "//")
                                    For X = 0 To UBound(tempARR2)
                                        If tempARR2(X) = dipStr Then
                                            isExsits = True
                                        End If
                                    Next
                                    If Not isExsits Then
                                        tempSTR = dipDic(tempARR(0) & "#" & tempARR(1)) & "//" & dipStr
                                        dipDic(tempARR(0) & "#" & tempARR(1)) = tempSTR
                                    End If
                                Else
                                    tempSTR = dipDic(tempARR(0) & "#" & tempARR(1)) & "//" & dipStr
                                    dipDic(tempARR(0) & "#" & tempARR(1)) = tempSTR
                                End If
                            End If
                        Else
                            dipDic.Add(tempARR(0) & "#" & tempARR(1), dipStr)
                        End If
                    End If
                Loop
                '               getDIP = dipDic
            Catch ex As Exception
                MsgBox("读取DIP信息失败！请检查CDD中是否存在RXAPP指令!" & vbCrLf & ex.ToString, MsgBoxStyle.Critical)
                '           getDIP = New Dictionary(Of String, String)
                dipDic.Clear()
                Exit Sub
            Finally
                FileClose(fileNo)
                '               dipDic.Clear()
            End Try
        End If
    End Sub
    Private Sub getDIPState(filepath As String)
        Dim fileNo As Integer
        Dim lineNo As Integer = 0
        Dim lineStr As String
        Dim dipState As String
        Dim tempARR As Object
        Dim tempARR2 As Object
        Dim isExsits As Boolean
        If fso.fileExists(filepath) Then
            fileNo = FreeFile()
            FileOpen(fileNo, filepath, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
            Try
                Do While Not EOF(fileNo)
                    lineStr = LineInput(fileNo)
                    lineNo = lineNo + 1
                    If lineNo <> 1 Then
                        tempARR = Split(lineStr, ",")
                        dipState = tempARR(3)
                        If dipStateDic.ContainsKey(tempARR(0) & "#" & tempARR(1)) Then
                            If dipState <> dipStateDic(tempARR(0) & "#" & tempARR(1)) Then
                                If InStr(dipStateDic(tempARR(0) & "#" & tempARR(1)), "//") > 0 Then
                                    isExsits = False
                                    tempARR2 = Split(dipStateDic(tempARR(0) & "#" & tempARR(1)), "//")
                                    For X = 0 To UBound(tempARR2)
                                        If tempARR2(X) = dipState Then
                                            isExsits = True
                                        End If
                                    Next
                                    If Not isExsits Then
                                        dipStateDic(tempARR(0) & "#" & tempARR(1)) = dipStateDic(tempARR(0) & "#" & tempARR(1)) & "//" & dipState
                                    End If
                                Else
                                    dipStateDic(tempARR(0) & "#" & tempARR(1)) = dipStateDic(tempARR(0) & "#" & tempARR(1)) & "//" & dipState
                                End If
                            End If
                        Else
                            dipStateDic.Add(tempARR(0) & "#" & tempARR(1), dipState)
                        End If
                    End If
                Loop
                '         getDIPState = dipDic
            Catch ex As Exception
                MsgBox("读取DIP状态信息失败！请检查CDD中是否存在DTSTP指令!", MsgBoxStyle.Critical)
                '         getDIPState = New Dictionary(Of String, String)
                Exit Sub
            Finally
                FileClose(fileNo)
                '           dipDic.Clear()
            End Try
        End If
    End Sub
    Private Sub doSeeSiteFile(outFolder As String, fileCreateTime As String, sender As Object, e As DoWorkEventArgs)
        Dim outFile As Object
        Dim outFilewithTRXS As Object
        Dim templine As String, temp_arr As Object, m As Integer = 0, i As Integer = 0
        Dim Ncell_TEMP As String, Ncell_lac_ci_temp As String, TCH_TEMP As String, TRX_TEMP As String
        Dim Ncell_arr As Object
        Dim CELL As String, BSIC As String, LAC As String, CI As String, ANT_DIRECTION As String
        Dim NCC As String, BCC As String, NoSiteCell As String
        Dim BSC As String, Dtile As String, Height As String, Type As String
        Dim Region As String, TCH As String, Traffic As String, BTS As String, Hop As String, X_long As String, Y_lat As String, bts_name As String
        'BSC	LAC	CI	站号	bts_name	Dir	NCC	BCC	Dtile	height	type	TRX1	TRX2	TRX3	TRX4	TRX5	TRX6	TRX7	TRX8	TRX9	TRX10	TRX11	TRX12	TRX13	TRX14	TRX15	TRX16	adj1	adj2	adj3	adj4	adj5	adj6	adj7	adj8	adj9	adj10	adj11	adj12	adj13	adj14	adj15	adj16	adj17	adj18	adj19	adj20	adj21	adj22	adj23	adj24	adj25	adj26	adj27	adj28	adj29	adj30	adj31	adj32	Region	TCH	Traffic	BTS	hop	X	Y
        BSC = ""
        CELL = ""
        LAC = ""
        CI = ""
        ANT_DIRECTION = ""
        NCC = ""
        BCC = ""
        Dtile = ""
        Height = ""
        Type = ""
        X_long = ""
        Y_lat = ""
        Region = ""
        TCH = ""
        Traffic = ""
        BTS = ""
        Hop = ""
        TCH_TEMP = ""
        NoSiteCell = ""
        Ncell_TEMP = ""
        Ncell_lac_ci_temp = ""
        bts_name = ""
        sender.reportProgress(0, "制作SeeSite文件..." & "0%")
        fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.folderExists(outFolder & "SeeSite_工参\") Then
            fso.createFolder(outFolder & "SeeSite_工参\")
        End If
        Try
            outFile = fso.OpenTextFile(outFolder & "SeeSite_工参\RFCON_" & fileCreateTime & ".CSV", 2, 1, 0)
        Catch ex As Exception
            MsgBox("RFCON_" & fileCreateTime & ".CSV文件正处于打开状态，请关闭后，重试!", MsgBoxStyle.Critical)
            outFile = Nothing
            Exit Sub
        End Try
        For Y = 1 To 16
            If Y = 1 Then
                TCH_TEMP = "TRX" & Y
            Else
                TCH_TEMP = Trim(TCH_TEMP & "," & "TRX" & Y)
            End If
        Next
        For x = 1 To MaxNeighborNo
            If x = 1 Then
                Ncell_TEMP = "adj" & x
            Else
                Ncell_TEMP = Trim(Ncell_TEMP & "," & "adj" & x)
            End If
        Next
        outFile.writeLine("BSC" & "," & "LAC" & "," & "CI" & "," & "站号" & "," & "bts_name" & "," & "X" & "," & "Y" & "," & "Dir" & "," & "type" & "," & "NCC" & "," & "BCC" & "," & TCH_TEMP & "," & Ncell_TEMP & "," & "Region" & "," & "TCH" & "," & "Traffic" & "," & "BTS" & "," & "hop" & "," & "Dtile" & "," & "height")
        If siteFileName <> "" Then
            FileClose(3)
            FileOpen(3, siteFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
            Try
                Do While Not EOF(3)
                    m = m + 1
                    templine = LineInput(3)
                    temp_arr = Split(templine, "	")
                    If m = 1 Then
                        If UCase(temp_arr(0)) <> "CELL" Or UCase(temp_arr(2)) <> "LONGITUDE" Or UCase(temp_arr(3)) <> "LATITUDE" Or UCase(temp_arr(4)) <> "DIR" Then
                            MsgBox("Mcom Site文件不正确，请重新导入！", MsgBoxStyle.Critical)
                            m = 0
                            FileClose(3)
                            Exit Sub
                        End If
                    Else
                        cellOfSite_arr.Add(temp_arr(0))
                        Lat_arr.Add(temp_arr(3))
                        Lon_arr.Add(temp_arr(2))
                        Dir_arr.Add(temp_arr(4))
                        Height_arr.Add(temp_arr(5))
                        Dtile_arr.Add(temp_arr(6))
                        SiteName_arr.Add(temp_arr(12))
                    End If
                Loop
            Catch ex As Exception
                MsgBox("Mcom Site文件不正确，请重新导入！", MsgBoxStyle.Critical)
                Exit Sub
            Finally
                m = 0
                FileClose(3)
            End Try
        End If
        
        'BSC	LAC	CI	站号	bts_name	Dir	NCC	BCC	Dtile	height	type	TRX1	TRX2	TRX3	TRX4	TRX5	TRX6	TRX7	TRX8	TRX9	TRX10	TRX11	TRX12	TRX13	TRX14	TRX15	TRX16	adj1	adj2	adj3	adj4	adj5	adj6	adj7	adj8	adj9	adj10	adj11	adj12	adj13	adj14	adj15	adj16	adj17	adj18	adj19	adj20	adj21	adj22	adj23	adj24	adj25	adj26	adj27	adj28	adj29	adj30	adj31	adj32	Region	TCH	Traffic	BTS	hop	X	Y

        Try
            For X = 0 To DEP_BSC_CELL.Count - 1
                BSC = Strings.Split(DEP_BSC_CELL(X), "#")(0)
                CELL = Strings.Split(DEP_BSC_CELL(X), "#")(1)
                BSIC = BSIC_arr(X)
                Type = CSYSTYPE_ARR(X)
                If Len(BSIC) = 1 Then
                    NCC = 0
                    BCC = BSIC
                Else
                    NCC = Strings.Left(BSIC, 1)
                    BCC = Strings.Right(BSIC, 1)
                End If
                LAC = LAI_arr(X)
                CI = CI_arr(X)
                If S_cell_Neighbour.Contains(CELL) Then
                    If InStr(N_cell_Neighbour(S_cell_Neighbour.IndexOf(CELL)), " ") > 0 Then
                        Ncell_arr = Split(N_cell_Neighbour(S_cell_Neighbour.IndexOf(CELL)), " ")
                        For i = 0 To UBound(Ncell_arr)
                            If i = 0 Then
                                '             Ncell_lac_ci_temp = Ncell_arr(i)
                                If DEP_Cell_arr.Contains(Ncell_arr(i)) Then
                                    Ncell_lac_ci_temp = CI_arr(DEP_Cell_arr.IndexOf(Ncell_arr(i)))
                                Else
                                    Ncell_lac_ci_temp = Ncell_arr(i)
                                End If
                            Else
                                ' Ncell_lac_ci_temp = Ncell_lac_ci_temp & "," & Ncell_arr(i)
                                If DEP_Cell_arr.Contains(Ncell_arr(i)) Then
                                    Ncell_lac_ci_temp = Ncell_lac_ci_temp & "," & CI_arr(DEP_Cell_arr.IndexOf(Ncell_arr(i)))
                                Else
                                    Ncell_lac_ci_temp = Ncell_lac_ci_temp & "," & Ncell_arr(i)
                                End If
                            End If
                        Next
                    Else
                        '   Ncell_lac_ci_temp = N_cell_Neighbour(S_cell_Neighbour.IndexOf(CELL))
                        If DEP_Cell_arr.Contains(N_cell_Neighbour(S_cell_Neighbour.IndexOf(CELL))) Then
                            Ncell_lac_ci_temp = CI_arr(DEP_Cell_arr.IndexOf(N_cell_Neighbour(S_cell_Neighbour.IndexOf(CELL))))
                        Else
                            Ncell_lac_ci_temp = N_cell_Neighbour(S_cell_Neighbour.IndexOf(CELL))
                        End If
                    End If
                Else
                    Ncell_lac_ci_temp = ""
                End If
                If MaxNeighborNo > UBound(Split(Ncell_lac_ci_temp, ",")) + 1 Then
                    For d = 1 To MaxNeighborNo - UBound(Split(Ncell_lac_ci_temp, ",")) - 1
                        Ncell_lac_ci_temp = Ncell_lac_ci_temp & "," & ""
                    Next
                End If
                If CFP_BSC_CELL.Contains(DEP_BSC_CELL(X)) Then
                    TRX_TEMP = Trim(TRX_ARR(CFP_BSC_CELL.IndexOf(DEP_BSC_CELL(X))))
                    TCH_TEMP = BCCH_arr(X)
                    If InStr(TRX_TEMP, " ") > 0 Then
                        For p = 0 To 15
                            If p <= UBound(Split(TRX_TEMP, " ")) Then
                                If Split(TRX_TEMP, " ")(p) <> BCCH_arr(X) Then
                                    If TCH_TEMP = "" Then
                                        TCH_TEMP = Split(TRX_TEMP, " ")(p)
                                    Else
                                        TCH_TEMP = TCH_TEMP & "," & Split(TRX_TEMP, " ")(p)
                                    End If
                                End If
                            ElseIf p <= 15 And p > UBound(Split(TRX_TEMP, " ")) Then
                                TCH_TEMP = TCH_TEMP & "," & ""
                            End If
                        Next
                        'If UBound(Split(TCH_TEMP, "	")) < MaxTCHNo - 1 Then
                        '    TCH_TEMP = TCH_TEMP & "	" & ""
                        'End If
                    Else
                        For Y = 1 To 16
                            If Y = 1 Then
                                If TRX_TEMP <> BCCH_arr(X) Then
                                    TCH_TEMP = TCH_TEMP & "," & TRX_TEMP
                                End If
                            Else
                                TCH_TEMP = TCH_TEMP & "," & ""
                            End If
                        Next
                    End If
                Else
                    For Y = 1 To 15
                        TCH_TEMP = TCH_TEMP & "	" & ""
                    Next
                End If
                If cellOfSite_arr.Contains(CELL) Then
                    Y_lat = Lat_arr(cellOfSite_arr.IndexOf(CELL))
                    X_long = Lon_arr(cellOfSite_arr.IndexOf(CELL))
                    Dtile = Dtile_arr(cellOfSite_arr.IndexOf(CELL))
                    Height = Height_arr(cellOfSite_arr.IndexOf(CELL))
                    ANT_DIRECTION = Dir_arr(cellOfSite_arr.IndexOf(CELL))
                    If ANT_DIRECTION = "360" Then
                        ANT_DIRECTION = "-1"
                    End If
                    bts_name = SiteName_arr(cellOfSite_arr.IndexOf(CELL))

                Else
                    Y_lat = ""
                    X_long = ""
                    ANT_DIRECTION = ""
                    Dtile = ""
                    Height = ""
                    bts_name = ""
                    NoSiteCell = Trim("=>" & CELL & "在Mcom site中无数据！" & vbCrLf & NoSiteCell)
                End If

                'BSC	LAC	CI	站号	bts_name	X	Y	Dir	type	NCC	BCC	TRX1	TRX2	TRX3	TRX4	TRX5	TRX6	TRX7	TRX8	TRX9	TRX10	TRX11	TRX12	TRX13	TRX14	TRX15	TRX16	adj1	adj2	adj3	adj4	adj5	adj6	adj7	adj8	adj9	adj10	adj11	adj12	adj13	adj14	adj15	adj16	adj17	adj18	adj19	adj20	adj21	adj22	adj23	adj24	adj25	adj26	adj27	adj28	adj29	adj30	adj31	adj32	Region	TCH	Traffic	BTS	hop	Dtile	height
                outFile.writeLine(BSC & "," & LAC & "," & CI & "," & CELL & "," & bts_name & "," & X_long & "," & Y_lat & "," & ANT_DIRECTION & "," & Type & "," & NCC & "," & BCC & "," & TCH_TEMP & "," & Ncell_lac_ci_temp & "," & Region & "," & TCH & "," & Traffic & "," & BTS & "," & Hop & "," & Dtile & "," & Height)
                TCH_TEMP = ""
                Ncell_lac_ci_temp = ""
                sender.reportProgress(X / (DEP_BSC_CELL.Count - 1) * 100, "制作.CELL文件..." & Mid(X / (DEP_BSC_CELL.Count - 1) * 100, 1, 2) & "%")
            Next
            sender.reportProgress(100, "制作SeeSite文件完成..100% @=>SeeSite文件制作完成..100%@" & NoSiteCell)
        Catch ex As Exception
            MsgBox("SeeSite文件制作失败！请检查数据！" & ex.ToString, MsgBoxStyle.Critical)
            Exit Sub
        Finally
            outFile.close()
            outFile = Nothing
            outFilewithTRXS = Nothing
            TCH_TEMP = ""
            Ncell_lac_ci_temp = ""
        End Try
        csv2xlsx(outFolder & "SeeSite_工参\RFCON_" & fileCreateTime & ".CSV")
    End Sub
    private sub doIOEXP(sender As Object)
        Dim templine As String
        Dim filepath As String
        Dim outputFile As Object
        filepath = outPath & "IOEXP_" & CDDTime & ".CSV"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "IDENTITY")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If useIOEXP.Checked Then
                        If InStr(templine, "/") Then
                            BSCNAME = Strings.Left(templine, InStr(templine, "/") - 1)
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 100)))
                            outputFile.close()
                            outputFile = Nothing
                            Exit Sub
                        ElseIf InStr(templine, "_") Then
                            BSCNAME = Strings.Left(templine, InStr(templine, "_") - 1)
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 100)))
                            outputFile.close()
                            outputFile = Nothing
                            Exit Sub
                        End If
                    Else
                        If UCase(Trim(templine)) = "IDENTITY" Then
                            templine = getNowLine()
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 100)))
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
  
    Private Sub doALLIP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim alarmType_time As String, alarmDetails As String
        alarmType_time = ""
        alarmDetails = ""
        filepath = outPath & "ALLIP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "Alarm_Type&Time" & "," & "Details")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If InStr(Trim(templine), "/") Then
                        If alarmType_time <> "" Then
                            outputFile.writeline(BSCNAME & "," & alarmType_time & "," & alarmDetails)
                            alarmDetails = ""
                        End If
                        alarmType_time = Trim(templine)
                    ElseIf UCase(Trim(templine)) <> "ALARM LIST" Then
                        alarmDetails = templine
                        If InStr(alarmDetails, ",") Then
                            alarmDetails = Replace(alarmDetails, ",", ".")
                        End If
                        outputFile.writeline(BSCNAME & "," & alarmType_time & "," & alarmDetails)
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    private sub doCACLP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String

        filepath = outPath & "CACLP_" & CDDtime & ".CSV"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "DATE" & "," & "TIME" & "," & "SUMMERTIME" & "," & "DAY" & "," & "DCAT")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Strings.left(Trim(templine), 13)) = "DATE     TIME" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'DATE     TIME     SUMMERTIME     DAY      DCAT
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 9)) & "," & Trim(Mid(templine, 19, 15)) & "," & Trim(Mid(templine, 34, 9)) & "," & Trim(Mid(templine, 43, 6)))
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    private sub doDBTSP_AXEPARS(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim NAME As String, SETNAME As String, PARID As String, VALUE As String, UNIT As String, DBTSP_CLASS As String, DISTRIB As String
        NAME = ""
        SETNAME = ""
        PARID = ""
        VALUE = ""
        UNIT = ""
        DBTSP_CLASS = ""
        DISTRIB = ""
        'NAME            SETNAME         PARID      VALUE UNIT CLASS   DISTRIB
        '                                STATUS  FCVSET FCVALUE DCINFO FCODE
        filepath = outPath & "DBTSP_AXEPARS_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "NAME" & "," & "SETNAME" & "," & "PARID" & "," & "VALUE" & "," & "UNIT" & "," & "CLASS" & "," & "DISTRIB" & "," & "STATUS" & "," & "FCVSET" & "," & "FCVALUE" & "," & "DCINFO" & "," & "FCODE")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Strings.left(Trim(templine), 23)) = "NAME            SETNAME" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            NAME = Trim(Mid(templine, 1, 16))
                            SETNAME = Trim(Mid(templine, 17, 16))
                            PARID = Trim(Mid(templine, 33, 11))
                            VALUE = Trim(Mid(templine, 44, 6))
                            UNIT = Trim(Mid(templine, 50, 5))
                            DBTSP_CLASS = Trim(Mid(templine, 55, 8))
                            DISTRIB = Trim(Mid(templine, 63, 9))
                        End If
                    ElseIf UCase(Strings.left(templine, 46)) = "                                STATUS  FCVSET" Then
                        templine = getNowLine()
                        outputFile.writeLine(BSCNAME & "," & NAME & "," & SETNAME & "," & PARID & "," & VALUE & "," & UNIT & "," & DBTSP_CLASS & "," & DISTRIB _
                            & "," & Trim(Mid(templine, 33, 8)) & "," & Trim(Mid(templine, 41, 7)) & "," & Trim(Mid(templine, 48, 8)) & "," & Trim(Mid(templine, 56, 7)) & "," & Trim(Mid(templine, 63, 9)))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    private sub doDTSTP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String

        'DIP      TYPE    STATE  LOOP TSLOTL DIPEND   FAULT                SECTION
        filepath = outPath & "DTSTP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "DIP" & "," & "TYPE" & "," & "STATE" & "," & "LOOP" & "," & "TSLOTL" & "," & "DIPEND" & "," & "FAULT" & "," & "SECTION")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "DIP      TYPE    STATE  LOOP TSLOTL DIPEND   FAULT                SECTION" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 8)) & "," & Trim(Mid(templine, 18, 7)) & "," & Trim(Mid(templine, 25, 5)) & "," & Trim(Mid(templine, 30, 7)) & "," & Trim(Mid(templine, 37, 9)) & "," & Trim(Mid(templine, 46, 21)) & "," & Trim(Mid(templine, 67, 8)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "DIGITAL PATH STATE" And Len(Trim(Mid(templine, 10, 8))) <= 4 Then
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 8)) & "," & Trim(Mid(templine, 18, 7)) & "," & Trim(Mid(templine, 25, 5)) & "," & Trim(Mid(templine, 30, 7)) & "," & Trim(Mid(templine, 37, 9)) & "," & Trim(Mid(templine, 46, 21)) & "," & Trim(Mid(templine, 67, 8)))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doDTQUP(sender As Object)
        Dim outputFile As Object
        Dim DKoutPutFile As Object
        Dim DKfilePath As String
        Dim filepath As String
        Dim templine As String
        Dim LEVEL_UN As String, LEVEL_DEG As String, Direction_IN As String, Direction_OUT As String
        Dim DIP As String, SECTION As String, T1 As String, T2 As String
        Dim UN_N_ES As String, UN_N_SES As String, UN_N_UAS As String, UN_N_UAV As String, UN_SLIP As String, UN_IN_SMI As String
        Dim UN_F_ES As String, UN_F_SES As String, UN_F_UAS As String, UN_F_UAV As String, UN_OUT_SMI As String
        Dim DEG_N_ES As String, DEG_N_SES As String, DEG_N_UAS As String, DEG_N_UAV As String, DEG_SLIP As String, DEG_IN_SMI As String
        Dim DEG_F_ES As String, DEG_F_SES As String, DEG_F_UAS As String, DEG_F_UAV As String, DEG_OUT_SMI As String
        Dim SFV As String, SFTI As String
        Dim DK_DIP As String, DK_T1 As String, DK_T2 As String, DK_SLIP As String, DK_SLIP2 As String, DK_UAS As String, DK_UASR As String, DK_UAV1 As String, DK_UASB1 As String, DK_UAV2 As String, DK_UASB2 As String, DK_SECTION As String, DK_ESV As String, DK_SESV As String, DK_DMV As String, DK_ESVR As String, DK_SESVR As String, DK_DMVR As String, DK_SFV As String, DK_SFTI As String, DK_ES2V As String, DK_SES2V As String, DK_DM2V As String, DK_ES2VR As String, DK_SES2VR As String, DK_DM2VR As String, DK_SMI As String

        '''''''光口
        'UNACCEPTABLE PERFORMANCE LEVEL
        'INCOMING DIRECTION
        'DIP      SECTION  T1    N-ES   N-SES  N-UAS  N-UAV  SLIP  SMI
        'OUTGOING DIRECTION
        'DIP      SECTION  T1    F-ES   F-SES  F-UAS  F-UAV  SMI
        'DEGRADED PERFORMANCE LEVEL
        'INCOMING DIRECTION
        'DIP      SECTION  T2    N-ES   N-SES  N-UAS  N-UAV  SLIP  SMI
        'OUTGOING DIRECTION
        'DIP      SECTION  T2    F-ES   F-SES  F-UAS  F-UAV  SMI
        'FRAME SLIPS
        'DIP      SECTION  SFV   SFTI
        LEVEL_UN = ""
        LEVEL_DEG = ""
        Direction_IN = ""
        Direction_OUT = ""
        DIP = ""
        SECTION = ""
        T1 = ""
        T2 = ""
        UN_N_ES = ""
        UN_N_SES = ""
        UN_N_UAS = ""
        UN_N_UAV = ""
        UN_SLIP = ""
        UN_IN_SMI = ""
        UN_F_ES = ""
        UN_F_SES = ""
        UN_F_UAS = ""
        UN_F_UAV = ""
        UN_OUT_SMI = ""
        DEG_N_ES = ""
        DEG_N_SES = ""
        DEG_N_UAS = ""
        DEG_N_UAV = ""
        DEG_SLIP = ""
        DEG_IN_SMI = ""
        DEG_F_ES = ""
        DEG_F_SES = ""
        DEG_F_UAS = ""
        DEG_F_UAV = ""
        DEG_OUT_SMI = ""
        SFV = ""
        SFTI = ""

        '''''电口
        'DIP      T1   T2  SLIP   SLIP2  UAS    UASR  UAV1  UASB1  UAV2  UASB2
        'SECTION  ESV    SESV   DMV    ESVR   SESVR  DMVR   SFV  SFTI
        'SECTION  ES2V   SES2V  DM2V   ES2VR  SES2VR DM2VR  SMI
        DK_DIP = ""
        DK_T1 = ""
        DK_T2 = ""
        DK_SLIP = ""
        DK_SLIP2 = ""
        DK_UAS = ""
        DK_UASR = ""
        DK_UAV1 = ""
        DK_UASB1 = ""
        DK_UAV2 = ""
        DK_UASB2 = ""
        DK_SECTION = ""
        DK_ESV = ""
        DK_SESV = ""
        DK_DMV = ""
        DK_ESVR = ""
        DK_SESVR = ""
        DK_DMVR = ""
        DK_SFV = ""
        DK_SFTI = ""
        DK_ES2V = ""
        DK_SES2V = ""
        DK_DM2V = ""
        DK_ES2VR = ""
        DK_SES2VR = ""
        DK_DM2VR = ""
        DK_SMI = ""

        filepath = outPath & "DTQUP_光口_" & CDDTime & ".csv"
        DKfilePath = outPath & "DTQUP_电口_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "DIP" & "," & "SECTION" & "," & "LEVEL" & "," & "Direction" & "," & "T1" & "," & "T2" & "," & "N-ES" & "," & "N-SES" & "," & "N-UAS" & "," & "N-UAV" & "," & "SLIP" & "," & "SMI" & "," & "Direction" & "," & "F-ES" & "," & "F-SES" & "," & "F-UAS" & "," & "F-UAV" & "," & "SMI" & "," & "SFV" & "," & "SFTI")
        End If
        If fso.fileExists(DKfilePath) Then
            DKoutPutFile = fso.OpenTextFile(DKfilePath, 8, 1, 0)
        Else
            DKoutPutFile = fso.OpenTextFile(DKfilePath, 2, 1, 0)
            DKoutPutFile.writeline("BSC" & "," & "DIP" & "," & "T1" & "," & "T2" & "," & "SLIP" & "," & "SLIP2" & "," & "UAS" & "," & "UASR" & "," & "UAV1" & "," & "UASB1" & "," & "UAV2" & "," & "UASB2" & "," & "SECTION" & "," & "ESV" & "," & "SESV" & "," & "DMV" & "," & "ESVR" & "," & "SESVR" & "," & "DMVR" & "," & "SFV" & "," & "SFTI" & "," & "ES2V" & "," & "SES2V" & "," & "DM2V" & "," & "ES2VR" & "," & "SES2VR" & "," & "DM2VR" & "," & "SMI")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        If DIP <> "" Then
                            outputFile.writeLine(BSCNAME & "," & DIP & "," & SECTION & "," & LEVEL_UN & "," & Direction_IN & "," & T1 & "," & "" & "," & UN_N_ES & "," & UN_N_SES & "," & UN_N_UAS & "," & UN_N_UAV & "," & UN_SLIP & "," & UN_IN_SMI & "," & Direction_OUT & "," & UN_F_ES & "," & UN_F_SES & "," & UN_F_UAS & "," & UN_F_UAV & "," & UN_OUT_SMI & "," & SFV & "," & SFTI)
                            outputFile.writeLine(BSCNAME & "," & DIP & "," & SECTION & "," & LEVEL_DEG & "," & Direction_IN & "," & "" & "," & T2 & "," & DEG_N_ES & "," & DEG_N_SES & "," & DEG_N_UAS & "," & DEG_N_UAV & "," & DEG_SLIP & "," & DEG_IN_SMI & "," & Direction_OUT & "," & DEG_F_ES & "," & DEG_F_SES & "," & DEG_F_UAS & "," & DEG_F_UAV & "," & DEG_OUT_SMI & "," & SFV & "," & SFTI)
                            LEVEL_UN = ""
                            LEVEL_DEG = ""
                            Direction_IN = ""
                            Direction_OUT = ""
                            DIP = ""
                            SECTION = ""
                            T1 = ""
                            T2 = ""
                            UN_N_ES = ""
                            UN_N_SES = ""
                            UN_N_UAS = ""
                            UN_N_UAV = ""
                            UN_SLIP = ""
                            UN_IN_SMI = ""
                            UN_F_ES = ""
                            UN_F_SES = ""
                            UN_F_UAS = ""
                            UN_F_UAV = ""
                            UN_OUT_SMI = ""
                            DEG_N_ES = ""
                            DEG_N_SES = ""
                            DEG_N_UAS = ""
                            DEG_N_UAV = ""
                            DEG_SLIP = ""
                            DEG_IN_SMI = ""
                            DEG_F_ES = ""
                            DEG_F_SES = ""
                            DEG_F_UAS = ""
                            DEG_F_UAV = ""
                            DEG_OUT_SMI = ""
                            SFV = ""
                            SFTI = ""
                        End If
                        If DK_DIP <> "" Then
                            DKoutPutFile.writeLine(BSCNAME & "," & DK_DIP & "," & DK_T1 & "," & DK_T2 & "," & DK_SLIP & "," & DK_SLIP2 & "," & DK_UAS & "," & DK_UASR & "," & DK_UAV1 & "," & DK_UASB1 & "," & DK_UAV2 & "," & DK_UASB2 & "," & DK_SECTION & "," & DK_ESV & "," & DK_SESV & "," & DK_DMV & "," & DK_ESVR & "," & DK_SESVR & "," & DK_DMVR & "," & DK_SFV & "," & DK_SFTI & "," & DK_ES2V & "," & DK_SES2V & "," & DK_DM2V & "," & DK_ES2VR & "," & DK_SES2VR & "," & DK_DM2VR & "," & DK_SMI)
                            DK_DIP = ""
                            DK_T1 = ""
                            DK_T2 = ""
                            DK_SLIP = ""
                            DK_SLIP2 = ""
                            DK_UAS = ""
                            DK_UASR = ""
                            DK_UAV1 = ""
                            DK_UASB1 = ""
                            DK_UAV2 = ""
                            DK_UASB2 = ""
                            DK_SECTION = ""
                            DK_ESV = ""
                            DK_SESV = ""
                            DK_DMV = ""
                            DK_ESVR = ""
                            DK_SESVR = ""
                            DK_DMVR = ""
                            DK_SFV = ""
                            DK_SFTI = ""
                            DK_ES2V = ""
                            DK_SES2V = ""
                            DK_DM2V = ""
                            DK_ES2VR = ""
                            DK_SES2VR = ""
                            DK_DM2VR = ""
                            DK_SMI = ""
                        End If
                        outputFile.close()
                        DKoutPutFile.close()
                        DKoutPutFile = Nothing
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If DIP <> "" Then
                            outputFile.writeLine(BSCNAME & "," & DIP & "," & SECTION & "," & LEVEL_UN & "," & Direction_IN & "," & T1 & "," & "" & "," & UN_N_ES & "," & UN_N_SES & "," & UN_N_UAS & "," & UN_N_UAV & "," & UN_SLIP & "," & UN_IN_SMI & "," & Direction_OUT & "," & UN_F_ES & "," & UN_F_SES & "," & UN_F_UAS & "," & UN_F_UAV & "," & UN_OUT_SMI & "," & SFV & "," & SFTI)
                            outputFile.writeLine(BSCNAME & "," & DIP & "," & SECTION & "," & LEVEL_DEG & "," & Direction_IN & "," & "" & "," & T2 & "," & DEG_N_ES & "," & DEG_N_SES & "," & DEG_N_UAS & "," & DEG_N_UAV & "," & DEG_SLIP & "," & DEG_IN_SMI & "," & Direction_OUT & "," & DEG_F_ES & "," & DEG_F_SES & "," & DEG_F_UAS & "," & DEG_F_UAV & "," & DEG_OUT_SMI & "," & SFV & "," & SFTI)
                            LEVEL_UN = ""
                            LEVEL_DEG = ""
                            Direction_IN = ""
                            Direction_OUT = ""
                            DIP = ""
                            SECTION = ""
                            T1 = ""
                            T2 = ""
                            UN_N_ES = ""
                            UN_N_SES = ""
                            UN_N_UAS = ""
                            UN_N_UAV = ""
                            UN_SLIP = ""
                            UN_IN_SMI = ""
                            UN_F_ES = ""
                            UN_F_SES = ""
                            UN_F_UAS = ""
                            UN_F_UAV = ""
                            UN_OUT_SMI = ""
                            DEG_N_ES = ""
                            DEG_N_SES = ""
                            DEG_N_UAS = ""
                            DEG_N_UAV = ""
                            DEG_SLIP = ""
                            DEG_IN_SMI = ""
                            DEG_F_ES = ""
                            DEG_F_SES = ""
                            DEG_F_UAS = ""
                            DEG_F_UAV = ""
                            DEG_OUT_SMI = ""
                            SFV = ""
                            SFTI = ""
                        End If
                        If DK_DIP <> "" Then
                            DKoutPutFile.writeLine(BSCNAME & "," & DK_DIP & "," & DK_T1 & "," & DK_T2 & "," & DK_SLIP & "," & DK_SLIP2 & "," & DK_UAS & "," & DK_UASR & "," & DK_UAV1 & "," & DK_UASB1 & "," & DK_UAV2 & "," & DK_UASB2 & "," & DK_SECTION & "," & DK_ESV & "," & DK_SESV & "," & DK_DMV & "," & DK_ESVR & "," & DK_SESVR & "," & DK_DMVR & "," & DK_SFV & "," & DK_SFTI & "," & DK_ES2V & "," & DK_SES2V & "," & DK_DM2V & "," & DK_ES2VR & "," & DK_SES2VR & "," & DK_DM2VR & "," & DK_SMI)
                            DK_DIP = ""
                            DK_T1 = ""
                            DK_T2 = ""
                            DK_SLIP = ""
                            DK_SLIP2 = ""
                            DK_UAS = ""
                            DK_UASR = ""
                            DK_UAV1 = ""
                            DK_UASB1 = ""
                            DK_UAV2 = ""
                            DK_UASB2 = ""
                            DK_SECTION = ""
                            DK_ESV = ""
                            DK_SESV = ""
                            DK_DMV = ""
                            DK_ESVR = ""
                            DK_SESVR = ""
                            DK_DMVR = ""
                            DK_SFV = ""
                            DK_SFTI = ""
                            DK_ES2V = ""
                            DK_SES2V = ""
                            DK_DM2V = ""
                            DK_ES2VR = ""
                            DK_SES2VR = ""
                            DK_DM2VR = ""
                            DK_SMI = ""
                        End If
                        outputFile.close()
                        DKoutPutFile.close()
                        DKoutPutFile = Nothing
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "UNACCEPTABLE PERFORMANCE LEVEL" Then
                        'UNACCEPTABLE PERFORMANCE LEVEL
                        If DIP <> "" Then
                            outputFile.writeLine(BSCNAME & "," & DIP & "," & SECTION & "," & LEVEL_UN & "," & Direction_IN & "," & T1 & "," & "" & "," & UN_N_ES & "," & UN_N_SES & "," & UN_N_UAS & "," & UN_N_UAV & "," & UN_SLIP & "," & UN_IN_SMI & "," & Direction_OUT & "," & UN_F_ES & "," & UN_F_SES & "," & UN_F_UAS & "," & UN_F_UAV & "," & UN_OUT_SMI & "," & SFV & "," & SFTI)
                            outputFile.writeLine(BSCNAME & "," & DIP & "," & SECTION & "," & LEVEL_DEG & "," & Direction_IN & "," & "" & "," & T2 & "," & DEG_N_ES & "," & DEG_N_SES & "," & DEG_N_UAS & "," & DEG_N_UAV & "," & DEG_SLIP & "," & DEG_IN_SMI & "," & Direction_OUT & "," & DEG_F_ES & "," & DEG_F_SES & "," & DEG_F_UAS & "," & DEG_F_UAV & "," & DEG_OUT_SMI & "," & SFV & "," & SFTI)
                            LEVEL_UN = ""
                            LEVEL_DEG = ""
                            Direction_IN = ""
                            Direction_OUT = ""
                            DIP = ""
                            SECTION = ""
                            T1 = ""
                            T2 = ""
                            UN_N_ES = ""
                            UN_N_SES = ""
                            UN_N_UAS = ""
                            UN_N_UAV = ""
                            UN_SLIP = ""
                            UN_IN_SMI = ""
                            UN_F_ES = ""
                            UN_F_SES = ""
                            UN_F_UAS = ""
                            UN_F_UAV = ""
                            UN_OUT_SMI = ""
                            DEG_N_ES = ""
                            DEG_N_SES = ""
                            DEG_N_UAS = ""
                            DEG_N_UAV = ""
                            DEG_SLIP = ""
                            DEG_IN_SMI = ""
                            DEG_F_ES = ""
                            DEG_F_SES = ""
                            DEG_F_UAS = ""
                            DEG_F_UAV = ""
                            DEG_OUT_SMI = ""
                            SFV = ""
                            SFTI = ""
                        End If
                        LEVEL_UN = UCase(Trim(templine))
                        Direction_IN = "INCOMING"
                        Direction_OUT = "OUTGOING"

                    ElseIf UCase(Trim(templine)) = "DEGRADED PERFORMANCE LEVEL" Then
                        'DEGRADED PERFORMANCE LEVEL
                        LEVEL_DEG = UCase(Trim(templine))
                        Direction_IN = "INCOMING"
                        Direction_OUT = "OUTGOING"
                    ElseIf UCase(Trim(templine)) = "DIP      SECTION  T1    N-ES   N-SES  N-UAS  N-UAV  SLIP  SMI" Then
                        'DIP      SECTION  T1    N-ES   N-SES  N-UAS  N-UAV  SLIP  SMI
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            DIP = Trim(Mid(templine, 1, 9))
                            SECTION = Trim(Mid(templine, 10, 9))
                            T1 = Trim(Mid(templine, 19, 6))
                            UN_N_ES = Trim(Mid(templine, 25, 7))
                            UN_N_SES = Trim(Mid(templine, 32, 7))
                            UN_N_UAS = Trim(Mid(templine, 39, 7))
                            UN_N_UAV = Trim(Mid(templine, 46, 7))
                            UN_SLIP = Trim(Mid(templine, 53, 6))
                            UN_IN_SMI = Trim(Mid(templine, 59, 6))
                        End If
                    ElseIf UCase(Trim(templine)) = "DIP      SECTION  T1    F-ES   F-SES  F-UAS  F-UAV  SMI" Then
                        'DIP      SECTION  T1    F-ES   F-SES  F-UAS  F-UAV  SMI
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            UN_F_ES = Trim(Mid(templine, 25, 7))
                            UN_F_SES = Trim(Mid(templine, 32, 7))
                            UN_F_UAS = Trim(Mid(templine, 39, 7))
                            UN_F_UAV = Trim(Mid(templine, 46, 7))
                            UN_OUT_SMI = Trim(Mid(templine, 53, 6))
                        End If
                    ElseIf UCase(Trim(templine)) = "DIP      SECTION  T2    N-ES   N-SES  N-UAS  N-UAV  SLIP  SMI" Then
                        'DIP      SECTION  T2    N-ES   N-SES  N-UAS  N-UAV  SLIP  SMI
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            DIP = Trim(Mid(templine, 1, 9))
                            SECTION = Trim(Mid(templine, 10, 9))
                            T2 = Trim(Mid(templine, 19, 6))
                            DEG_N_ES = Trim(Mid(templine, 25, 7))
                            DEG_N_SES = Trim(Mid(templine, 32, 7))
                            DEG_N_UAS = Trim(Mid(templine, 39, 7))
                            DEG_N_UAV = Trim(Mid(templine, 46, 7))
                            DEG_SLIP = Trim(Mid(templine, 53, 6))
                            DEG_IN_SMI = Trim(Mid(templine, 59, 6))
                        End If
                    ElseIf UCase(Trim(templine)) = "DIP      SECTION  T2    F-ES   F-SES  F-UAS  F-UAV  SMI" Then
                        'DIP      SECTION  T2    F-ES   F-SES  F-UAS  F-UAV  SMI
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            DEG_F_ES = Trim(Mid(templine, 25, 7))
                            DEG_F_SES = Trim(Mid(templine, 32, 7))
                            DEG_F_UAS = Trim(Mid(templine, 39, 7))
                            DEG_F_UAV = Trim(Mid(templine, 46, 7))
                            DEG_OUT_SMI = Trim(Mid(templine, 53, 6))
                        End If
                    ElseIf UCase(Trim(templine)) = "DIP      SECTION  SFV   SFTI" Then
                        'DIP      SECTION  SFV   SFTI
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            SFV = Trim(Mid(templine, 19, 6))
                            SFTI = Trim(Mid(templine, 25, 7))
                        End If
                        outputFile.writeLine(BSCNAME & "," & DIP & "," & SECTION & "," & LEVEL_UN & "," & Direction_IN & "," & T1 & "," & "" & "," & UN_N_ES & "," & UN_N_SES & "," & UN_N_UAS & "," & UN_N_UAV & "," & UN_SLIP & "," & UN_IN_SMI & "," & Direction_OUT & "," & UN_F_ES & "," & UN_F_SES & "," & UN_F_UAS & "," & UN_F_UAV & "," & UN_OUT_SMI & "," & SFV & "," & SFTI)
                        outputFile.writeLine(BSCNAME & "," & DIP & "," & SECTION & "," & LEVEL_DEG & "," & Direction_IN & "," & "" & "," & T2 & "," & DEG_N_ES & "," & DEG_N_SES & "," & DEG_N_UAS & "," & DEG_N_UAV & "," & DEG_SLIP & "," & DEG_IN_SMI & "," & Direction_OUT & "," & DEG_F_ES & "," & DEG_F_SES & "," & DEG_F_UAS & "," & DEG_F_UAV & "," & DEG_OUT_SMI & "," & SFV & "," & SFTI)
                        LEVEL_UN = ""
                        LEVEL_DEG = ""
                        Direction_IN = ""
                        Direction_OUT = ""
                        DIP = ""
                        SECTION = ""
                        T1 = ""
                        T2 = ""
                        UN_N_ES = ""
                        UN_N_SES = ""
                        UN_N_UAS = ""
                        UN_N_UAV = ""
                        UN_SLIP = ""
                        UN_IN_SMI = ""
                        UN_F_ES = ""
                        UN_F_SES = ""
                        UN_F_UAS = ""
                        UN_F_UAV = ""
                        UN_OUT_SMI = ""
                        DEG_N_ES = ""
                        DEG_N_SES = ""
                        DEG_N_UAS = ""
                        DEG_N_UAV = ""
                        DEG_SLIP = ""
                        DEG_IN_SMI = ""
                        DEG_F_ES = ""
                        DEG_F_SES = ""
                        DEG_F_UAS = ""
                        DEG_F_UAV = ""
                        DEG_OUT_SMI = ""
                        SFV = ""
                        SFTI = ""
                    ElseIf UCase(Trim(templine)) = "DIP      T1   T2  SLIP   SLIP2  UAS    UASR  UAV1  UASB1  UAV2  UASB2" Then
                        'DIP      T1   T2  SLIP   SLIP2  UAS    UASR  UAV1  UASB1  UAV2  UASB2
                        If DK_DIP <> "" Then
                            DKoutPutFile.writeLine(BSCNAME & "," & DK_DIP & "," & DK_T1 & "," & DK_T2 & "," & DK_SLIP & "," & DK_SLIP2 & "," & DK_UAS & "," & DK_UASR & "," & DK_UAV1 & "," & DK_UASB1 & "," & DK_UAV2 & "," & DK_UASB2 & "," & DK_SECTION & "," & DK_ESV & "," & DK_SESV & "," & DK_DMV & "," & DK_ESVR & "," & DK_SESVR & "," & DK_DMVR & "," & DK_SFV & "," & DK_SFTI & "," & DK_ES2V & "," & DK_SES2V & "," & DK_DM2V & "," & DK_ES2VR & "," & DK_SES2VR & "," & DK_DM2VR & "," & DK_SMI)
                            DK_DIP = ""
                            DK_T1 = ""
                            DK_T2 = ""
                            DK_SLIP = ""
                            DK_SLIP2 = ""
                            DK_UAS = ""
                            DK_UASR = ""
                            DK_UAV1 = ""
                            DK_UASB1 = ""
                            DK_UAV2 = ""
                            DK_UASB2 = ""
                            DK_SECTION = ""
                            DK_ESV = ""
                            DK_SESV = ""
                            DK_DMV = ""
                            DK_ESVR = ""
                            DK_SESVR = ""
                            DK_DMVR = ""
                            DK_SFV = ""
                            DK_SFTI = ""
                            DK_ES2V = ""
                            DK_SES2V = ""
                            DK_DM2V = ""
                            DK_ES2VR = ""
                            DK_SES2VR = ""
                            DK_DM2VR = ""
                            DK_SMI = ""
                        End If
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            DK_DIP = Trim(Mid(templine, 1, 9))
                            DK_T1 = Trim(Mid(templine, 10, 5))
                            DK_T2 = Trim(Mid(templine, 15, 4))
                            DK_SLIP = Trim(Mid(templine, 19, 7))
                            DK_SLIP2 = Trim(Mid(templine, 26, 7))
                            DK_UAS = Trim(Mid(templine, 33, 7))
                            DK_UASR = Trim(Mid(templine, 40, 6))
                            DK_UAV1 = Trim(Mid(templine, 46, 6))
                            DK_UASB1 = Trim(Mid(templine, 52, 7))
                            DK_UAV2 = Trim(Mid(templine, 59, 6))
                            DK_UASB2 = Trim(Mid(templine, 65, 6))
                        End If
                    ElseIf UCase(Trim(templine)) = "SECTION  ESV    SESV   DMV    ESVR   SESVR  DMVR   SFV  SFTI" Then
                        'SECTION  ESV    SESV   DMV    ESVR   SESVR  DMVR   SFV  SFTI
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            DK_SECTION = Trim(Mid(templine, 1, 9))
                            DK_ESV = Trim(Mid(templine, 10, 7))
                            DK_SESV = Trim(Mid(templine, 17, 7))
                            DK_DMV = Trim(Mid(templine, 24, 7))
                            DK_ESVR = Trim(Mid(templine, 31, 7))
                            DK_SESVR = Trim(Mid(templine, 38, 7))
                            DK_DMVR = Trim(Mid(templine, 45, 7))
                            DK_SFV = Trim(Mid(templine, 52, 5))
                            DK_SFTI = Trim(Mid(templine, 57, 6))
                        End If
                    ElseIf UCase(Trim(templine)) = "SECTION  ES2V   SES2V  DM2V   ES2VR  SES2VR DM2VR  SMI" Then
                        'SECTION  ES2V   SES2V  DM2V   ES2VR  SES2VR DM2VR  SMI
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            DK_ES2V = Trim(Mid(templine, 10, 7))
                            DK_SES2V = Trim(Mid(templine, 17, 7))
                            DK_DM2V = Trim(Mid(templine, 24, 7))
                            DK_ES2VR = Trim(Mid(templine, 31, 7))
                            DK_SES2VR = Trim(Mid(templine, 38, 7))
                            DK_DM2VR = Trim(Mid(templine, 45, 7))
                            DK_SMI = Trim(Mid(templine, 52, 5))
                        End If
                        DKoutPutFile.writeLine(BSCNAME & "," & DK_DIP & "," & DK_T1 & "," & DK_T2 & "," & DK_SLIP & "," & DK_SLIP2 & "," & DK_UAS & "," & DK_UASR & "," & DK_UAV1 & "," & DK_UASB1 & "," & DK_UAV2 & "," & DK_UASB2 & "," & DK_SECTION & "," & DK_ESV & "," & DK_SESV & "," & DK_DMV & "," & DK_ESVR & "," & DK_SESVR & "," & DK_DMVR & "," & DK_SFV & "," & DK_SFTI & "," & DK_ES2V & "," & DK_SES2V & "," & DK_DM2V & "," & DK_ES2VR & "," & DK_SES2VR & "," & DK_DM2VR & "," & DK_SMI)
                        DK_DIP = ""
                        DK_T1 = ""
                        DK_T2 = ""
                        DK_SLIP = ""
                        DK_SLIP2 = ""
                        DK_UAS = ""
                        DK_UASR = ""
                        DK_UAV1 = ""
                        DK_UASB1 = ""
                        DK_UAV2 = ""
                        DK_UASB2 = ""
                        DK_SECTION = ""
                        DK_ESV = ""
                        DK_SESV = ""
                        DK_DMV = ""
                        DK_ESVR = ""
                        DK_SESVR = ""
                        DK_DMVR = ""
                        DK_SFV = ""
                        DK_SFTI = ""
                        DK_ES2V = ""
                        DK_SES2V = ""
                        DK_DM2V = ""
                        DK_ES2VR = ""
                        DK_SES2VR = ""
                        DK_DM2VR = ""
                        DK_SMI = ""
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        DKoutPutFile.close()
        DKoutPutFile = Nothing
        outputFile = Nothing
    End Sub
    private sub doEXRPP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'RP    STATE  TYPE     TWIN  STATE   DS     MAINT.STATE

        filepath = outPath & "EXRPP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "RP" & "," & "STATE" & "," & "TYPE" & "," & "TWIN" & "," & "STATE" & "," & "DS" & "," & "MAINT.STATE")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "RP    STATE  TYPE     TWIN  STATE   DS     MAINT.STATE" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 6)) & "," & Trim(Mid(templine, 7, 7)) & "," & Trim(Mid(templine, 14, 9)) & "," & Trim(Mid(templine, 23, 6)) & "," & Trim(Mid(templine, 29, 8)) & "," & Trim(Mid(templine, 37, 7)) & "," & Trim(Mid(templine, 44, 11)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "RP DATA" And Len(Trim(Mid(templine, 1, 6))) <= 5 Then
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 6)) & "," & Trim(Mid(templine, 7, 7)) & "," & Trim(Mid(templine, 14, 9)) & "," & Trim(Mid(templine, 23, 6)) & "," & Trim(Mid(templine, 29, 8)) & "," & Trim(Mid(templine, 37, 7)) & "," & Trim(Mid(templine, 44, 11)))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doEXEMP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'RP    TYPE   EM  EQM                       TWIN  CNTRL  PP     STATE

        filepath = outPath & "EXEMP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "RP" & "," & "TYPE" & "," & "EM" & "," & "EQM" & "," & "TWIN" & "," & "CNTRL" & "," & "PP" & "," & "STATE")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "RP    TYPE   EM  EQM                       TWIN  CNTRL  PP     STATE" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 6)) & "," & Trim(Mid(templine, 7, 7)) & "," & Trim(Mid(templine, 14, 4)) & "," & Trim(Mid(templine, 18, 26)) & "," & Trim(Mid(templine, 44, 6)) _
                                                 & "," & Trim(Mid(templine, 50, 7)) & "," & Trim(Mid(templine, 57, 7)) & "," & Trim(Mid(templine, 64, 11)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "EM DATA" And Len(Trim(Mid(templine, 1, 6))) <= 5 Then
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 6)) & "," & Trim(Mid(templine, 7, 7)) & "," & Trim(Mid(templine, 14, 4)) & "," & Trim(Mid(templine, 18, 26)) & "," & Trim(Mid(templine, 44, 6)) _
                                             & "," & Trim(Mid(templine, 50, 7)) & "," & Trim(Mid(templine, 57, 7)) & "," & Trim(Mid(templine, 64, 11)))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub

    Private Sub doMGBRP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String

        'IMSI             RTY      RR       ACCT
        filepath = outPath & "MGBRP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("MSC" & "," & "IMSI" & "," & "RTY" & "," & "RR" & "," & "ACCT")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "IMSI             RTY      RR       ACCT" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" And Len(Trim(Mid(templine, 1, 17))) <> 0 Then
                            'IMSI             RTY      RR       ACCT
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 17)) & "," & Trim(Mid(templine, 18, 9)) & "," & Trim(Mid(templine, 27, 9)) & "," & Trim(Mid(templine, 36, 4)))
                        End If
                    ElseIf Trim(templine) <> "MT BSC RECORDING INITIATION DATA" And Len(Trim(Mid(templine, 1, 17))) <= 17 And Len(Trim(Mid(templine, 1, 14))) <> 0 Then
                        outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 17)) & "," & Trim(Mid(templine, 18, 9)) & "," & Trim(Mid(templine, 27, 9)) & "," & Trim(Mid(templine, 36, 4)))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doMGLAP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String

        'LAI           PFC  PRL  POOL  AIDX
        filepath = outPath & "MGLAP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("MSC" & "," & "LAI" & "," & "PFC" & "," & "PRL" & "," & "POOL" & "," & "AIDX")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "LAI           PFC  PRL  POOL  AIDX" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" And Len(Trim(Mid(templine, 1, 14))) <> 0 Then
                            'LAI           PFC  PRL  POOL  AIDX
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 14)) & "," & Trim(Mid(templine, 15, 5)) & "," & Trim(Mid(templine, 20, 5)) & "," & Trim(Mid(templine, 25, 6)) & "," & Trim(Mid(templine, 31, 5)))
                        End If
                    ElseIf Trim(templine) <> "MT LOCATION AREA DATA" And Len(Trim(Mid(templine, 1, 14))) <= 12 And Len(Trim(Mid(templine, 1, 14))) <> 0 Then
                        outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 14)) & "," & Trim(Mid(templine, 15, 5)) & "," & Trim(Mid(templine, 20, 5)) & "," & Trim(Mid(templine, 25, 6)) & "," & Trim(Mid(templine, 31, 5)))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private sub doMGCEP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String

        'CELL     CGI                 BSC      CO     RO     NCS  EA
        filepath = outPath & "MGCEP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("MSC" & "," & "CELL" & "," & "CGI" & "," & "BSC" & "," & "CO" & "," & "RO" & "," & "NCS" & "," & "EA")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Strings.left(Trim(templine), 12) = "CELL     CGI" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            'CELL     CGI                 BSC      CO     RO     NCS  EA
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 20)) & "," & Trim(Mid(templine, 30, 9)) & "," & Trim(Mid(templine, 39, 7)) & "," & Trim(Mid(templine, 46, 7)) & "," & Trim(Mid(templine, 53, 5)) & "," & Trim(Mid(templine, 58, 9)))
                        End If
                    ElseIf Trim(templine) <> "MT CELL DATA" And Len(Trim(Mid(templine, 1, 9))) <= 7 Then
                        outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 20)) & "," & Trim(Mid(templine, 30, 9)) & "," & Trim(Mid(templine, 39, 7)) & "," & Trim(Mid(templine, 46, 7)) & "," & Trim(Mid(templine, 53, 5)) & "," & Trim(Mid(templine, 58, 9)))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    private sub doMGEPP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String

        'PROP                   TYPE
        filepath = outPath & "MGEPP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("MSC" & "," & "PROP" & "," & "TYPE")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "PROP                   TYPE" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            'PROP                   TYPE
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 23)) & "," & Trim(Mid(templine, 24, 13)))
                        End If
                    ElseIf Trim(templine) <> "MT EXCHANGE PROPERTY DATA" Then
                        outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 23)) & "," & Trim(Mid(templine, 24, 13)))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    private sub doMGIDP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String

        'BTDM        GTDM

        filepath = outPath & "MGIDP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("MSC" & "," & "BTDM" & "," & "GTDM")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "BTDM        GTDM" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            'BTDM        GTDM
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 12)) & "," & Trim(Mid(templine, 13, 9)))
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    private sub doMGOCP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String

        'CELL     AREAID              MSC      NCS

        filepath = outPath & "MGOCP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("MSC" & "," & "CELL" & "," & "AREAID" & "," & "MSC" & "," & "NCS")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Strings.left(Trim(templine), 12) = "CELL     ARE" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            'CELL     AREAID              MSC      NCS
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 20)) & "," & Trim(Mid(templine, 30, 9)) & "," & Trim(Mid(templine, 39, 7)))
                        End If
                    ElseIf Trim(templine) <> "MT OUTER CELL DATA" And Len(Trim(Mid(templine, 1, 9))) <= 7 Then
                        outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 20)) & "," & Trim(Mid(templine, 30, 9)) & "," & Trim(Mid(templine, 39, 7)))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doPCORP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim BLOCK As String, SUID As String, CA As String, CAF As String
        Dim CI As String, S As String, TYPE As String, POSITION As String, SIZE As String

        BLOCK = ""
        SUID = ""
        CA = ""
        CAF = ""

        'BLOCK    SUID                               CA    CAF
        'CI               S  TYPE  POSITION         SIZE
        filepath = outPath & "PCORP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "BLOCK" & "," & "SUID" & "," & "CA" & "," & "CAF" _
                                 & "," & "CI" & "," & "S" & "," & "TYPE" & "," & "POSITION" & "," & "SIZE")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "BLOCK    SUID                               CA    CAF" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            'BLOCK    SUID                               CA    CAF
                            BLOCK = Trim(Mid(templine, 1, 9))
                            SUID = Trim(Mid(templine, 10, 35))
                            CA = Trim(Mid(templine, 45, 6))
                            CAF = Trim(Mid(templine, 51, 6))
                        End If
                    ElseIf Trim(templine) = "CI               S  TYPE  POSITION         SIZE" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            'CI               S  TYPE  POSITION         SIZE
                            CI = Trim(Mid(templine, 1, 17))
                            S = Trim(Mid(templine, 18, 3))
                            TYPE = Trim(Mid(templine, 21, 6))
                            POSITION = Trim(Mid(templine, 27, 17))
                            SIZE = Trim(Mid(templine, 44, 5))
                            outputFile.writeline(BSCNAME & "," & BLOCK & "," & SUID & "," & CA & "," & CAF _
                                                 & "," & CI & "," & S & "," & TYPE & "," & POSITION & "," & SIZE)
                        End If
                    ElseIf Trim(templine) <> "CI               S  TYPE  POSITION         SIZE" And Trim(templine) <> "BLOCK    SUID                               CA    CAF" And Trim(templine) <> "PROGRAM CORRECTIONS" Then
                        If Trim(templine) <> "NONE" Then
                            'CI               S  TYPE  POSITION         SIZE
                            CI = Trim(Mid(templine, 1, 17))
                            S = Trim(Mid(templine, 18, 3))
                            TYPE = Trim(Mid(templine, 21, 6))
                            POSITION = Trim(Mid(templine, 27, 17))
                            SIZE = Trim(Mid(templine, 44, 5))
                            outputFile.writeline(BSCNAME & "," & BLOCK & "," & SUID & "," & CA & "," & CAF _
                                                 & "," & CI & "," & S & "," & TYPE & "," & POSITION & "," & SIZE)
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    private sub doPLLDP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim INT As String, PLOAD As String, CALIM As String, OFFDO As String, OFFDI As String, FTCHDO As String, FTCHDI As String, OFFMPH As String, OFFMPL As String, FTCHMPH As String, FTCHMPL As String
        Dim OFFTCAP As String, FTDTCAP As String
        'INT PLOAD CALIM OFFDO OFFDI FTCHDO FTCHDI OFFMPH OFFMPL FTCHMPH FTCHMPL
        'INT OFFTCAP FTDTCAP
        INT = ""
        PLOAD = ""
        CALIM = ""
        OFFDO = ""
        OFFDI = ""
        FTCHDO = ""
        FTCHDI = ""
        OFFMPH = ""
        OFFMPL = ""
        FTCHMPH = ""
        FTCHMPL = ""
        OFFTCAP = ""
        FTDTCAP = ""

        filepath = outPath & "PLLDP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "INT" & "," & "PLOAD" & "," & "CALIM" & "," & "OFFDO" & "," & "OFFDI" & "," & "FTCHDO" & "," & "FTCHDI" & "," & "OFFMPH" & "," & "OFFMPL" & "," & "FTCHMPH" & "," & "FTCHMPL")

        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "INT PLOAD CALIM OFFDO OFFDI FTCHDO FTCHDI OFFMPH OFFMPL FTCHMPH FTCHMPL" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            'INT PLOAD CALIM OFFDO OFFDI FTCHDO FTCHDI OFFMPH OFFMPL FTCHMPH FTCHMPL
                            INT = Trim(Mid(templine, 1, 4))
                            PLOAD = Trim(Mid(templine, 5, 6))
                            CALIM = Trim(Mid(templine, 11, 6))
                            OFFDO = Trim(Mid(templine, 17, 6))
                            OFFDI = Trim(Mid(templine, 23, 6))
                            FTCHDO = Trim(Mid(templine, 29, 7))
                            FTCHDI = Trim(Mid(templine, 36, 7))
                            OFFMPH = Trim(Mid(templine, 43, 7))
                            OFFMPL = Trim(Mid(templine, 50, 7))
                            FTCHMPH = Trim(Mid(templine, 57, 8))
                            FTCHMPL = Trim(Mid(templine, 65, 8))
                            outputFile.writeline(BSCNAME & "," & INT & "," & PLOAD & "," & CALIM & "," & OFFDO & "," & OFFDI & "," & FTCHDO & "," & FTCHDI & "," & OFFMPH & "," & OFFMPL & "," & FTCHMPH & "," & FTCHMPL)
                        End If
                        'ElseIf Trim(templine) = "INT OFFTCAP FTDTCAP" Then
                        '    'INT OFFTCAP FTDTCAP
                        '    If Trim(templine) <> "NONE" Then
                        '        OFFTCAP = Trim(Mid(templine, 5, 8))
                        '        FTDTCAP = Trim(Mid(templine, 13, 8))
                        '        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 23)) & "," & Trim(Mid(templine, 24, 12)))
                        '    End If
                    ElseIf Trim(templine) <> "INT OFFTCAP FTDTCAP" And Trim(templine) <> "PROCESSOR LOAD DATA" And Len(Trim(templine)) >= 60 Then
                        INT = Trim(Mid(templine, 1, 4))
                        PLOAD = Trim(Mid(templine, 5, 6))
                        CALIM = Trim(Mid(templine, 11, 6))
                        OFFDO = Trim(Mid(templine, 17, 6))
                        OFFDI = Trim(Mid(templine, 23, 6))
                        FTCHDO = Trim(Mid(templine, 29, 7))
                        FTCHDI = Trim(Mid(templine, 36, 7))
                        OFFMPH = Trim(Mid(templine, 43, 7))
                        OFFMPL = Trim(Mid(templine, 50, 7))
                        FTCHMPH = Trim(Mid(templine, 57, 8))
                        FTCHMPL = Trim(Mid(templine, 65, 8))
                        outputFile.writeline(BSCNAME & "," & INT & "," & PLOAD & "," & CALIM & "," & OFFDO & "," & OFFDI & "," & FTCHDO & "," & FTCHDI & "," & OFFMPH & "," & OFFMPL & "," & FTCHMPH & "," & FTCHMPL)
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRACLP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String

        'CLNAME           CLAL      CLSUP  CLEX  CLUSE  CLTHR  CLMAX

        'G14
        'CLNAME           CLAL      CLSUP  CLEX  CLUSE  CLTHR  CLMAX  CLABSUSE

        filepath = outPath & "RACLP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CLNAME" & "," & "CLAL" & "," & "CLSUP" & "," & "CLEX" & "," & "CLUSE" & "," & "CLTHR" & "," & "CLMAX" & "," & "CLABSUSE")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Mid(UCase(Trim(templine)), 1, 35) = "CLNAME           CLAL      CLSUP  C" Then
                        'CLNAME           CLAL      CLSUP  CLEX  CLUSE  CLTHR  CLMAX
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 17)) & "," & Trim(Mid(templine, 18, 10)) & "," & Trim(Mid(templine, 28, 7)) & "," & Trim(Mid(templine, 35, 6)) & "," & Trim(Mid(templine, 41, 7)) & "," & Trim(Mid(templine, 48, 7)) & "," & Trim(Mid(templine, 55, 7)) & "," & Trim(Mid(templine, 62, 10)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "BSC CAPACITY LOCK ALARM DATA" And Mid(UCase(Trim(templine)), 1, 35) <> "CLNAME           CLAL      CLSUP  C" Then
                        outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 17)) & "," & Trim(Mid(templine, 18, 10)) & "," & Trim(Mid(templine, 28, 7)) & "," & Trim(Mid(templine, 35, 6)) & "," & Trim(Mid(templine, 41, 7)) & "," & Trim(Mid(templine, 48, 7)) & "," & Trim(Mid(templine, 55, 7)) & "," & Trim(Mid(templine, 62, 10)))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    private sub doRAEPP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String

        'PROP                   TYPE
        filepath = outPath & "RAEPP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "PROP" & "," & "TYPE")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "PROP                   TYPE" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 23)) & "," & Trim(Mid(templine, 24, 12)))
                        End If
                        'PROP                   TYPE
                    ElseIf Trim(templine) <> "BSC EXCHANGE PROPERTY DATA" And Trim(Mid(templine, 22, 1)) = "" Then
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 23)) & "," & Trim(Mid(templine, 24, 12)))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRAHHP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim HHP_DATE As String, HHP_TIME As String, AVECAP As String
        Dim SNT As String, CAP As String, HIGHCAP As String, LOWCAP As String
        HHP_DATE = ""
        HHP_TIME = ""
        AVECAP = ""
        SNT = ""
        CAP = ""
        HIGHCAP = ""
        LOWCAP = ""
        'DATE    TIME  AVECAP
        'SNT          CAP  HIGHCAP  LOWCAP
        filepath = outPath & "RAHHP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "DATE" & "," & "TIME" & "," & "AVECAP" & "," & "SNT" & "," & "CAP" & "," & "HIGHCAP" & "," & "LOWCAP")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "DATE    TIME  AVECAP" Then
                        'DATE    TIME  AVECAP
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            HHP_DATE = Trim(Mid(templine, 1, 8))
                            HHP_TIME = Trim(Mid(templine, 9, 6))
                            AVECAP = Trim(Mid(templine, 15, 6))
                        End If
                    ElseIf Trim(templine) = "SNT          CAP  HIGHCAP  LOWCAP" Then
                        'SNT          CAP  HIGHCAP  LOWCAP
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            SNT = Trim(Mid(templine, 1, 13))
                            CAP = Trim(Mid(templine, 14, 5))
                            HIGHCAP = Trim(Mid(templine, 19, 9))
                            LOWCAP = Trim(Mid(templine, 28, 7))
                            outputFile.writeLINE(BSCNAME & "," & HHP_DATE & "," & HHP_TIME & "," & AVECAP & "," & SNT & "," & CAP & "," & HIGHCAP & "," & LOWCAP)
                            SNT = ""
                            CAP = ""
                            HIGHCAP = ""
                            LOWCAP = ""
                        End If
                    ElseIf UCase(Trim(templine)) <> "RADIO CONTROL ADMINISTRATION" And UCase(Trim(templine)) <> "TRH LOAD TABLE" And UCase(Trim(templine)) <> "DATE    TIME  AVECAP" And UCase(Trim(templine)) <> "SNT          CAP  HIGHCAP  LOWCAP" Then
                        SNT = Trim(Mid(templine, 1, 13))
                        CAP = Trim(Mid(templine, 14, 5))
                        HIGHCAP = Trim(Mid(templine, 19, 9))
                        LOWCAP = Trim(Mid(templine, 28, 7))
                        outputFile.writeLINE(BSCNAME & "," & HHP_DATE & "," & HHP_TIME & "," & AVECAP & "," & SNT & "," & CAP & "," & HIGHCAP & "," & LOWCAP)
                        SNT = ""
                        CAP = ""
                        HIGHCAP = ""
                        LOWCAP = ""
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRASAP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'Dim n_dev As String, n_state As String
        'Dim c_dev As String, c_state As String
        'Dim s_dev As String, s_state As String
        Dim type As String
        Dim DEV As String, STATE As String
        'n_dev = ""
        'n_state = ""
        'c_dev = ""
        'c_state = ""
        's_dev = ""
        's_state = ""
        type = ""
        DEV = ""
        STATE = ""
        'NEVER USED DEVICES
        'CONTINUOUSLY BUSY DEVICES
        'SUSPECTED FAULTY DEVICES

        filepath = outPath & "RASAP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "TYPE" & "," & "DEV" & "," & "STATE")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "NEVER USED DEVICES" Then
                        type = "NEVER USED DEVICES"
                    ElseIf UCase(Trim(templine)) = "CONTINUOUSLY BUSY DEVICES" Then
                        type = "CONTINUOUSLY BUSY DEVICES"
                    ElseIf UCase(Trim(templine)) = "SUSPECTED FAULTY DEVICES" Then
                        type = "SUSPECTED FAULTY DEVICES"
                    ElseIf UCase(Trim(templine)) = "SUSPECTED FAULTY DEVICES" Then
                        type = "SUSPECTED FAULTY DEVICES"
                    ElseIf UCase(Trim(templine)) = "DEV         STATE" Then
                        templine = getNowLine()
                        DEV = Trim(Mid(templine, 1, 12))
                        STATE = Trim(Mid(templine, 13, 7))
                        outputFile.writeLINE(BSCNAME & "," & type & "," & DEV & "," & STATE)
                    ElseIf UCase(Trim(templine)) <> "SEIZURE SUPERVISION OF DEVICES IN BSC ALARMED DEVICES" And Len(Trim(Mid(templine, 1, 12))) > 5 And UCase(Trim(templine)) <> "DEV         STATE" Then
                        DEV = Trim(Mid(templine, 1, 12))
                        STATE = Trim(Mid(templine, 13, 7))
                        outputFile.writeLINE(BSCNAME & "," & type & "," & DEV & "," & STATE)
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    private sub doRFCAP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String

        'CELL      AFRCSCC  AHRCSCC  AWBCSCC  EFRCSCC  HRCSCC  VCSCC

        filepath = outPath & "RFCAP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "AFRCSCC" & "," & "AHRCSCC" & "," & "AWBCSCC" & "," & "EFRCSCC" & "," & "HRCSCC" & "," & "VCSCC")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Mid(UCase(Trim(templine)), 1, 35) = "CELL      AFRCSCC  AHRCSCC  AWBCSCC" Then
                        'CELL      AFRCSCC  AHRCSCC  AWBCSCC  EFRCSCC  HRCSCC              (G10B)
                        'CELL      AFRCSCC  AHRCSCC  AWBCSCC  EFRCSCC  HRCSCC  VCSCC       (G12A)
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 10)) & "," & Trim(Mid(templine, 11, 9)) & "," & Trim(Mid(templine, 20, 9)) & "," & Trim(Mid(templine, 29, 9)) & "," & Trim(Mid(templine, 38, 9)) & "," & Trim(Mid(templine, 47, 8)) & "," & Trim(Mid(templine, 55, 8)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL ALLOCATION CAPACITY LOCK DATA" And Len(Trim(Mid(templine, 1, 10))) <= 8 And Mid(UCase(Trim(templine)), 1, 35) <> "CELL      AFRCSCC  AHRCSCC  AWBCSCC" Then
                        outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 10)) & "," & Trim(Mid(templine, 11, 9)) & "," & Trim(Mid(templine, 20, 9)) & "," & Trim(Mid(templine, 29, 9)) & "," & Trim(Mid(templine, 38, 9)) & "," & Trim(Mid(templine, 47, 8)) & "," & Trim(Mid(templine, 55, 8)))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRFPCP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, DLPCG As String, DLPCE As String, DLPCE2A As String, INITDLPCG As String, INITDLPCE As String, INITDLPCE2A As String, DLPCBCCH As String
        'CELL     DLPCG     DLPCE     DLPCE2A   INITDLPCG  INITDLPCE  INITDLPCE2A
        CELL = ""
        DLPCG = ""
        DLPCE = ""
        DLPCE2A = ""
        INITDLPCG = ""
        INITDLPCE = ""
        INITDLPCE2A = ""
        DLPCBCCH = ""

        filepath = outPath & "RFPCP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "DLPCG" & "," & "DLPCE" & "," & "DLPCE2A" & "," & "INITDLPCG" & "," & "INITDLPCE" & "," & "INITDLPCE2A" & "," & "DLPCBCCH")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If CELL <> "" Then
                        outputFile.writeLINE(BSCNAME & "," & CELL & "," & DLPCG & "," & DLPCE & "," & DLPCE2A & "," & INITDLPCG & "," & INITDLPCE & "," & INITDLPCE2A & "," & DLPCBCCH)
                        DLPCBCCH = ""
                        CELL = ""
                    End If
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     DLPCG     DLPCE     DLPCE2A   INITDLPCG  INITDLPCE  INITDLPCE2A" Then
                        If CELL <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & CELL & "," & DLPCG & "," & DLPCE & "," & DLPCE2A & "," & INITDLPCG & "," & INITDLPCE & "," & INITDLPCE2A & "," & DLPCBCCH)
                            CELL = ""
                            DLPCBCCH = ""
                        End If
                        templine = getNowLine()
                        CELL = Trim(Mid(templine, 1, 9))
                        DLPCG = Trim(Mid(templine, 10, 10))
                        DLPCE = Trim(Mid(templine, 20, 10))
                        DLPCE2A = Trim(Mid(templine, 30, 10))
                        INITDLPCG = Trim(Mid(templine, 40, 11))
                        INITDLPCE = Trim(Mid(templine, 51, 11))
                        INITDLPCE2A = Trim(Mid(templine, 62, 11))
                    ElseIf UCase(Trim(templine)) = "DLPCBCCH" Then
                        outputFile.writeLINE(BSCNAME & "," & CELL & "," & DLPCG & "," & DLPCE & "," & DLPCE2A & "," & INITDLPCG & "," & INITDLPCE & "," & INITDLPCE2A & "," & DLPCBCCH)
                        DLPCBCCH = ""
                        CELL = ""
                    ElseIf UCase(Trim(templine)) <> "PACKET SWITCHED DOWNLINK POWER CONTROL DATA" And Len(Trim(Mid(templine, 1, 9))) < 9 And UCase(Trim(templine)) <> "CELL     DLPCG     DLPCE     DLPCE2A   INITDLPCG  INITDLPCE  INITDLPCE2A" And UCase(Trim(templine)) <> "DLPCBCCH" Then
                        If CELL <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & CELL & "," & DLPCG & "," & DLPCE & "," & DLPCE2A & "," & INITDLPCG & "," & INITDLPCE & "," & INITDLPCE2A & "," & DLPCBCCH)
                            DLPCBCCH = ""
                            CELL = ""
                        End If
                        CELL = Trim(Mid(templine, 1, 9))
                        DLPCG = Trim(Mid(templine, 10, 10))
                        DLPCE = Trim(Mid(templine, 20, 10))
                        DLPCE2A = Trim(Mid(templine, 30, 10))
                        INITDLPCG = Trim(Mid(templine, 40, 11))
                        INITDLPCE = Trim(Mid(templine, 51, 11))
                        INITDLPCE2A = Trim(Mid(templine, 62, 11))
                        outputFile.writeLINE(BSCNAME & "," & CELL & "," & DLPCG & "," & DLPCE & "," & DLPCE2A & "," & INITDLPCG & "," & INITDLPCE & "," & INITDLPCE2A)
                        CELL = ""
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub

    Private Sub doRFVCP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, VAMOSCELLSTATE As String, AQPSKONBCCH As String, ASSV As String, QBAHRV As String, QBAWBV As String, QBNAV As String
        Dim QDESDLOVN As String, QDESDLOVP As String, QDESULOVN As String, QDESULOVP As String, SSTHRASSV As String, SSTHRV As String
        Dim TSCSET1PAIRS As String, FULLAQPSK As String, VHOSUCCESS As String, CLTHV As String, THRAV As String, THRVP As String, THRDV As String
        Dim QBAFRV As String

        'CELL     VAMOSCELLSTATE  AQPSKONBCCH  ASSV  QBAHRV  QBAWBV  QBNAV
        'QDESDLOVN  QDESDLOVP  QDESULOVN  QDESULOVP  SSTHRASSV  SSTHRV
        'TSCSET1PAIRS  FULLAQPSK  VHOSUCCESS  CLTHV  THRAV  THRVP  THRDV
        'QBAFRV

        CELL = ""
        VAMOSCELLSTATE = ""
        AQPSKONBCCH = ""
        ASSV = ""
        QBAHRV = ""
        QBAWBV = ""
        QBNAV = ""
        QDESDLOVN = ""
        QDESDLOVP = ""
        QDESULOVN = ""
        QDESULOVP = ""
        SSTHRASSV = ""
        SSTHRV = ""
        TSCSET1PAIRS = ""
        FULLAQPSK = ""
        VHOSUCCESS = ""
        CLTHV = ""
        THRAV = ""
        THRVP = ""
        THRDV = ""
        QBAFRV = ""

        filepath = outPath & "RFVCP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "VAMOSCELLSTATE" & "," & "AQPSKONBCCH" & "," & "ASSV" & "," & "QBAHRV" & "," & "QBAWBV" & "," & "")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     DLPCG     DLPCE     DLPCE2A   INITDLPCG  INITDLPCE  INITDLPCE2A" Then
                        templine = getNowLine()
                        CELL = Trim(Mid(templine, 1, 9))
                        'DLPCG = Trim(Mid(templine, 10, 10))
                        'DLPCE = Trim(Mid(templine, 20, 10))
                        'DLPCE2A = Trim(Mid(templine, 30, 10))
                        'INITDLPCG = Trim(Mid(templine, 40, 11))
                        'INITDLPCE = Trim(Mid(templine, 51, 11))
                        'INITDLPCE2A = Trim(Mid(templine, 62, 11))
                        'outputFile.writeLINE(BSCNAME & "," & CELL & "," & DLPCG & "," & DLPCE & "," & DLPCE2A & "," & INITDLPCG & "," & INITDLPCE & "," & INITDLPCE2A)
                    ElseIf UCase(Trim(templine)) <> "PACKET SWITCHED DOWNLINK POWER CONTROL DATA" And Len(Trim(Mid(templine, 1, 9))) < 9 And UCase(Trim(templine)) <> "CELL     DLPCG     DLPCE     DLPCE2A   INITDLPCG  INITDLPCE  INITDLPCE2A" Then
                        CELL = Trim(Mid(templine, 1, 9))

                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub

    Private sub doRLACP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'CELL     ACSTATE  SLEVEL  STIME    EXCHGR
        filepath = outPath & "RLACP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "ACSTATE" & "," & "SLEVEL" & "," & "STIME" & "," & "EXCHGR")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "CELL     ACSTATE  SLEVEL  STIME    EXCHGR" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            'CELL     ACSTATE  SLEVEL  STIME    EXCHGR
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 9)) & "," & Trim(Mid(templine, 19, 8)) & "," & Trim(Mid(templine, 27, 9)) & "," & Trim(Mid(templine, 36, 9)))
                        End If
                    ElseIf Trim(templine) <> "CELL ADAPTIVE LOGICAL CHANNEL CONFIGURATION" And Len(Trim(Mid(templine, 10, 9))) <= 3 Then
                        outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 9)) & "," & Trim(Mid(templine, 19, 8)) & "," & Trim(Mid(templine, 27, 9)) & "," & Trim(Mid(templine, 36, 9)))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    private sub doRLADP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim ADP_SET As String, MODE As String, THR As String, HYST As String, ICMODE As String
        'SET  MODE  THR  HYST  ICMODE
        ADP_SET = ""
        MODE = ""
        THR = ""
        HYST = ""
        ICMODE = ""

        filepath = outPath & "RLADP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "SET" & "," & "MODE" & "," & "THR" & "," & "HYST" & "," & "ICMODE")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "SET  MODE  THR  HYST  ICMODE" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            'SET  MODE  THR  HYST  ICMODE
                            ADP_SET = Trim(Mid(templine, 1, 5))
                            MODE = Trim(Mid(templine, 6, 6))
                            THR = Trim(Mid(templine, 12, 5))
                            HYST = Trim(Mid(templine, 17, 6))
                            ICMODE = Trim(Mid(templine, 23, 6))
                            outputFile.writeline(BSCNAME & "," & ADP_SET & "," & MODE & "," & THR & "," & HYST & "," & ICMODE)
                        End If
                    ElseIf Trim(templine) <> "AMR CODEC SET DATA" Then
                        If Trim(Mid(templine, 1, 5)) <> "" Then
                            ADP_SET = Trim(Mid(templine, 1, 5))
                        End If
                        MODE = Trim(Mid(templine, 6, 6))
                        THR = Trim(Mid(templine, 12, 5))
                        HYST = Trim(Mid(templine, 17, 6))
                        ICMODE = Trim(Mid(templine, 23, 6))
                        outputFile.writeline(BSCNAME & "," & ADP_SET & "," & MODE & "," & THR & "," & HYST & "," & ICMODE)
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    private sub doRLAPP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, AMRPCSTATE As String
        Dim SCTYPE As String, SSDESDLAFR As String, SSDESULAFR As String, QDESDLAFR As String, QDESULAFR As String
        Dim SSDESDLAHR As String, SSDESULAHR As String, QDESDLAHR As String, QDESULAHR As String
        Dim SSDESDLAWB As String, SSDESULAWB As String, QDESDLAWB As String, QDESULAWB As String
        'CELL     AMRPCSTATE
        '         SCTYPE  SSDESDLAFR  SSDESULAFR  QDESDLAFR  QDESULAFR
        '                 SSDESDLAHR  SSDESULAHR  QDESDLAHR  QDESULAHR
        '                 SSDESDLAWB  SSDESULAWB  QDESDLAWB  QDESULAWB
        CELL = ""
        AMRPCSTATE = ""
        SCTYPE = ""
        SSDESDLAFR = ""
        SSDESULAFR = ""
        QDESDLAFR = ""
        QDESULAFR = ""
        SSDESDLAHR = ""
        SSDESULAHR = ""
        QDESDLAHR = ""
        QDESULAHR = ""
        SSDESDLAWB = ""
        SSDESULAWB = ""
        QDESDLAWB = ""
        QDESULAWB = ""


        filepath = outPath & "RLAPP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "AMRPCSTATE" _
                                 & "," & "SCTYPE" & "," & "SSDESDLAFR" & "," & "SSDESULAFR" & "," & "QDESDLAFR" & "," & "QDESULAFR" _
                                 & "," & "SSDESDLAHR" & "," & "SSDESULAHR" & "," & "QDESDLAHR" & "," & "QDESULAHR" _
                                 & "," & "SSDESDLAWB" & "," & "SSDESULAWB" & "," & "QDESDLAWB" & "," & "QDESULAWB")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & AMRPCSTATE _
                                                  & "," & SCTYPE & "," & SSDESDLAFR & "," & SSDESULAFR & "," & QDESDLAFR & "," & QDESULAFR _
                                                  & "," & SSDESDLAHR & "," & SSDESULAHR & "," & QDESDLAHR & "," & QDESULAHR _
                                                  & "," & SSDESDLAWB & "," & SSDESULAWB & "," & QDESDLAWB & "," & QDESULAWB)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & AMRPCSTATE _
                                                  & "," & SCTYPE & "," & SSDESDLAFR & "," & SSDESULAFR & "," & QDESDLAFR & "," & QDESULAFR _
                                                  & "," & SSDESDLAHR & "," & SSDESULAHR & "," & QDESDLAHR & "," & QDESULAHR _
                                                  & "," & SSDESDLAWB & "," & SSDESULAWB & "," & QDESDLAWB & "," & QDESULAWB)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     AMRPCSTATE" Then
                        'CELL     AMRPCSTATE
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            If CELL <> "" Then
                                outputFile.writeline(BSCNAME & "," & CELL & "," & AMRPCSTATE _
                                                     & "," & SCTYPE & "," & SSDESDLAFR & "," & SSDESULAFR & "," & QDESDLAFR & "," & QDESULAFR _
                                                     & "," & SSDESDLAHR & "," & SSDESULAHR & "," & QDESDLAHR & "," & QDESULAHR _
                                                     & "," & SSDESDLAWB & "," & SSDESULAWB & "," & QDESDLAWB & "," & QDESULAWB)
                                CELL = ""
                                AMRPCSTATE = ""
                                SCTYPE = ""
                                SSDESDLAFR = ""
                                SSDESULAFR = ""
                                QDESDLAFR = ""
                                QDESULAFR = ""
                                SSDESDLAHR = ""
                                SSDESULAHR = ""
                                QDESDLAHR = ""
                                QDESULAHR = ""
                                SSDESDLAWB = ""
                                SSDESULAWB = ""
                                QDESDLAWB = ""
                                QDESULAWB = ""
                            End If
                            CELL = Trim(Mid(templine, 1, 9))
                            AMRPCSTATE = Trim(Mid(templine, 10, 10))
                        End If
                    ElseIf UCase(templine) = "         SCTYPE  SSDESDLAFR  SSDESULAFR  QDESDLAFR  QDESULAFR" Then
                        '         SCTYPE  SSDESDLAFR  SSDESULAFR  QDESDLAFR  QDESULAFR
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            SCTYPE = Trim(Mid(templine, 10, 8))
                            SSDESDLAFR = Trim(Mid(templine, 18, 12))
                            SSDESULAFR = Trim(Mid(templine, 30, 12))
                            QDESDLAFR = Trim(Mid(templine, 42, 11))
                            QDESULAFR = Trim(Mid(templine, 53, 11))
                        End If
                    ElseIf UCase(templine) = "                 SSDESDLAHR  SSDESULAHR  QDESDLAHR  QDESULAHR" Then
                        '                 SSDESDLAHR  SSDESULAHR  QDESDLAHR  QDESULAHR
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            SSDESDLAHR = Trim(Mid(templine, 18, 12))
                            SSDESULAHR = Trim(Mid(templine, 30, 12))
                            QDESDLAHR = Trim(Mid(templine, 42, 11))
                            QDESULAHR = Trim(Mid(templine, 53, 11))
                        End If
                    ElseIf UCase(templine) = "                 SSDESDLAWB  SSDESULAWB  QDESDLAWB  QDESULAWB" Then
                        '                 SSDESDLAWB  SSDESULAWB  QDESDLAWB  QDESULAWB
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            SSDESDLAWB = Trim(Mid(templine, 18, 12))
                            SSDESULAWB = Trim(Mid(templine, 30, 12))
                            QDESDLAWB = Trim(Mid(templine, 42, 11))
                            QDESULAWB = Trim(Mid(templine, 53, 11))
                        End If
                        outputFile.writeline(BSCNAME & "," & CELL & "," & AMRPCSTATE _
                                                    & "," & SCTYPE & "," & SSDESDLAFR & "," & SSDESULAFR & "," & QDESDLAFR & "," & QDESULAFR _
                                                    & "," & SSDESDLAHR & "," & SSDESULAHR & "," & QDESDLAHR & "," & QDESULAHR _
                                                    & "," & SSDESDLAWB & "," & SSDESULAWB & "," & QDESDLAWB & "," & QDESULAWB)
                        If SCTYPE <> "UL" Then
                            CELL = ""
                            AMRPCSTATE = ""
                        End If
                        SCTYPE = ""
                        SSDESDLAFR = ""
                        SSDESULAFR = ""
                        QDESDLAFR = ""
                        QDESULAFR = ""
                        SSDESDLAHR = ""
                        SSDESULAHR = ""
                        QDESDLAHR = ""
                        QDESULAHR = ""
                        SSDESDLAWB = ""
                        SSDESULAWB = ""
                        QDESDLAWB = ""
                        QDESULAWB = ""
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLBAP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim REGINTDL As String, SSLENDL As String, QLENDL As String
        REGINTDL = ""
        SSLENDL = ""
        QLENDL = ""

        'REGINTDL  SSLENDL  QLENDL

        filepath = outPath & "RLBAP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "REGINTDL" & "," & "SSLENDL" & "," & "QLENDL")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "REGINTDL  SSLENDL  QLENDL" Then
                        templine = getNowLine()
                        REGINTDL = Trim(Mid(templine, 1, 10))
                        SSLENDL = Trim(Mid(templine, 11, 9))
                        QLENDL = Trim(Mid(templine, 20, 6))
                        outputFile.writeLINE(BSCNAME & "," & REGINTDL & "," & SSLENDL & "," & QLENDL)
                    ElseIf UCase(Trim(templine)) <> "DYNAMIC BTS POWER CONTROL ADDITIONAL BSC DATA" And Len(Trim(Mid(templine, 1, 10))) < 10 And UCase(Trim(templine)) <> "REGINTDL  SSLENDL  QLENDL" Then
                        REGINTDL = Trim(Mid(templine, 1, 10))
                        SSLENDL = Trim(Mid(templine, 11, 9))
                        QLENDL = Trim(Mid(templine, 20, 6))
                        outputFile.writeLINE(BSCNAME & "," & REGINTDL & "," & SSLENDL & "," & QLENDL)
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    private sub doRLBCP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim Cell As String, DBPSTATE As String
        Dim SCTYPE As String, SSDESDL As String, QDESDL As String, LCOMPDL As String, QCOMPDL As String, DBPBCCHSTATE As String
        'CELL     DBPSTATE  DBPBCCHSTATE(g14)
        'SCTYPE   SSDESDL  QDESDL  LCOMPDL  QCOMPDL
        '         BSPWRMINP  BSPWRMINN
        Cell = ""
        DBPSTATE = ""
        SCTYPE = ""
        SSDESDL = ""
        QDESDL = ""
        LCOMPDL = ""
        QCOMPDL = ""
        DBPBCCHSTATE = ""
        filepath = outPath & "RLBCP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "DBPSTATE" & "," & "SCTYPE" & "," & "SSDESDL" & "," & "QDESDL" & "," & "LCOMPDL" & "," & "QCOMPDL" & "," & "BSPWRMINP" & "," & "BSPWRMINN" & "," & "DBPBCCHSTATE")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "CELL     DBPSTATE" Or Trim(templine) = "CELL     DBPSTATE  DBPBCCHSTATE" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            Cell = Trim(Mid(templine, 1, 9))
                            DBPSTATE = Trim(Mid(templine, 10, 10))
                            DBPBCCHSTATE = Trim(Mid(templine, 20, 10))
                        End If
                    ElseIf Trim(templine) = "SCTYPE   SSDESDL  QDESDL  LCOMPDL  QCOMPDL" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            'SCTYPE   SSDESDL  QDESDL  LCOMPDL  QCOMPDL
                            SCTYPE = Trim(Mid(templine, 1, 9))
                            SSDESDL = Trim(Mid(templine, 10, 9))
                            QDESDL = Trim(Mid(templine, 19, 8))
                            LCOMPDL = Trim(Mid(templine, 27, 9))
                            QCOMPDL = Trim(Mid(templine, 36, 9))
                        End If
                    ElseIf Trim(templine) = "BSPWRMINP  BSPWRMINN" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            '         BSPWRMINP  BSPWRMINN
                            outputFile.writeLine(BSCNAME & "," & Cell & "," & DBPSTATE & "," & SCTYPE & "," & SSDESDL & "," & QDESDL & "," & LCOMPDL & "," & QCOMPDL _
                                & "," & Trim(Mid(templine, 1, 20)) & "," & Trim(Mid(templine, 21, 9)) & "," & DBPBCCHSTATE)
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    private sub doRLBDP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim Cell As String
        Dim CHGR As String, NUMREQBPC As String, NUMREQEGPRSBPC As String, NUMREQCS3CS4BPC As String, TN7BCCH As String, EACPREF As String
        Dim ETCHTN As String, TNBCCH As String, NUMREQE2ABPC As String
        CHGR = ""
        NUMREQBPC = ""
        NUMREQEGPRSBPC = ""
        NUMREQCS3CS4BPC = ""
        TN7BCCH = ""
        EACPREF = ""
        ETCHTN = ""
        TNBCCH = ""
        NUMREQE2ABPC = ""
        'CELL
        'CHGR   NUMREQBPC  NUMREQEGPRSBPC  NUMREQCS3CS4BPC  TN7BCCH  EACPREF
        '       ETCHTN           TNBCCH
        '       ETCHTN           TNBCCH  NUMREQE2ABPC

        Cell = ""

        filepath = outPath & "RLBDP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "CHGR" & "," & "NUMREQBPC" & "," & "NUMREQEGPRSBPC" & "," & "NUMREQCS3CS4BPC" & "," & "TN7BCCH" & "," & "EACPREF" & "," & "ETCHTN" & "," & "TNBCCH" & "," & "NUMREQE2ABPC")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "CELL" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            Cell = Trim(Mid(templine, 1, 9))
                        End If
                    ElseIf Trim(templine) = "CHGR   NUMREQBPC  NUMREQEGPRSBPC  NUMREQCS3CS4BPC  TN7BCCH  EACPREF" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            CHGR = Trim(Mid(templine, 1, 7))
                            NUMREQBPC = Trim(Mid(templine, 8, 11))
                            NUMREQEGPRSBPC = Trim(Mid(templine, 19, 16))
                            NUMREQCS3CS4BPC = Trim(Mid(templine, 35, 17))
                            TN7BCCH = Trim(Mid(templine, 52, 9))
                            EACPREF = Trim(Mid(templine, 61, 9))
                            'CHGR   NUMREQBPC  NUMREQEGPRSBPC  NUMREQCS3CS4BPC  TN7BCCH  EACPREF
                        End If
                    ElseIf Trim(templine) = "ETCHTN           TNBCCH" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            '       ETCHTN           TNBCCH
                            ETCHTN = Trim(Mid(templine, 1, 24))
                            TNBCCH = Trim(Mid(templine, 25, 9))
                            outputFile.writeLine(BSCNAME & "," & Cell & "," & CHGR & "," & NUMREQBPC & "," & NUMREQEGPRSBPC & "," & NUMREQCS3CS4BPC & "," & TN7BCCH & "," & EACPREF _
                                & "," & ETCHTN & "," & TNBCCH)
                            CHGR = ""
                            NUMREQBPC = ""
                            NUMREQEGPRSBPC = ""
                            NUMREQCS3CS4BPC = ""
                            TN7BCCH = ""
                            EACPREF = ""
                            ETCHTN = ""
                            TNBCCH = ""
                            NUMREQE2ABPC = ""
                        End If
                    ElseIf Trim(templine) = "ETCHTN           TNBCCH  NUMREQE2ABPC" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            '       ETCHTN           TNBCCH  NUMREQE2ABPC
                            ETCHTN = Trim(Mid(templine, 1, 24))
                            TNBCCH = Trim(Mid(templine, 25, 8))
                            NUMREQE2ABPC = Trim(Mid(templine, 33, 12))
                            outputFile.writeLine(BSCNAME & "," & Cell & "," & CHGR & "," & NUMREQBPC & "," & NUMREQEGPRSBPC & "," & NUMREQCS3CS4BPC & "," & TN7BCCH & "," & EACPREF _
                                & "," & ETCHTN & "," & TNBCCH & "," & NUMREQE2ABPC)
                            CHGR = ""
                            NUMREQBPC = ""
                            NUMREQEGPRSBPC = ""
                            NUMREQCS3CS4BPC = ""
                            TN7BCCH = ""
                            EACPREF = ""
                            ETCHTN = ""
                            TNBCCH = ""
                            NUMREQE2ABPC = ""
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    private sub doRLCAP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String

        filepath = outPath & "RLCAP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "ALG")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "ALG" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            'ALG
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 9)))
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    private sub doRLCCP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, CHGR As String, SDCCH As String, SDCCHAC As String, TN As String, CBCH As String, CCCH As String
        'CELL
        'CHGR  SDCCH  SDCCHAC  TN  CBCH  CCCH
        CELL = ""
        CHGR = ""
        SDCCH = ""
        SDCCHAC = ""
        TN = ""
        CBCH = ""
        CCCH = ""
        filepath = outPath & "RLCCP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "CHGR" & "," & "SDCCH" & "," & "SDCCHAC" & "," & "TN" & "," & "CBCH" & "," & "CCCH")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If CELL <> "" Then
                        outputFile.writeLine(BSCNAME & "," & CELL & "," & CHGR & "," & SDCCH & "," & SDCCHAC & "," & TN & "," & CBCH & "," & CCCH)
                    End If
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL" Then
                        If CELL <> "" And UCase(Trim(CELL)) <> "NONE" Then
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & CHGR & "," & SDCCH & "," & SDCCHAC & "," & TN & "," & CBCH & "," & CCCH)
                            CELL = ""
                            CHGR = ""
                            SDCCH = ""
                            SDCCHAC = ""
                            TN = ""
                            CBCH = ""
                            CCCH = ""
                        End If
                        CELL = getNowLine()
                    ElseIf UCase(Strings.left(Trim(templine), 18)) = "CHGR  SDCCH  SDCCH" Then
                        'CHGR  SDCCH  SDCCHAC  TN  CBCH  CCCH
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            CHGR = Trim(Mid(templine, 1, 6))
                            SDCCH = Trim(Mid(templine, 7, 7))
                            SDCCHAC = Trim(Mid(templine, 14, 9))
                            TN = Trim(Mid(templine, 23, 4))
                            CBCH = Trim(Mid(templine, 27, 6))
                            CCCH = Trim(Mid(templine, 33, 4))
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL CONFIGURATION CONTROL CHANNEL DATA" Then
                        If Trim(Mid(templine, 1, 7)) <> "" And Len(Trim(Mid(templine, 1, 7))) = 1 Then
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & CHGR & "," & SDCCH & "," & SDCCHAC & "," & TN & "," & CBCH & "," & CCCH)
                            CHGR = Trim(Mid(templine, 1, 6))
                            SDCCH = Trim(Mid(templine, 7, 7))
                            SDCCHAC = Trim(Mid(templine, 14, 9))
                            TN = Trim(Mid(templine, 23, 4))
                            CBCH = Trim(Mid(templine, 27, 6))
                            CCCH = Trim(Mid(templine, 33, 4))
                        Else
                            TN = Trim(TN & " " & Trim(Mid(templine, 23, 4)))
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    private sub doRLCDP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim BSCMC As String

        BSCMC = ""
        filepath = outPath & "RLCDP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "MC" & "," & "BSCMC")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "BSCMC" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            'BSCMC
                            BSCMC = Trim(Mid(templine, 1, 9))
                        End If
                    ElseIf Trim(templine) = "CELL     MC" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            'CELL     MC
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 5)) & "," & BSCMC)
                        End If
                    ElseIf Trim(templine) <> "CELL MULTIPLE CHANNEL ALLOCATION DATA" And UCase(Trim(templine)) <> "NONE" Then
                        If Trim(templine) <> "NONE" Then
                            'CELL     MC
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 5)) & "," & BSCMC)
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    private sub doRLCFP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim filePath1 As String
        Dim outputFile1 As Object
        Dim CELL As String, CHGR As String, SCTYPE As String, SDCCH As String, SDCCHAC As String, TN As String, CBCH As String, HSN As String, HOP As String, DCHNO As String
        Dim HOP_TEMP As String, HSN_TEMP As String, TCH_TEMP As String, TRX_TEMP As String, CHGR_TEMP As String

        HOP_TEMP = ""
        HSN_TEMP = ""
        TCH_TEMP = ""
        TRX_TEMP = ""
        CHGR_TEMP = ""
        filepath = outPath & "RLCFP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "CHGR" & "," & "SCTYPE" & "," & "SDCCH" & "," & "SDCCHAC" & "," & "TN" & "," & "CBCH" & "," & "HSN" & "," & "HOP" & "," & "DCHNO")
        End If
        filePath1 = outPath & "RLCFP_original_" & CDDTime & ".csv"
        If fso.fileExists(filePath1) Then
            outputFile1 = fso.OpenTextFile(filePath1, 8, 1, 0)
        Else
            outputFile1 = fso.OpenTextFile(filePath1, 2, 1, 0)
            outputFile1.writeline("BSC" & "," & "CELL" & "," & "CHGR" & "," & "SCTYPE" & "," & "SDCCH" & "," & "SDCCHAC" & "," & "TN" & "," & "CBCH" & "," & "HSN" & "," & "HOP" & "," & "DCHNO")
        End If
        CELL = ""
        CHGR = ""
        SCTYPE = ""
        SDCCH = ""
        SDCCHAC = ""
        TN = ""
        CBCH = ""
        HSN = ""
        HOP = ""
        DCHNO = ""

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If CELL <> "" Then
                        outputFile.writeLine(BSCNAME & "," & CELL & "," & CHGR & "," & SCTYPE & "," & SDCCH & "," & SDCCHAC & "," & TN & "," & CBCH & "," & HSN & "," & HOP & "," & DCHNO)
                        CFP_BSC_CELL.Add(BSCNAME & "#" & CELL)
                        HOP_arr.Add(HOP_TEMP)
                        HSN_arr.Add(HSN_TEMP)
                        TCH_arr.Add(TCH_TEMP)
                        TRX_ARR.Add(TRX_TEMP)
                        If InStr(TRX_TEMP, " ") Then
                            If UBound(Split(TRX_TEMP, " ")) > MaxTCHNo Then
                                MaxTCHNo = UBound(Split(TRX_TEMP, " "))
                            End If
                        End If
                        CHGR_ARR.Add(CHGR_TEMP)
                        If Trim(SCTYPE) = "OL" Then
                            MBC_CELL_ARR.Add(CELL)
                        End If
                    End If
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        outputFile1.close()
                        outputFile1 = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        outputFile1.close()
                        outputFile1 = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL" Then
                        If CELL <> "" And UCase(Trim(CELL)) <> "NONE" Then
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & CHGR & "," & SCTYPE & "," & SDCCH & "," & SDCCHAC & "," & TN & "," & CBCH & "," & HSN & "," & HOP & "," & DCHNO)
                            CFP_BSC_CELL.Add(BSCNAME & "#" & CELL)
                            HOP_arr.Add(HOP_TEMP)
                            HSN_arr.Add(HSN_TEMP)
                            TCH_arr.Add(TCH_TEMP)
                            TRX_ARR.Add(TRX_TEMP)
                            If InStr(TRX_TEMP, " ") > 0 Then
                                If UBound(Split(TRX_TEMP, " ")) > MaxTCHNo Then
                                    MaxTCHNo = UBound(Split(TRX_TEMP, " "))
                                End If
                            End If
                            CHGR_ARR.Add(CHGR_TEMP)
                            If Trim(SCTYPE) = "OL" Then
                                MBC_CELL_ARR.Add(CELL)
                            End If
                            CELL = ""
                            CHGR = ""
                            SCTYPE = ""
                            SDCCH = ""
                            SDCCHAC = ""
                            TN = ""
                            CBCH = ""
                            HSN = ""
                            HOP = ""
                            DCHNO = ""
                        End If
                        CELL = Trim(getNowLine())
                    ElseIf UCase(Strings.Left(Trim(templine), 18)) = "CHGR   SCTYPE    S" Then
                        templine = getNowLine()
                        CHGR = Trim(Mid(templine, 1, 7))
                        SCTYPE = Trim(Mid(templine, 8, 10))
                        SDCCH = Trim(Mid(templine, 18, 8))
                        SDCCHAC = Trim(Mid(templine, 26, 10))
                        TN = Trim(Mid(templine, 36, 5))
                        CBCH = Trim(Mid(templine, 41, 9))
                        HSN = Trim(Mid(templine, 50, 6))
                        HOP = Trim(Mid(templine, 56, 5))
                        DCHNO = Trim(Mid(templine, 61, 5))
                        HOP_TEMP = "[" & CHGR & "] " & HOP
                        HSN_TEMP = "[" & CHGR & "] " & HSN
                        TCH_TEMP = "[" & CHGR & "] " & DCHNO
                        TRX_TEMP = DCHNO
                        CHGR_TEMP = CHGR
                        outputFile1.writeLine(BSCNAME & "," & CELL & "," & CHGR & "," & Trim(Mid(templine, 8, 10)) & "," & Trim(Mid(templine, 18, 8)) & "," & Trim(Mid(templine, 26, 10)) & "," & Trim(Mid(templine, 36, 5)) & "," & Trim(Mid(templine, 41, 9)) & "," & Trim(Mid(templine, 50, 6)) & "," & Trim(Mid(templine, 56, 5)) & "," & Trim(Mid(templine, 61, 5)))
                    ElseIf UCase(Trim(templine)) <> "CELL CONFIGURATION FREQUENCY DATA" Then
                        If Trim(Mid(templine, 1, 7)) <> "" And Len(Trim(Mid(templine, 1, 7))) = 1 Then
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & CHGR & "," & SCTYPE & "," & SDCCH & "," & SDCCHAC & "," & TN & "," & CBCH & "," & HSN & "," & HOP & "," & DCHNO)
                            CHGR = Trim(Mid(templine, 1, 7))
                            SCTYPE = Trim(Mid(templine, 8, 10))
                            SDCCH = Trim(Mid(templine, 18, 8))
                            SDCCHAC = Trim(Mid(templine, 26, 10))
                            TN = Trim(Mid(templine, 36, 5))
                            CBCH = Trim(Mid(templine, 41, 9))
                            HSN = Trim(Mid(templine, 50, 6))
                            HOP = Trim(Mid(templine, 56, 5))
                            DCHNO = Trim(Mid(templine, 61, 5))
                            HOP_TEMP = Trim(HOP_TEMP & " [" & CHGR & "] " & HOP)
                            HSN_TEMP = Trim(HSN_TEMP & " [" & CHGR & "] " & HSN)
                            TCH_TEMP = Trim(TCH_TEMP & " [" & CHGR & "] " & DCHNO)
                            TRX_TEMP = Trim(TRX_TEMP & " " & DCHNO)
                            CHGR_TEMP = Trim(CHGR_TEMP & " " & CHGR)
                        Else
                            TN = Trim(TN & " " & Trim(Mid(templine, 36, 5)))
                            DCHNO = Trim(DCHNO & " " & Trim(Mid(templine, 61, 5)))
                            TCH_TEMP = Trim(TCH_TEMP & " " & Trim(Mid(templine, 61, 5)))
                            TRX_TEMP = Trim(TRX_TEMP & " " & Trim(Mid(templine, 61, 5)))
                        End If
                        outputFile1.writeLine(BSCNAME & "," & CELL & "," & CHGR & "," & Trim(Mid(templine, 8, 10)) & "," & Trim(Mid(templine, 18, 8)) & "," & Trim(Mid(templine, 26, 10)) & "," & Trim(Mid(templine, 36, 5)) & "," & Trim(Mid(templine, 41, 9)) & "," & Trim(Mid(templine, 50, 6)) & "," & Trim(Mid(templine, 56, 5)) & "," & Trim(Mid(templine, 61, 5)))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile1.close()
        outputFile = Nothing
        outputFile1 = Nothing
    End Sub
    private sub doRLCHP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, CHGR As String, HSN As String, HOP As String, MAIO As String, BCCD As String

        CELL = ""
        CHGR = ""
        HSN = ""
        HOP = ""
        MAIO = ""
        BCCD = ""
        'CELL
        'CHGR  HSN  HOP  MAIO      BCCD
        filepath = outPath & "RLCHP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "CHGR" & "," & "HSN" & "," & "HOP" & "," & "MAIO" & "," & "BCCD")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        If CELL <> "" Then
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & CHGR & "," & HSN & "," & HOP & "," & MAIO & "," & BCCD)
                            CELL = ""
                            CHGR = ""
                            HSN = ""
                            HOP = ""
                            MAIO = ""
                            BCCD = ""
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        If CELL <> "" Then
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & CHGR & "," & HSN & "," & HOP & "," & MAIO & "," & BCCD)
                            CELL = ""
                            CHGR = ""
                            HSN = ""
                            HOP = ""
                            MAIO = ""
                            BCCD = ""
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL" Then
                        templine = getNowLine()
                        If CELL <> UCase(Trim(templine)) And CELL <> "" Then
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & CHGR & "," & HSN & "," & HOP & "," & MAIO & "," & BCCD)
                            CELL = ""
                            CHGR = ""
                            HSN = ""
                            HOP = ""
                            MAIO = ""
                            BCCD = ""
                        End If
                        If Not Trim(templine) = "NONE" Then
                            'CELL
                            CELL = Trim(Mid(templine, 1, 9))
                        End If
                    ElseIf UCase(Trim(templine)) = "CHGR  HSN  HOP  MAIO      BCCD" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'CHGR  HSN  HOP  MAIO      BCCD
                            CHGR = Trim(Mid(templine, 1, 6))
                            HSN = Trim(Mid(templine, 7, 5))
                            HOP = Trim(Mid(templine, 12, 5))
                            MAIO = Trim(Mid(templine, 17, 10))
                            BCCD = Trim(Mid(templine, 27, 5))
                            '                       outputFile.writeLine(BSCNAME & "," & CELL & "," & Trim(Mid(templine, 1, 6)) & "," & Trim(Mid(templine, 7, 5)) & "," & Trim(Mid(templine, 12, 5)) & "," & Trim(Mid(templine, 17, 10)) & "," & Trim(Mid(templine, 27, 5)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL CONFIGURATION FREQUENCY HOPPING DATA" And Len(Trim(Mid(templine, 1, 5))) <= 2 And UCase(Trim(templine)) <> "CELL" And Len(Trim(Mid(templine, 1, 15))) <> 0 Then
                        If CHGR <> "" Then
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & CHGR & "," & HSN & "," & HOP & "," & MAIO & "," & BCCD)
                            CHGR = ""
                            HSN = ""
                            HOP = ""
                            MAIO = ""
                            BCCD = ""
                        End If
                        If Not Trim(templine) = "NONE" Then
                            'CHGR  HSN  HOP  MAIO      BCCD
                            CHGR = Trim(Mid(templine, 1, 6))
                            HSN = Trim(Mid(templine, 7, 5))
                            HOP = Trim(Mid(templine, 12, 5))
                            MAIO = Trim(Mid(templine, 17, 10))
                            BCCD = Trim(Mid(templine, 27, 5))
                            '                       outputFile.writeLine(BSCNAME & "," & CELL & "," & Trim(Mid(templine, 1, 6)) & "," & Trim(Mid(templine, 7, 5)) & "," & Trim(Mid(templine, 12, 5)) & "," & Trim(Mid(templine, 17, 10)) & "," & Trim(Mid(templine, 27, 5)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL CONFIGURATION FREQUENCY HOPPING DATA" And Len(Trim(Mid(templine, 1, 15))) <= 1 Then
                        If Not Trim(templine) = "NONE" Then
                            MAIO = MAIO & " " & Trim(templine)
                            '                 outputFile.writeLine(BSCNAME & "," & CELL & "," & Trim(Mid(templine, 1, 6)) & "," & Trim(Mid(templine, 7, 5)) & "," & Trim(Mid(templine, 12, 5)) & "," & Trim(Mid(templine, 17, 10)) & "," & Trim(Mid(templine, 27, 5)))
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    private sub doRLCLP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String

        'CELL      CSPSALLOC        CSPSPRIO   GPRSPRIO  CODECREST  HRTRX  AHRTRX
        filepath = outPath & "RLCLP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "CSPSALLOC" & "," & "CSPSPRIO" & "," & "GPRSPRIO" & "," & "CODECREST" & "," & "HRTRX" & "," & "AHRTRX")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Strings.left(Trim(templine), 19)) = "CELL      CSPSALLOC" Then
                        templine = getNowLine()
                        'CELL      CSPSALLOC        CSPSPRIO   GPRSPRIO  CODECREST  HRTRX  AHRTRX
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 10)) & "," & Trim(Mid(templine, 11, 17)) & "," & Trim(Mid(templine, 28, 11)) & "," & Trim(Mid(templine, 39, 10)) & "," & Trim(Mid(templine, 49, 11)) & "," & Trim(Mid(templine, 60, 7)) & "," & Trim(Mid(templine, 67, 6)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL CHANNEL ALLOCATION DATA" And Len(Trim(Mid(templine, 1, 10))) <= 8 Then
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 10)) & "," & Trim(Mid(templine, 11, 17)) & "," & Trim(Mid(templine, 28, 11)) & "," & Trim(Mid(templine, 39, 10)) & "," & Trim(Mid(templine, 49, 11)) & "," & Trim(Mid(templine, 60, 7)) & "," & Trim(Mid(templine, 67, 6)))
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    private sub doRLCPP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, BSPWRB As String, TYPE As String

        CELL = ""
        BSPWRB = ""
        TYPE = ""
        filepath = outPath & "RLCPP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "TYPE" & "," & "BSPWRB" & "," & "BSPWRT" & "," & "MSTXPWR" & "," & "SCTYPE")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "CELL      TYPE BSPWRB BSPWRT MSTXPWR SCTYPE" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            'CELL      TYPE BSPWRB BSPWRT MSTXPWR SCTYPE
                            CELL = Trim(Mid(templine, 1, 10))
                            TYPE = Trim(Mid(templine, 11, 5))
                            BSPWRB = Trim(Mid(templine, 16, 7))
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 10)) & "," & Trim(Mid(templine, 11, 5)) & "," & Trim(Mid(templine, 16, 7)) & "," & Trim(Mid(templine, 23, 7)) & "," & Trim(Mid(templine, 30, 8)) & "," & Trim(Mid(templine, 38, 7)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL CONFIGURATION POWER DATA" And Len(Trim(Mid(templine, 11, 5))) <= 3 And Len(Trim(Mid(templine, 11, 5))) <> 0 Then
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 10)) & "," & Trim(Mid(templine, 11, 5)) & "," & Trim(Mid(templine, 16, 7)) & "," & Trim(Mid(templine, 23, 7)) & "," & Trim(Mid(templine, 30, 8)) & "," & Trim(Mid(templine, 38, 7)))
                        CELL = Trim(Mid(templine, 1, 10))
                        TYPE = Trim(Mid(templine, 11, 5))
                        BSPWRB = Trim(Mid(templine, 16, 7))
                    ElseIf UCase(Trim(templine)) <> "CELL CONFIGURATION POWER DATA" And Len(Trim(Mid(templine, 11, 5))) <= 3 And Len(Trim(Mid(templine, 11, 5))) = 0 Then
                        outputFile.writeLINE(BSCNAME & "," & CELL & "," & TYPE & "," & BSPWRB & "," & Trim(Mid(templine, 23, 7)) & "," & Trim(Mid(templine, 30, 8)) & "," & Trim(Mid(templine, 38, 7)))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    private sub doRLCPP_EXT(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, TYPE As String, MSTXPWR As String

        CELL = ""
        TYPE = ""
        MSTXPWR = ""
        filepath = outPath & "RLCPP_EXT_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "TYPE" & "," & "MSTXPWR")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "CELL      TYPE MSTXPWR" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            'CELL      TYPE BSPWRB BSPWRT MSTXPWR SCTYPE
                            CELL = Trim(Mid(templine, 1, 10))
                            TYPE = Trim(Mid(templine, 11, 5))
                            MSTXPWR = Trim(Mid(templine, 16, 7))
                            outputFile.writeLINE(BSCNAME & "," & CELL & "," & TYPE & "," & MSTXPWR)
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL CONFIGURATION POWER DATA" And Len(Trim(Mid(templine, 11, 5))) <= 3 And Len(Trim(Mid(templine, 11, 5))) <> 0 Then
                        CELL = Trim(Mid(templine, 1, 10))
                        TYPE = Trim(Mid(templine, 11, 5))
                        MSTXPWR = Trim(Mid(templine, 16, 7))
                        outputFile.writeLINE(BSCNAME & "," & CELL & "," & TYPE & "," & MSTXPWR)
                    ElseIf UCase(Trim(templine)) <> "CELL CONFIGURATION POWER DATA" And Len(Trim(Mid(templine, 11, 5))) <= 3 And Len(Trim(Mid(templine, 11, 5))) = 0 Then
                        CELL = Trim(Mid(templine, 1, 10))
                        TYPE = Trim(Mid(templine, 11, 5))
                        MSTXPWR = Trim(Mid(templine, 16, 7))
                        outputFile.writeLINE(BSCNAME & "," & CELL & "," & TYPE & "," & MSTXPWR)
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    private sub doRLCRP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String

        'CELL      BCCH  CBCH  SDCCH  NOOFTCH  QUEUED
        filepath = outPath & "RLCRP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "BCCH" & "," & "CBCH" & "," & "SDCCH" & "," & "NOOFTCH(F--H)" & "," & "QUEUED")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "CELL      BCCH  CBCH  SDCCH  NOOFTCH  QUEUED" Then
                        templine = getNowLine()
                        'CELL      BCCH  CBCH  SDCCH  NOOFTCH  QUEUED
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 10)) & "," & Trim(Mid(templine, 11, 6)) & "," & Trim(Mid(templine, 17, 6)) & "," & Trim(Mid(templine, 23, 7)) & "," & Replace(Trim(Mid(templine, 30, 9)), "- ", "--") & "," & Trim(Mid(templine, 39, 9)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL RESOURCES" And Len(Trim(Mid(templine, 1, 10))) <= 8 Then
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 10)) & "," & Trim(Mid(templine, 11, 6)) & "," & Trim(Mid(templine, 17, 6)) & "," & Trim(Mid(templine, 23, 7)) & "," & Replace(Trim(Mid(templine, 30, 9)), "- ", "--") & "," & Trim(Mid(templine, 39, 9)))
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLCXP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String

        'CELL      DTXD
        filepath = outPath & "RLCXP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "DTXD")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL      DTXD" Then
                        templine = getNowLine()
                        'CELL      DTXD
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 10)) & "," & Trim(Mid(templine, 11, 6)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL CONFIGURATION DTX DOWNLINK DATA" And Len(Trim(Mid(templine, 1, 10))) <= 8 Then
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 10)) & "," & Trim(Mid(templine, 11, 6)))
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLDCP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String

        'STATE   EMERGPRL
        filepath = outPath & "RLDCP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "STATE" & "," & "EMERGPRL")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "STATE   EMERGPRL" Then
                        templine = getNowLine()
                        'STATE   EMERGPRL
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 8)) & "," & Trim(Mid(templine, 9, 9)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "BSC DIFFERENTIAL CHANNEL ALLOCATION DATA" And Len(Trim(Mid(templine, 1, 8))) <= 3 Then
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 8)) & "," & Trim(Mid(templine, 9, 9)))
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLDEP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, CGI As String, BSIC As String, BCCHNO As String, AGBLK As String, MFRMS As String, IRC As String, _
        TYPE As String, BCCHTYPE As String, FNOFFSET As String, XRANGE As String, CSYSTYPE As String
        Dim CELLIND As String, RAC As String, RIMNACC As String, GAN As String, DFI As String
        'CELL     CGI                  BSIC  BCCHNO  AGBLK  MFRMS  IRC
        'TYPE     BCCHTYPE             FNOFFSET      XRANGE  CSYSTYPE
        'CELLIND  RAC  RIMNACC  GAN  DFI
        CELL = ""
        CGI = ""
        BSIC = ""
        BCCHNO = ""
        AGBLK = ""
        MFRMS = ""
        IRC = ""
        TYPE = ""
        BCCHTYPE = ""
        FNOFFSET = ""
        XRANGE = ""
        CSYSTYPE = ""
        CELLIND = ""
        RAC = ""
        RIMNACC = ""
        GAN = ""
        DFI = ""

        filepath = outPath & "RLDEP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "CGI" & "," & "BSIC" & "," & "BCCHNO" & "," & "AGBLK" & "," & "MFRMS" & "," & "IRC" & "," & "TYPE" & "," & "BCCHTYPE" & "," & "FNOFFSET" & "," & "XRANGE" & "," & "CSYSTYPE" & "," & "CELLIND" & "," & "RAC" & "," & "RIMNACC" & "," & "GAN" & "," & "DFI" & "," & "NCC" & "," & "BCC")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & CGI & "," & BSIC & "," & BCCHNO & "," & AGBLK & "," & MFRMS & "," & IRC _
                                                 & "," & TYPE & "," & BCCHTYPE & "," & FNOFFSET & "," & XRANGE & "," & CSYSTYPE _
                                                 & "," & CELLIND & "," & RAC & "," & RIMNACC & "," & GAN & "," & DFI & "," & Strings.Left(BSIC, 1) & "," & Strings.Right(BSIC, 1))

                            DEP_BSC_CELL.Add(BSCNAME & "#" & CELL)
                            DEP_Cell_arr.Add(CELL)
                            If CGI <> "" Then
                                CGI_arr.Add(CGI)
                                LAI_arr.Add(Strings.Split(CGI, "-")(2))
                                CI_arr.Add(Strings.Right(CGI, Len(CGI) - InStrRev(CGI, "-")))
                            Else
                                CGI_arr.Add("")
                                LAI_arr.Add("")
                                CI_arr.Add("")
                            End If
                            BCCH_arr.Add(BCCHNO)
                            BSIC_arr.Add(BSIC)
                            CSYSTYPE_ARR.Add(CSYSTYPE)
                            CELL = ""
                            CGI = ""
                            BSIC = ""
                            BCCHNO = ""
                            AGBLK = ""
                            MFRMS = ""
                            IRC = ""
                            TYPE = ""
                            BCCHTYPE = ""
                            FNOFFSET = ""
                            XRANGE = ""
                            CSYSTYPE = ""
                            CELLIND = ""
                            RAC = ""
                            RIMNACC = ""
                            GAN = ""
                            DFI = ""
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & CGI & "," & BSIC & "," & BCCHNO & "," & AGBLK & "," & MFRMS & "," & IRC _
                                                 & "," & TYPE & "," & BCCHTYPE & "," & FNOFFSET & "," & XRANGE & "," & CSYSTYPE _
                                                 & "," & CELLIND & "," & RAC & "," & RIMNACC & "," & GAN & "," & DFI & "," & Strings.Left(BSIC, 1) & "," & Strings.Right(BSIC, 1))

                            DEP_BSC_CELL.Add(BSCNAME & "#" & CELL)
                            DEP_Cell_arr.Add(CELL)

                            If CGI <> "" Then
                                CGI_arr.Add(CGI)
                                LAI_arr.Add(Strings.Split(CGI, "-")(2))
                                CI_arr.Add(Strings.Right(CGI, Len(CGI) - InStrRev(CGI, "-")))
                            Else
                                CGI_arr.Add("")
                                LAI_arr.Add("")
                                CI_arr.Add("")
                            End If
                            BCCH_arr.Add(BCCHNO)
                            BSIC_arr.Add(BSIC)
                            CSYSTYPE_ARR.Add(CSYSTYPE)
                            CELL = ""
                            CGI = ""
                            BSIC = ""
                            BCCHNO = ""
                            AGBLK = ""
                            MFRMS = ""
                            IRC = ""
                            TYPE = ""
                            BCCHTYPE = ""
                            FNOFFSET = ""
                            XRANGE = ""
                            CSYSTYPE = ""
                            CELLIND = ""
                            RAC = ""
                            RIMNACC = ""
                            GAN = ""
                            DFI = ""
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "CELL     CGI                  BSIC  BCCHNO  AGBLK  MFRMS  IRC" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'CELL     CGI                  BSIC  BCCHNO  AGBLK  MFRMS  IRC
                            CELL = Trim(Mid(templine, 1, 9))
                            CGI = Trim(Mid(templine, 10, 21))
                            BSIC = Trim(Mid(templine, 31, 6))
                            BCCHNO = Trim(Mid(templine, 37, 8))
                            AGBLK = Trim(Mid(templine, 45, 7))
                            MFRMS = Trim(Mid(templine, 52, 7))
                            IRC = Trim(Mid(templine, 59, 4))
                        End If
                    ElseIf Trim(templine) = "TYPE     BCCHTYPE             FNOFFSET      XRANGE  CSYSTYPE" Then
                        templine = getNowLine()
                        TYPE = Trim(Mid(templine, 1, 9))
                        BCCHTYPE = Trim(Mid(templine, 10, 21))
                        FNOFFSET = Trim(Mid(templine, 31, 14))
                        XRANGE = Trim(Mid(templine, 45, 8))
                        CSYSTYPE = Trim(Mid(templine, 53, 8))
                    ElseIf Mid(Trim(templine), 1, 26) = "CELLIND  RAC  RIMNACC  GAN" Then
                        templine = getNowLine()
                        CELLIND = Trim(Strings.Left(templine, 9))
                        RAC = Trim(Mid(templine, 10, 5))
                        RIMNACC = Trim(Mid(templine, 15, 9))
                        GAN = Trim(Mid(templine, 24, 5))
                        DFI = Trim(Mid(templine, 29, 15))
                        outputFile.writeline(BSCNAME & "," & CELL & "," & CGI & "," & BSIC & "," & BCCHNO & "," & AGBLK & "," & MFRMS & "," & IRC _
                                             & "," & TYPE & "," & BCCHTYPE & "," & FNOFFSET & "," & XRANGE & "," & CSYSTYPE _
                                             & "," & CELLIND & "," & RAC & "," & RIMNACC & "," & GAN & "," & DFI & "," & Strings.Left(BSIC, 1) & "," & Strings.Right(BSIC, 1))

                        DEP_BSC_CELL.Add(BSCNAME & "#" & CELL)
                        DEP_Cell_arr.Add(CELL)

                        If CGI <> "" Then
                            CGI_arr.Add(CGI)
                            LAI_arr.Add(Strings.Split(CGI, "-")(2))
                            CI_arr.Add(Strings.Right(CGI, Len(CGI) - InStrRev(CGI, "-")))
                        Else
                            CGI_arr.Add("")
                            LAI_arr.Add("")
                            CI_arr.Add("")
                        End If
                        BCCH_arr.Add(BCCHNO)
                        BSIC_arr.Add(BSIC)
                        CSYSTYPE_ARR.Add(CSYSTYPE)
                        CELL = ""
                        CGI = ""
                        BSIC = ""
                        BCCHNO = ""
                        AGBLK = ""
                        MFRMS = ""
                        IRC = ""
                        TYPE = ""
                        BCCHTYPE = ""
                        FNOFFSET = ""
                        XRANGE = ""
                        CSYSTYPE = ""
                        CELLIND = ""
                        RAC = ""
                        RIMNACC = ""
                        GAN = ""
                        DFI = ""
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLDEP_ext(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, CGI As String, BSIC As String, BCCHNO As String, AGBLK As String, MFRMS As String, IRC As String, _
        TYPE As String, BCCHTYPE As String, FNOFFSET As String, XRANGE As String, CSYSTYPE As String
        Dim CELLIND As String, RAC As String, RIMNACC As String, GAN As String, DFI As String
        'CELL     CGI                  BSIC  BCCHNO  AGBLK  MFRMS  IRC
        'TYPE     BCCHTYPE             FNOFFSET      XRANGE  CSYSTYPE
        'CELLIND  RAC  RIMNACC  GAN  DFI
        CELL = ""
        CGI = ""
        BSIC = ""
        BCCHNO = ""
        AGBLK = ""
        MFRMS = ""
        IRC = ""
        TYPE = ""
        BCCHTYPE = ""
        FNOFFSET = ""
        XRANGE = ""
        CSYSTYPE = ""
        CELLIND = ""
        RAC = ""
        RIMNACC = ""
        GAN = ""
        DFI = ""
        filepath = outPath & "RLDEP_EXT_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "CGI" & "," & "BSIC" & "," & "BCCHNO" & "," & "AGBLK" & "," & "MFRMS" & "," & "IRC" & "," & "TYPE" & "," & "BCCHTYPE" & "," & "FNOFFSET" & "," & "XRANGE" & "," & "CSYSTYPE" & "," & "CELLIND" & "," & "RAC" & "," & "RIMNACC" & "," & "GAN" & "," & "DFI")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & CGI & "," & BSIC & "," & BCCHNO & "," & AGBLK & "," & MFRMS & "," & IRC _
                                             & "," & TYPE & "," & BCCHTYPE & "," & FNOFFSET & "," & XRANGE & "," & CSYSTYPE _
                                              & "," & CELLIND & "," & RAC & "," & RIMNACC & "," & GAN & "," & DFI)
                            CELL = ""
                            CGI = ""
                            BSIC = ""
                            BCCHNO = ""
                            AGBLK = ""
                            MFRMS = ""
                            IRC = ""
                            TYPE = ""
                            BCCHTYPE = ""
                            FNOFFSET = ""
                            XRANGE = ""
                            CSYSTYPE = ""
                            CELLIND = ""
                            RAC = ""
                            RIMNACC = ""
                            GAN = ""
                            DFI = ""
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & CGI & "," & BSIC & "," & BCCHNO & "," & AGBLK & "," & MFRMS & "," & IRC _
                                             & "," & TYPE & "," & BCCHTYPE & "," & FNOFFSET & "," & XRANGE & "," & CSYSTYPE _
                                              & "," & CELLIND & "," & RAC & "," & RIMNACC & "," & GAN & "," & DFI)
                            CELL = ""
                            CGI = ""
                            BSIC = ""
                            BCCHNO = ""
                            AGBLK = ""
                            MFRMS = ""
                            IRC = ""
                            TYPE = ""
                            BCCHTYPE = ""
                            FNOFFSET = ""
                            XRANGE = ""
                            CSYSTYPE = ""
                            CELLIND = ""
                            RAC = ""
                            RIMNACC = ""
                            GAN = ""
                            DFI = ""
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "CELL     CGI                  BSIC  BCCHNO  AGBLK  MFRMS  IRC" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'CELL     CGI                  BSIC  BCCHNO  AGBLK  MFRMS  IRC
                            CELL = Trim(Mid(templine, 1, 9))
                            CGI = Trim(Mid(templine, 10, 21))
                            BSIC = Trim(Mid(templine, 31, 6))
                            BCCHNO = Trim(Mid(templine, 37, 8))
                            AGBLK = Trim(Mid(templine, 45, 7))
                            MFRMS = Trim(Mid(templine, 52, 7))
                            IRC = Trim(Mid(templine, 59, 4))
                        End If
                    ElseIf Trim(templine) = "TYPE     BCCHTYPE             FNOFFSET      XRANGE  CSYSTYPE" Then
                        templine = getNowLine()
                        TYPE = Trim(Mid(templine, 1, 9))
                        BCCHTYPE = Trim(Mid(templine, 10, 21))
                        FNOFFSET = Trim(Mid(templine, 31, 14))
                        XRANGE = Trim(Mid(templine, 45, 8))
                        CSYSTYPE = Trim(Mid(templine, 53, 8))
                    ElseIf Mid(Trim(templine), 1, 26) = "CELLIND  RAC  RIMNACC  GAN" Then
                        templine = getNowLine()
                        CELLIND = Trim(Strings.Left(templine, 9))
                        RAC = Trim(Mid(templine, 10, 5))
                        RIMNACC = Trim(Mid(templine, 15, 9))
                        GAN = Trim(Mid(templine, 24, 5))
                        DFI = Trim(Mid(templine, 29, 15))
                        outputFile.writeline(BSCNAME & "," & CELL & "," & CGI & "," & BSIC & "," & BCCHNO & "," & AGBLK & "," & MFRMS & "," & IRC _
                                             & "," & TYPE & "," & BCCHTYPE & "," & FNOFFSET & "," & XRANGE & "," & CSYSTYPE _
                                              & "," & CELLIND & "," & RAC & "," & RIMNACC & "," & GAN & "," & DFI)
                        CELL = ""
                        CGI = ""
                        BSIC = ""
                        BCCHNO = ""
                        AGBLK = ""
                        MFRMS = ""
                        IRC = ""
                        TYPE = ""
                        BCCHTYPE = ""
                        FNOFFSET = ""
                        XRANGE = ""
                        CSYSTYPE = ""
                        CELLIND = ""
                        RAC = ""
                        RIMNACC = ""
                        GAN = ""
                        DFI = ""
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLDEP_ext_UTRAN(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, UTRANID As String, FDDARFCN As String, SCRCODE As String, MRSL As String, _
        CELLIND As String, RAC As String, RIMNACC As String, GAN As String
        'CELL     UTRANID                   FDDARFCN  SCRCODE  MRSL
        'CELLIND  RAC  RIMNACC  GAN
        CELL = ""
        UTRANID = ""
        FDDARFCN = ""
        SCRCODE = ""
        MRSL = ""
        CELLIND = ""
        RAC = ""
        RIMNACC = ""
        GAN = ""

        filepath = outPath & "RLDEP_EXT_UTRAN_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "UTRANID" & "," & "FDDARFCN" & "," & "SCRCODE" & "," & "MRSL" & "," & "CELLIND" & "," & "RAC" & "," & "RIMNACC" & "," & "GAN")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "CELL     UTRANID                   FDDARFCN  SCRCODE  MRSL" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'CELL     UTRANID                   FDDARFCN  SCRCODE  MRSL
                            CELL = Trim(Mid(templine, 1, 9))
                            UTRANID = Trim(Mid(templine, 10, 26))
                            FDDARFCN = Trim(Mid(templine, 36, 10))
                            SCRCODE = Trim(Mid(templine, 46, 9))
                            MRSL = Trim(Mid(templine, 55, 7))
                        End If
                    ElseIf Trim(Mid(Trim(templine), 1, 9)) = "CELLIND" Then
                        templine = getNowLine()
                        outputFile.writeline(BSCNAME & "," & CELL & "," & UTRANID & "," & FDDARFCN & "," & SCRCODE & "," & MRSL _
                                             & "," & Trim(Strings.Left(templine, 9)) & "," & Trim(Mid(templine, 10, 5)) & "," & Trim(Mid(templine, 15, 9)) & "," & Trim(Mid(templine, 24, 3)))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLDGP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String

        CELL = ""
        'CELL      CHGR   SCTYPE   BAND
        filepath = outPath & "RLDGP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "CHGR" & "," & "SCTYPE" & "," & "BAND")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL      CHGR   SCTYPE   BAND" Then
                        templine = getNowLine()
                        'CELL      CHGR   SCTYPE   BAND
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 10)) & "," & Trim(Mid(templine, 11, 7)) & "," & Trim(Mid(templine, 18, 9)) & "," & Trim(Mid(templine, 27, 7)))
                            CELL = Trim(Mid(templine, 1, 10))
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL CHANNEL GROUP DATA" And Trim(Mid(templine, 1, 10)) <> "" Then
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 10)) & "," & Trim(Mid(templine, 11, 7)) & "," & Trim(Mid(templine, 18, 9)) & "," & Trim(Mid(templine, 27, 7)))
                            CELL = Trim(Mid(templine, 1, 10))
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL CHANNEL GROUP DATA" And Trim(Mid(templine, 1, 10)) = "" Then
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & Trim(Mid(templine, 11, 7)) & "," & Trim(Mid(templine, 18, 9)) & "," & Trim(Mid(templine, 27, 7)))
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLDHP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String
        Dim DHA As String
        Dim DHPR As String
        DHPR = ""
        CELL = ""
        DHA = ""
        filepath = outPath & "RLDHP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "DHA" & "," & "DTHAMR" & "," & "DTHNAMR" & "," & "DTHAMRWB" & "," & "DHPRL" & "," & "DHPR")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL      DHA  DTHAMR  DTHNAMR  DTHAMRWB  DHPRL  DHPR" Then
                        templine = getNowLine()
                        'CELL DYNAMIC HR ALLOCATION DATA
                        'CELL      DHA  DTHAMR  DTHNAMR  DTHAMRWB  DHPRL  DHPR
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 10)) & "," & Trim(Mid(templine, 11, 5)) & "," & Trim(Mid(templine, 16, 8)) & "," & Trim(Mid(templine, 24, 9)) & "," & Trim(Mid(templine, 33, 10)) & "," & Trim(Mid(templine, 43, 7)) & "," & Trim(Mid(templine, 50, 7)))
                            CELL = Trim(Mid(templine, 1, 10))
                            DHA = Trim(Mid(templine, 11, 5))
                            DHPR = Trim(Mid(templine, 50, 7))
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL DYNAMIC HR ALLOCATION DATA" And Len(Trim(Mid(templine, 1, 10))) = 0 Then
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & DHA & "," & Trim(Mid(templine, 16, 8)) & "," & Trim(Mid(templine, 24, 9)) & "," & Trim(Mid(templine, 33, 10)) & "," & Trim(Mid(templine, 43, 7)) & "," & DHPR)
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLDMP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, DMSUPP As String, DMTHAMR As String, DMTHNAMR As String, DMTFAMR As String, DMTFNAMR As String, DMQB As String, DMQBAMR As String, DMQBNAMR As String, DMQG As String, DMQGAMR As String, DMQGNAMR As String, DMPR As String, DMPRL As String
        Dim DMQBAMR_2 As String, DMQBNAMR_2 As String, DMQGAMR_2 As String, DMQGNAMR_2 As String
        'CELL      DMSUPP  DMTHAMR  DMTHNAMR  DMTFAMR  DMTFNAMR
        '          DMQB    DMQBAMR  DMQBNAMR
        '          DMQG    DMQGAMR  DMQGNAMR
        '          DMPR
        '          DMPRL   DMQBAMR  DMQBNAMR  DMQGAMR  DMQGNAMR
        CELL = ""
        DMSUPP = ""
        DMTHAMR = ""
        DMTHNAMR = ""
        DMTFAMR = ""
        DMTFNAMR = ""
        DMQB = ""
        DMQBAMR = ""
        DMQBNAMR = ""
        DMQG = ""
        DMQGAMR = ""
        DMQGNAMR = ""
        DMPR = ""
        DMPRL = ""
        DMQBAMR_2 = ""
        DMQBNAMR_2 = ""
        DMQGAMR_2 = ""
        DMQGNAMR_2 = ""

        filepath = outPath & "RLDMP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "DMSUPP" & "," & "DMTHAMR" & "," & "DMTHNAMR" & "," & "DMTFAMR" & "," & "DMTFNAMR" & "," & "DMQB" & "," & "DMQBAMR" & "," & "DMQBNAMR" & "," & "DMQG" & "," & "DMQGAMR" & "," & "DMQGNAMR" & "," & "DMPR" & "," & "DMPRL" & "," & "DMQBAMR" & "," & "DMQBNAMR" & "," & "DMQGAMR" & "," & "DMQGNAMR")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL      DMSUPP  DMTHAMR  DMTHNAMR  DMTFAMR  DMTFNAMR" Then
                        templine = getNowLine()
                        'CELL      DMSUPP  DMTHAMR  DMTHNAMR  DMTFAMR  DMTFNAMR
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 10))
                            DMSUPP = Trim(Mid(templine, 11, 8))
                            DMTHAMR = Trim(Mid(templine, 19, 9))
                            DMTHNAMR = Trim(Mid(templine, 28, 10))
                            DMTFAMR = Trim(Mid(templine, 38, 9))
                            DMTFNAMR = Trim(Mid(templine, 47, 8))
                        End If
                    ElseIf UCase(Trim(templine)) = "DMQB    DMQBAMR  DMQBNAMR" Then
                        templine = getNowLine()
                        '          DMQB    DMQBAMR  DMQBNAMR
                        If Not Trim(templine) = "NONE" Then
                            DMQB = Trim(Mid(templine, 11, 8))
                            DMQBAMR = Trim(Mid(templine, 19, 9))
                            DMQBNAMR = Trim(Mid(templine, 28, 8))
                        End If
                    ElseIf UCase(Trim(templine)) = "DMQG    DMQGAMR  DMQGNAMR" Then
                        templine = getNowLine()
                        '          DMQG    DMQGAMR  DMQGNAMR
                        If Not Trim(templine) = "NONE" Then
                            DMQG = Trim(Mid(templine, 11, 8))
                            DMQGAMR = Trim(Mid(templine, 19, 9))
                            DMQGNAMR = Trim(Mid(templine, 28, 8))
                        End If
                    ElseIf UCase(Trim(templine)) = "DMPR" Then
                        templine = getNowLine()
                        '          DMPR
                        If Not Trim(templine) = "NONE" Then
                            DMPR = Trim(Mid(templine, 11, 8))
                        End If
                    ElseIf UCase(Trim(templine)) = "DMPRL   DMQBAMR  DMQBNAMR  DMQGAMR  DMQGNAMR" Then
                        templine = getNowLine()
                        '          DMPRL   DMQBAMR  DMQBNAMR  DMQGAMR  DMQGNAMR
                        If Not Trim(templine) = "NONE" Then
                            DMPRL = Trim(Mid(templine, 11, 8))
                            DMQBAMR_2 = Trim(Mid(templine, 19, 9))
                            DMQBNAMR_2 = Trim(Mid(templine, 28, 10))
                            DMQGAMR_2 = Trim(Mid(templine, 38, 9))
                            DMQGNAMR_2 = Trim(Mid(templine, 47, 8))
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & DMSUPP & "," & DMTHAMR & "," & DMTHNAMR & "," & DMTFAMR & "," & DMTFNAMR & "," & DMQB & "," & DMQBAMR & "," & DMQBNAMR & "," & DMQG & "," & DMQGAMR & "," & DMQGNAMR & "," & DMPR & "," & DMPRL & "," & DMQBAMR_2 & "," & DMQBNAMR_2 & "," & DMQGAMR_2 & "," & DMQGNAMR_2)
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL DYNAMIC FR/HR MODE ADAPTATION DATA" And Trim(Mid(templine, 1, 10)) = "" Then
                        If Not Trim(templine) = "NONE" Then
                            DMPRL = Trim(Mid(templine, 11, 8))
                            DMQBAMR_2 = Trim(Mid(templine, 19, 9))
                            DMQBNAMR_2 = Trim(Mid(templine, 28, 10))
                            DMQGAMR_2 = Trim(Mid(templine, 38, 9))
                            DMQGNAMR_2 = Trim(Mid(templine, 47, 8))
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & DMSUPP & "," & DMTHAMR & "," & DMTHNAMR & "," & DMTFAMR & "," & DMTFNAMR & "," & DMQB & "," & DMQBAMR & "," & DMQBNAMR & "," & DMQG & "," & DMQGAMR & "," & DMQGNAMR & "," & DMPR & "," & DMPRL & "," & DMQBAMR_2 & "," & DMQBNAMR_2 & "," & DMQGAMR_2 & "," & DMQGNAMR_2)
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLDRP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String
        Dim CMDR As String
        'CELL     CMDR
        CELL = ""
        CMDR = ""
        filepath = outPath & "RLDRP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "CMDR")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     CMDR" Then
                        templine = getNowLine()
                        'CELL     CMDR
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 9))
                            CMDR = Trim(Mid(templine, 10, 7))
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & CMDR)
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL CHANNEL DATA RATE" And Len(Trim(Mid(templine, 1, 9))) > 0 And Len(Trim(Mid(templine, 1, 9))) < 8 Then
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 9))
                            CMDR = Trim(Mid(templine, 10, 7))
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & CMDR)
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLDTP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String
        Dim CHGR As String, TSC As String
        CELL = ""
        CHGR = ""
        TSC = ""
        'CELL
        'CHGR     TSC

        filepath = outPath & "RLDTP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "CHGR" & "," & "TSC")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL" Then
                        templine = getNowLine()
                        'CELL
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(templine)
                        End If
                    ElseIf UCase(Trim(templine)) = "CHGR     TSC" Then
                        templine = getNowLine()
                        'CHGR     TSC
                        If Not Trim(templine) = "NONE" Then
                            CHGR = Trim(Mid(templine, 1, 9))
                            TSC = Trim(Mid(templine, 10, 6))
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & CHGR & "," & TSC)
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL TRAINING SEQUENCE CODE DATA" And Len(Trim(Mid(templine, 1, 9))) <= 1 Then
                        If Not Trim(templine) = "NONE" Then
                            CHGR = Trim(Mid(templine, 1, 9))
                            TSC = Trim(Mid(templine, 10, 6))
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & CHGR & "," & TSC)
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLDUP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, DTMSTATE As String, MAXLAPDM As String, ALLOCPREF As String, OPTMSCLASS As String
        'CELL     DTMSTATE  MAXLAPDM  ALLOCPREF  OPTMSCLASS

        CELL = ""
        DTMSTATE = ""
        MAXLAPDM = ""
        ALLOCPREF = ""
        OPTMSCLASS = ""

        filepath = outPath & "RLDUP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "DTMSTATE" & "," & "MAXLAPDM" & "," & "ALLOCPREF" & "," & "OPTMSCLASS")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     DTMSTATE  MAXLAPDM  ALLOCPREF  OPTMSCLASS" Then
                        templine = getNowLine()
                        'CELL     DTMSTATE  MAXLAPDM  ALLOCPREF  OPTMSCLASS
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 9))
                            DTMSTATE = Trim(Mid(templine, 10, 10))
                            MAXLAPDM = Trim(Mid(templine, 20, 10))
                            ALLOCPREF = Trim(Mid(templine, 30, 11))
                            OPTMSCLASS = Trim(Mid(templine, 41, 11))
                            outputFile.writeline(BSCNAME & "," & CELL & "," & DTMSTATE & "," & MAXLAPDM & "," & ALLOCPREF & "," & OPTMSCLASS)
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL DUAL TRANSFER MODE DATA" And Len(Trim(Mid(templine, 1, 9))) <= 8 Then
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 9))
                            DTMSTATE = Trim(Mid(templine, 10, 10))
                            MAXLAPDM = Trim(Mid(templine, 20, 10))
                            ALLOCPREF = Trim(Mid(templine, 30, 11))
                            OPTMSCLASS = Trim(Mid(templine, 41, 11))
                            outputFile.writeline(BSCNAME & "," & CELL & "," & DTMSTATE & "," & MAXLAPDM & "," & ALLOCPREF & "," & OPTMSCLASS)
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLEFP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, EARFCN As String, LISTTYPE As String, EBCAST As String, EFREE As String, EREQ As String
        'CELL
        'EARFCN

        'CELL
        'D57390P

        'LISTTYPE
        'IDLE

        '     EARFCN
        '     38355 

        'LISTTYPE    EBCAST   EFREE  EREQ
        'ACTIVE      UNKNOWN

        '     EARFCN

        CELL = ""
        EARFCN = ""
        LISTTYPE = ""
        EBCAST = ""
        EFREE = ""
        EREQ = ""

        filepath = outPath & "RLEFP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "EARFCN" & "," & "LISTTYPE" & "," & "EBCAST" & "," & "EFREE" & "," & "EREQ")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        If CELL <> "" And EARFCN <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & EARFCN & "," & LISTTYPE & "," & EBCAST & "," & EFREE & "," & EREQ)
                            CELL = ""
                            EARFCN = ""
                            LISTTYPE = ""
                            EBCAST = ""
                            EFREE = ""
                            EREQ = ""
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        If CELL <> "" And EARFCN <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & EARFCN & "," & LISTTYPE & "," & EBCAST & "," & EFREE & "," & EREQ)
                            CELL = ""
                            EARFCN = ""
                            LISTTYPE = ""
                            EBCAST = ""
                            EFREE = ""
                            EREQ = ""
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL" Then
                        If CELL <> "" And EARFCN <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & EARFCN & "," & LISTTYPE & "," & EBCAST & "," & EFREE & "," & EREQ)
                            CELL = ""
                            EARFCN = ""
                            LISTTYPE = ""
                            EBCAST = ""
                            EFREE = ""
                            EREQ = ""
                        End If
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 9))
                        End If
                    ElseIf UCase(Trim(Mid(templine, 1, 12))) = "LISTTYPE" Then
                        templine = getNowLine()
                        LISTTYPE = Trim(Mid(templine, 1, 12))
                        EBCAST = Trim(Mid(templine, 13, 9))
                        EFREE = Trim(Mid(templine, 22, 7))
                        EREQ = Trim(Mid(templine, 29, 9))
                    ElseIf UCase(Trim(templine)) = "EARFCN" Then
                        templine = getNowLine()
                        EARFCN = Trim(templine)
                        outputFile.writeline(BSCNAME & "," & CELL & "," & EARFCN & "," & LISTTYPE & "," & EBCAST & "," & EFREE & "," & EREQ)
                        EARFCN = ""
                    ElseIf UCase(Trim(templine)) <> "CELL E-UTRAN MEASUREMENT FREQUENCY DATA" And Len(Trim(Mid(templine, 1, 9))) <= 6 Then
                        If Not Trim(templine) = "NONE" Then
                            EARFCN = Trim(templine)
                            outputFile.writeline(BSCNAME & "," & CELL & "," & EARFCN & "," & LISTTYPE & "," & EBCAST & "," & EFREE & "," & EREQ)
                            EARFCN = ""
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing

    End Sub
    Private Sub doRLGAP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String
        CELL = ""
        'CELL     CHGR     SAS       ODPDCHLIMIT

        filepath = outPath & "RLGAP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "CHGR" & "," & "SAS" & "," & "ODPDCHLIMIT")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     CHGR     SAS       ODPDCHLIMIT" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 9))
                            outputFile.writeLINE(BSCNAME & "," & CELL & "," & Trim(Mid(templine, 10, 9)) & "," & Trim(Mid(templine, 19, 10)) & "," & Trim(Mid(templine, 29, 11)))
                        End If
                    ElseIf Len(Trim(Mid(templine, 1, 9))) = 0 Then
                        outputFile.writeLINE(BSCNAME & "," & CELL & "," & Trim(Mid(templine, 10, 9)) & "," & Trim(Mid(templine, 19, 10)) & "," & Trim(Mid(templine, 29, 11)))
                    ElseIf UCase(Trim(templine)) <> "CELL CHANNEL GROUP ALLOCATION DATA" And Len(Trim(Mid(templine, 1, 9))) <= 7 And Len(Trim(Mid(templine, 1, 9))) > 0 Then
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 9)) & "," & Trim(Mid(templine, 19, 10)) & "," & Trim(Mid(templine, 29, 11)))
                        CELL = Trim(Mid(templine, 1, 9))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLGRP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, NOPDCH As String, RP As String
        CELL = ""
        NOPDCH = ""
        RP = ""
        ' CELL     NOPDCH  RP
        'BPC    PDCHTYPE  SPDCHSTATE  PDCHCAP   PDCHIND  PSETIND  TN


        filepath = outPath & "RLGRP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "NOPDCH" & "," & "RP" & "," & "BPC" & "," & "PDCHTYPE" & "," & "SPDCHSTATE" _
                                 & "," & "PDCHCAP" & "," & "PDCHIND" & "," & "PSETIND" & "," & "TN")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     NOPDCH  RP" Then
                        templine = getNowLine()
                        'CELL     NOPDCH  RP
                        'BPC    PDCHTYPE  SPDCHSTATE  PDCHCAP  PDCHIND  PSETIND  TN
                        'BPC    PDCHTYPE  SPDCHSTATE  PDCHCAP   PDCHIND  PSETIND  TN
                        'CELL GENERAL PACKET RADIO SERVICE RESOURCES DATA
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 9))
                            NOPDCH = Trim(Mid(templine, 10, 8))
                            RP = Trim(Mid(templine, 18, 5))
                        End If
                    ElseIf UCase(Mid(Trim(templine), 1, 27)) = "BPC    PDCHTYPE  SPDCHSTATE" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" And Trim(templine) <> "" Then
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & NOPDCH & "," & RP & "," & Trim(Mid(templine, 1, 7)) & "," & Trim(Mid(templine, 8, 10)) & "," & Trim(Mid(templine, 18, 12)) _
                                                  & "," & Trim(Mid(templine, 30, 9)) & "," & Trim(Mid(templine, 39, 9)) & "," & Trim(Mid(templine, 48, 9)) & "," & Trim(Mid(templine, 57, 4)))
                        ElseIf Trim(templine) = "" Then
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & NOPDCH & "," & RP)
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL GENERAL PACKET RADIO SERVICE RESOURCES DATA" And Len(Trim(Mid(templine, 1, 7))) <= 6 And UCase(Mid(Trim(templine), 1, 27)) <> "BPC    PDCHTYPE  SPDCHSTATE" Then
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & NOPDCH & "," & RP & "," & Trim(Mid(templine, 1, 7)) & "," & Trim(Mid(templine, 8, 10)) & "," & Trim(Mid(templine, 18, 12)) _
                                                  & "," & Trim(Mid(templine, 30, 9)) & "," & Trim(Mid(templine, 39, 9)) & "," & Trim(Mid(templine, 48, 9)) & "," & Trim(Mid(templine, 57, 4)))
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLGSP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, GPRSSUP As String, BVCI As String, FPDCH As String, GAMMA As String, PSKONBCCH As String, LA As String, CHCSDL As String
        Dim SCALLOC As String, PDCHPREEMPT As String, MPDCH As String, PRIMPLIM As String, SPDCH As String, FLEXHIGHGPRS As String, EFACTOR As String
        Dim RTTI As String, IAN As String, E2AFACTOR As String, E2ALQC As String, EINITMCS As String, E2AINITMCS As String, RTTIINITMCS As String
        Dim TBFDLLIMIT As String, TBFULLIMIT As String
        Dim TBFDLLIM As String, TBFULLIM As String, OSTBFRTHR As String
        CELL = ""
        GPRSSUP = ""
        BVCI = ""
        FPDCH = ""
        GAMMA = ""
        PSKONBCCH = ""
        LA = ""
        CHCSDL = ""
        SCALLOC = ""
        PDCHPREEMPT = ""
        MPDCH = ""
        PRIMPLIM = ""
        SPDCH = ""
        FLEXHIGHGPRS = ""
        EFACTOR = ""
        RTTI = ""
        IAN = ""
        E2AFACTOR = ""
        E2ALQC = ""
        EINITMCS = ""
        E2AINITMCS = ""
        RTTIINITMCS = ""
        TBFDLLIMIT = ""
        TBFULLIMIT = ""
        TBFDLLIM = ""
        TBFULLIM = ""
        OSTBFRTHR = ""
        'CELL     GPRSSUP  BVCI   FPDCH  GAMMA  PSKONBCCH  LA   CHCSDL
        'SCALLOC  PDCHPREEMPT  MPDCH  PRIMPLIM  SPDCH  FLEXHIGHGPRS  EFACTOR
        'RTTI      IAN       E2AFACTOR  E2ALQC  EINITMCS  E2AINITMCS  RTTIINITMCS
        'RTTI      IAN    /RTTI      IAN       TBFDLLIMIT  TBFULLIMIT
        'TBFDLLIM  TBFULLIM
        'TBFDLLIM  TBFULLIM   OSTBFRTHR
        filepath = outPath & "RLGSP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeLine("BSC" & "," & "CELL" & "," & "GPRSSUP" & "," & "BVCI" & "," & "FPDCH" & "," & "GAMMA" & "," & "PSKONBCCH" & "," & "LA" & "," & "CHCSDL" _
                                & "," & "SCALLOC" & "," & "PDCHPREEMPT" & "," & "MPDCH" & "," & "PRIMPLIM" & "," & "SPDCH" & "," & "FLEXHIGHGPRS" & "," & "EFACTOR" _
                                 & "," & "RTTI" & "," & "IAN" & "," & "E2AFACTOR" & "," & "E2ALQC" & "," & "EINITMCS" & "," & "E2AINITMCS" & "," & "RTTIINITMCS" _
                                 & "," & "TBFDLLIMIT" & "," & "TBFULLIMIT" & "," & "TBFDLLIM" & "," & "TBFULLIM" & "," & "OSTBFRTHR")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If CELL <> "" Then
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & GPRSSUP & "," & BVCI & "," & FPDCH & "," & GAMMA & "," & PSKONBCCH & "," & LA & "," & CHCSDL _
                                                & "," & SCALLOC & "," & PDCHPREEMPT & "," & MPDCH & "," & PRIMPLIM & "," & SPDCH & "," & FLEXHIGHGPRS & "," & EFACTOR _
                                                 & "," & RTTI & "," & IAN & "," & E2AFACTOR & "," & E2ALQC & "," & EINITMCS & "," & E2AINITMCS & "," & RTTIINITMCS _
                                                & "," & TBFDLLIMIT & "," & TBFULLIMIT & "," & TBFDLLIM & "," & TBFULLIM & "," & OSTBFRTHR)
                            CELL = ""
                            GPRSSUP = ""
                            BVCI = ""
                            FPDCH = ""
                            GAMMA = ""
                            PSKONBCCH = ""
                            LA = ""
                            CHCSDL = ""
                            SCALLOC = ""
                            PDCHPREEMPT = ""
                            MPDCH = ""
                            PRIMPLIM = ""
                            SPDCH = ""
                            FLEXHIGHGPRS = ""
                            EFACTOR = ""
                            RTTI = ""
                            IAN = ""
                            E2AFACTOR = ""
                            E2ALQC = ""
                            EINITMCS = ""
                            E2AINITMCS = ""
                            RTTIINITMCS = ""
                            TBFDLLIMIT = ""
                            TBFULLIMIT = ""
                            TBFDLLIM = ""
                            TBFULLIM = ""
                            OSTBFRTHR = ""
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If CELL <> "" Then
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & GPRSSUP & "," & BVCI & "," & FPDCH & "," & GAMMA & "," & PSKONBCCH & "," & LA & "," & CHCSDL _
                                                & "," & SCALLOC & "," & PDCHPREEMPT & "," & MPDCH & "," & PRIMPLIM & "," & SPDCH & "," & FLEXHIGHGPRS & "," & EFACTOR _
                                                 & "," & RTTI & "," & IAN & "," & E2AFACTOR & "," & E2ALQC & "," & EINITMCS & "," & E2AINITMCS & "," & RTTIINITMCS _
                                                & "," & TBFDLLIMIT & "," & TBFULLIMIT & "," & TBFDLLIM & "," & TBFULLIM & "," & OSTBFRTHR)
                            CELL = ""
                            GPRSSUP = ""
                            BVCI = ""
                            FPDCH = ""
                            GAMMA = ""
                            PSKONBCCH = ""
                            LA = ""
                            CHCSDL = ""
                            SCALLOC = ""
                            PDCHPREEMPT = ""
                            MPDCH = ""
                            PRIMPLIM = ""
                            SPDCH = ""
                            FLEXHIGHGPRS = ""
                            EFACTOR = ""
                            RTTI = ""
                            IAN = ""
                            E2AFACTOR = ""
                            E2ALQC = ""
                            EINITMCS = ""
                            E2AINITMCS = ""
                            RTTIINITMCS = ""
                            TBFDLLIMIT = ""
                            TBFULLIMIT = ""
                            TBFDLLIM = ""
                            TBFULLIM = ""
                            OSTBFRTHR = ""
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     GPRSSUP  BVCI   FPDCH  GAMMA  PSKONBCCH  LA   CHCSDL" Then
                        templine = getNowLine()
                        If CELL <> "" Then
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & GPRSSUP & "," & BVCI & "," & FPDCH & "," & GAMMA & "," & PSKONBCCH & "," & LA & "," & CHCSDL _
                                                & "," & SCALLOC & "," & PDCHPREEMPT & "," & MPDCH & "," & PRIMPLIM & "," & SPDCH & "," & FLEXHIGHGPRS & "," & EFACTOR _
                                                 & "," & RTTI & "," & IAN & "," & E2AFACTOR & "," & E2ALQC & "," & EINITMCS & "," & E2AINITMCS & "," & RTTIINITMCS _
                                                & "," & TBFDLLIMIT & "," & TBFULLIMIT & "," & TBFDLLIM & "," & TBFULLIM & "," & OSTBFRTHR)
                            CELL = ""
                            GPRSSUP = ""
                            BVCI = ""
                            FPDCH = ""
                            GAMMA = ""
                            PSKONBCCH = ""
                            LA = ""
                            CHCSDL = ""
                            SCALLOC = ""
                            PDCHPREEMPT = ""
                            MPDCH = ""
                            PRIMPLIM = ""
                            SPDCH = ""
                            FLEXHIGHGPRS = ""
                            EFACTOR = ""
                            RTTI = ""
                            IAN = ""
                            E2AFACTOR = ""
                            E2ALQC = ""
                            EINITMCS = ""
                            E2AINITMCS = ""
                            RTTIINITMCS = ""
                            TBFDLLIMIT = ""
                            TBFULLIMIT = ""
                            TBFDLLIM = ""
                            TBFULLIM = ""
                        End If
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 9))
                            GPRSSUP = Trim(Mid(templine, 10, 9))
                            BVCI = Trim(Mid(templine, 19, 7))
                            FPDCH = Trim(Mid(templine, 26, 7))
                            GAMMA = Trim(Mid(templine, 33, 7))
                            PSKONBCCH = Trim(Mid(templine, 40, 11))
                            LA = Trim(Mid(templine, 51, 5))
                            CHCSDL = Trim(Mid(templine, 56, 8))
                        End If
                    ElseIf UCase(Trim(templine)) = "SCALLOC  PDCHPREEMPT  MPDCH  PRIMPLIM  SPDCH  FLEXHIGHGPRS  EFACTOR" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            SCALLOC = Trim(Mid(templine, 1, 9))
                            PDCHPREEMPT = Trim(Mid(templine, 10, 13))
                            MPDCH = Trim(Mid(templine, 23, 7))
                            PRIMPLIM = Trim(Mid(templine, 30, 10))
                            SPDCH = Trim(Mid(templine, 40, 7))
                            FLEXHIGHGPRS = Trim(Mid(templine, 47, 14))
                            EFACTOR = Trim(Mid(templine, 61, 9))
                        End If
                    ElseIf UCase(Trim(templine)) = "RTTI      IAN       E2AFACTOR  E2ALQC  EINITMCS  E2AINITMCS  RTTIINITMCS" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            RTTI = Trim(Mid(templine, 1, 10))
                            IAN = Trim(Mid(templine, 11, 10))
                            E2AFACTOR = Trim(Mid(templine, 21, 11))
                            E2ALQC = Trim(Mid(templine, 32, 8))
                            EINITMCS = Trim(Mid(templine, 40, 10))
                            E2AINITMCS = Trim(Mid(templine, 50, 12))
                            RTTIINITMCS = Trim(Mid(templine, 62, 11))
                        End If
                    ElseIf UCase(Trim(templine)) = "TBFDLLIM  TBFULLIM" Or UCase(Trim(templine)) = "TBFDLLIM  TBFULLIM   OSTBFRTHR" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            TBFDLLIM = Trim(Mid(templine, 1, 10))
                            TBFULLIM = Trim(Mid(templine, 11, 11))
                            OSTBFRTHR = Trim(Mid(templine, 22, 11))
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & GPRSSUP & "," & BVCI & "," & FPDCH & "," & GAMMA & "," & PSKONBCCH & "," & LA & "," & CHCSDL _
                                                & "," & SCALLOC & "," & PDCHPREEMPT & "," & MPDCH & "," & PRIMPLIM & "," & SPDCH & "," & FLEXHIGHGPRS & "," & EFACTOR _
                                                 & "," & RTTI & "," & IAN & "," & E2AFACTOR & "," & E2ALQC & "," & EINITMCS & "," & E2AINITMCS & "," & RTTIINITMCS _
                                                & "," & TBFDLLIMIT & "," & TBFULLIMIT & "," & TBFDLLIM & "," & TBFULLIM & "," & OSTBFRTHR)
                            CELL = ""
                            GPRSSUP = ""
                            BVCI = ""
                            FPDCH = ""
                            GAMMA = ""
                            PSKONBCCH = ""
                            LA = ""
                            CHCSDL = ""
                            SCALLOC = ""
                            PDCHPREEMPT = ""
                            MPDCH = ""
                            PRIMPLIM = ""
                            SPDCH = ""
                            FLEXHIGHGPRS = ""
                            EFACTOR = ""
                            RTTI = ""
                            IAN = ""
                            E2AFACTOR = ""
                            E2ALQC = ""
                            EINITMCS = ""
                            E2AINITMCS = ""
                            RTTIINITMCS = ""
                            TBFDLLIMIT = ""
                            TBFULLIMIT = ""
                            TBFDLLIM = ""
                            TBFULLIM = ""
                            OSTBFRTHR = ""
                        End If
                    ElseIf UCase(Trim(templine)) = "RTTI      IAN       TBFDLLIMIT  TBFULLIMIT" Then
                        'RTTI      IAN       TBFDLLIMIT  TBFULLIMIT
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            RTTI = Trim(Mid(templine, 1, 10))
                            IAN = Trim(Mid(templine, 11, 10))
                            TBFDLLIMIT = Trim(Mid(templine, 21, 12))
                            TBFULLIMIT = Trim(Mid(templine, 33, 10))
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & GPRSSUP & "," & BVCI & "," & FPDCH & "," & GAMMA & "," & PSKONBCCH & "," & LA & "," & CHCSDL _
                                                & "," & SCALLOC & "," & PDCHPREEMPT & "," & MPDCH & "," & PRIMPLIM & "," & SPDCH & "," & FLEXHIGHGPRS & "," & EFACTOR _
                                                 & "," & RTTI & "," & IAN & "," & E2AFACTOR & "," & E2ALQC & "," & EINITMCS & "," & E2AINITMCS & "," & RTTIINITMCS _
                                                & "," & TBFDLLIMIT & "," & TBFULLIMIT & "," & TBFDLLIM & "," & TBFULLIM & "," & OSTBFRTHR)
                            CELL = ""
                            GPRSSUP = ""
                            BVCI = ""
                            FPDCH = ""
                            GAMMA = ""
                            PSKONBCCH = ""
                            LA = ""
                            CHCSDL = ""
                            SCALLOC = ""
                            PDCHPREEMPT = ""
                            MPDCH = ""
                            PRIMPLIM = ""
                            SPDCH = ""
                            FLEXHIGHGPRS = ""
                            EFACTOR = ""
                            RTTI = ""
                            IAN = ""
                            E2AFACTOR = ""
                            E2ALQC = ""
                            EINITMCS = ""
                            E2AINITMCS = ""
                            RTTIINITMCS = ""
                            TBFDLLIMIT = ""
                            TBFULLIMIT = ""
                            TBFDLLIM = ""
                            TBFULLIM = ""
                            OSTBFRTHR = ""
                        End If
                    ElseIf UCase(Trim(templine)) = "RTTI      IAN" Then
                        'RTTI      IAN
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            RTTI = Trim(Mid(templine, 1, 10))
                            IAN = Trim(Mid(templine, 11, 10))
                            TBFDLLIMIT = Trim(Mid(templine, 21, 12))
                            TBFULLIMIT = Trim(Mid(templine, 33, 10))
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & GPRSSUP & "," & BVCI & "," & FPDCH & "," & GAMMA & "," & PSKONBCCH & "," & LA & "," & CHCSDL _
                                                & "," & SCALLOC & "," & PDCHPREEMPT & "," & MPDCH & "," & PRIMPLIM & "," & SPDCH & "," & FLEXHIGHGPRS & "," & EFACTOR _
                                                 & "," & RTTI & "," & IAN & "," & E2AFACTOR & "," & E2ALQC & "," & EINITMCS & "," & E2AINITMCS & "," & RTTIINITMCS _
                                                & "," & TBFDLLIMIT & "," & TBFULLIMIT & "," & TBFDLLIM & "," & TBFULLIM & "," & OSTBFRTHR)
                            CELL = ""
                            GPRSSUP = ""
                            BVCI = ""
                            FPDCH = ""
                            GAMMA = ""
                            PSKONBCCH = ""
                            LA = ""
                            CHCSDL = ""
                            SCALLOC = ""
                            PDCHPREEMPT = ""
                            MPDCH = ""
                            PRIMPLIM = ""
                            SPDCH = ""
                            FLEXHIGHGPRS = ""
                            EFACTOR = ""
                            RTTI = ""
                            IAN = ""
                            E2AFACTOR = ""
                            E2ALQC = ""
                            EINITMCS = ""
                            E2AINITMCS = ""
                            RTTIINITMCS = ""
                            TBFDLLIMIT = ""
                            TBFULLIMIT = ""
                            TBFDLLIM = ""
                            TBFULLIM = ""
                            OSTBFRTHR = ""
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLHBP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim HCSBANDHYST As String
        Dim HCSBANDTHR As String, HCSBAND As String
        HCSBANDHYST = ""
        HCSBAND = ""
        HCSBANDTHR = ""
        filepath = outPath & "RLHBP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "HCSBANDHYST" & "," & "HCSBAND" & "," & "HCSBANDTHR" & "," & "LAYER")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "HCSBANDHYST" Then
                        templine = getNowLine()
                        'HCSBANDHYST
                        'HCSBAND  HCSBANDTHR  LAYER

                        If Not Trim(templine) = "NONE" Then
                            HCSBANDHYST = Trim(Mid(templine, 1, 10))
                        End If
                    ElseIf UCase(Trim(templine)) = "HCSBAND  HCSBANDTHR  LAYER" Then
                        templine = getNowLine()
                        outputFile.writeLine(BSCNAME & "," & HCSBANDHYST & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 12)) & "," & Trim(Mid(templine, 22, 9)))
                        HCSBAND = Trim(Mid(templine, 1, 9))
                        HCSBANDTHR = Trim(Mid(templine, 10, 12))
                    ElseIf Len(Trim(Mid(templine, 1, 19))) = 0 Then
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLine(BSCNAME & "," & HCSBANDHYST & "," & HCSBAND & "," & HCSBANDTHR & "," & Trim(Mid(templine, 22, 9)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "BSC LOCATING HIERARCHICAL CELL STRUCTURE BAND DATA" And Len(Trim(Mid(templine, 1, 9))) = 1 Then
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLine(BSCNAME & "," & HCSBANDHYST & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 12)) & "," & Trim(Mid(templine, 22, 9)))
                            HCSBAND = Trim(Mid(templine, 1, 9))
                            HCSBANDTHR = Trim(Mid(templine, 10, 12))
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLHPP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'CELL     CHAP

        filepath = outPath & "RLHPP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "CHAP")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     CHAP" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 5)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "CONNECTION OF CELL TO CHANNEL ALLOCATION PROFILE DATA" And Len(Trim(Mid(templine, 1, 9))) <= 7 And Len(Trim(Mid(templine, 1, 9))) > 0 Then
                        outputFile.writeLine(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 5)))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLIHP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, SCTYPE As String, IHO As String, MAXIHO As String, TMAXIHO As String, TIHO As String, SSOFFSETULP As String, SSOFFSETULN As String
        Dim SSOFFSETDLP As String, SSOFFSETDLN As String, QOFFSETULP As String, QOFFSETULN As String, QOFFSETDLP As String, QOFFSETDLN As String
        Dim SSOFFSETULAFRP As String, SSOFFSETULAFRN As String, SSOFFSETDLAFRP As String, SSOFFSETDLAFRN As String
        Dim QOFFSETULAFRP As String, QOFFSETULAFRN As String, QOFFSETDLAFRP As String, QOFFSETDLAFRN As String
        Dim SSOFFSETULAWBP As String, SSOFFSETULAWBN As String, SSOFFSETDLAWBP As String, SSOFFSETDLAWBN As String
        Dim QOFFSETULAWBP As String, QOFFSETULAWBN As String, QOFFSETDLAWBP As String, QOFFSETDLAWBN As String

        CELL = ""
        SCTYPE = ""
        IHO = ""
        MAXIHO = ""
        TMAXIHO = ""
        TIHO = ""
        SSOFFSETULP = ""
        SSOFFSETULN = ""
        SSOFFSETDLP = ""
        SSOFFSETDLN = ""
        QOFFSETULP = ""
        QOFFSETULN = ""
        QOFFSETDLP = ""
        QOFFSETDLN = ""
        SSOFFSETULAFRP = ""
        SSOFFSETULAFRN = ""
        SSOFFSETDLAFRP = ""
        SSOFFSETDLAFRN = ""
        QOFFSETULAFRP = ""
        QOFFSETULAFRN = ""
        QOFFSETDLAFRP = ""
        QOFFSETDLAFRN = ""
        SSOFFSETULAWBP = ""
        SSOFFSETULAWBN = ""
        SSOFFSETDLAWBP = ""
        SSOFFSETDLAWBN = ""
        QOFFSETULAWBP = ""
        QOFFSETULAWBN = ""
        QOFFSETDLAWBP = ""
        QOFFSETDLAWBN = ""
        filepath = outPath & "RLIHP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "SCTYPE" & "," & "IHO" & "," & "MAXIHO" & "," & "TMAXIHO" & "," & "TIHO" & "," & "SSOFFSETULP" & "," & "SSOFFSETULN" _
                                & "," & "SSOFFSETDLP" & "," & "SSOFFSETDLN" & "," & "QOFFSETULP" & "," & "QOFFSETULN" & "," & "QOFFSETDLP" & "," & "QOFFSETDLN" _
                                & "," & "SSOFFSETULAFRP" & "," & "SSOFFSETULAFRN" & "," & "SSOFFSETDLAFRP" & "," & "SSOFFSETDLAFRN" _
                                & "," & "QOFFSETULAFRP" & "," & "QOFFSETULAFRN" & "," & "QOFFSETDLAFRP" & "," & "QOFFSETDLAFRN" _
                                & "," & "SSOFFSETULAWBP" & "," & "SSOFFSETULAWBN" & "," & "SSOFFSETDLAWBP" & "," & "SSOFFSETDLAWBN" _
                                & "," & "QOFFSETULAWBP" & "," & "QOFFSETULAWBN" & "," & "QOFFSETDLAWBP" & "," & "QOFFSETDLAWBN")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If CELL <> "" Then
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & SCTYPE & "," & IHO & "," & MAXIHO & "," & TMAXIHO & "," & TIHO & "," & SSOFFSETULP & "," & SSOFFSETULN _
                            & "," & SSOFFSETDLP & "," & SSOFFSETDLN & "," & QOFFSETULP & "," & QOFFSETULN & "," & QOFFSETDLP & "," & QOFFSETDLN _
                            & "," & SSOFFSETULAFRP & "," & SSOFFSETULAFRN & "," & SSOFFSETDLAFRP & "," & SSOFFSETDLAFRN _
                            & "," & QOFFSETULAFRP & "," & QOFFSETULAFRN & "," & QOFFSETDLAFRP & "," & QOFFSETDLAFRN _
                            & "," & SSOFFSETULAWBP & "," & SSOFFSETULAWBN & "," & SSOFFSETDLAWBP & "," & SSOFFSETDLAWBN _
                            & "," & QOFFSETULAWBP & "," & QOFFSETULAWBN & "," & QOFFSETDLAWBP & "," & QOFFSETDLAWBN)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If CELL <> "" Then
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & SCTYPE & "," & IHO & "," & MAXIHO & "," & TMAXIHO & "," & TIHO & "," & SSOFFSETULP & "," & SSOFFSETULN _
                            & "," & SSOFFSETDLP & "," & SSOFFSETDLN & "," & QOFFSETULP & "," & QOFFSETULN & "," & QOFFSETDLP & "," & QOFFSETDLN _
                            & "," & SSOFFSETULAFRP & "," & SSOFFSETULAFRN & "," & SSOFFSETDLAFRP & "," & SSOFFSETDLAFRN _
                            & "," & QOFFSETULAFRP & "," & QOFFSETULAFRN & "," & QOFFSETDLAFRP & "," & QOFFSETDLAFRN _
                            & "," & SSOFFSETULAWBP & "," & SSOFFSETULAWBN & "," & SSOFFSETDLAWBP & "," & SSOFFSETDLAWBN _
                            & "," & QOFFSETULAWBP & "," & QOFFSETULAWBN & "," & QOFFSETDLAWBP & "," & QOFFSETDLAWBN)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     SCTYPE  IHO  MAXIHO  TMAXIHO  TIHO  SSOFFSETULP  SSOFFSETULN" Then
                        'CELL     SCTYPE  IHO  MAXIHO  TMAXIHO  TIHO  SSOFFSETULP  SSOFFSETULN
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 9))
                            SCTYPE = Trim(Mid(templine, 10, 8))
                            IHO = Trim(Mid(templine, 18, 5))
                            MAXIHO = Trim(Mid(templine, 23, 8))
                            TMAXIHO = Trim(Mid(templine, 31, 9))
                            TIHO = Trim(Mid(templine, 40, 6))
                            SSOFFSETULP = Trim(Mid(templine, 46, 13))
                            SSOFFSETULN = Trim(Mid(templine, 59, 11))
                        End If
                    ElseIf UCase(Trim(templine)) = "SSOFFSETDLP  SSOFFSETDLN  QOFFSETULP  QOFFSETULN  QOFFSETDLP  QOFFSETDLN" Then
                        'SSOFFSETDLP  SSOFFSETDLN  QOFFSETULP  QOFFSETULN  QOFFSETDLP  QOFFSETDLN
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            SSOFFSETDLP = Trim(Mid(templine, 1, 13))
                            SSOFFSETDLN = Trim(Mid(templine, 14, 13))
                            QOFFSETULP = Trim(Mid(templine, 27, 12))
                            QOFFSETULN = Trim(Mid(templine, 39, 12))
                            QOFFSETDLP = Trim(Mid(templine, 51, 12))
                            QOFFSETDLN = Trim(Mid(templine, 63, 12))
                        End If
                    ElseIf UCase(Trim(templine)) = "SSOFFSETULAFRP  SSOFFSETULAFRN  SSOFFSETDLAFRP  SSOFFSETDLAFRN" Then
                        'SSOFFSETULAFRP  SSOFFSETULAFRN  SSOFFSETDLAFRP  SSOFFSETDLAFRN
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            SSOFFSETULAFRP = Trim(Mid(templine, 1, 16))
                            SSOFFSETULAFRN = Trim(Mid(templine, 17, 16))
                            SSOFFSETDLAFRP = Trim(Mid(templine, 33, 16))
                            SSOFFSETDLAFRN = Trim(Mid(templine, 49, 16))
                        End If
                    ElseIf UCase(Trim(templine)) = "QOFFSETULAFRP  QOFFSETULAFRN  QOFFSETDLAFRP  QOFFSETDLAFRN" Then
                        'QOFFSETULAFRP  QOFFSETULAFRN  QOFFSETDLAFRP  QOFFSETDLAFRN
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            QOFFSETULAFRP = Trim(Mid(templine, 1, 15))
                            QOFFSETULAFRN = Trim(Mid(templine, 16, 15))
                            QOFFSETDLAFRP = Trim(Mid(templine, 31, 15))
                            QOFFSETDLAFRN = Trim(Mid(templine, 46, 15))
                        End If
                    ElseIf UCase(Trim(templine)) = "SSOFFSETULAWBP  SSOFFSETULAWBN  SSOFFSETDLAWBP  SSOFFSETDLAWBN" Then
                        'SSOFFSETULAWBP  SSOFFSETULAWBN  SSOFFSETDLAWBP  SSOFFSETDLAWBN
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            SSOFFSETULAWBP = Trim(Mid(templine, 1, 16))
                            SSOFFSETULAWBN = Trim(Mid(templine, 17, 16))
                            SSOFFSETDLAWBP = Trim(Mid(templine, 33, 16))
                            SSOFFSETDLAWBN = Trim(Mid(templine, 49, 16))
                        End If
                    ElseIf UCase(Trim(templine)) = "QOFFSETULAWBP  QOFFSETULAWBN  QOFFSETDLAWBP  QOFFSETDLAWBN" Then
                        'QOFFSETULAWBP  QOFFSETULAWBN  QOFFSETDLAWBP  QOFFSETDLAWBN
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            QOFFSETULAWBP = Trim(Mid(templine, 1, 15))
                            QOFFSETULAWBN = Trim(Mid(templine, 16, 15))
                            QOFFSETDLAWBP = Trim(Mid(templine, 31, 15))
                            QOFFSETDLAWBN = Trim(Mid(templine, 46, 15))
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & SCTYPE & "," & IHO & "," & MAXIHO & "," & TMAXIHO & "," & TIHO & "," & SSOFFSETULP & "," & SSOFFSETULN _
                                                        & "," & SSOFFSETDLP & "," & SSOFFSETDLN & "," & QOFFSETULP & "," & QOFFSETULN & "," & QOFFSETDLP & "," & QOFFSETDLN _
                                                        & "," & SSOFFSETULAFRP & "," & SSOFFSETULAFRN & "," & SSOFFSETDLAFRP & "," & SSOFFSETDLAFRN _
                                                        & "," & QOFFSETULAFRP & "," & QOFFSETULAFRN & "," & QOFFSETDLAFRP & "," & QOFFSETDLAFRN _
                                                        & "," & SSOFFSETULAWBP & "," & SSOFFSETULAWBN & "," & SSOFFSETDLAWBP & "," & SSOFFSETDLAWBN _
                                                        & "," & QOFFSETULAWBP & "," & QOFFSETULAWBN & "," & QOFFSETDLAWBP & "," & QOFFSETDLAWBN)
                            CELL = ""
                            SCTYPE = ""
                            IHO = ""
                            MAXIHO = ""
                            TMAXIHO = ""
                            TIHO = ""
                            SSOFFSETULP = ""
                            SSOFFSETULN = ""
                            SSOFFSETDLP = ""
                            SSOFFSETDLN = ""
                            QOFFSETULP = ""
                            QOFFSETULN = ""
                            QOFFSETDLP = ""
                            QOFFSETDLN = ""
                            SSOFFSETULAFRP = ""
                            SSOFFSETULAFRN = ""
                            SSOFFSETDLAFRP = ""
                            SSOFFSETDLAFRN = ""
                            QOFFSETULAFRP = ""
                            QOFFSETULAFRN = ""
                            QOFFSETDLAFRP = ""
                            QOFFSETDLAFRN = ""
                            SSOFFSETULAWBP = ""
                            SSOFFSETULAWBN = ""
                            SSOFFSETDLAWBP = ""
                            SSOFFSETDLAWBN = ""
                            QOFFSETULAWBP = ""
                            QOFFSETULAWBN = ""
                            QOFFSETDLAWBP = ""
                            QOFFSETDLAWBN = ""
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLIMP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'CELL     ICMSTATE  INTAVE  LIMIT1  LIMIT2  LIMIT3  LIMIT4
        filepath = outPath & "RLIMP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "ICMSTATE" & "," & "INTAVE" & "," & "LIMIT1" & "," & "LIMIT2" & "," & "LIMIT3" & "," & "LIMIT4")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     ICMSTATE  INTAVE  LIMIT1  LIMIT2  LIMIT3  LIMIT4" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 10)) & "," & Trim(Mid(templine, 20, 8)) & "," & Trim(Mid(templine, 28, 8)) & "," & Trim(Mid(templine, 36, 8)) & "," & Trim(Mid(templine, 44, 8)) & "," & Trim(Mid(templine, 52, 8)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL IDLE CHANNEL MEASUREMENT DATA" And Len(Trim(Mid(templine, 1, 9))) <= 7 Then
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 10)) & "," & Trim(Mid(templine, 20, 8)) & "," & Trim(Mid(templine, 28, 8)) & "," & Trim(Mid(templine, 36, 8)) & "," & Trim(Mid(templine, 44, 8)) & "," & Trim(Mid(templine, 52, 8)))
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLLBP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim SYSTYPE As String
        Dim TAAVELEN As String, TINIT As String, TALLOC As String, TURGEN As String, EVALTYPE As String, THO As String, NHO As String
        SYSTYPE = ""
        TAAVELEN = ""
        TINIT = ""
        TALLOC = ""
        TURGEN = ""
        EVALTYPE = ""
        THO = ""
        NHO = ""
        filepath = outPath & "RLLBP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "SYSTYPE" & "," & "TAAVELEN" & "," & "TINIT" & "," & "TALLOC" & "," & "TURGEN" & "," & "EVALTYPE" & "," & "THO" & "," & "NHO" _
                                 & "," & "ASSOC" & "," & "IBHOASS" & "," & "IBHOSICH" & "," & "IHOSICH")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "SYSTYPE" Then
                        templine = getNowLine()
                        'SYSTYPE
                        'TAAVELEN TINIT TALLOC TURGEN EVALTYPE THO NHO
                        'ASSOC IBHOASS  IBHOSICH IHOSICH

                        If Not Trim(templine) = "NONE" Then
                            SYSTYPE = Trim(Mid(templine, 1, 10))
                        End If
                    ElseIf UCase(Trim(templine)) = "TAAVELEN TINIT TALLOC TURGEN EVALTYPE THO NHO" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            TAAVELEN = Trim(Mid(templine, 1, 9))
                            TINIT = Trim(Mid(templine, 10, 6))
                            TALLOC = Trim(Mid(templine, 16, 7))
                            TURGEN = Trim(Mid(templine, 23, 7))
                            EVALTYPE = Trim(Mid(templine, 30, 9))
                            THO = Trim(Mid(templine, 39, 4))
                            NHO = Trim(Mid(templine, 43, 6))
                        End If
                    ElseIf UCase(Trim(templine)) = "ASSOC IBHOASS  IBHOSICH IHOSICH" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLine(BSCNAME & "," & SYSTYPE & "," & TAAVELEN & "," & TINIT & "," & TALLOC & "," & TURGEN & "," & EVALTYPE & "," & THO & "," & NHO _
                                                 & "," & Trim(Mid(templine, 1, 6)) & "," & Trim(Mid(templine, 7, 9)) & "," & Trim(Mid(templine, 16, 9)) & "," & Trim(Mid(templine, 25, 7)) & "," & Trim(Mid(templine, 32, 7)))
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLLCP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'CELL     CLSSTATE  CLSLEVEL  CLSACC    HOCLSACC  RHYST     CLSRAMP
        filepath = outPath & "RLLCP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "CLSSTATE" & "," & "CLSLEVEL" & "," & "CLSACC" & "," & "HOCLSACC" & "," & "RHYST" & "," & "CLSRAMP")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     CLSSTATE  CLSLEVEL  CLSACC    HOCLSACC  RHYST     CLSRAMP" Then
                        'CELL     CLSSTATE  CLSLEVEL  CLSACC    HOCLSACC  RHYST     CLSRAMP
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 10)) & "," & Trim(Mid(templine, 20, 10)) & "," & Trim(Mid(templine, 30, 10)) & "," & Trim(Mid(templine, 40, 10)) & "," & Trim(Mid(templine, 50, 10)) & "," & Trim(Mid(templine, 60, 7)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL LOAD SHARING DATA" And Len(Trim(Mid(templine, 1, 9))) <= 7 Then
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 10)) & "," & Trim(Mid(templine, 20, 10)) & "," & Trim(Mid(templine, 30, 10)) & "," & Trim(Mid(templine, 40, 10)) & "," & Trim(Mid(templine, 50, 10)) & "," & Trim(Mid(templine, 60, 7)))
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLLDP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'CELL     MAXTA  RLINKUP  RLINKUPAFR  RLINKUPAHR  RLINKUPAWB
        filepath = outPath & "RLLDP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "MAXTA" & "," & "RLINKUP" & "," & "RLINKUPAFR" & "," & "RLINKUPAHR" & "," & "RLINKUPAWB")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     MAXTA  RLINKUP  RLINKUPAFR  RLINKUPAHR  RLINKUPAWB" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 7)) & "," & Trim(Mid(templine, 17, 9)) & "," & Trim(Mid(templine, 26, 12)) & "," & Trim(Mid(templine, 36, 12)) & "," & Trim(Mid(templine, 48, 10)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL LOCATING DISCONNECT DATA" And Len(Trim(Mid(templine, 1, 9))) <= 7 Then
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 7)) & "," & Trim(Mid(templine, 17, 9)) & "," & Trim(Mid(templine, 26, 12)) & "," & Trim(Mid(templine, 36, 12)) & "," & Trim(Mid(templine, 48, 10)))
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLLFP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, SSEVALSD As String, QEVALSD As String, SSEVALSI As String, QEVALSI As String
        Dim SSLENSD As String, QLENSD As String, SSLENSI As String, QLENSI As String, SSRAMPSD As String, SSRAMPSI As String, FERLEN As String
        'CELL     SSEVALSD QEVALSD SSEVALSI QEVALSI
        'SSLENSD QLENSD SSLENSI QLENSI SSRAMPSD SSRAMPSI FERLEN
        CELL = ""
        SSEVALSD = ""
        QEVALSD = ""
        SSEVALSI = ""
        QEVALSI = ""
        SSLENSD = ""
        QLENSD = ""
        SSLENSI = ""
        QLENSI = ""
        SSRAMPSD = ""
        SSRAMPSI = ""
        FERLEN = ""
        filepath = outPath & "RLLFP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "SSEVALSD" & "," & "QEVALSD" & "," & "SSEVALSI" & "," & "QEVALSI" _
                                  & "," & "SSLENSD" & "," & "QLENSD" & "," & "SSLENSI" & "," & "QLENSI" & "," & "SSRAMPSD" & "," & "SSRAMPSI" & "," & "FERLEN")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If CELL <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & CELL & "," & SSEVALSD & "," & QEVALSD & "," & SSEVALSI & "," & QEVALSI _
                                             & "," & SSLENSD & "," & QLENSD & "," & SSLENSI & "," & QLENSI & "," & SSRAMPSD & "," & SSRAMPSI & "," & FERLEN)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If CELL <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & CELL & "," & SSEVALSD & "," & QEVALSD & "," & SSEVALSI & "," & QEVALSI _
                                             & "," & SSLENSD & "," & QLENSD & "," & SSLENSI & "," & QLENSI & "," & SSRAMPSD & "," & SSRAMPSI & "," & FERLEN)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     SSEVALSD QEVALSD SSEVALSI QEVALSI" Then
                        'CELL     SSEVALSD QEVALSD SSEVALSI QEVALSI
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 9))
                            SSEVALSD = Trim(Mid(templine, 10, 9))
                            QEVALSD = Trim(Mid(templine, 19, 8))
                            SSEVALSI = Trim(Mid(templine, 27, 9))
                            QEVALSI = Trim(Mid(templine, 36, 7))
                        End If
                    ElseIf UCase(Trim(templine)) = "SSLENSD QLENSD SSLENSI QLENSI SSRAMPSD SSRAMPSI FERLEN" Then
                        'SSLENSD QLENSD SSLENSI QLENSI SSRAMPSD SSRAMPSI FERLEN
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            SSLENSD = Trim(Mid(templine, 1, 8))
                            QLENSD = Trim(Mid(templine, 9, 7))
                            SSLENSI = Trim(Mid(templine, 16, 8))
                            QLENSI = Trim(Mid(templine, 24, 7))
                            SSRAMPSD = Trim(Mid(templine, 31, 9))
                            SSRAMPSI = Trim(Mid(templine, 40, 9))
                            FERLEN = Trim(Mid(templine, 49, 7))
                            outputFile.writeLINE(BSCNAME & "," & CELL & "," & SSEVALSD & "," & QEVALSD & "," & SSEVALSI & "," & QEVALSI _
                                             & "," & SSLENSD & "," & QLENSD & "," & SSLENSI & "," & QLENSI & "," & SSRAMPSD & "," & SSRAMPSI & "," & FERLEN)
                            CELL = ""
                            SSEVALSD = ""
                            QEVALSD = ""
                            SSEVALSI = ""
                            QEVALSI = ""
                            SSLENSD = ""
                            QLENSD = ""
                            SSLENSI = ""
                            QLENSI = ""
                            SSRAMPSD = ""
                            SSRAMPSI = ""
                            FERLEN = ""
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLLHP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, TYPE As String, LAYER As String, LAYERTHR As String, LAYERHYST As String, PSSTEMP As String, PTIMTEMP As String, FASTMSREG As String
        Dim HCSIN As String, HCSOUT As String
        'CELL     TYPE  LAYER  LAYERTHR  LAYERHYST  PSSTEMP  PTIMTEMP  FASTMSREG
        '         HCSIN  HCSOUT
        CELL = ""
        TYPE = ""
        LAYER = ""
        LAYERTHR = ""
        LAYERHYST = ""
        PSSTEMP = ""
        PTIMTEMP = ""
        FASTMSREG = ""
        HCSIN = ""
        HCSOUT = ""
        filepath = outPath & "RLLHP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "TYPE" & "," & "LAYER" & "," & "LAYERTHR" & "," & "LAYERHYST" & "," & "PSSTEMP" & "," & "PTIMTEMP" & "," & "FASTMSREG" _
                                 & "," & "HCSIN" & "," & "HCSOUT")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If CELL <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & CELL & "," & TYPE & "," & LAYER & "," & LAYERTHR & "," & LAYERHYST & "," & PSSTEMP & "," & PTIMTEMP & "," & FASTMSREG _
                     & "," & HCSIN & "," & HCSOUT)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If CELL <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & CELL & "," & TYPE & "," & LAYER & "," & LAYERTHR & "," & LAYERHYST & "," & PSSTEMP & "," & PTIMTEMP & "," & FASTMSREG _
                     & "," & HCSIN & "," & HCSOUT)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     TYPE  LAYER  LAYERTHR  LAYERHYST  PSSTEMP  PTIMTEMP  FASTMSREG" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 9))
                            TYPE = Trim(Mid(templine, 10, 6))
                            LAYER = Trim(Mid(templine, 16, 7))
                            LAYERTHR = Trim(Mid(templine, 23, 10))
                            LAYERHYST = Trim(Mid(templine, 33, 11))
                            PSSTEMP = Trim(Mid(templine, 44, 9))
                            PTIMTEMP = Trim(Mid(templine, 53, 10))
                            FASTMSREG = Trim(Mid(templine, 63, 9))
                        End If
                    ElseIf UCase(Trim(templine)) = "HCSIN  HCSOUT" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            HCSIN = Trim(Mid(templine, 10, 7))
                            HCSOUT = Trim(Mid(templine, 17, 6))
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & TYPE & "," & LAYER & "," & LAYERTHR & "," & LAYERHYST & "," & PSSTEMP & "," & PTIMTEMP & "," & FASTMSREG _
                                                 & "," & HCSIN & "," & HCSOUT)
                            CELL = ""
                            TYPE = ""
                            LAYER = ""
                            LAYERTHR = ""
                            LAYERHYST = ""
                            PSSTEMP = ""
                            PTIMTEMP = ""
                            FASTMSREG = ""
                            HCSIN = ""
                            HCSOUT = ""
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLLHP_EXT(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, TYPE As String, LAYER As String, LAYERTHR As String, LAYERHYST As String, PSSTEMP As String, PTIMTEMP As String, FASTMSREG As String
        Dim HCSIN As String, HCSOUT As String
        'CELL     TYPE  LAYER  LAYERTHR  LAYERHYST  PSSTEMP  PTIMTEMP  FASTMSREG
        '         HCSIN  HCSOUT
        CELL = ""
        TYPE = ""
        LAYER = ""
        LAYERTHR = ""
        LAYERHYST = ""
        PSSTEMP = ""
        PTIMTEMP = ""
        FASTMSREG = ""
        HCSIN = ""
        HCSOUT = ""
        filepath = outPath & "RLLHP_EXT_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "TYPE" & "," & "LAYER" & "," & "LAYERTHR" & "," & "LAYERHYST" & "," & "PSSTEMP" & "," & "PTIMTEMP" & "," & "FASTMSREG" _
                                 & "," & "HCSIN" & "," & "HCSOUT")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If CELL <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & CELL & "," & TYPE & "," & LAYER & "," & LAYERTHR & "," & LAYERHYST & "," & PSSTEMP & "," & PTIMTEMP & "," & FASTMSREG _
                     & "," & HCSIN & "," & HCSOUT)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If CELL <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & CELL & "," & TYPE & "," & LAYER & "," & LAYERTHR & "," & LAYERHYST & "," & PSSTEMP & "," & PTIMTEMP & "," & FASTMSREG _
                     & "," & HCSIN & "," & HCSOUT)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     TYPE  LAYER  LAYERTHR  LAYERHYST  PSSTEMP  PTIMTEMP  FASTMSREG" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 9))
                            TYPE = Trim(Mid(templine, 10, 6))
                            LAYER = Trim(Mid(templine, 16, 7))
                            LAYERTHR = Trim(Mid(templine, 23, 10))
                            LAYERHYST = Trim(Mid(templine, 33, 11))
                            PSSTEMP = Trim(Mid(templine, 44, 9))
                            PTIMTEMP = Trim(Mid(templine, 53, 10))
                            FASTMSREG = Trim(Mid(templine, 63, 9))
                        End If
                    ElseIf UCase(Trim(templine)) = "HCSIN  HCSOUT" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            HCSIN = Trim(Mid(templine, 10, 7))
                            HCSOUT = Trim(Mid(templine, 17, 6))
                            outputFile.writeLINE(BSCNAME & "," & CELL & "," & TYPE & "," & LAYER & "," & LAYERTHR & "," & LAYERHYST & "," & PSSTEMP & "," & PTIMTEMP & "," & FASTMSREG _
                                                 & "," & HCSIN & "," & HCSOUT)
                            CELL = ""
                            TYPE = ""
                            LAYER = ""
                            LAYERTHR = ""
                            LAYERHYST = ""
                            PSSTEMP = ""
                            PTIMTEMP = ""
                            FASTMSREG = ""
                            HCSIN = ""
                            HCSOUT = ""
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLLLP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'CELL     SCLD  SCLDLUL SCLDLOL SCLDSC
        filepath = outPath & "RLLLP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "SCLD" & "," & "SCLDLUL" & "," & "SCLDLOL" & "," & "SCLDSC")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     SCLD  SCLDLUL SCLDLOL SCLDSC" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 6)) & "," & Trim(Mid(templine, 16, 8)) & "," & Trim(Mid(templine, 24, 8)) & "," & Trim(Mid(templine, 32, 6)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "SUBCELL LOAD DISTRIBUTION DATA" And Len(Trim(Mid(templine, 1, 9))) <= 7 Then
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 6)) & "," & Trim(Mid(templine, 16, 8)) & "," & Trim(Mid(templine, 24, 8)) & "," & Trim(Mid(templine, 32, 6)))
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLLOP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, BSPWR As String, BSRXMIN As String, BSRXSUFF As String, MSRXMIN As String, MSRXSUFF As String, SCHO As String, MISSNM As String, AW As String
        Dim BCCHREUSE As String, BCCHLOSS As String, BCCHDTCBP As String, BCCHDTCBN As String, BCCHLOSSHYST As String
        Dim BCCHDTCBHYST As String
        Dim SCTYPE As String, BSTXPWR As String, EXTPEN As String, HYSTSEP As String, ISHOLEV As String, MAXISHO As String, FBOFFSP As String, FBOFFSN As String
        'CELL     BSPWR  BSRXMIN  BSRXSUFF  MSRXMIN  MSRXSUFF  SCHO  MISSNM  AW
        '         BCCHREUSE  BCCHLOSS  BCCHDTCBP  BCCHDTCBN  BCCHLOSSHYST
        '         BCCHDTCBHYST
        'SCTYPE   BSTXPWR  EXTPEN  HYSTSEP  ISHOLEV  MAXISHO  FBOFFSP  FBOFFSN
        CELL = ""
        BSPWR = ""
        BSRXMIN = ""
        BSRXSUFF = ""
        MSRXMIN = ""
        MSRXSUFF = ""
        SCHO = ""
        MISSNM = ""
        AW = ""
        BCCHREUSE = ""
        BCCHLOSS = ""
        BCCHDTCBP = ""
        BCCHDTCBN = ""
        BCCHLOSSHYST = ""
        BCCHDTCBHYST = ""
        SCTYPE = ""
        BSTXPWR = ""
        EXTPEN = ""
        HYSTSEP = ""
        ISHOLEV = ""
        MAXISHO = ""
        FBOFFSP = ""
        FBOFFSN = ""

        filepath = outPath & "RLLOP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "BSPWR" & "," & "BSRXMIN" & "," & "BSRXSUFF" & "," & "MSRXMIN" & "," & "MSRXSUFF" & "," & "SCHO" & "," & "MISSNM" & "," & "AW" _
                                 & "," & "BCCHREUSE" & "," & "BCCHLOSS" & "," & "BCCHDTCBP" & "," & "BCCHDTCBN" & "," & "BCCHLOSSHYST" _
                                 & "," & "BCCHDTCBHYST" _
                                 & "," & "SCTYPE" & "," & "BSTXPWR" & "," & "EXTPEN" & "," & "HYSTSEP" & "," & "ISHOLEV" & "," & "MAXISHO" & "," & "FBOFFSP" & "," & "FBOFFSN")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & BSPWR & "," & BSRXMIN & "," & BSRXSUFF & "," & MSRXMIN & "," & MSRXSUFF & "," & SCHO & "," & MISSNM & "," & AW _
                                                 & "," & BCCHREUSE & "," & BCCHLOSS & "," & BCCHDTCBP & "," & BCCHDTCBN & "," & BCCHLOSSHYST _
                                                 & "," & BCCHDTCBHYST _
                                                 & "," & SCTYPE & "," & BSTXPWR & "," & EXTPEN & "," & HYSTSEP & "," & ISHOLEV & "," & MAXISHO & "," & FBOFFSP & "," & FBOFFSN)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & BSPWR & "," & BSRXMIN & "," & BSRXSUFF & "," & MSRXMIN & "," & MSRXSUFF & "," & SCHO & "," & MISSNM & "," & AW _
                                                 & "," & BCCHREUSE & "," & BCCHLOSS & "," & BCCHDTCBP & "," & BCCHDTCBN & "," & BCCHLOSSHYST _
                                                 & "," & BCCHDTCBHYST _
                                                 & "," & SCTYPE & "," & BSTXPWR & "," & EXTPEN & "," & HYSTSEP & "," & ISHOLEV & "," & MAXISHO & "," & FBOFFSP & "," & FBOFFSN)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Strings.left(Trim(templine), 20)) = "CELL     BSPWR  BSRX" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'CELL     BSPWR  BSRXMIN  BSRXSUFF  MSRXMIN  MSRXSUFF  SCHO  MISSNM  AW
                            CELL = Trim(Mid(templine, 1, 9))
                            BSPWR = Trim(Mid(templine, 10, 7))
                            BSRXMIN = Trim(Mid(templine, 17, 9))
                            BSRXSUFF = Trim(Mid(templine, 26, 10))
                            MSRXMIN = Trim(Mid(templine, 36, 9))
                            MSRXSUFF = Trim(Mid(templine, 45, 10))
                            SCHO = Trim(Mid(templine, 55, 6))
                            MISSNM = Trim(Mid(templine, 61, 8))
                            AW = Trim(Mid(templine, 69, 4))
                        End If
                    ElseIf UCase(Trim(templine)) = "BCCHREUSE  BCCHLOSS  BCCHDTCBP  BCCHDTCBN  BCCHLOSSHYST" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            '         BCCHREUSE  BCCHLOSS  BCCHDTCBP  BCCHDTCBN  BCCHLOSSHYST
                            BCCHREUSE = Trim(Mid(templine, 10, 11))
                            BCCHLOSS = Trim(Mid(templine, 21, 10))
                            BCCHDTCBP = Trim(Mid(templine, 31, 11))
                            BCCHDTCBN = Trim(Mid(templine, 42, 11))
                            BCCHLOSSHYST = Trim(Mid(templine, 53, 12))
                        End If
                    ElseIf UCase(Trim(templine)) = "BCCHDTCBHYST" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            '         BCCHDTCBHYST
                            BCCHDTCBHYST = Trim(Mid(templine, 10, 12))

                        End If
                    ElseIf UCase(Trim(templine)) = "SCTYPE   BSTXPWR  EXTPEN  HYSTSEP  ISHOLEV  MAXISHO  FBOFFSP  FBOFFSN" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'SCTYPE   BSTXPWR  EXTPEN  HYSTSEP  ISHOLEV  MAXISHO  FBOFFSP  FBOFFSN
                            SCTYPE = Trim(Mid(templine, 1, 9))
                            BSTXPWR = Trim(Mid(templine, 10, 9))
                            EXTPEN = Trim(Mid(templine, 19, 8))
                            HYSTSEP = Trim(Mid(templine, 27, 9))
                            ISHOLEV = Trim(Mid(templine, 36, 9))
                            MAXISHO = Trim(Mid(templine, 45, 9))
                            FBOFFSP = Trim(Mid(templine, 54, 9))
                            FBOFFSN = Trim(Mid(templine, 63, 9))
                            outputFile.writeline(BSCNAME & "," & CELL & "," & BSPWR & "," & BSRXMIN & "," & BSRXSUFF & "," & MSRXMIN & "," & MSRXSUFF & "," & SCHO & "," & MISSNM & "," & AW _
                                                 & "," & BCCHREUSE & "," & BCCHLOSS & "," & BCCHDTCBP & "," & BCCHDTCBN & "," & BCCHLOSSHYST _
                                                 & "," & BCCHDTCBHYST & "," & SCTYPE & "," & BSTXPWR & "," & EXTPEN & "," & HYSTSEP & "," & ISHOLEV & "," & MAXISHO & "," & FBOFFSP & "," & FBOFFSN)
                            If SCTYPE = "OL" Or SCTYPE = "" Then
                                CELL = ""
                                BSPWR = ""
                                BSRXMIN = ""
                                BSRXSUFF = ""
                                MSRXMIN = ""
                                MSRXSUFF = ""
                                SCHO = ""
                                MISSNM = ""
                                AW = ""
                                BCCHREUSE = ""
                                BCCHLOSS = ""
                                BCCHDTCBP = ""
                                BCCHDTCBN = ""
                                BCCHLOSSHYST = ""
                                BCCHDTCBHYST = ""
                                SCTYPE = ""
                            End If
                            BSTXPWR = ""
                            EXTPEN = ""
                            HYSTSEP = ""
                            ISHOLEV = ""
                            MAXISHO = ""
                            FBOFFSP = ""
                            FBOFFSN = ""
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL LOCATING DATA" And SCTYPE = "UL" Then
                        If Not Trim(templine) = "NONE" Then
                            'SCTYPE   BSTXPWR  EXTPEN  HYSTSEP  ISHOLEV  MAXISHO  FBOFFSP  FBOFFSN
                            SCTYPE = Trim(Mid(templine, 1, 9))
                            BSTXPWR = Trim(Mid(templine, 10, 9))
                            EXTPEN = Trim(Mid(templine, 19, 8))
                            HYSTSEP = Trim(Mid(templine, 27, 9))
                            ISHOLEV = Trim(Mid(templine, 36, 9))
                            MAXISHO = Trim(Mid(templine, 45, 9))
                            FBOFFSP = Trim(Mid(templine, 54, 9))
                            FBOFFSN = Trim(Mid(templine, 63, 9))
                            outputFile.writeline(BSCNAME & "," & CELL & "," & BSPWR & "," & BSRXMIN & "," & BSRXSUFF & "," & MSRXMIN & "," & MSRXSUFF & "," & SCHO & "," & MISSNM & "," & AW _
                                                  & "," & BCCHREUSE & "," & BCCHLOSS & "," & BCCHDTCBP & "," & BCCHDTCBN & "," & BCCHLOSSHYST _
                                                  & "," & BCCHDTCBHYST _
                                                  & "," & SCTYPE & "," & BSTXPWR & "," & EXTPEN & "," & HYSTSEP & "," & ISHOLEV & "," & MAXISHO & "," & FBOFFSP & "," & FBOFFSN)
                            If SCTYPE = "OL" Or SCTYPE = "" Then
                                CELL = ""
                                BSPWR = ""
                                BSRXMIN = ""
                                BSRXSUFF = ""
                                MSRXMIN = ""
                                MSRXSUFF = ""
                                SCHO = ""
                                MISSNM = ""
                                AW = ""
                                BCCHREUSE = ""
                                BCCHLOSS = ""
                                BCCHDTCBP = ""
                                BCCHDTCBN = ""
                                BCCHLOSSHYST = ""
                                BCCHDTCBHYST = ""
                            End If
                            SCTYPE = ""
                            BSTXPWR = ""
                            EXTPEN = ""
                            HYSTSEP = ""
                            ISHOLEV = ""
                            MAXISHO = ""
                            FBOFFSP = ""
                            FBOFFSN = ""
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLLOP_EXT(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, BSPWR As String, BSRXMIN As String, BSRXSUFF As String, MSRXMIN As String, MSRXSUFF As String, SCHO As String, MISSNM As String, AW As String
        Dim BCCHREUSE As String, BCCHLOSS As String, BCCHDTCBP As String, BCCHDTCBN As String, BCCHLOSSHYST As String
        Dim BCCHDTCBHYST As String
        Dim SCTYPE As String, BSTXPWR As String, EXTPEN As String, HYSTSEP As String, ISHOLEV As String, MAXISHO As String, FBOFFSP As String, FBOFFSN As String
        'CELL     BSPWR  BSRXMIN  BSRXSUFF  MSRXMIN  MSRXSUFF  SCHO  MISSNM  AW
        '         BCCHREUSE  BCCHLOSS  BCCHDTCBP  BCCHDTCBN  BCCHLOSSHYST
        '         BCCHDTCBHYST
        'SCTYPE   BSTXPWR  EXTPEN  HYSTSEP  ISHOLEV  MAXISHO  FBOFFSP  FBOFFSN
        CELL = ""
        BSPWR = ""
        BSRXMIN = ""
        BSRXSUFF = ""
        MSRXMIN = ""
        MSRXSUFF = ""
        SCHO = ""
        MISSNM = ""
        AW = ""
        BCCHREUSE = ""
        BCCHLOSS = ""
        BCCHDTCBP = ""
        BCCHDTCBN = ""
        BCCHLOSSHYST = ""
        BCCHDTCBHYST = ""
        SCTYPE = ""
        BSTXPWR = ""
        EXTPEN = ""
        HYSTSEP = ""
        ISHOLEV = ""
        MAXISHO = ""
        FBOFFSP = ""
        FBOFFSN = ""

        filepath = outPath & "RLLOP_EXT_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "BSPWR" & "," & "BSRXMIN" & "," & "BSRXSUFF" & "," & "MSRXMIN" & "," & "MSRXSUFF" & "," & "SCHO" & "," & "MISSNM" & "," & "AW" _
                                 & "," & "BCCHREUSE" & "," & "BCCHLOSS" & "," & "BCCHDTCBP" & "," & "BCCHDTCBN" & "," & "BCCHLOSSHYST" _
                                 & "," & "BCCHDTCBHYST" _
                                 & "," & "SCTYPE" & "," & "BSTXPWR" & "," & "EXTPEN" & "," & "HYSTSEP" & "," & "ISHOLEV" & "," & "MAXISHO" & "," & "FBOFFSP" & "," & "FBOFFSN")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & BSPWR & "," & BSRXMIN & "," & BSRXSUFF & "," & MSRXMIN & "," & MSRXSUFF & "," & SCHO & "," & MISSNM & "," & AW _
                                                 & "," & BCCHREUSE & "," & BCCHLOSS & "," & BCCHDTCBP & "," & BCCHDTCBN & "," & BCCHLOSSHYST _
                                                 & "," & BCCHDTCBHYST _
                                                 & "," & SCTYPE & "," & BSTXPWR & "," & EXTPEN & "," & HYSTSEP & "," & ISHOLEV & "," & MAXISHO & "," & FBOFFSP & "," & FBOFFSN)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & BSPWR & "," & BSRXMIN & "," & BSRXSUFF & "," & MSRXMIN & "," & MSRXSUFF & "," & SCHO & "," & MISSNM & "," & AW _
                                                 & "," & BCCHREUSE & "," & BCCHLOSS & "," & BCCHDTCBP & "," & BCCHDTCBN & "," & BCCHLOSSHYST _
                                                 & "," & BCCHDTCBHYST _
                                                 & "," & SCTYPE & "," & BSTXPWR & "," & EXTPEN & "," & HYSTSEP & "," & ISHOLEV & "," & MAXISHO & "," & FBOFFSP & "," & FBOFFSN)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Strings.left(Trim(templine), 20)) = "CELL     BSPWR  BSRX" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'CELL     BSPWR  BSRXMIN  BSRXSUFF  MSRXMIN  MSRXSUFF  SCHO  MISSNM  AW
                            CELL = Trim(Mid(templine, 1, 9))
                            BSPWR = Trim(Mid(templine, 10, 7))
                            BSRXMIN = Trim(Mid(templine, 17, 9))
                            BSRXSUFF = Trim(Mid(templine, 26, 10))
                            MSRXMIN = Trim(Mid(templine, 36, 9))
                            MSRXSUFF = Trim(Mid(templine, 45, 10))
                            SCHO = Trim(Mid(templine, 55, 6))
                            MISSNM = Trim(Mid(templine, 61, 8))
                            AW = Trim(Mid(templine, 69, 4))
                        End If
                    ElseIf UCase(Trim(templine)) = "BCCHREUSE  BCCHLOSS  BCCHDTCBP  BCCHDTCBN  BCCHLOSSHYST" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            '         BCCHREUSE  BCCHLOSS  BCCHDTCBP  BCCHDTCBN  BCCHLOSSHYST
                            BCCHREUSE = Trim(Mid(templine, 10, 11))
                            BCCHLOSS = Trim(Mid(templine, 21, 10))
                            BCCHDTCBP = Trim(Mid(templine, 31, 11))
                            BCCHDTCBN = Trim(Mid(templine, 42, 11))
                            BCCHLOSSHYST = Trim(Mid(templine, 53, 12))
                        End If
                    ElseIf UCase(Trim(templine)) = "BCCHDTCBHYST" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            '         BCCHDTCBHYST
                            BCCHDTCBHYST = Trim(Mid(templine, 10, 12))

                        End If
                    ElseIf UCase(Trim(templine)) = "SCTYPE   BSTXPWR  EXTPEN  HYSTSEP  ISHOLEV  MAXISHO  FBOFFSP  FBOFFSN" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'SCTYPE   BSTXPWR  EXTPEN  HYSTSEP  ISHOLEV  MAXISHO  FBOFFSP  FBOFFSN
                            SCTYPE = Trim(Mid(templine, 1, 9))
                            BSTXPWR = Trim(Mid(templine, 10, 9))
                            EXTPEN = Trim(Mid(templine, 19, 8))
                            HYSTSEP = Trim(Mid(templine, 27, 9))
                            ISHOLEV = Trim(Mid(templine, 36, 9))
                            MAXISHO = Trim(Mid(templine, 45, 9))
                            FBOFFSP = Trim(Mid(templine, 54, 9))
                            FBOFFSN = Trim(Mid(templine, 63, 9))
                            outputFile.writeline(BSCNAME & "," & CELL & "," & BSPWR & "," & BSRXMIN & "," & BSRXSUFF & "," & MSRXMIN & "," & MSRXSUFF & "," & SCHO & "," & MISSNM & "," & AW _
                                                 & "," & BCCHREUSE & "," & BCCHLOSS & "," & BCCHDTCBP & "," & BCCHDTCBN & "," & BCCHLOSSHYST _
                                                 & "," & BCCHDTCBHYST & "," & SCTYPE & "," & BSTXPWR & "," & EXTPEN & "," & HYSTSEP & "," & ISHOLEV & "," & MAXISHO & "," & FBOFFSP & "," & FBOFFSN)
                            If SCTYPE = "OL" Or SCTYPE = "" Then
                                CELL = ""
                                BSPWR = ""
                                BSRXMIN = ""
                                BSRXSUFF = ""
                                MSRXMIN = ""
                                MSRXSUFF = ""
                                SCHO = ""
                                MISSNM = ""
                                AW = ""
                                BCCHREUSE = ""
                                BCCHLOSS = ""
                                BCCHDTCBP = ""
                                BCCHDTCBN = ""
                                BCCHLOSSHYST = ""
                                BCCHDTCBHYST = ""
                                SCTYPE = ""
                            End If
                            BSTXPWR = ""
                            EXTPEN = ""
                            HYSTSEP = ""
                            ISHOLEV = ""
                            MAXISHO = ""
                            FBOFFSP = ""
                            FBOFFSN = ""
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL LOCATING DATA" And SCTYPE = "UL" Then
                        If Not Trim(templine) = "NONE" Then
                            'SCTYPE   BSTXPWR  EXTPEN  HYSTSEP  ISHOLEV  MAXISHO  FBOFFSP  FBOFFSN
                            SCTYPE = Trim(Mid(templine, 1, 9))
                            BSTXPWR = Trim(Mid(templine, 10, 9))
                            EXTPEN = Trim(Mid(templine, 19, 8))
                            HYSTSEP = Trim(Mid(templine, 27, 9))
                            ISHOLEV = Trim(Mid(templine, 36, 9))
                            MAXISHO = Trim(Mid(templine, 45, 9))
                            FBOFFSP = Trim(Mid(templine, 54, 9))
                            FBOFFSN = Trim(Mid(templine, 63, 9))
                            outputFile.writeline(BSCNAME & "," & CELL & "," & BSPWR & "," & BSRXMIN & "," & BSRXSUFF & "," & MSRXMIN & "," & MSRXSUFF & "," & SCHO & "," & MISSNM & "," & AW _
                                                  & "," & BCCHREUSE & "," & BCCHLOSS & "," & BCCHDTCBP & "," & BCCHDTCBN & "," & BCCHLOSSHYST _
                                                  & "," & BCCHDTCBHYST _
                                                  & "," & SCTYPE & "," & BSTXPWR & "," & EXTPEN & "," & HYSTSEP & "," & ISHOLEV & "," & MAXISHO & "," & FBOFFSP & "," & FBOFFSN)
                            If SCTYPE = "OL" Or SCTYPE = "" Then
                                CELL = ""
                                BSPWR = ""
                                BSRXMIN = ""
                                BSRXSUFF = ""
                                MSRXMIN = ""
                                MSRXSUFF = ""
                                SCHO = ""
                                MISSNM = ""
                                AW = ""
                                BCCHREUSE = ""
                                BCCHLOSS = ""
                                BCCHDTCBP = ""
                                BCCHDTCBN = ""
                                BCCHLOSSHYST = ""
                                BCCHDTCBHYST = ""
                            End If
                            SCTYPE = ""
                            BSTXPWR = ""
                            EXTPEN = ""
                            HYSTSEP = ""
                            ISHOLEV = ""
                            MAXISHO = ""
                            FBOFFSP = ""
                            FBOFFSN = ""
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLLPP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'CELL     PTIMHF PTIMBQ PTIMTA PSSHF PSSBQ PSSTA
        filepath = outPath & "RLLPP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "PTIMHF" & "," & "PTIMBQ" & "," & "PTIMTA" & "," & "PSSHF" & "," & "PSSBQ" & "," & "PSSTA")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     PTIMHF PTIMBQ PTIMTA PSSHF PSSBQ PSSTA" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 7)) & "," & Trim(Mid(templine, 17, 7)) & "," & Trim(Mid(templine, 24, 7)) & "," & Trim(Mid(templine, 31, 6)) & "," & Trim(Mid(templine, 37, 6)) & "," & Trim(Mid(templine, 42, 5)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL LOCATING PENALTY DATA" And Len(Trim(Mid(templine, 1, 9))) <= 7 Then
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 7)) & "," & Trim(Mid(templine, 17, 7)) & "," & Trim(Mid(templine, 24, 7)) & "," & Trim(Mid(templine, 31, 6)) & "," & Trim(Mid(templine, 37, 6)) & "," & Trim(Mid(templine, 42, 5)))
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLLSP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'LSSTATE
        filepath = outPath & "RLLSP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "LSSTATE")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "LSSTATE" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 9)))
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLLUP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, TALIM As String, CELLQ As String
        Dim SCTYPE As String, QLIMUL As String, QLIMDL As String, QLIMULAFR As String, QLIMDLAFR As String, QLIMULAWB As String, QLIMDLAWB As String
        'CELL     TALIM  CELLQ
        'SCTYPE  QLIMUL  QLIMDL  QLIMULAFR  QLIMDLAFR   QLIMULAWB  QLIMDLAWB
        CELL = ""
        TALIM = ""
        CELLQ = ""
        SCTYPE = ""
        QLIMUL = ""
        QLIMDL = ""
        QLIMULAFR = ""
        QLIMDLAFR = ""
        QLIMULAWB = ""
        QLIMDLAWB = ""
        filepath = outPath & "RLLUP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "TALIM" & "," & "CELLQ" _
                                 & "," & "SCTYPE" & "," & "QLIMUL" & "," & "QLIMDL" & "," & "QLIMULAFR" & "," & "QLIMDLAFR" & "," & "QLIMULAWB" & "," & "QLIMDLAWB")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & TALIM & "," & CELLQ _
                                                 & "," & SCTYPE & "," & QLIMUL & "," & QLIMDL & "," & QLIMULAFR & "," & QLIMDLAFR & "," & QLIMULAWB & "," & QLIMDLAWB)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & TALIM & "," & CELLQ _
                                                 & "," & SCTYPE & "," & QLIMUL & "," & QLIMDL & "," & QLIMULAFR & "," & QLIMDLAFR & "," & QLIMULAWB & "," & QLIMDLAWB)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     TALIM  CELLQ" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'CELL     TALIM  CELLQ
                            CELL = Trim(Mid(templine, 1, 9))
                            TALIM = Trim(Mid(templine, 10, 7))
                            CELLQ = Trim(Mid(templine, 17, 5))
                        End If
                    ElseIf UCase(Trim(templine)) = "SCTYPE  QLIMUL  QLIMDL  QLIMULAFR  QLIMDLAFR   QLIMULAWB  QLIMDLAWB" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'SCTYPE  QLIMUL  QLIMDL  QLIMULAFR  QLIMDLAFR   QLIMULAWB  QLIMDLAWB
                            SCTYPE = Trim(Mid(templine, 1, 8))
                            QLIMUL = Trim(Mid(templine, 9, 8))
                            QLIMDL = Trim(Mid(templine, 17, 8))
                            QLIMULAFR = Trim(Mid(templine, 25, 11))
                            QLIMDLAFR = Trim(Mid(templine, 36, 11))
                            QLIMULAWB = Trim(Mid(templine, 47, 11))
                            QLIMDLAWB = Trim(Mid(templine, 58, 11))
                            outputFile.writeline(BSCNAME & "," & CELL & "," & TALIM & "," & CELLQ _
                                                 & "," & SCTYPE & "," & QLIMUL & "," & QLIMDL & "," & QLIMULAFR & "," & QLIMDLAFR & "," & QLIMULAWB & "," & QLIMDLAWB)
                            If SCTYPE = "OL" Or SCTYPE = "" Then
                                CELL = ""
                                TALIM = ""
                                CELLQ = ""
                            End If
                            SCTYPE = ""
                            QLIMUL = ""
                            QLIMDL = ""
                            QLIMULAFR = ""
                            QLIMDLAFR = ""
                            QLIMULAWB = ""
                            QLIMDLAWB = ""
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL LOCATING URGENCY DATA" And UCase(Trim(Mid(templine, 1, 8))) = "OL" Then
                        If Not Trim(templine) = "NONE" Then
                            'SCTYPE  QLIMUL  QLIMDL  QLIMULAFR  QLIMDLAFR   QLIMULAWB  QLIMDLAWB
                            SCTYPE = Trim(Mid(templine, 1, 8))
                            QLIMUL = Trim(Mid(templine, 9, 8))
                            QLIMDL = Trim(Mid(templine, 17, 8))
                            QLIMULAFR = Trim(Mid(templine, 25, 11))
                            QLIMDLAFR = Trim(Mid(templine, 36, 11))
                            QLIMULAWB = Trim(Mid(templine, 47, 11))
                            QLIMDLAWB = Trim(Mid(templine, 58, 11))
                            outputFile.writeline(BSCNAME & "," & CELL & "," & TALIM & "," & CELLQ _
                                                 & "," & SCTYPE & "," & QLIMUL & "," & QLIMDL & "," & QLIMULAFR & "," & QLIMDLAFR & "," & QLIMULAWB & "," & QLIMDLAWB)
                            If SCTYPE = "OL" Or SCTYPE = "" Then
                                CELL = ""
                                TALIM = ""
                                CELLQ = ""
                            End If
                            SCTYPE = ""
                            QLIMUL = ""
                            QLIMDL = ""
                            QLIMULAFR = ""
                            QLIMDLAFR = ""
                            QLIMULAWB = ""
                            QLIMDLAWB = ""
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLMAP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim DTXFUL As String, REGINTUL As String, SSLENUL As String, QLENUL As String
        DTXFUL = ""
        REGINTUL = ""
        SSLENUL = ""
        QLENUL = ""

        'DTXFUL  REGINTUL  SSLENUL  QLENUL

        filepath = outPath & "RLMAP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "DTXFUL" & "," & "REGINTUL" & "," & "SSLENUL" & "," & "QLENUL")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "DTXFUL  REGINTUL  SSLENUL  QLENUL" Then
                        'DTXFUL  REGINTUL  SSLENUL  QLENUL

                        templine = getNowLine()
                        DTXFUL = Trim(Mid(templine, 1, 8))
                        REGINTUL = Trim(Mid(templine, 9, 10))
                        SSLENUL = Trim(Mid(templine, 19, 9))
                        QLENUL = Trim(Mid(templine, 28, 7))

                        outputFile.writeLINE(BSCNAME & "," & DTXFUL & "," & REGINTUL & "," & SSLENUL & "," & QLENUL)
                    ElseIf UCase(Trim(templine)) <> "DYNAMIC MS POWER CONTROL ADDITIONAL BSC DATA" And Len(Trim(Mid(templine, 1, 8))) < 8 And UCase(Trim(templine)) <> "DTXFUL  REGINTUL  SSLENUL  QLENUL" Then
                        DTXFUL = Trim(Mid(templine, 1, 8))
                        REGINTUL = Trim(Mid(templine, 9, 10))
                        SSLENUL = Trim(Mid(templine, 19, 9))
                        QLENUL = Trim(Mid(templine, 28, 7))

                        outputFile.writeLINE(BSCNAME & "," & DTXFUL & "," & REGINTUL & "," & SSLENUL & "," & QLENUL)
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLMFP(sender As Object)
        Dim filepath As String
        Dim templine As String
        Dim outputFile As Object
        Dim CELL As String, MF As String, LISTTYPE As String

        CELL = ""
        MF = ""
        LISTTYPE = ""
        filepath = outPath & "RLMFP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "LISTTYPE" & "," & "MF1" & "," & "MF2" & "," & "MF3" & "," & "MF4" & "," & "MF5" & "," & "MF6" & "," & "MF7" & "," & "MF8" & "," & "MF9" & "," & "MF10" & "," & "MF11" & "," & "MF12" & "," & "MF13" & "," & "MF14" & "," & "MF15" & "," & "MF16" & "," & "MF17" & "," & "MF18" & "," & "MF19" & "," & "MF20" & "," & "MF21" & "," & "MF22" & "," & "MF23" & "," & "MF24" & "," & "MF25" & "," & "MF26" & "," & "MF27" & "," & "MF28" & "," & "MF29" & "," & "MF30" & "," & "MF31" & "," & "MF32")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & LISTTYPE & "," & MF)
                        End If

                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & LISTTYPE & "," & MF)
                        End If

                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL" Then
                        templine = getNowLine()
                        If Trim(templine) = "NONE" Then
                            outputFile.CLOSE()
                            outputFile = Nothing
                            Exit Sub
                        End If
                        CELL = templine
                        MF = ""
                    ElseIf UCase(Trim(templine)) = "LISTTYPE" Then
                        templine = getNowLine()
                        If UCase(Trim(templine)) = "IDLE" Then
                            LISTTYPE = "IDLE"
                        ElseIf UCase(Trim(templine)) = "ACTIVE" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & LISTTYPE & "," & MF)
                            MF = ""
                            LISTTYPE = "ACTIVE"
                        End If
                    ElseIf UCase(Trim(templine)) <> "TEST MEASUREMENT FREQUENCIES" And Len(Trim(Mid(templine, 1, 5))) <= 4 And UCase(Trim(templine)) <> "ACTIVE" And UCase(Trim(templine)) <> "LISTTYPE" And UCase(Trim(templine)) <> "IDLE" And UCase(Trim(templine)) <> "MBCCHNO" And UCase(Trim(templine)) <> "CELL MEASUREMENT FREQUENCIES" Then
                        If MF = "" Then
                            MF = Trim(Mid(templine, 1, 5)) & "," & Trim(Mid(templine, 6, 5)) & "," & Trim(Mid(templine, 11, 5)) & "," & Trim(Mid(templine, 16, 5)) & "," & Trim(Mid(templine, 21, 5)) & "," & Trim(Mid(templine, 26, 5)) & "," & Trim(Mid(templine, 31, 5)) & "," & Trim(Mid(templine, 36, 5)) & "," & Trim(Mid(templine, 41, 5)) & "," & Trim(Mid(templine, 46, 5)) & "," & Trim(Mid(templine, 51, 5)) & "," & Trim(Mid(templine, 56, 5)) & "," & Trim(Mid(templine, 61, 5)) & "," & Trim(Mid(templine, 66, 5))
                        Else
                            MF = MF & "," & Trim(Mid(templine, 1, 5)) & "," & Trim(Mid(templine, 6, 5)) & "," & Trim(Mid(templine, 11, 5)) & "," & Trim(Mid(templine, 16, 5)) & "," & Trim(Mid(templine, 21, 5)) & "," & Trim(Mid(templine, 26, 5)) & "," & Trim(Mid(templine, 31, 5)) & "," & Trim(Mid(templine, 36, 5)) & "," & Trim(Mid(templine, 41, 5)) & "," & Trim(Mid(templine, 46, 5)) & "," & Trim(Mid(templine, 51, 5)) & "," & Trim(Mid(templine, 56, 5)) & "," & Trim(Mid(templine, 61, 5)) & "," & Trim(Mid(templine, 66, 5))
                        End If
                    ElseIf UCase(Trim(templine)) = "TEST MEASUREMENT FREQUENCIES" Then
                        outputFile.writeline(BSCNAME & "," & CELL & "," & LISTTYPE & "," & MF)
                        MF = ""
                        CELL = ""
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLNMP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, NCSTAT As String, NCRPT As String
        'CELL      NCSTAT  NCRPT
        CELL = ""
        NCSTAT = ""
        NCRPT = ""

        filepath = outPath & "RLNMP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "NCSTAT" & "," & "NCRPT")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL      NCSTAT  NCRPT" Then
                        templine = getNowLine()
                        CELL = Trim(Mid(templine, 1, 10))
                        NCSTAT = Trim(Mid(templine, 11, 8))
                        NCRPT = Trim(Mid(templine, 19, 6))
                        outputFile.writeLINE(BSCNAME & "," & CELL & "," & NCSTAT & "," & NCRPT)
                    ElseIf UCase(Trim(templine)) <> "CELL NETWORK CONTROLLED CELL RESELECTION MODE DATA" And Len(Trim(Mid(templine, 1, 9))) < 9 And UCase(Trim(templine)) <> "CELL      NCSTAT  NCRPT" Then
                        CELL = Trim(Mid(templine, 1, 10))
                        NCSTAT = Trim(Mid(templine, 11, 8))
                        NCRPT = Trim(Mid(templine, 19, 6))
                        outputFile.writeLINE(BSCNAME & "," & CELL & "," & NCSTAT & "," & NCRPT)
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLNPP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, NCPROF As String
        'CELL      NCPROF
        CELL = ""
        NCPROF = ""

        filepath = outPath & "RLNPP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "NCPROF")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL      NCPROF" Then
                        templine = getNowLine()
                        CELL = Trim(Mid(templine, 1, 10))
                        NCPROF = Trim(Mid(templine, 11, 15))
                        outputFile.writeLINE(BSCNAME & "," & CELL & "," & NCPROF)
                    ElseIf UCase(Trim(templine)) <> "CELL NETWORK CONTROLLED CELL RESELECTION PROFILE DATA" And Len(Trim(Mid(templine, 1, 9))) < 9 And UCase(Trim(templine)) <> "CELL      NCPROF" Then
                        CELL = Trim(Mid(templine, 1, 10))
                        CELL = Trim(Mid(templine, 1, 10))
                        NCPROF = Trim(Mid(templine, 11, 15))
                        outputFile.writeLINE(BSCNAME & "," & CELL & "," & NCPROF)
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLNRP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String
        Dim CELLR As String, DIR As String, CAND As String, CS As String
        Dim KHYST As String, KOFFSETP As String, KOFFSETN As String, LHYST As String, LOFFSETP As String, LOFFSETN As String
        Dim TRHYST As String, TROFFSETP As String, TROFFSETN As String, AWOFFSET As String, BQOFFSET As String
        Dim HIHYST As String, LOHYST As String, OFFSETP As String, OFFSETN As String, BQOFFSETAFR As String, BQOFFSETAWB As String
        Dim Ncell_temp As String, Scell_temp As String
        Ncell_temp = ""
        Scell_temp = ""
        CELLR = ""
        DIR = ""
        CAND = ""
        CS = ""
        KHYST = ""
        KOFFSETP = ""
        KOFFSETN = ""
        LHYST = ""
        LOFFSETN = ""
        LOFFSETP = ""
        TRHYST = ""
        TROFFSETN = ""
        TROFFSETP = ""
        AWOFFSET = ""
        BQOFFSET = ""
        HIHYST = ""
        LOHYST = ""
        OFFSETP = ""
        OFFSETN = ""
        BQOFFSETAFR = ""
        BQOFFSETAWB = ""
        filepath = outPath & "RLNRP_" & CDDtime & ".csv"

        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "CELLR" & "," & "DIR" & "," & "CAND" & "," & "CS" & "," & "KHYST" & "," & "KOFFSETP" & "," & "KOFFSETN" & "," & "LHYST" & "," & "LOFFSETP" & "," & "LOFFSETN" & "," & "TRHYST" & "," & "TROFFSETP" & "," & "TROFFSETN" & "," & "AWOFFSET" & "," & "BQOFFSET" & "," & "HIHYST" & "," & "LOHYST" & "," & "OFFSETP" & "," & "OFFSETN" & "," & "BQOFFSETAFR" & "," & "BQOFFSETAWB")
        End If
        CELL = ""

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        Try
                            outputFile.writeline(BSCNAME & "," & CELL & "," & CELLR & "," & DIR & "," & CAND & "," & CS _
                                             & "," & KHYST & "," & KOFFSETP & "," & KOFFSETN & "," & LHYST & "," & LOFFSETP & "," & LOFFSETN _
                                             & "," & TRHYST & "," & TROFFSETP & "," & TROFFSETN & "," & AWOFFSET & "," & BQOFFSET _
                                             & "," & HIHYST & "," & LOHYST & "," & OFFSETP & "," & OFFSETN & "," & BQOFFSETAFR & "," & BQOFFSETAWB)
                            outputFile.close()
                            outputFile = Nothing
                            If InStr(Trim(Ncell_temp), " ") Then
                                If UBound(Split(Trim(Ncell_temp), " ")) + 1 > MaxNeighborNo Then
                                    MaxNeighborNo = UBound(Split(Trim(Ncell_temp), " ")) + 1
                                End If
                            End If
                            S_cell_Neighbour.Add(Scell_temp)
                            N_cell_Neighbour.Add(Ncell_temp)
                            Exit Sub
                        Catch ex As Exception
                        Finally
                            doWhichMethod(templine, sender)
                        End Try
                    Else
                        If InStr(Trim(Ncell_temp), " ") Then
                            If UBound(Split(Trim(Ncell_temp), " ")) + 1 > MaxNeighborNo Then
                                MaxNeighborNo = UBound(Split(Trim(Ncell_temp), " ")) + 1
                            End If
                        End If
                        S_cell_Neighbour.Add(Scell_temp)
                        N_cell_Neighbour.Add(Ncell_temp)
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            CELL = Trim(templine)
                            If Scell_temp = "" Then
                                Scell_temp = CELL
                            End If
                        End If
                    ElseIf Trim(templine) = "CELLR     DIR     CAND   CS" Then
                        templine = getNowLine()
                        'CELLR     DIR     CAND   CS
                        If Trim(templine) <> "NONE" Then
                            CELLR = Trim(Mid(templine, 1, 10))
                            DIR = Trim(Mid(templine, 11, 8))
                            CAND = Trim(Mid(templine, 19, 7))
                            CS = Trim(Mid(templine, 26, 3))
                            If Scell_temp <> CELL Then
                                S_cell_Neighbour.Add(Scell_temp)
                                N_cell_Neighbour.Add(Ncell_temp)
                                If InStr(Trim(Ncell_temp), " ") Then
                                    If UBound(Split(Trim(Ncell_temp), " ")) + 1 > MaxNeighborNo Then
                                        MaxNeighborNo = UBound(Split(Trim(Ncell_temp), " ")) + 1
                                    End If
                                End If
                                Scell_temp = CELL
                                Ncell_temp = CELLR
                            Else
                                Ncell_temp = Trim(Ncell_temp & " " & CELLR)
                            End If
                        Else
                            outputFile.writeline(BSCNAME & "," & CELL & "," & "NONE")
                        End If
                    ElseIf Trim(templine) = "KHYST   KOFFSETP  KOFFSETN   LHYST   LOFFSETP  LOFFSETN" Then
                        templine = getNowLine()
                        'KHYST   KOFFSETP  KOFFSETN   LHYST   LOFFSETP  LOFFSETN
                        KHYST = Trim(Mid(templine, 1, 8))
                        KOFFSETP = Trim(Mid(templine, 9, 10))
                        KOFFSETN = Trim(Mid(templine, 19, 11))
                        LHYST = Trim(Mid(templine, 30, 8))
                        LOFFSETP = Trim(Mid(templine, 38, 10))
                        LOFFSETN = Trim(Mid(templine, 48, 8))
                    ElseIf Trim(templine) = "TRHYST  TROFFSETP  TROFFSETN  AWOFFSET  BQOFFSET" Then
                        templine = getNowLine()
                        'TRHYST  TROFFSETP  TROFFSETN  AWOFFSET  BQOFFSET
                        TRHYST = Trim(Mid(templine, 1, 8))
                        TROFFSETP = Trim(Mid(templine, 9, 11))
                        TROFFSETN = Trim(Mid(templine, 20, 11))
                        AWOFFSET = Trim(Mid(templine, 31, 10))
                        BQOFFSET = Trim(Mid(templine, 41, 8))
                    ElseIf Trim(templine) = "HIHYST  LOHYST  OFFSETP  OFFSETN  BQOFFSETAFR  BQOFFSETAWB" Then
                        templine = getNowLine()
                        'HIHYST  LOHYST  OFFSETP  OFFSETN  BQOFFSETAFR  BQOFFSETAWB
                        HIHYST = Trim(Mid(templine, 1, 8))
                        LOHYST = Trim(Mid(templine, 9, 8))
                        OFFSETP = Trim(Mid(templine, 17, 9))
                        OFFSETN = Trim(Mid(templine, 26, 9))
                        BQOFFSETAFR = Trim(Mid(templine, 35, 13))
                        BQOFFSETAWB = Trim(Mid(templine, 48, 11))
                        outputFile.writeline(BSCNAME & "," & CELL & "," & CELLR & "," & DIR & "," & CAND & "," & CS _
                                         & "," & KHYST & "," & KOFFSETP & "," & KOFFSETN & "," & LHYST & "," & LOFFSETP & "," & LOFFSETN _
                                         & "," & TRHYST & "," & TROFFSETP & "," & TROFFSETN & "," & AWOFFSET & "," & BQOFFSET _
                                         & "," & HIHYST & "," & LOHYST & "," & OFFSETP & "," & OFFSETN & "," & BQOFFSETAFR & "," & BQOFFSETAWB)
                        CELLR = ""
                        DIR = ""
                        CAND = ""
                        CS = ""
                        KHYST = ""
                        KOFFSETP = ""
                        KOFFSETN = ""
                        LHYST = ""
                        LOFFSETN = ""
                        LOFFSETP = ""
                        TRHYST = ""
                        TROFFSETN = ""
                        TROFFSETP = ""
                        AWOFFSET = ""
                        BQOFFSET = ""
                        HIHYST = ""
                        LOHYST = ""
                        OFFSETP = ""
                        OFFSETN = ""
                        BQOFFSETAFR = ""
                        BQOFFSETAWB = ""
                        'Ncell_temp = ""
                        'Scell_temp = ""
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLNRP_NODATA(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String
        Dim CELLR As String, DIR As String, CAND As String, CS As String
        CELLR = ""
        DIR = ""
        CAND = ""
        CS = ""
        filepath = outPath & "RLNRP_NODATA_" & CDDTime & ".csv"

        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "CELLR" & "," & "DIR" & "," & "CAND" & "," & "CS")
        End If
        CELL = ""

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        Try
                            If CELLR <> "" Then
                                outputFile.writeline(BSCNAME & "," & CELL & "," & CELLR & "," & DIR & "," & CAND & "," & CS)
                            End If
                            outputFile.close()
                            outputFile = Nothing
                        Catch ex As Exception
                        Finally
                            doWhichMethod(templine, sender)
                        End Try
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            CELL = Trim(templine)
                        End If
                    ElseIf Trim(templine) = "CELLR     DIR     CAND   CS" Then
                        templine = getNowLine()
                        'CELLR     DIR     CAND   CS
                        If Trim(templine) <> "NONE" Then
                            CELLR = Trim(Mid(templine, 1, 10))
                            DIR = Trim(Mid(templine, 11, 8))
                            CAND = Trim(Mid(templine, 19, 7))
                            CS = Trim(Mid(templine, 26, 3))
                            outputFile.writeline(BSCNAME & "," & CELL & "," & CELLR & "," & DIR & "," & CAND & "," & CS)
                            CELLR = ""
                            DIR = ""
                            CAND = ""
                            CS = ""
                        Else
                            outputFile.writeline(BSCNAME & "," & CELL & "," & "NONE")
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLNRP_UTRAN(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String
        Dim CELLR As String, DIR As String, CAND As String, CS As String
        Dim KHYST As String, KOFFSETP As String, KOFFSETN As String, LHYST As String, LOFFSETP As String, LOFFSETN As String
        Dim TRHYST As String, TROFFSETP As String, TROFFSETN As String, AWOFFSET As String, BQOFFSET As String
        Dim HIHYST As String, LOHYST As String, OFFSETP As String, OFFSETN As String, BQOFFSETAFR As String, BQOFFSETAWB As String
        Dim Ncell_temp As String, Scell_temp As String
        Ncell_temp = ""
        Scell_temp = ""
        CELLR = ""
        DIR = ""
        CAND = ""
        CS = ""
        KHYST = ""
        KOFFSETP = ""
        KOFFSETN = ""
        LHYST = ""
        LOFFSETN = ""
        LOFFSETP = ""
        TRHYST = ""
        TROFFSETN = ""
        TROFFSETP = ""
        AWOFFSET = ""
        BQOFFSET = ""
        HIHYST = ""
        LOHYST = ""
        OFFSETP = ""
        OFFSETN = ""
        BQOFFSETAFR = ""
        BQOFFSETAWB = ""
        filepath = outPath & "RLNRP_UTRAN_" & CDDTime & ".csv"

        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "CELLR" & "," & "DIR")
        End If
        CELL = ""

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        Try
                            If CELL <> "" And CELLR <> "" Then
                                outputFile.writeline(BSCNAME & "," & CELL & "," & CELLR & "," & DIR)
                                CELL = ""
                                CELLR = ""
                                DIR = ""
                            End If
                            outputFile.close()
                            outputFile = Nothing
                            'If InStr(Trim(Ncell_temp), " ") Then
                            '    If UBound(Split(Trim(Ncell_temp), " ")) + 1 > MaxNeighborNo Then
                            '        MaxNeighborNo = UBound(Split(Trim(Ncell_temp), " ")) + 1
                            '    End If
                            'End If
                            'S_cell_Neighbour.Add(Scell_temp)
                            'N_cell_Neighbour.Add(Ncell_temp)
                            Exit Sub
                        Catch ex As Exception

                        Finally
                            doWhichMethod(templine, sender)
                        End Try
                    Else
                        'If InStr(Trim(Ncell_temp), " ") Then
                        '    If UBound(Split(Trim(Ncell_temp), " ")) + 1 > MaxNeighborNo Then
                        '        MaxNeighborNo = UBound(Split(Trim(Ncell_temp), " ")) + 1
                        '    End If
                        'End If
                        'S_cell_Neighbour.Add(Scell_temp)
                        'N_cell_Neighbour.Add(Ncell_temp)
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL" Then
                        templine = getNowLine()
                        If CELL <> "" And CELLR <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & CELLR & "," & DIR)
                            CELL = ""
                            CELLR = ""
                            DIR = ""
                        End If
                        If Trim(templine) <> "NONE" Then
                            CELL = Trim(templine)
                            If Scell_temp = "" Then
                                Scell_temp = CELL
                            End If
                        End If
                    ElseIf Trim(templine) = "CELLR     DIR" Then
                        templine = getNowLine()
                        'CELLR     DIR     CAND   CS
                        If Trim(templine) <> "NONE" Then
                            CELLR = Trim(Mid(templine, 1, 10))
                            DIR = Trim(Mid(templine, 11, 8))
                            'If Scell_temp <> CELL Then
                            '    S_cell_Neighbour.Add(Scell_temp)
                            '    N_cell_Neighbour.Add(Ncell_temp)
                            '    If InStr(Trim(Ncell_temp), " ") Then
                            '        If UBound(Split(Trim(Ncell_temp), " ")) + 1 > MaxNeighborNo Then
                            '            MaxNeighborNo = UBound(Split(Trim(Ncell_temp), " ")) + 1
                            '        End If
                            '    End If
                            '    Scell_temp = CELL
                            '    Ncell_temp = CELLR
                            'Else
                            '    Ncell_temp = Trim(Ncell_temp & " " & CELLR)
                            'End If
                            outputFile.writeline(BSCNAME & "," & CELL & "," & CELLR & "," & DIR )
                            CELLR = ""
                            DIR = ""
                        Else
                            outputFile.writeline(BSCNAME & "," & CELL & "," & "NONE")
                        End If
                    ElseIf Trim(templine) = "KHYST   KOFFSETP  KOFFSETN   LHYST   LOFFSETP  LOFFSETN" Then
                        templine = getNowLine()
                        'KHYST   KOFFSETP  KOFFSETN   LHYST   LOFFSETP  LOFFSETN
                        KHYST = Trim(Mid(templine, 1, 8))
                        KOFFSETP = Trim(Mid(templine, 9, 10))
                        KOFFSETN = Trim(Mid(templine, 19, 11))
                        LHYST = Trim(Mid(templine, 30, 8))
                        LOFFSETP = Trim(Mid(templine, 38, 10))
                        LOFFSETN = Trim(Mid(templine, 48, 8))
                    ElseIf Trim(templine) = "TRHYST  TROFFSETP  TROFFSETN  AWOFFSET  BQOFFSET" Then
                        templine = getNowLine()
                        'TRHYST  TROFFSETP  TROFFSETN  AWOFFSET  BQOFFSET
                        TRHYST = Trim(Mid(templine, 1, 8))
                        TROFFSETP = Trim(Mid(templine, 9, 11))
                        TROFFSETN = Trim(Mid(templine, 20, 11))
                        AWOFFSET = Trim(Mid(templine, 31, 10))
                        BQOFFSET = Trim(Mid(templine, 41, 8))
                    ElseIf Trim(templine) = "HIHYST  LOHYST  OFFSETP  OFFSETN  BQOFFSETAFR  BQOFFSETAWB" Then
                        templine = getNowLine()
                        'HIHYST  LOHYST  OFFSETP  OFFSETN  BQOFFSETAFR  BQOFFSETAWB
                        HIHYST = Trim(Mid(templine, 1, 8))
                        LOHYST = Trim(Mid(templine, 9, 8))
                        OFFSETP = Trim(Mid(templine, 17, 9))
                        OFFSETN = Trim(Mid(templine, 26, 9))
                        BQOFFSETAFR = Trim(Mid(templine, 35, 13))
                        BQOFFSETAWB = Trim(Mid(templine, 48, 11))
                        outputFile.writeline(BSCNAME & "," & CELL & "," & CELLR & "," & DIR & "," & CAND & "," & CS _
                                         & "," & KHYST & "," & KOFFSETP & "," & KOFFSETN & "," & LHYST & "," & LOFFSETP & "," & LOFFSETN _
                                         & "," & TRHYST & "," & TROFFSETP & "," & TROFFSETN & "," & AWOFFSET & "," & BQOFFSET _
                                         & "," & HIHYST & "," & LOHYST & "," & OFFSETP & "," & OFFSETN & "," & BQOFFSETAFR & "," & BQOFFSETAWB)
                        CELLR = ""
                        DIR = ""
                        CAND = ""
                        CS = ""
                        KHYST = ""
                        KOFFSETP = ""
                        KOFFSETN = ""
                        LHYST = ""
                        LOFFSETN = ""
                        LOFFSETP = ""
                        TRHYST = ""
                        TROFFSETN = ""
                        TROFFSETP = ""
                        AWOFFSET = ""
                        BQOFFSET = ""
                        HIHYST = ""
                        LOHYST = ""
                        OFFSETP = ""
                        OFFSETN = ""
                        BQOFFSETAFR = ""
                        BQOFFSETAWB = ""
                        'Ncell_temp = ""
                        'Scell_temp = ""
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLNCP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String
        CELL = ""

        filepath = outPath & "RLNCP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "CELLR(IS NEIGHBOUR TO)")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 10))
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL" And UCase(Trim(templine)) <> "IS NEIGHBOUR TO" And UCase(Trim(templine)) <> "NEIGHBOUR TO CELLS" Then
                        If Not Trim(templine) = "NONE" Then
                            For x = 1 To 64 Step 9
                                If Trim(Mid(templine, x, 9)) <> "" Then
                                    outputFile.writeline(BSCNAME & "," & CELL & "," & Trim(Mid(templine, x, 9)))
                                End If
                            Next
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLOLP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, LOL As String, LOLHYST As String, TAOL As String, TAOLHYST As String
        Dim DTCBP As String, DTCBN As String, DTCBHYST As String, NDIST As String, NNCELLS As String
        'CELL     LOL  LOLHYST  TAOL  TAOLHYST
        'DTCBP  DTCBN  DTCBHYST  NDIST  NNCELLS
        CELL = ""
        LOL = ""
        LOLHYST = ""
        TAOL = ""
        TAOLHYST = ""
        DTCBP = ""
        DTCBN = ""
        DTCBHYST = ""
        NDIST = ""
        NNCELLS = ""

        filepath = outPath & "RLOLP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "LOL" & "," & "LOLHYST" & "," & "TAOL" & "," & "TAOLHYST" _
                                 & "," & "DTCBP" & "," & "DTCBN" & "," & "DTCBHYST" & "," & "NDIST" & "," & "NNCELLS")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & LOL & "," & LOLHYST & "," & TAOL & "," & TAOLHYST _
                                                 & "," & DTCBP & "," & DTCBN & "," & DTCBHYST & "," & NDIST & "," & NNCELLS)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & LOL & "," & LOLHYST & "," & TAOL & "," & TAOLHYST _
                                                 & "," & DTCBP & "," & DTCBN & "," & DTCBHYST & "," & NDIST & "," & NNCELLS)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     LOL  LOLHYST  TAOL  TAOLHYST" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'CELL     LOL  LOLHYST  TAOL  TAOLHYST
                            CELL = Trim(Mid(templine, 1, 9))
                            LOL = Trim(Mid(templine, 10, 5))
                            LOLHYST = Trim(Mid(templine, 15, 9))
                            TAOL = Trim(Mid(templine, 24, 6))
                            TAOLHYST = Trim(Mid(templine, 30, 8))
                        End If
                    ElseIf UCase(Trim(templine)) = "DTCBP  DTCBN  DTCBHYST  NDIST  NNCELLS" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'DTCBP  DTCBN  DTCBHYST  NDIST  NNCELLS
                            DTCBP = Trim(Mid(templine, 1, 7))
                            DTCBN = Trim(Mid(templine, 8, 7))
                            DTCBHYST = Trim(Mid(templine, 15, 10))
                            NDIST = Trim(Mid(templine, 25, 7))
                            NNCELLS = Trim(Mid(templine, 32, 7))
                            outputFile.writeline(BSCNAME & "," & CELL & "," & LOL & "," & LOLHYST & "," & TAOL & "," & TAOLHYST _
                                                 & "," & DTCBP & "," & DTCBN & "," & DTCBHYST & "," & NDIST & "," & NNCELLS)
                            CELL = ""
                            LOL = ""
                            LOLHYST = ""
                            TAOL = ""
                            TAOLHYST = ""
                            DTCBP = ""
                            DTCBN = ""
                            DTCBHYST = ""
                            NDIST = ""
                            NNCELLS = ""
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLOMP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'MODE
        filepath = outPath & "RLOMP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "MODE")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "MODE" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 9)))
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLPBP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'CELL     HPBSTATE
        filepath = outPath & "RLPBP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "HPBSTATE")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     HPBSTATE" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 10)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "HANDOVER POWER BOOST DATA" And Len(Trim(Mid(templine, 1, 9))) <= 7 Then
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 9)) & "," & Trim(Mid(templine, 10, 10)))
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLPCP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, DMPSTATE As String
        Dim SCTYPE As String, SSDESUL As String, QDESUL As String, LCOMPUL As String, QCOMPUL As String
        'CELL      DMPSTATE
        'SCTYPE    SSDESUL  QDESUL  LCOMPUL  QCOMPUL
        CELL = ""
        DMPSTATE = ""
        SCTYPE = ""
        SSDESUL = ""
        QDESUL = ""
        LCOMPUL = ""
        QCOMPUL = ""

        filepath = outPath & "RLPCP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "DMPSTATE" _
                                 & "," & "SCTYPE" & "," & "SSDESUL" & "," & "QDESUL" & "," & "LCOMPUL" & "," & "QCOMPUL")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & DMPSTATE _
                                                 & "," & SCTYPE & "," & SSDESUL & "," & QDESUL & "," & LCOMPUL & "," & QCOMPUL)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & DMPSTATE _
                                                 & "," & SCTYPE & "," & SSDESUL & "," & QDESUL & "," & LCOMPUL & "," & QCOMPUL)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL      DMPSTATE" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'CELL      DMPSTATE
                            CELL = Trim(Mid(templine, 1, 10))
                            DMPSTATE = Trim(Mid(templine, 11, 8))
                        End If
                    ElseIf UCase(Trim(templine)) = "SCTYPE    SSDESUL  QDESUL  LCOMPUL  QCOMPUL" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'SCTYPE    SSDESUL  QDESUL  LCOMPUL  QCOMPUL
                            SCTYPE = Trim(Mid(templine, 1, 10))
                            SSDESUL = Trim(Mid(templine, 11, 9))
                            QDESUL = Trim(Mid(templine, 20, 8))
                            LCOMPUL = Trim(Mid(templine, 28, 9))
                            QCOMPUL = Trim(Mid(templine, 37, 7))
                            outputFile.writeline(BSCNAME & "," & CELL & "," & DMPSTATE _
                                                 & "," & SCTYPE & "," & SSDESUL & "," & QDESUL & "," & LCOMPUL & "," & QCOMPUL)
                            If SCTYPE = "OL" Or SCTYPE = "" Then
                                CELL = ""
                                DMPSTATE = ""
                            End If
                            SCTYPE = ""
                            SSDESUL = ""
                            QDESUL = ""
                            LCOMPUL = ""
                            QCOMPUL = ""
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLPDP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, CBCHD As String, PRACHBLK As String, MAXSBLK As String, MAXSMSG As String, TRAFBLK As String
        'CELL     CBCHD   PRACHBLK  MAXSBLK   MAXSMSG   TRAFBLK
        CELL = ""
        CBCHD = ""
        PRACHBLK = ""
        MAXSBLK = ""
        MAXSMSG = ""
        TRAFBLK = ""
        filepath = outPath & "RLPDP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "CBCHD" & "," & "PRACHBLK" & "," & "MAXSBLK" & "," & "MAXSMSG" & "," & "TRAFBLK")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     CBCHD   PRACHBLK  MAXSBLK   MAXSMSG   TRAFBLK" Then
                        templine = getNowLine()
                        'CELL     CBCHD   PRACHBLK  MAXSBLK   MAXSMSG   TRAFBLK
                        CELL = Trim(Mid(templine, 1, 9))
                        CBCHD = Trim(Mid(templine, 10, 8))
                        PRACHBLK = Trim(Mid(templine, 18, 10))
                        MAXSBLK = Trim(Mid(templine, 28, 10))
                        MAXSMSG = Trim(Mid(templine, 38, 10))
                        TRAFBLK = Trim(Mid(templine, 48, 8))
                        outputFile.writeLINE(BSCNAME & "," & CELL & "," & CBCHD & "," & PRACHBLK & "," & MAXSBLK & "," & MAXSMSG & "," & TRAFBLK)
                    ElseIf UCase(Trim(templine)) <> "GPRS PACKET CONTROL CHANNEL DATA" And Len(Trim(Mid(templine, 1, 9))) < 9 And UCase(Trim(templine)) <> "CELL     CBCHD   PRACHBLK  MAXSBLK   MAXSMSG   TRAFBLK" Then
                        CELL = Trim(Mid(templine, 1, 9))
                        CBCHD = Trim(Mid(templine, 10, 8))
                        PRACHBLK = Trim(Mid(templine, 18, 10))
                        MAXSBLK = Trim(Mid(templine, 28, 10))
                        MAXSMSG = Trim(Mid(templine, 38, 10))
                        TRAFBLK = Trim(Mid(templine, 48, 8))
                        outputFile.writeLINE(BSCNAME & "," & CELL & "," & CBCHD & "," & PRACHBLK & "," & MAXSBLK & "," & MAXSMSG & "," & TRAFBLK)
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLPRP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, SCTYPE As String, CHTYPE As String, CHRATE As String, PP As String
        'CELL     SCTYPE  CHTYPE  CHRATE  PP

        CELL = ""
        SCTYPE = ""
        CHTYPE = ""
        CHRATE = ""
        PP = ""
        filepath = outPath & "RLPRP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "SCTYPE" & "," & "CHTYPE" & "," & "CHRATE" & "," & "PP")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Strings.left(Trim(templine), 20)) = "CELL     SCTYPE  CHT" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'CELL     SCTYPE  CHTYPE  CHRATE  PP
                            CELL = Trim(Mid(templine, 1, 9))
                            SCTYPE = Trim(Mid(templine, 10, 8))
                            CHTYPE = Trim(Mid(templine, 18, 8))
                            CHRATE = Trim(Mid(templine, 26, 8))
                            PP = Trim(Mid(templine, 34, 12))
                            outputFile.writeline(BSCNAME & "," & CELL & "," & SCTYPE & "," & CHTYPE & "," & CHRATE & "," & PP)
                        End If
                    ElseIf UCase(Trim(templine)) <> "BSC DIFFERENTIAL CHANNEL ALLOCATION" And UCase(Trim(templine)) <> "PRIORITY PROFILE RESOURCE TYPE DATA" And Len(Trim(Mid(templine, 1, 9))) > 0 And Len(Trim(Mid(templine, 1, 9))) <= 8 Then
                        If Not Trim(templine) = "NONE" Then
                            'SCTYPE    SSDESUL  QDESUL  LCOMPUL  QCOMPUL
                            CELL = Trim(Mid(templine, 1, 9))
                            SCTYPE = Trim(Mid(templine, 10, 8))
                            CHTYPE = Trim(Mid(templine, 18, 8))
                            CHRATE = Trim(Mid(templine, 26, 8))
                            PP = Trim(Mid(templine, 34, 12))
                            outputFile.writeline(BSCNAME & "," & CELL & "," & SCTYPE & "," & CHTYPE & "," & CHRATE & "," & PP)
                        End If
                    ElseIf UCase(Trim(templine)) <> "BSC DIFFERENTIAL CHANNEL ALLOCATION" And UCase(Trim(templine)) <> "PRIORITY PROFILE RESOURCE TYPE DATA" And Len(Trim(Mid(templine, 1, 9))) = 0 Then
                        If Not Trim(templine) = "NONE" Then
                            'SCTYPE    SSDESUL  QDESUL  LCOMPUL  QCOMPUL
                            SCTYPE = Trim(Mid(templine, 10, 8))
                            CHTYPE = Trim(Mid(templine, 18, 8))
                            CHRATE = Trim(Mid(templine, 26, 8))
                            PP = Trim(Mid(templine, 34, 12))
                            outputFile.writeline(BSCNAME & "," & CELL & "," & SCTYPE & "," & CHTYPE & "," & CHRATE & "," & PP)
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLRLP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, FASTRETLTE As String
        'CELL      FASTRETLTE

        CELL = ""
        FASTRETLTE = ""

        filepath = outPath & "RLRLP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "FASTRETLTE")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL      FASTRETLTE" Then
                        templine = getNowLine()
                        'CELL     DTMSTATE  MAXLAPDM  ALLOCPREF  OPTMSCLASS
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 10))
                            FASTRETLTE = Trim(Mid(templine, 11, 10))
                            outputFile.writeline(BSCNAME & "," & CELL & "," & FASTRETLTE)
                        End If
                    ElseIf UCase(Trim(templine)) <> "FAST RETURN TO LTE AFTER CALL RELEASE DATA" And Len(Trim(Mid(templine, 1, 10))) <= 8 Then
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 10))
                            FASTRETLTE = Trim(Mid(templine, 11, 10))
                            outputFile.writeline(BSCNAME & "," & CELL & "," & FASTRETLTE)
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLRLP_COVERAGEE(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, FASTRETLTE As String, EARFCN As String, COVERAGEE As String

        'CELL      FASTRETLTE  EARFCN  COVERAGEE

        CELL = ""
        FASTRETLTE = ""
        EARFCN = ""
        COVERAGEE = ""

        filepath = outPath & "RLRLP_COVERAGEE_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "FASTRETLTE" & "," & "EARFCN" & "," & "COVERAGEE")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL      FASTRETLTE  EARFCN  COVERAGEE" Then
                        templine = getNowLine()
                        'CELL      FASTRETLTE  EARFCN  COVERAGEE
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 10))
                            FASTRETLTE = Trim(Mid(templine, 11, 10))
                            EARFCN = Trim(Mid(templine, 21, 8))
                            COVERAGEE = Trim(Mid(templine, 29, 10))
                            outputFile.writeline(BSCNAME & "," & CELL & "," & FASTRETLTE & "," & EARFCN & "," & COVERAGEE)
                        End If
                    ElseIf UCase(Trim(templine)) <> "FAST RETURN TO LTE AFTER CALL RELEASE DATA" And Len(Trim(Mid(templine, 1, 10))) <= 8 And Len(Trim(Mid(templine, 1, 10))) <> 0 Then
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 10))
                            FASTRETLTE = Trim(Mid(templine, 11, 10))
                            EARFCN = Trim(Mid(templine, 21, 8))
                            COVERAGEE = Trim(Mid(templine, 29, 10))
                            outputFile.writeline(BSCNAME & "," & CELL & "," & FASTRETLTE & "," & EARFCN & "," & COVERAGEE)
                        End If
                    ElseIf UCase(Trim(templine)) <> "FAST RETURN TO LTE AFTER CALL RELEASE DATA" And Len(Trim(Mid(templine, 1, 10))) = 0 Then
                        If Not Trim(templine) = "NONE" Then
                            EARFCN = Trim(Mid(templine, 21, 8))
                            COVERAGEE = Trim(Mid(templine, 29, 10))
                            outputFile.writeline(BSCNAME & "," & CELL & "," & FASTRETLTE & "," & EARFCN & "," & COVERAGEE)
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLSBP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String
        Dim CB As String, MAXRET As String, TX As String, ATT As String, T3212 As String, CBQ As String, CRO As String, RLSBP_TO As String, PT As String, ECSC As String
        Dim ACC As String
        'CELL
        'CB   MAXRET  TX  ATT  T3212  CBQ   CRO  TO  PT  ECSC
        'ACC
        CELL = ""
        CB = ""
        MAXRET = ""
        TX = ""
        ATT = ""
        T3212 = ""
        CBQ = ""
        CRO = ""
        RLSBP_TO = ""
        PT = ""
        ECSC = ""
        ACC = ""
        filepath = outPath & "RLSBP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" _
                                  & "," & "CB" & "," & "MAXRET" & "," & "TX" & "," & "ATT" & "," & "T3212" & "," & "CBQ" & "," & "CRO" & "," & "TO" & "," & "PT" & "," & "ECSC" _
                                  & "," & "ACC")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL _
                                                  & "," & CB & "," & MAXRET & "," & TX & "," & ATT & "," & T3212 & "," & CBQ & "," & CRO & "," & RLSBP_TO & "," & PT & "," & ECSC _
                                                  & "," & ACC)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL _
                                                  & "," & CB & "," & MAXRET & "," & TX & "," & ATT & "," & T3212 & "," & CBQ & "," & CRO & "," & RLSBP_TO & "," & PT & "," & ECSC _
                                                  & "," & ACC)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL" Then
                        'CELL
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 9))
                        End If
                    ElseIf UCase(Trim(templine)) = "CB   MAXRET  TX  ATT  T3212  CBQ   CRO  TO  PT  ECSC" Then
                        'CB   MAXRET  TX  ATT  T3212  CBQ   CRO  TO  PT  ECSC
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            CB = Trim(Mid(templine, 1, 5))
                            MAXRET = Trim(Mid(templine, 6, 8))
                            TX = Trim(Mid(templine, 14, 4))
                            ATT = Trim(Mid(templine, 18, 5))
                            T3212 = Trim(Mid(templine, 23, 7))
                            CBQ = Trim(Mid(templine, 30, 6))
                            CRO = Trim(Mid(templine, 36, 5))
                            RLSBP_TO = Trim(Mid(templine, 41, 4))
                            PT = Trim(Mid(templine, 45, 4))
                            ECSC = Trim(Mid(templine, 49, 4))
                        End If
                    ElseIf UCase(Trim(templine)) = "ACC" Then
                        'ACC
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            ACC = Trim(templine)
                            outputFile.writeline(BSCNAME & "," & CELL _
                                                 & "," & CB & "," & MAXRET & "," & TX & "," & ATT & "," & T3212 & "," & CBQ & "," & CRO & "," & RLSBP_TO & "," & PT & "," & ECSC _
                                                 & "," & ACC)
                            CELL = ""
                            CB = ""
                            MAXRET = ""
                            TX = ""
                            ATT = ""
                            T3212 = ""
                            CBQ = ""
                            CRO = ""
                            RLSBP_TO = ""
                            PT = ""
                            ECSC = ""
                            ACC = ""
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLSCP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim STATE As String, STATSINT As String, TIME As String
        STATE = ""
        STATSINT = ""
        TIME = ""

        'STATE   STATSINT    TIME

        filepath = outPath & "RLSCP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "STATE" & "," & "STATSINT" & "," & "TIME")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "STATE   STATSINT    TIME" Then
                        templine = getNowLine()
                        STATE = Trim(Mid(templine, 1, 8))
                        STATSINT = Trim(Mid(templine, 9, 12))
                        TIME = Trim(Mid(templine, 21, 19))
                        'STATE   STATSINT    TIME
                        outputFile.writeLINE(BSCNAME & "," & STATE & "," & STATSINT & "," & TIME)
                    ElseIf UCase(Trim(templine)) <> "BSC DIFFERENTIAL CHANNEL ALLOCATION STATISTICS COLLECTION DATA" And Len(Trim(Mid(templine, 1, 8))) < 8 And UCase(Trim(templine)) <> "STATE   STATSINT    TIME" Then
                        STATE = Trim(Mid(templine, 1, 8))
                        STATSINT = Trim(Mid(templine, 9, 12))
                        TIME = Trim(Mid(templine, 21, 19))
                        outputFile.writeLINE(BSCNAME & "," & STATE & "," & STATSINT & "," & TIME)
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLSLP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, SCTYPE As String
        Dim ACTIVE As String, CHTYPE As String, CHRATE As String, SPV As String, LVA As String, ACL As String, NCH As String
        'CELL
        'ACTIVE     CHTYPE                     LVA    ACL   NCH
        'CELL       SCTYPE
        'ACTIVE     CHTYPE      CHRATE   SPV   LVA    ACL   NCH
        CELL = ""
        SCTYPE = ""
        ACTIVE = ""
        CHTYPE = ""
        CHRATE = ""
        SPV = ""
        LVA = ""
        ACL = ""
        NCH = ""
        filepath = outPath & "RLSLP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "SCTYPE" & "," & "ACTIVE" & "," & "CHTYPE" & "," & "CHRATE" & "," & "SPV" & "," & "LVA" & "," & "ACL" & "," & "NCH")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & SCTYPE & "," & ACTIVE & "," & CHTYPE & "," & CHRATE & "," & SPV & "," & LVA & "," & ACL & "," & NCH)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & SCTYPE & "," & ACTIVE & "," & CHTYPE & "," & CHRATE & "," & SPV & "," & LVA & "," & ACL & "," & NCH)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL" Then
                        'CELL
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 9))
                            SCTYPE = ""
                        End If
                    ElseIf UCase(Trim(templine)) = "ACTIVE     CHTYPE                     LVA    ACL   NCH" Then
                        'ACTIVE     CHTYPE                     LVA    ACL   NCH
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            ACTIVE = Trim(Mid(templine, 1, 11))
                            CHTYPE = Trim(Mid(templine, 12, 12))
                            CHRATE = Trim(Mid(templine, 24, 9))
                            SPV = Trim(Mid(templine, 33, 6))
                            LVA = Trim(Mid(templine, 39, 7))
                            ACL = Trim(Mid(templine, 46, 6))
                            NCH = Trim(Mid(templine, 52, 6))
                        End If
                        outputFile.writeline(BSCNAME & "," & CELL & "," & SCTYPE & "," & ACTIVE & "," & CHTYPE & "," & CHRATE & "," & SPV & "," & LVA & "," & ACL & "," & NCH)
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            CHTYPE = Trim(Mid(templine, 12, 12))
                            CHRATE = Trim(Mid(templine, 24, 9))
                            SPV = Trim(Mid(templine, 33, 6))
                            LVA = Trim(Mid(templine, 39, 7))
                            ACL = Trim(Mid(templine, 46, 6))
                            NCH = Trim(Mid(templine, 52, 6))
                        End If
                        outputFile.writeline(BSCNAME & "," & CELL & "," & SCTYPE & "," & ACTIVE & "," & CHTYPE & "," & CHRATE & "," & SPV & "," & LVA & "," & ACL & "," & NCH)
                    ElseIf UCase(Trim(templine)) = "CELL       SCTYPE" Then
                        'CELL       SCTYPE
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 11))
                            SCTYPE = Trim(Mid(templine, 12, 9))
                        End If
                    ElseIf UCase(Trim(templine)) = "ACTIVE     CHTYPE      CHRATE   SPV   LVA    ACL   NCH" Then
                        'ACTIVE     CHTYPE      CHRATE   SPV   LVA    ACL   NCH
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            ACTIVE = Trim(Mid(templine, 1, 11))
                            CHTYPE = Trim(Mid(templine, 12, 12))
                            CHRATE = Trim(Mid(templine, 24, 9))
                            SPV = Trim(Mid(templine, 33, 6))
                            LVA = Trim(Mid(templine, 39, 7))
                            ACL = Trim(Mid(templine, 46, 6))
                            NCH = Trim(Mid(templine, 52, 6))
                        End If
                        outputFile.writeline(BSCNAME & "," & CELL & "," & SCTYPE & "," & ACTIVE & "," & CHTYPE & "," & CHRATE & "," & SPV & "," & LVA & "," & ACL & "," & NCH)
                        For x = 1 To 8
                            templine = getNowLine()
                            If Not Trim(templine) = "NONE" Then
                                CHTYPE = Trim(Mid(templine, 12, 12))
                                CHRATE = Trim(Mid(templine, 24, 9))
                                SPV = Trim(Mid(templine, 33, 6))
                                LVA = Trim(Mid(templine, 39, 7))
                                ACL = Trim(Mid(templine, 46, 6))
                                NCH = Trim(Mid(templine, 52, 6))
                            End If
                            outputFile.writeline(BSCNAME & "," & CELL & "," & SCTYPE & "," & ACTIVE & "," & CHTYPE & "," & CHRATE & "," & SPV & "," & LVA & "," & ACL & "," & NCH)
                        Next
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLSMP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, SIMSG As String, MSGDIST As String
        'CELL     SIMSG  MSGDIST
        CELL = ""
        SIMSG = ""
        MSGDIST = ""
        filepath = outPath & "RLSMP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "SIMSG" & "," & "MSGDIST")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & SIMSG & "," & MSGDIST)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & SIMSG & "," & MSGDIST)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     SIMSG  MSGDIST" Then
                        'CELL     SIMSG  MSGDIST
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 9))
                            SIMSG = Trim(Mid(templine, 10, 7))
                            MSGDIST = Trim(Mid(templine, 17, 9))
                        End If
                        outputFile.writeline(BSCNAME & "," & CELL & "," & SIMSG & "," & MSGDIST)
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            SIMSG = Trim(Mid(templine, 10, 7))
                            MSGDIST = Trim(Mid(templine, 17, 9))
                        End If
                        outputFile.writeline(BSCNAME & "," & CELL & "," & SIMSG & "," & MSGDIST)
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            SIMSG = Trim(Mid(templine, 10, 7))
                            MSGDIST = Trim(Mid(templine, 17, 9))
                        End If
                        outputFile.writeline(BSCNAME & "," & CELL & "," & SIMSG & "," & MSGDIST)
                        CELL = ""
                        SIMSG = ""
                        MSGDIST = ""
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLSRP(sender As Object)
        Dim outputFile1 As Object
        Dim filepath1 As String
        Dim outputFile2 As Object
        Dim filepath2 As String
        Dim outputFile3 As Object
        Dim filepath3 As String
        Dim templine As String
        Dim CELL As String, PRIOCR As String, BCAST As String, FREE As String, REQ As String
        Dim RATPRIO As String, MEASTHR As String, PRIOTHR As String, HPRIO As String, TRES As String
        Dim FDDARFCN As String, FDD_RATPRIO As String, FDD_HPRIOTHR As String, FDD_LPRIOTHR As String, FDD_QRXLEVMINU As String
        Dim EARFCN As String, E_RATPRIO As String, E_HPRIOTHR As String, E_LPRIOTHR As String, E_QRXLEVMINE As String, E_MINCHBW As String
        Dim TDDARFCN As String, TDD_RATPRIO As String, TDD_HPRIOTHR As String, TDD_LPRIOTHR As String, TDD_QRXLEVMINU As String
        Dim mark As String
        'CELL     PRIOCR  BCAST    FREE  REQ
        'RATPRIO  MEASTHR  PRIOTHR  HPRIO  TRES
        'FDDARFCN  RATPRIO  HPRIOTHR  LPRIOTHR  QRXLEVMINU
        'EARFCN  RATPRIO  HPRIOTHR  LPRIOTHR  QRXLEVMINE  MINCHBW
        'TDDARFCN  RATPRIO  HPRIOTHR  LPRIOTHR  QRXLEVMINU
        mark = ""
        CELL = ""
        PRIOCR = ""
        BCAST = ""
        FREE = ""
        REQ = ""
        RATPRIO = ""
        MEASTHR = ""
        PRIOTHR = ""
        HPRIO = ""
        TRES = ""
        FDDARFCN = ""
        FDD_RATPRIO = ""
        FDD_HPRIOTHR = ""
        FDD_LPRIOTHR = ""
        FDD_QRXLEVMINU = ""
        EARFCN = ""
        E_RATPRIO = ""
        E_HPRIOTHR = ""
        E_LPRIOTHR = ""
        E_QRXLEVMINE = ""
        E_MINCHBW = ""
        TDDARFCN = ""
        TDD_RATPRIO = ""
        TDD_HPRIOTHR = ""
        TDD_LPRIOTHR = ""
        TDD_QRXLEVMINU = ""

        filepath1 = outPath & "RLSRP_FDD_" & CDDTime & ".csv"
        If fso.fileExists(filepath1) Then
            outputFile1 = fso.OpenTextFile(filepath1, 8, 1, 0)
        Else
            outputFile1 = fso.OpenTextFile(filepath1, 2, 1, 0)
            outputFile1.writeline("BSC" & "," & "CELL" & "," & "PRIOCR" & "," & "BCAST" & "," & "FREE" & "," & "REQ" & "," & _
                                  "RATPRIO" & "," & "MEASTHR" & "," & "PRIOTHR" & "," & "HPRIO" & "," & "TRES" & "," & _
                                  "FDDARFCN" & "," & "RATPRIO" & "," & "HPRIOTHR" & "," & "LPRIOTHR" & "," & "QRXLEVMINU")
        End If

        filepath2 = outPath & "RLSRP_EUTRAN_" & CDDTime & ".csv"
        If fso.fileExists(filepath2) Then
            outputFile2 = fso.OpenTextFile(filepath2, 8, 1, 0)
        Else
            outputFile2 = fso.OpenTextFile(filepath2, 2, 1, 0)
            outputFile2.writeline("BSC" & "," & "CELL" & "," & "PRIOCR" & "," & "BCAST" & "," & "FREE" & "," & "REQ" & "," & _
                                  "RATPRIO" & "," & "MEASTHR" & "," & "PRIOTHR" & "," & "HPRIO" & "," & "TRES" & "," & _
                                  "EARFCN" & "," & "RATPRIO" & "," & "HPRIOTHR" & "," & "LPRIOTHR" & "," & "QRXLEVMINE" & "," & "MINCHBW")
        End If

        filepath3 = outPath & "RLSRP_TDD_" & CDDTime & ".csv"
        If fso.fileExists(filepath3) Then
            outputFile3 = fso.OpenTextFile(filepath3, 8, 1, 0)
        Else
            outputFile3 = fso.OpenTextFile(filepath3, 2, 1, 0)
            outputFile3.writeline("BSC" & "," & "CELL" & "," & "PRIOCR" & "," & "BCAST" & "," & "FREE" & "," & "REQ" & "," & _
                                  "RATPRIO" & "," & "MEASTHR" & "," & "PRIOTHR" & "," & "HPRIO" & "," & "TRES" & "," & _
                                  "TDDARFCN" & "," & "RATPRIO" & "," & "HPRIOTHR" & "," & "LPRIOTHR" & "," & "QRXLEVMINU")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile1.close()
                        outputFile1 = Nothing
                        outputFile2.close()
                        outputFile2 = Nothing
                        outputFile3.close()
                        outputFile3 = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile1.close()
                        outputFile1 = Nothing
                        outputFile2.close()
                        outputFile2 = Nothing
                        outputFile3.close()
                        outputFile3 = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL     PRIOCR  BCAST    FREE  REQ" Then
                        templine = getNowLine()
                        CELL = Trim(Mid(templine, 1, 9))
                        PRIOCR = Trim(Mid(templine, 10, 8))
                        BCAST = Trim(Mid(templine, 18, 9))
                        FREE = Trim(Mid(templine, 27, 6))
                        REQ = Trim(Mid(templine, 33, 5))
                    ElseIf UCase(Trim(templine)) = "RATPRIO  MEASTHR  PRIOTHR  HPRIO  TRES" Then
                        templine = getNowLine()
                        RATPRIO = Trim(Mid(templine, 1, 9))
                        MEASTHR = Trim(Mid(templine, 10, 9))
                        PRIOTHR = Trim(Mid(templine, 19, 9))
                        HPRIO = Trim(Mid(templine, 28, 7))
                        TRES = Trim(Mid(templine, 35, 5))
                    ElseIf UCase(Trim(templine)) = "FDDARFCN  RATPRIO  HPRIOTHR  LPRIOTHR  QRXLEVMINU" Then
                        templine = getNowLine()
                        FDDARFCN = Trim(Mid(templine, 1, 10))
                        FDD_RATPRIO = Trim(Mid(templine, 11, 9))
                        FDD_HPRIOTHR = Trim(Mid(templine, 20, 10))
                        FDD_LPRIOTHR = Trim(Mid(templine, 30, 10))
                        FDD_QRXLEVMINU = Trim(Mid(templine, 40, 10))
                        outputFile1.writeLINE(BSCNAME & "," & CELL & "," & PRIOCR & "," & BCAST & "," & FREE & "," & REQ & "," &
                                             RATPRIO & "," & MEASTHR & "," & PRIOTHR & "," & HPRIO & "," & TRES & "," &
                                             FDDARFCN & "," & FDD_RATPRIO & "," & FDD_HPRIOTHR & "," & FDD_LPRIOTHR & "," & FDD_QRXLEVMINU)
                        mark = "FDD"
                    ElseIf UCase(Trim(templine)) = "EARFCN  RATPRIO  HPRIOTHR  LPRIOTHR  QRXLEVMINE  MINCHBW" Or UCase(Trim(templine)) = "EARFCN  RATPRIO  HPRIOTHR  LPRIOTHR  QRXLEVMINE" Then
                        templine = getNowLine()
                        EARFCN = Trim(Mid(templine, 1, 8))
                        E_RATPRIO = Trim(Mid(templine, 9, 9))
                        E_HPRIOTHR = Trim(Mid(templine, 18, 10))
                        E_LPRIOTHR = Trim(Mid(templine, 28, 10))
                        E_QRXLEVMINE = Trim(Mid(templine, 38, 12))
                        E_MINCHBW = Trim(Mid(templine, 50, 9))
                        outputFile2.writeLINE(BSCNAME & "," & CELL & "," & PRIOCR & "," & BCAST & "," & FREE & "," & REQ & "," & _
                                             RATPRIO & "," & MEASTHR & "," & PRIOTHR & "," & HPRIO & "," & TRES & "," & _
                                             EARFCN & "," & E_RATPRIO & "," & E_HPRIOTHR & "," & E_LPRIOTHR & "," & E_QRXLEVMINE & "," & E_MINCHBW)
                        mark = "EUTRAN"
                    ElseIf UCase(Trim(templine)) = "TDDARFCN  RATPRIO  HPRIOTHR  LPRIOTHR  QRXLEVMINU" Then
                        templine = getNowLine()
                        TDDARFCN = Trim(Mid(templine, 1, 10))
                        TDD_RATPRIO = Trim(Mid(templine, 11, 9))
                        TDD_HPRIOTHR = Trim(Mid(templine, 20, 10))
                        TDD_LPRIOTHR = Trim(Mid(templine, 30, 10))
                        TDD_QRXLEVMINU = Trim(Mid(templine, 40, 10))
                        outputFile3.writeLINE(BSCNAME & "," & CELL & "," & PRIOCR & "," & BCAST & "," & FREE & "," & REQ & "," & _
                                             RATPRIO & "," & MEASTHR & "," & PRIOTHR & "," & HPRIO & "," & TRES & "," & _
                                             TDDARFCN & "," & TDD_RATPRIO & "," & TDD_HPRIOTHR & "," & TDD_LPRIOTHR & "," & TDD_QRXLEVMINU)
                        mark = "TDD"
                    ElseIf UCase(Trim(templine)) <> "CELL SYSTEM INFORMATION RAT PRIORITY DATA" Then
                        If mark = "FDD" Then
                            templine = getNowLine()
                            FDDARFCN = Trim(Mid(templine, 1, 10))
                            FDD_RATPRIO = Trim(Mid(templine, 11, 9))
                            FDD_HPRIOTHR = Trim(Mid(templine, 20, 10))
                            FDD_LPRIOTHR = Trim(Mid(templine, 30, 10))
                            FDD_QRXLEVMINU = Trim(Mid(templine, 40, 10))
                            outputFile1.writeLINE(BSCNAME & "," & CELL & "," & PRIOCR & "," & BCAST & "," & FREE & "," & REQ & "," & _
                                                 RATPRIO & "," & MEASTHR & "," & PRIOTHR & "," & HPRIO & "," & TRES & "," & _
                                                 FDDARFCN & "," & FDD_RATPRIO & "," & FDD_HPRIOTHR & "," & FDD_LPRIOTHR & "," & FDD_QRXLEVMINU)

                        ElseIf mark = "EUTRAN" Then
                            EARFCN = Trim(Mid(templine, 1, 8))
                            E_RATPRIO = Trim(Mid(templine, 9, 9))
                            E_HPRIOTHR = Trim(Mid(templine, 18, 10))
                            E_LPRIOTHR = Trim(Mid(templine, 28, 10))
                            E_QRXLEVMINE = Trim(Mid(templine, 38, 12))
                            E_MINCHBW = Trim(Mid(templine, 50, 9))
                            outputFile2.writeLINE(BSCNAME & "," & CELL & "," & PRIOCR & "," & BCAST & "," & FREE & "," & REQ & "," & _
                                                 RATPRIO & "," & MEASTHR & "," & PRIOTHR & "," & HPRIO & "," & TRES & "," & _
                                                 EARFCN & "," & E_RATPRIO & "," & E_HPRIOTHR & "," & E_LPRIOTHR & "," & E_QRXLEVMINE & "," & E_MINCHBW)

                        ElseIf mark = "TDD" Then
                            TDDARFCN = Trim(Mid(templine, 1, 10))
                            TDD_RATPRIO = Trim(Mid(templine, 11, 9))
                            TDD_HPRIOTHR = Trim(Mid(templine, 20, 10))
                            TDD_LPRIOTHR = Trim(Mid(templine, 30, 10))
                            TDD_QRXLEVMINU = Trim(Mid(templine, 40, 10))
                            outputFile3.writeLINE(BSCNAME & "," & CELL & "," & PRIOCR & "," & BCAST & "," & FREE & "," & REQ & "," & _
                                                 RATPRIO & "," & MEASTHR & "," & PRIOTHR & "," & HPRIO & "," & TRES & "," & _
                                                 TDDARFCN & "," & TDD_RATPRIO & "," & TDD_HPRIOTHR & "," & TDD_LPRIOTHR & "," & TDD_QRXLEVMINU)


                        End If
                    End If
                End If
            End If
        Loop
        outputFile1.close()
        outputFile1 = Nothing
        outputFile2.close()
        outputFile2 = Nothing
        outputFile3.close()
        outputFile3 = Nothing
    End Sub
    Private Sub doRLSQP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim FERTHR1 As String, FERTHR2 As String, FERTHR3 As String, FERTHR4 As String, WINSIZE As String
        FERTHR1 = ""
        FERTHR2 = ""
        FERTHR3 = ""
        FERTHR4 = ""
        WINSIZE = ""
        'FERTHR1  FERTHR2  FERTHR3  FERTHR4  WINSIZE

        filepath = outPath & "RLSQP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "FERTHR1" & "," & "FERTHR2" & "," & "FERTHR3" & "," & "FERTHR4" & "," & "WINSIZE")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "FERTHR1  FERTHR2  FERTHR3  FERTHR4  WINSIZE" Then
                        templine = getNowLine()
                        FERTHR1 = Trim(Mid(templine, 1, 9))
                        FERTHR2 = Trim(Mid(templine, 10, 9))
                        FERTHR3 = Trim(Mid(templine, 19, 9))
                        FERTHR4 = Trim(Mid(templine, 28, 9))
                        WINSIZE = Trim(Mid(templine, 37, 9))
                        'FERTHR1  FERTHR2  FERTHR3  FERTHR4  WINSIZE

                        outputFile.writeLINE(BSCNAME & "," & FERTHR1 & "," & FERTHR2 & "," & FERTHR3 & "," & FERTHR4 & "," & WINSIZE)
                    ElseIf UCase(Trim(templine)) <> "BSC SPEECH QUALITY SUPERVISION DATA" And Len(Trim(Mid(templine, 1, 9))) < 9 And UCase(Trim(templine)) <> "FERTHR1  FERTHR2  FERTHR3  FERTHR4  WINSIZE" Then
                        FERTHR1 = Trim(Mid(templine, 1, 9))
                        FERTHR2 = Trim(Mid(templine, 10, 9))
                        FERTHR3 = Trim(Mid(templine, 19, 9))
                        FERTHR4 = Trim(Mid(templine, 28, 9))
                        WINSIZE = Trim(Mid(templine, 37, 9))
                        outputFile.writeLINE(BSCNAME & "," & FERTHR1 & "," & FERTHR2 & "," & FERTHR3 & "," & FERTHR4 & "," & WINSIZE)
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLSSP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String
        Dim ACCMIN As String, CCHPWR As String, CRH As String, DTXU As String, RLINKT As String, NECI As String, MBCR As String
        Dim NCCPERM As String
        Dim RLINKTAFR As String, RLINKTAHR As String, RLINKTAWB As String
        'CELL
        'ACCMIN  CCHPWR  CRH  DTXU  RLINKT  NECI  MBCR
        'NCCPERM
        'RLINKTAFR  RLINKTAHR  RLINKTAWB
        CELL = ""
        ACCMIN = ""
        CCHPWR = ""
        CRH = ""
        DTXU = ""
        RLINKT = ""
        NECI = ""
        MBCR = ""
        NCCPERM = ""
        RLINKTAFR = ""
        RLINKTAHR = ""
        RLINKTAWB = ""
        filepath = outPath & "RLSSP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" _
                                  & "," & "ACCMIN" & "," & "CCHPWR" & "," & "CRH" & "," & "DTXU" & "," & "RLINKT" & "," & "NECI" & "," & "MBCR" _
                                  & "," & "NCCPERM" _
                                  & "," & "RLINKTAFR" & "," & "RLINKTAHR" & "," & "RLINKTAWB")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL _
                                                    & "," & ACCMIN & "," & CCHPWR & "," & CRH & "," & DTXU & "," & RLINKT & "," & NECI & "," & MBCR _
                                                    & "," & NCCPERM _
                                                    & "," & RLINKTAFR & "," & RLINKTAHR & "," & RLINKTAWB)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL _
                                                    & "," & ACCMIN & "," & CCHPWR & "," & CRH & "," & DTXU & "," & RLINKT & "," & NECI & "," & MBCR _
                                                    & "," & NCCPERM _
                                                    & "," & RLINKTAFR & "," & RLINKTAHR & "," & RLINKTAWB)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL" Then
                        'CELL
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 9))
                        End If
                    ElseIf UCase(Trim(templine)) = "ACCMIN  CCHPWR  CRH  DTXU  RLINKT  NECI  MBCR" Then
                        'ACCMIN  CCHPWR  CRH  DTXU  RLINKT  NECI  MBCR
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            ACCMIN = Trim(Mid(templine, 1, 8))
                            CCHPWR = Trim(Mid(templine, 9, 8))
                            CRH = Trim(Mid(templine, 17, 5))
                            DTXU = Trim(Mid(templine, 22, 6))
                            RLINKT = Trim(Mid(templine, 28, 8))
                            NECI = Trim(Mid(templine, 36, 6))
                            MBCR = Trim(Mid(templine, 42, 4))
                        End If
                    ElseIf UCase(Trim(templine)) = "NCCPERM" Then
                        'NCCPERM
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            NCCPERM = Trim(templine)
                        End If
                    ElseIf UCase(Trim(templine)) = "RLINKTAFR  RLINKTAHR  RLINKTAWB" Then
                        'RLINKTAFR  RLINKTAHR  RLINKTAWB
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            RLINKTAFR = Trim(Mid(templine, 1, 11))
                            RLINKTAHR = Trim(Mid(templine, 12, 11))
                            RLINKTAWB = Trim(Mid(templine, 23, 11))
                            outputFile.writeline(BSCNAME & "," & CELL _
                                                    & "," & ACCMIN & "," & CCHPWR & "," & CRH & "," & DTXU & "," & RLINKT & "," & NECI & "," & MBCR _
                                                    & "," & NCCPERM _
                                                    & "," & RLINKTAFR & "," & RLINKTAHR & "," & RLINKTAWB)
                            CELL = ""
                            ACCMIN = ""
                            CCHPWR = ""
                            CRH = ""
                            DTXU = ""
                            RLINKT = ""
                            NECI = ""
                            MBCR = ""
                            NCCPERM = ""
                            RLINKTAFR = ""
                            RLINKTAHR = ""
                            RLINKTAWB = ""
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLSTP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, STATE As String, CHGR As String, CHGR_STATE As String
        Dim CHGR_STATE_TEMP As String

        'CELL      STATE  CHGR STATE
        'CELL      STATE
        CELL = ""
        STATE = ""
        CHGR = ""
        CHGR_STATE = ""
        CHGR_STATE_TEMP = ""
        filepath = outPath & "RLSTP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "STATE" & "," & "CHGR" & "," & "STATE")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL      STATE" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 10))
                            STATE = Trim(Mid(templine, 11, 7))
                            CHGR = Trim(Mid(templine, 18, 5))
                            CHGR_STATE = Trim(Mid(templine, 23, 7))
                            outputFile.writeLINE(BSCNAME & "," & CELL & "," & STATE & "," & CHGR & "," & CHGR_STATE)
                        End If
                        If STP_BSC_CELL_STATE.ContainsKey(BSCNAME & "#" & CELL) Then
                            STP_BSC_CELL_STATE(BSCNAME & "#" & CELL) = STATE
                        Else
                            STP_BSC_CELL_STATE.Add(BSCNAME & "#" & CELL, STATE)
                        End If
                    ElseIf UCase(Trim(templine)) = "CELL      STATE  CHGR STATE" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 10))
                            STATE = Trim(Mid(templine, 11, 7))
                            CHGR = Trim(Mid(templine, 18, 5))
                            CHGR_STATE = Trim(Mid(templine, 23, 7))
                            CHGR_STATE_TEMP = "[" & CHGR & "] " & CHGR_STATE
                            outputFile.writeLINE(BSCNAME & "," & CELL & "," & STATE & "," & CHGR & "," & CHGR_STATE)
                        End If
                        If STP_BSC_CELL_STATE.ContainsKey(BSCNAME & "#" & CELL) Then
                            STP_BSC_CELL_STATE(BSCNAME & "#" & CELL) = CHGR_STATE_TEMP
                        Else
                            STP_BSC_CELL_STATE.Add(BSCNAME & "#" & CELL, CHGR_STATE_TEMP)
                        End If
                    ElseIf UCase(Trim(Strings.left(templine, 10))) = "" Then
                        CHGR = Trim(Mid(templine, 18, 5))
                        CHGR_STATE = Trim(Mid(templine, 23, 7))
                        outputFile.writeLINE(BSCNAME & "," & CELL & "," & STATE & "," & CHGR & "," & CHGR_STATE)
                        CHGR_STATE_TEMP = CHGR_STATE_TEMP & " [" & CHGR & "] " & CHGR_STATE
                        If STP_BSC_CELL_STATE.ContainsKey(BSCNAME & "#" & CELL) Then
                            STP_BSC_CELL_STATE(BSCNAME & "#" & CELL) = CHGR_STATE_TEMP
                        Else
                            STP_BSC_CELL_STATE.Add(BSCNAME & "#" & CELL, CHGR_STATE_TEMP)
                        End If
                    ElseIf Trim(Mid(templine, 1, 10)) <> "" And Trim(Mid(templine, 1, 10)) <> "CELL STATU" And Len(Trim(templine)) < 17 Then
                        CELL = Trim(Mid(templine, 1, 10))
                        STATE = Trim(Mid(templine, 11, 7))
                        CHGR = Trim(Mid(templine, 18, 5))
                        CHGR_STATE = Trim(Mid(templine, 23, 7))
                        outputFile.writeLINE(BSCNAME & "," & CELL & "," & STATE & "," & CHGR & "," & CHGR_STATE)
                        If STP_BSC_CELL_STATE.ContainsKey(BSCNAME & "#" & CELL) Then
                            STP_BSC_CELL_STATE(BSCNAME & "#" & CELL) = STATE
                        Else
                            STP_BSC_CELL_STATE.Add(BSCNAME & "#" & CELL, STATE)
                        End If
                    ElseIf Trim(Mid(templine, 1, 10)) <> "" And Trim(Mid(templine, 1, 10)) <> "CELL STATU" And Len(Trim(templine)) > 17 Then
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 10))
                            STATE = Trim(Mid(templine, 11, 7))
                            CHGR = Trim(Mid(templine, 18, 5))
                            CHGR_STATE = Trim(Mid(templine, 23, 7))
                        End If
                        outputFile.writeLINE(BSCNAME & "," & CELL & "," & STATE & "," & CHGR & "," & CHGR_STATE)
                        CHGR_STATE_TEMP = "[" & CHGR & "] " & CHGR_STATE
                        If STP_BSC_CELL_STATE.ContainsKey(BSCNAME & "#" & CELL) Then
                            STP_BSC_CELL_STATE(BSCNAME & "#" & CELL) = CHGR_STATE_TEMP
                        Else
                            STP_BSC_CELL_STATE.Add(BSCNAME & "#" & CELL, CHGR_STATE_TEMP)
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLSUP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String
        Dim QSC As String, QSCI As String, QSI As String, FDDMRR As String, FDDQMIN As String, FDDQOFF As String, SPRIO As String, FDDQMINOFF As String, FDDRSCPMIN As String
        Dim FDDREPTHR2 As String, TQSI As String, TDDMRR As String, TDDQOFF As String, TSPRIO As String
        Dim TQSC As String, TQSCI As String, TDDTHR As String
        CELL = ""
        QSC = ""
        QSCI = ""
        QSI = ""
        FDDMRR = ""
        FDDQMIN = ""
        FDDQOFF = ""
        SPRIO = ""
        FDDQMINOFF = ""
        FDDRSCPMIN = ""
        FDDREPTHR2 = ""
        TQSI = ""
        TDDMRR = ""
        TDDQOFF = ""
        TSPRIO = ""

        TQSC = ""
        TQSCI = ""
        TDDTHR = ""

        'CELL
        'QSC  QSCI  QSI  FDDMRR  FDDQMIN  FDDQOFF  SPRIO  FDDQMINOFF  FDDRSCPMIN
        'FDDREPTHR2
        'TQSI   TDDMRR   TDDQOFF   TSPRIO

        'G14
        'TQSI   TDDMRR   TDDQOFF   TQSC   TQSCI   TDDTHR

        filepath = outPath & "RLSUP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "QSC" & "," & "QSCI" & "," & "QSI" & "," & "FDDMRR" & "," & "FDDQMIN" & "," & "FDDQOFF" & "," & "SPRIO" & "," & "FDDQMINOFF" & "," & "FDDRSCPMIN" & "," & "FDDREPTHR2" _
                                  & "," & "TQSI" & "," & "TDDMRR " & "," & "  TDDQOFF   " & "," & "TSPRIO" _
                                  & "," & "TQSC" & "," & "TQSCI" & "," & "TDDTHR")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If CELL <> "" Then
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & QSC & "," & QSCI & "," & QSI & "," & FDDMRR & "," & FDDQMIN & "," & FDDQOFF & "," & SPRIO & "," & FDDQMINOFF & "," & FDDRSCPMIN _
                          & "," & FDDREPTHR2 & "," & TQSI & "," & TDDMRR & "," & TDDQOFF & "," & TSPRIO & "," & TQSC & "," & TQSCI & "," & TDDTHR)

                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If CELL <> "" Then
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & QSC & "," & QSCI & "," & QSI & "," & FDDMRR & "," & FDDQMIN & "," & FDDQOFF & "," & SPRIO & "," & FDDQMINOFF & "," & FDDRSCPMIN _
                          & "," & FDDREPTHR2 & "," & TQSI & "," & TDDMRR & "," & TDDQOFF & "," & TSPRIO & "," & TQSC & "," & TQSCI & "," & TDDTHR)

                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL" Then
                        templine = getNowLine()
                        If CELL <> "" Then
                            outputFile.writeLine(BSCNAME & "," & CELL & "," & QSC & "," & QSCI & "," & QSI & "," & FDDMRR & "," & FDDQMIN & "," & FDDQOFF & "," & SPRIO & "," & FDDQMINOFF & "," & FDDRSCPMIN _
                          & "," & FDDREPTHR2 & "," & TQSI & "," & TDDMRR & "," & TDDQOFF & "," & TSPRIO & "," & TQSC & "," & TQSCI & "," & TDDTHR)
                            CELL = ""
                            QSC = ""
                            QSCI = ""
                            QSI = ""
                            FDDMRR = ""
                            FDDQMIN = ""
                            FDDQOFF = ""
                            SPRIO = ""
                            FDDQMINOFF = ""
                            FDDRSCPMIN = ""
                            FDDREPTHR2 = ""
                            TQSI = ""
                            TDDMRR = ""
                            TDDQOFF = ""
                            TSPRIO = ""
                            TQSC = ""
                            TQSCI = ""
                            TDDTHR = ""
                        End If
                        If Not Trim(templine) = "NONE" Then
                            CELL = Trim(Mid(templine, 1, 10))
                        End If
                    ElseIf UCase(Trim(templine)) = "QSC  QSCI  QSI  FDDMRR  FDDQMIN  FDDQOFF  SPRIO  FDDQMINOFF  FDDRSCPMIN" Then
                        templine = getNowLine()
                        QSC = Trim(Mid(templine, 1, 5))
                        QSCI = Trim(Mid(templine, 6, 6))
                        QSI = Trim(Mid(templine, 12, 5))
                        FDDMRR = Trim(Mid(templine, 17, 8))
                        FDDQMIN = Trim(Mid(templine, 25, 9))
                        FDDQOFF = Trim(Mid(templine, 34, 9))
                        SPRIO = Trim(Mid(templine, 43, 7))
                        FDDQMINOFF = Trim(Mid(templine, 50, 12))
                        FDDRSCPMIN = Trim(Mid(templine, 62, 10))
                    ElseIf UCase(Trim(templine)) = "FDDREPTHR2" Then
                        templine = getNowLine()
                        FDDREPTHR2 = Trim(Mid(templine, 1, 13))
                    ElseIf UCase(Trim(templine)) = "TQSI   TDDMRR   TDDQOFF   TSPRIO" Then
                        templine = getNowLine()
                        TQSI = Trim(Mid(templine, 1, 7))
                        TDDMRR = Trim(Mid(templine, 8, 9))
                        TDDQOFF = Trim(Mid(templine, 17, 10))
                        TSPRIO = Trim(Mid(templine, 27, 6))
                    ElseIf UCase(Trim(templine)) = "TQSI   TDDMRR   TDDQOFF   TQSC   TQSCI   TDDTHR" Then
                        'G14
                        templine = getNowLine()
                        TQSI = Trim(Mid(templine, 1, 7))
                        TDDMRR = Trim(Mid(templine, 8, 9))
                        TDDQOFF = Trim(Mid(templine, 17, 10))
                        TQSC = Trim(Mid(templine, 27, 7))
                        TQSCI = Trim(Mid(templine, 34, 8))
                        TDDTHR = Trim(Mid(templine, 42, 12))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLSVP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'CELL      BTSPS      BTSPSHYST  MINREQTCH  NDCTRX
        filepath = outPath & "RLSVP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "BTSPS" & "," & "BTSPSHYST" & "," & "MINREQTCH" & "," & "NDCTRX")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL      BTSPS      BTSPSHYST  MINREQTCH  NDCTRX" Then
                        'CELL      BTSPS      BTSPSHYST  MINREQTCH  NDCTRX
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 10)) & "," & Trim(Mid(templine, 11, 11)) & "," & Trim(Mid(templine, 22, 11)) & "," & Trim(Mid(templine, 33, 11)) & "," & Trim(Mid(templine, 44, 11)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL POWER SAVINGS DATA" And Len(Trim(Mid(templine, 1, 9))) <= 7 Then
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 10)) & "," & Trim(Mid(templine, 11, 11)) & "," & Trim(Mid(templine, 22, 11)) & "," & Trim(Mid(templine, 33, 11)) & "," & Trim(Mid(templine, 44, 11)))
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLTYP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'GSYSTYPE
        filepath = outPath & "RLTYP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "GSYSTYPE")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "GSYSTYPE" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 9)))
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLTDP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'MSC       MODE      CAP  RELCAP  RAND  PART  PLMNNAME

        filepath = outPath & "RLTDP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "MSC" & "," & "MODE" & "," & "CAP" & "," & "RELCAP" & "," & "RAND" & "," & "PART" & "," & "PLMNNAME")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Strings.left(Trim(templine), 20)) = "MSC       MODE      " Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'MSC       MODE      CAP  RELCAP  RAND  PART  PLMNNAME
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 10)) & "," & Trim(Mid(templine, 11, 10)) & "," & Trim(Mid(templine, 21, 5)) & "," & Trim(Mid(templine, 26, 8)) & "," & Trim(Mid(templine, 34, 6)) & "," & Trim(Mid(templine, 40, 6)) & "," & Trim(Mid(templine, 46, 10)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "MSC TRAFFIC DISTRIBUTION DATA" Then
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 10)) & "," & Trim(Mid(templine, 11, 10)) & "," & Trim(Mid(templine, 21, 5)) & "," & Trim(Mid(templine, 26, 8)) & "," & Trim(Mid(templine, 34, 6)) & "," & Trim(Mid(templine, 40, 6)) & "," & Trim(Mid(templine, 46, 10)))
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRLUMP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String
        Dim LISTTYPE As String
        CELL = ""
        LISTTYPE = ""
        filepath = outPath & "RLUMP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "LISTTYPE" & "," & "UMFI/TMFI")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CELL" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            CELL = Trim(Mid(templine, 1, 10))
                        End If
                    ElseIf UCase(Trim(templine)) = "LISTTYPE" Then
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            LISTTYPE = Trim(Mid(templine, 1, 10))
                        End If
                    ElseIf UCase(Trim(templine)) = "UMFI" Or UCase(Trim(templine)) = "TMFI" Then
                        templine = getNowLine()
                        outputFile.writeLINE(BSCNAME & "," & CELL & "," & LISTTYPE & "," & Trim(templine))
                    ElseIf UCase(Trim(templine)) <> "CELL UTRAN MEASUREMENT FREQUENCY DATA" And Len(Trim(Mid(templine, 1, 5))) = 0 And Len(Trim(templine)) > 4 Then
                        outputFile.writeLINE(BSCNAME & "," & CELL & "," & LISTTYPE & "," & Trim(templine))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLVAP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CHANNEL As String, CELL As String, STATE As String
        CHANNEL = ""
        CELL = ""
        STATE = ""
        'CHANNEL         CELL       STATE

        filepath = outPath & "RLVAP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CHANNEL" & "," & "CELL" & "," & "STATE")
        End If

        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CHANNEL         CELL       STATE" Then
                        templine = getNowLine()
                        CHANNEL = Trim(Mid(templine, 1, 16))
                        CELL = Trim(Mid(templine, 17, 11))
                        STATE = Trim(Mid(templine, 28, 8))
                        'CHANNEL         CELL       STATE
                        outputFile.writeLINE(BSCNAME & "," & CHANNEL & "," & CELL & "," & STATE)
                    ElseIf UCase(Trim(templine)) <> "CELL SEIZURE SUPERVISION OF LOGICAL CHANNELS ALARMED OBJECTS DATA" And Len(Trim(Mid(templine, 1, 16))) < 16 And UCase(Trim(templine)) <> "CHANNEL         CELL       STATE" Then
                        CHANNEL = Trim(Mid(templine, 1, 16))
                        CELL = Trim(Mid(templine, 17, 11))
                        STATE = Trim(Mid(templine, 28, 8))
                        outputFile.writeLINE(BSCNAME & "," & CHANNEL & "," & CELL & "," & STATE)
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRLVLP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim CELL As String, TCH As String, SDCCH As String, ACL_TCH As String, PL_TCH As String, STATUS_TCH As String, ACL_SDCCH As String, PL_SDCCH As String, STATUS_SDCCH As String
        'CHTYPE  ACL     PL     STATUS
        'CELL          TCH            SDCCH
        CELL = ""
        TCH = ""
        SDCCH = ""
        ACL_TCH = ""
        ACL_SDCCH = ""
        PL_TCH = ""
        PL_SDCCH = ""
        STATUS_TCH = ""
        STATUS_SDCCH = ""
        filepath = outPath & "RLVLP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "CELL" & "," & "TCH" & "," & "SDCCH" & "," & "ACL_TCH" & "," & "PL_TCH" & "," & "STATUS_TCH" & "," & "ACL_SDCCH" & "," & "PL_SDCCH" & "," & "STATUS_SDCCH")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & TCH & "," & SDCCH & "," & ACL_TCH & "," & PL_TCH & "," & STATUS_TCH & "," & ACL_SDCCH & "," & PL_SDCCH & "," & STATUS_SDCCH)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If CELL <> "" Then
                            outputFile.writeline(BSCNAME & "," & CELL & "," & TCH & "," & SDCCH & "," & ACL_TCH & "," & PL_TCH & "," & STATUS_TCH & "," & ACL_SDCCH & "," & PL_SDCCH & "," & STATUS_SDCCH)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "CHTYPE  ACL     PL     STATUS" Then
                        'CHTYPE  ACL     PL     STATUS
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            If UCase(Trim(Mid(templine, 1, 8))) = "TCH" Then
                                ACL_TCH = Trim(Mid(templine, 9, 8))
                                PL_TCH = Trim(Mid(templine, 17, 7))
                                STATUS_TCH = Trim(Mid(templine, 24, 10))
                            ElseIf UCase(Trim(Mid(templine, 1, 8))) = "SDCCH" Then
                                ACL_SDCCH = Trim(Mid(templine, 9, 8))
                                PL_SDCCH = Trim(Mid(templine, 17, 7))
                                STATUS_SDCCH = Trim(Mid(templine, 24, 10))
                            End If
                        End If
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            If UCase(Trim(Mid(templine, 1, 8))) = "TCH" Then
                                ACL_TCH = Trim(Mid(templine, 9, 8))
                                PL_TCH = Trim(Mid(templine, 17, 7))
                                STATUS_TCH = Trim(Mid(templine, 24, 10))
                            ElseIf UCase(Trim(Mid(templine, 1, 8))) = "SDCCH" Then
                                ACL_SDCCH = Trim(Mid(templine, 9, 8))
                                PL_SDCCH = Trim(Mid(templine, 17, 7))
                                STATUS_SDCCH = Trim(Mid(templine, 24, 10))
                            End If
                        End If
                    ElseIf UCase(Trim(templine)) = "CELL          TCH            SDCCH" Then
                        'CELL          TCH            SDCCH
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            CELL = Trim(Mid(templine, 1, 14))
                            TCH = Trim(Mid(templine, 15, 15))
                            SDCCH = Trim(Mid(templine, 30, 15))
                            outputFile.writeline(BSCNAME & "," & CELL & "," & TCH & "," & SDCCH & "," & ACL_TCH & "," & PL_TCH & "," & STATUS_TCH & "," & ACL_SDCCH & "," & PL_SDCCH & "," & STATUS_SDCCH)
                            CELL = ""
                            TCH = ""
                            SDCCH = ""
                        End If
                    ElseIf UCase(Trim(templine)) <> "CELL SEIZURE SUPERVISION OF LOGICAL CHANNELS DATA" And UCase(Trim(templine)) <> "CHANNEL TYPE SUPERVISION DATA" And UCase(Trim(templine)) <> "CELLS INCLUDED IN SUPERVISION" And Len(Trim(templine)) > 31 Then
                        If Trim(templine) <> "NONE" Then
                            CELL = Trim(Mid(templine, 1, 14))
                            TCH = Trim(Mid(templine, 15, 15))
                            SDCCH = Trim(Mid(templine, 30, 15))
                            outputFile.writeline(BSCNAME & "," & CELL & "," & TCH & "," & SDCCH & "," & ACL_TCH & "," & PL_TCH & "," & STATUS_TCH & "," & ACL_SDCCH & "," & PL_SDCCH & "," & STATUS_SDCCH)
                            CELL = ""
                            TCH = ""
                            SDCCH = ""
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRRMBP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim MSC As String, SP As String, CNID As String

        'MSC      SP           CNID
        MSC = ""
        SP = ""
        CNID = ""

        filepath = outPath & "RRMBP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "MSC" & "," & "SP" & "," & "CNID")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "MSC      SP           CNID" Then
                        'MSC      SP           CNID
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            MSC = Trim(Mid(templine, 1, 9))
                            SP = Trim(Mid(templine, 10, 13))
                            CNID = Trim(Mid(templine, 23, 5))
                            outputFile.writeline(BSCNAME & "," & MSC & "," & SP & "," & CNID)
                        End If
                    ElseIf UCase(Trim(templine)) <> "MSC DEFINITION IN BSC DATA" Then
                        MSC = Trim(Mid(templine, 1, 9))
                        SP = Trim(Mid(templine, 10, 13))
                        CNID = Trim(Mid(templine, 23, 5))
                        outputFile.writeline(BSCNAME & "," & MSC & "," & SP & "," & CNID)
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRRSCP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim SSGR As String
        Dim SC As String
        Dim DEV As String
        Dim DEV1 As String
        Dim NUMDEV As String
        Dim DCP As String
        Dim STATE As String
        Dim REASON As String

        'SCGR  SC  DEV            DEV1           NUMDEV  DCP  STATE  REASON
        SSGR = ""
        SC = ""
        DEV = ""
        DEV1 = ""
        NUMDEV = ""
        DCP = ""
        STATE = ""
        REASON = ""

        filepath = outPath & "RRSCP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "SCGR" & "," & "SC" & "," & "DEV" & "," & "DEV1" & "," & "NUMDEV" & "," & "DCP" & "," & "STATE" & "," & "REASON")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "SCGR  SC  DEV            DEV1           NUMDEV  DCP  STATE  REASON" Then
                        'SCGR  SC  DEV            DEV1           NUMDEV  DCP  STATE  REASON
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            SSGR = Trim(Mid(templine, 1, 6))
                            SC = Trim(Mid(templine, 7, 4))
                            DEV = Trim(Mid(templine, 11, 15))
                            DEV1 = Trim(Mid(templine, 26, 15))
                            NUMDEV = Trim(Mid(templine, 41, 8))
                            DCP = Trim(Mid(templine, 49, 5))
                            STATE = Trim(Mid(templine, 54, 7))
                            REASON = Trim(Mid(templine, 61, 15))
                            outputFile.writeline(BSCNAME & "," & SSGR & "," & SC & "," & DEV & "," & DEV1 & "," & NUMDEV & "," & DCP & "," & STATE & "," & REASON)
                        End If
                    ElseIf UCase(Trim(templine)) <> "RADIO TRANSMISSION SUPER CHANNEL DATA" Then
                        If Trim(Mid(templine, 1, 6)) <> "" Then
                            SSGR = Trim(Mid(templine, 1, 6))
                        End If
                        SC = Trim(Mid(templine, 7, 4))
                        DEV = Trim(Mid(templine, 11, 15))
                        DEV1 = Trim(Mid(templine, 26, 15))
                        NUMDEV = Trim(Mid(templine, 41, 8))
                        DCP = Trim(Mid(templine, 49, 5))
                        STATE = Trim(Mid(templine, 54, 7))
                        REASON = Trim(Mid(templine, 61, 15))
                        outputFile.writeline(BSCNAME & "," & SSGR & "," & SC & "," & DEV & "," & DEV1 & "," & NUMDEV & "," & DCP & "," & STATE & "," & REASON)
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRRTTP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim TRAPOOL As String, CHRATE As String, SPV As String, RNOTRA As String, POOLACT As String, POOLIDLE As String, POOLTRAF As String
        Dim SUBPOOL As String, SUBACT As String, SUBIDLE As String, SUBTRAF As String
        'TRAPOOL  CHRATE  SPV  RNOTRA  POOLACT  POOLIDLE  POOLTRAF
        'SUBPOOL                       SUBACT   SUBIDLE   SUBTRAF 
        TRAPOOL = ""
        CHRATE = ""
        SPV = ""
        RNOTRA = ""
        POOLACT = ""
        POOLIDLE = ""
        POOLTRAF = ""
        SUBPOOL = ""
        SUBACT = ""
        SUBIDLE = ""
        SUBTRAF = ""

        filepath = outPath & "RRTTP_RXOTG_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "TRAPOOL" & "," & "CHRATE" & "," & "SPV" & "," & "RNOTRA" & "," & "POOLACT" & "," & "POOLIDLE" & "," & "POOLTRAF" _
                                 & "," & "SUBPOOL" & "," & "SUBACT" & "," & "SUBIDLE" & "," & "SUBTRAF")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If TRAPOOL <> "" Then
                            outputFile.writeline(BSCNAME & "," & TRAPOOL & "," & CHRATE & "," & SPV & "," & RNOTRA & "," & POOLACT & "," & POOLIDLE & "," & POOLTRAF _
                                                 & "," & SUBPOOL & "," & SUBACT & "," & SUBIDLE & "," & SUBTRAF)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If TRAPOOL <> "" Then
                            outputFile.writeline(BSCNAME & "," & TRAPOOL & "," & CHRATE & "," & SPV & "," & RNOTRA & "," & POOLACT & "," & POOLIDLE & "," & POOLTRAF _
                                                 & "," & SUBPOOL & "," & SUBACT & "," & SUBIDLE & "," & SUBTRAF)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "TRAPOOL  CHRATE  SPV  RNOTRA  POOLACT  POOLIDLE  POOLTRAF" Then
                        'TRAPOOL  CHRATE  SPV  RNOTRA  POOLACT  POOLIDLE  POOLTRAF
                        templine = getNowLine()
                        If TRAPOOL <> "" Then
                            outputFile.writeline(BSCNAME & "," & TRAPOOL & "," & CHRATE & "," & SPV & "," & RNOTRA & "," & POOLACT & "," & POOLIDLE & "," & POOLTRAF _
                                                 & "," & SUBPOOL & "," & SUBACT & "," & SUBIDLE & "," & SUBTRAF)
                        End If
                        If Trim(templine) <> "NONE" Then
                            TRAPOOL = Trim(Mid(templine, 1, 9))
                            CHRATE = Trim(Mid(templine, 10, 8))
                            SPV = Trim(Mid(templine, 18, 5))
                            RNOTRA = Trim(Mid(templine, 23, 8))
                            POOLACT = Trim(Mid(templine, 31, 9))
                            POOLIDLE = Trim(Mid(templine, 40, 10))
                            POOLTRAF = Trim(Mid(templine, 50, 8))
                        End If
                    ElseIf UCase(Trim(templine)) = "SUBPOOL                       SUBACT   SUBIDLE   SUBTRAF" Then
                        'SUBPOOL                       SUBACT   SUBIDLE   SUBTRAF
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            SUBPOOL = Trim(Mid(templine, 1, 30))
                            SUBACT = Trim(Mid(templine, 31, 9))
                            SUBIDLE = Trim(Mid(templine, 40, 10))
                            SUBTRAF = Trim(Mid(templine, 50, 7))
                        End If
                        outputFile.writeline(BSCNAME & "," & TRAPOOL & "," & CHRATE & "," & SPV & "," & RNOTRA & "," & POOLACT & "," & POOLIDLE & "," & POOLTRAF _
                     & "," & SUBPOOL & "," & SUBACT & "," & SUBIDLE & "," & SUBTRAF)
                        TRAPOOL = ""
                        CHRATE = ""
                        SPV = ""
                        RNOTRA = ""
                        POOLACT = ""
                        POOLIDLE = ""
                        POOLTRAF = ""
                        SUBPOOL = ""
                        SUBACT = ""
                        SUBIDLE = ""
                        SUBTRAF = ""
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRXAPP_RXOTG(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim MO As String, DEV As String, DCP As String, APUSAGE As String, APSTATE As String, RXAPP_64K As String, TEI As String

        MO = ""
        DEV = ""
        DCP = ""
        APUSAGE = ""
        APSTATE = ""
        RXAPP_64K = ""
        TEI = ""

        'MO
        'DEV            DCP  APUSAGE  APSTATE           64K  TEI

        filepath = outPath & "RXAPP_RXOTG_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "MO" & "," & "DEV" & "," & "DCP" & "," & "APUSAGE" & "," & "APSTATE" & "," & "RXAPP_64K" & "," & "TEI")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "MO" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            MO = Trim(templine)
                        End If
                    ElseIf UCase(Trim(templine)) = "DEV            DCP  APUSAGE  APSTATE           64K  TEI" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            DEV = Trim(Mid(templine, 1, 15))
                            DCP = Trim(Mid(templine, 16, 5))
                            APUSAGE = Trim(Mid(templine, 21, 9))
                            APSTATE = Trim(Mid(templine, 30, 18))
                            RXAPP_64K = Trim(Mid(templine, 48, 5))
                            TEI = Trim(Mid(templine, 53, 15))
                            outputFile.writeLine(BSCNAME & "," & MO & "," & DEV & "," & DCP & "," & APUSAGE & "," & APSTATE & "," & RXAPP_64K & "," & TEI)
                        End If
                    ElseIf UCase(Trim(templine)) <> "RADIO X-CEIVER ADMINISTRATION" And UCase(Trim(templine)) <> "ABIS PATH STATUS" And Len(Trim(Mid(templine, 1, 15))) >= 6 And Len(Trim(Mid(templine, 1, 15))) <= 12 Then
                        DEV = Trim(Mid(templine, 1, 15))
                        DCP = Trim(Mid(templine, 16, 5))
                        APUSAGE = Trim(Mid(templine, 21, 9))
                        APSTATE = Trim(Mid(templine, 30, 18))
                        RXAPP_64K = Trim(Mid(templine, 48, 5))
                        TEI = Trim(Mid(templine, 53, 15))
                        outputFile.writeLine(BSCNAME & "," & MO & "," & DEV & "," & DCP & "," & APUSAGE & "," & APSTATE & "," & RXAPP_64K & "," & TEI)
                    ElseIf UCase(Trim(templine)) <> "RADIO X-CEIVER ADMINISTRATION" And UCase(Trim(templine)) <> "ABIS PATH STATUS" And Len(Trim(Mid(templine, 1, 15))) = 0 Then
                        APUSAGE = Trim(Mid(templine, 21, 9))
                        APSTATE = Trim(Mid(templine, 30, 18))
                        RXAPP_64K = Trim(Mid(templine, 48, 5))
                        TEI = Trim(Mid(templine, 53, 15))
                        outputFile.writeLine(BSCNAME & "," & MO & "," & DEV & "," & DCP & "," & APUSAGE & "," & APSTATE & "," & RXAPP_64K & "," & TEI)
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRXASP_RXOTG(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'MO                SCGR  SC         RSITE           ALARM SITUATION
        filepath = outPath & "RXASP_RXOTG_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "MO" & "," & "SCGR" & "," & "SC" & "," & "RSITE" & "," & "ALARM SITUATION")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Strings.left(Trim(templine), 20)) = "MO                SC" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'MO                SCGR  SC         RSITE           ALARM SITUATION
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 18)) & "," & Trim(Mid(templine, 19, 6)) & "," & Trim(Mid(templine, 25, 11)) & "," & Trim(Mid(templine, 36, 16)) & "," & Trim(Mid(templine, 52, 25)))
                        End If
                    ElseIf UCase(Strings.left(Trim(templine), 10)) <> "NO ALARM S" And UCase(Trim(templine)) <> "RADIO X-CEIVER ADMINISTRATION" And UCase(Trim(templine)) <> "MANAGED OBJECT ALARM SITUATIONS" Then
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 18)) & "," & Trim(Mid(templine, 19, 6)) & "," & Trim(Mid(templine, 25, 11)) & "," & Trim(Mid(templine, 36, 16)) & "," & Trim(Mid(templine, 52, 25)))
                    ElseIf UCase(Strings.left(Trim(templine), 10)) = "NO ALARM S" Then
                        outputFile.writeLINE(BSCNAME & "," & Trim(templine))
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRXBFP_RXOTG(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim MO As String, AHOP As String, BTSAHOP As String, FFTAD As String, BTSFFTAD As String, PSUPS As String, BTSPSUPS As String
        Dim TSPS As String, BTSTSPS As String, CSTMA As String, BTSCSTMA As String, AISGTMA As String, BTSAISGTMA As String
        Dim AISGRET As String, BTSAISGRET As String

        'MO               AHOP   BTSAHOP  FFTAD  BTSFFTAD  PSUPS  BTSPSUPS
        '                 TSPS  BTSTSPS  CSTMA  BTSCSTMA  AISGTMA  BTSAISGTMA
        '                 AISGRET  BTSAISGRET
        MO = ""
        AHOP = ""
        BTSAHOP = ""
        FFTAD = ""
        BTSFFTAD = ""
        PSUPS = ""
        BTSPSUPS = ""
        TSPS = ""
        BTSTSPS = ""
        CSTMA = ""
        BTSCSTMA = ""
        AISGTMA = ""
        BTSAISGTMA = ""
        AISGRET = ""
        BTSAISGRET = ""

        filepath = outPath & "RXBFP_RXOTG_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "MO" & "," & "AHOP" & "," & "BTSAHOP" & "," & "FFTAD" & "," & "BTSFFTAD" & "," & "PSUPS" & "," & "BTSPSUPS" _
                                 & "," & "TSPS" & "," & "BTSTSPS" & "," & "CSTMA" & "," & "BTSCSTMA" & "," & "AISGTMA" & "," & "BTSAISGTMA" _
                                 & "," & "AISGRET" & "," & "BTSAISGRET")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        If MO <> "" Then
                            outputFile.writeline(BSCNAME & "," & MO & "," & AHOP & "," & BTSAHOP & "," & FFTAD & "," & BTSFFTAD & "," & PSUPS & "," & BTSPSUPS _
                                    & "," & TSPS & "," & BTSTSPS & "," & CSTMA & "," & BTSCSTMA & "," & AISGTMA & "," & BTSAISGTMA _
                                    & "," & AISGRET & "," & BTSAISGRET)
                            MO = ""
                            AHOP = ""
                            BTSAHOP = ""
                            FFTAD = ""
                            BTSFFTAD = ""
                            PSUPS = ""
                            BTSPSUPS = ""
                            TSPS = ""
                            BTSTSPS = ""
                            CSTMA = ""
                            BTSCSTMA = ""
                            AISGTMA = ""
                            BTSAISGTMA = ""
                            AISGRET = ""
                            BTSAISGRET = ""
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If MO <> "" Then
                            outputFile.writeline(BSCNAME & "," & MO & "," & AHOP & "," & BTSAHOP & "," & FFTAD & "," & BTSFFTAD & "," & PSUPS & "," & BTSPSUPS _
                                    & "," & TSPS & "," & BTSTSPS & "," & CSTMA & "," & BTSCSTMA & "," & AISGTMA & "," & BTSAISGTMA _
                                    & "," & AISGRET & "," & BTSAISGRET)
                            MO = ""
                            AHOP = ""
                            BTSAHOP = ""
                            FFTAD = ""
                            BTSFFTAD = ""
                            PSUPS = ""
                            BTSPSUPS = ""
                            TSPS = ""
                            BTSTSPS = ""
                            CSTMA = ""
                            BTSCSTMA = ""
                            AISGTMA = ""
                            BTSAISGTMA = ""
                            AISGRET = ""
                            BTSAISGRET = ""
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "MO               AHOP   BTSAHOP  FFTAD  BTSFFTAD  PSUPS  BTSPSUPS" Then
                        'MO               AHOP   BTSAHOP  FFTAD  BTSFFTAD  PSUPS  BTSPSUPS
                        If MO <> "" Then
                            outputFile.writeline(BSCNAME & "," & MO & "," & AHOP & "," & BTSAHOP & "," & FFTAD & "," & BTSFFTAD & "," & PSUPS & "," & BTSPSUPS _
                                    & "," & TSPS & "," & BTSTSPS & "," & CSTMA & "," & BTSCSTMA & "," & AISGTMA & "," & BTSAISGTMA _
                                    & "," & AISGRET & "," & BTSAISGRET)
                            MO = ""
                            AHOP = ""
                            BTSAHOP = ""
                            FFTAD = ""
                            BTSFFTAD = ""
                            PSUPS = ""
                            BTSPSUPS = ""
                            TSPS = ""
                            BTSTSPS = ""
                            CSTMA = ""
                            BTSCSTMA = ""
                            AISGTMA = ""
                            BTSAISGTMA = ""
                            AISGRET = ""
                            BTSAISGRET = ""
                        End If
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            MO = Trim(Mid(templine, 1, 17))
                            AHOP = Trim(Mid(templine, 18, 7))
                            BTSAHOP = Trim(Mid(templine, 25, 9))
                            FFTAD = Trim(Mid(templine, 34, 7))
                            BTSFFTAD = Trim(Mid(templine, 41, 10))
                            PSUPS = Trim(Mid(templine, 51, 7))
                            BTSPSUPS = Trim(Mid(templine, 58, 8))
                        End If
                    ElseIf UCase(Trim(templine)) = "TSPS  BTSTSPS  CSTMA  BTSCSTMA  AISGTMA  BTSAISGTMA" Then
                        '                 TSPS  BTSTSPS  CSTMA  BTSCSTMA  AISGTMA  BTSAISGTMA
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            TSPS = Trim(Mid(templine, 18, 6))
                            BTSTSPS = Trim(Mid(templine, 24, 9))
                            CSTMA = Trim(Mid(templine, 31, 7))
                            BTSCSTMA = Trim(Mid(templine, 38, 10))
                            AISGTMA = Trim(Mid(templine, 48, 9))
                            BTSAISGTMA = Trim(Mid(templine, 57, 10))
                        End If
                    ElseIf UCase(Trim(templine)) = "AISGRET  BTSAISGRET" Then
                        '                 AISGRET  BTSAISGRET
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            AISGRET = Trim(Mid(templine, 18, 9))
                            BTSAISGRET = Trim(Mid(templine, 27, 10))
                        End If
                        outputFile.writeline(BSCNAME & "," & MO & "," & AHOP & "," & BTSAHOP & "," & FFTAD & "," & BTSFFTAD & "," & PSUPS & "," & BTSPSUPS _
                                    & "," & TSPS & "," & BTSTSPS & "," & CSTMA & "," & BTSCSTMA & "," & AISGTMA & "," & BTSAISGTMA _
                                    & "," & AISGRET & "," & BTSAISGRET)
                        MO = ""
                        AHOP = ""
                        BTSAHOP = ""
                        FFTAD = ""
                        BTSFFTAD = ""
                        PSUPS = ""
                        BTSPSUPS = ""
                        TSPS = ""
                        BTSTSPS = ""
                        CSTMA = ""
                        BTSCSTMA = ""
                        AISGTMA = ""
                        BTSAISGTMA = ""
                        AISGRET = ""
                        BTSAISGRET = ""
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRXELP_RXOTG(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim MO As String, BTSSWVER As String, ELP_DATE As String, TIME As String, BTS As String
        Dim REPLMAP As String, AMAP1 As String
        Dim BMAP1 As String, AMAP2 As String
        Dim EXT1BMAP As String, EXT2BMAP As String, EXT2BMAPEXT As String

        'MO                BTSSWVER         DATE      TIME       BTS
        'REPLMAP                       1AMAP
        '1BMAP                         2AMAP
        'EXT1BMAP  EXT2BMAP  EXT2BMAPEXT
        MO = ""
        BTSSWVER = ""
        ELP_DATE = ""
        TIME = ""
        BTS = ""
        REPLMAP = ""
        AMAP1 = ""
        BMAP1 = ""
        AMAP2 = ""
        EXT1BMAP = ""
        EXT2BMAP = ""
        EXT2BMAPEXT = ""


        filepath = outPath & "RXELP_RXOTG_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "MO" & "," & "BTSSWVER" & "," & "DATE" & "," & "TIME" & "," & "BTS" _
                                 & "," & "REPLMAP" & "," & "1AMAP" _
                                 & "," & "1BMAP" & "," & "2AMAP" _
                                 & "," & "EXT1BMAP" & "," & "EXT2BMAP" & "," & "EXT2BMAPEXT")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If MO <> "" Then
                            outputFile.writeline(BSCNAME & "," & MO & "," & BTSSWVER & "," & ELP_DATE & "," & TIME & "," & BTS _
                     & "," & REPLMAP & "," & AMAP1 _
                     & "," & BMAP1 & "," & AMAP2 _
                     & "," & EXT1BMAP & "," & EXT2BMAP & "," & EXT2BMAPEXT)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If MO <> "" Then
                            outputFile.writeline(BSCNAME & "," & MO & "," & BTSSWVER & "," & ELP_DATE & "," & TIME & "," & BTS _
                     & "," & REPLMAP & "," & AMAP1 _
                     & "," & BMAP1 & "," & AMAP2 _
                     & "," & EXT1BMAP & "," & EXT2BMAP & "," & EXT2BMAPEXT)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "MO                BTSSWVER         DATE      TIME       BTS" Then
                        'MO                BTSSWVER         DATE      TIME       BTS
                        templine = getNowLine()
                        If MO <> "" Then
                            outputFile.writeline(BSCNAME & "," & MO & "," & BTSSWVER & "," & ELP_DATE & "," & TIME & "," & BTS _
                                                 & "," & REPLMAP & "," & AMAP1 _
                                                 & "," & BMAP1 & "," & AMAP2 _
                                                 & "," & EXT1BMAP & "," & EXT2BMAP & "," & EXT2BMAPEXT)
                        End If
                        If Trim(templine) <> "NONE" Then
                            MO = Trim(Mid(templine, 1, 18))
                            BTSSWVER = Trim(Mid(templine, 19, 17))
                            ELP_DATE = Trim(Mid(templine, 36, 10))
                            TIME = Trim(Mid(templine, 46, 11))
                            BTS = Trim(Mid(templine, 57, 7))
                        End If
                    ElseIf UCase(Trim(templine)) = "REPLMAP                       1AMAP" Then
                        'REPLMAP                       1AMAP
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            REPLMAP = Trim(Mid(templine, 1, 30))
                            AMAP1 = Trim(Mid(templine, 31, 30))
                        End If
                    ElseIf UCase(Trim(templine)) = "1BMAP                         2AMAP" Then
                        '1BMAP                         2AMAP
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            BMAP1 = Trim(Mid(templine, 1, 30))
                            AMAP2 = Trim(Mid(templine, 31, 30))
                        End If
                    ElseIf UCase(Trim(templine)) = "EXT1BMAP  EXT2BMAP  EXT2BMAPEXT" Then
                        'EXT1BMAP  EXT2BMAP  EXT2BMAPEXT
                        templine = getNowLine()
                        If Trim(templine) <> "NONE" Then
                            EXT1BMAP = Trim(Mid(templine, 1, 10))
                            EXT2BMAP = Trim(Mid(templine, 11, 10))
                            EXT2BMAPEXT = Trim(Mid(templine, 21, 11))
                        End If
                        outputFile.writeline(BSCNAME & "," & MO & "," & BTSSWVER & "," & ELP_DATE & "," & TIME & "," & BTS _
                                             & "," & REPLMAP & "," & AMAP1 _
                                             & "," & BMAP1 & "," & AMAP2 _
                                             & "," & EXT1BMAP & "," & EXT2BMAP & "," & EXT2BMAPEXT)
                        MO = ""
                        BTSSWVER = ""
                        ELP_DATE = ""
                        TIME = ""
                        BTS = ""
                        REPLMAP = ""
                        AMAP1 = ""
                        BMAP1 = ""
                        AMAP2 = ""
                        EXT1BMAP = ""
                        EXT2BMAP = ""
                        EXT2BMAPEXT = ""
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRXMDP_RXOTRX(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'MO                DEVS           DEVT           SDEV  RPS   RPT
        filepath = outPath & "RXMDP_RXOTRX_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "MO" & "," & "DEVS" & "," & "DEVT" & "," & "SDEV" & "," & "RPS" & "," & "RPT")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "MO                DEVS           DEVT           SDEV  RPS   RPT" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 18)) & "," & Trim(Mid(templine, 19, 15)) & "," & Trim(Mid(templine, 34, 15)) & "," & Trim(Mid(templine, 49, 6)) & "," & Trim(Mid(templine, 55, 6)) & "," & Trim(Mid(templine, 61, 6)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "RADIO X-CEIVER ADMINISTRATION" And UCase(Trim(templine)) <> "MANAGED OBJECT DEVICE INFORMATION" And Len(Trim(Mid(templine, 1, 18))) >= 10 And Len(Trim(Mid(templine, 1, 18))) < 16 Then
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 18)) & "," & Trim(Mid(templine, 19, 15)) & "," & Trim(Mid(templine, 34, 15)) & "," & Trim(Mid(templine, 49, 6)) & "," & Trim(Mid(templine, 55, 6)) & "," & Trim(Mid(templine, 61, 6)))
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRXMDP_RXOCF(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'MO                DEVS           DEVT           SDEV  RPS   RPT
        filepath = outPath & "RXMDP_RXOCF_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "MO" & "," & "DEVS" & "," & "DEVT" & "," & "SDEV" & "," & "RPS" & "," & "RPT")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    'MO                DEVS           DEVT           SDEV  RPS   RPT
                    If UCase(Trim(templine)) = "MO                DEVS           DEVT           SDEV  RPS   RPT" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 18)) & "," & Trim(Mid(templine, 19, 15)) & "," & Trim(Mid(templine, 34, 15)) & "," & Trim(Mid(templine, 49, 6)) & "," & Trim(Mid(templine, 55, 6)) & "," & Trim(Mid(templine, 61, 6)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "RADIO X-CEIVER ADMINISTRATION" And UCase(Trim(templine)) <> "MANAGED OBJECT DEVICE INFORMATION" And Len(Trim(Mid(templine, 1, 18))) >= 7 And Len(Trim(Mid(templine, 1, 18))) < 16 Then
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 18)) & "," & Trim(Mid(templine, 19, 15)) & "," & Trim(Mid(templine, 34, 15)) & "," & Trim(Mid(templine, 49, 6)) & "," & Trim(Mid(templine, 55, 6)) & "," & Trim(Mid(templine, 61, 6)))
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doRXMFP_FAULTY(sender As Object, type As String)
        Dim outputFile1 As Object
        Dim filepath1 As String
        Dim templine As String
        Dim tempFaultyCode As String
        Dim MO As String, BTSSWVER As String, STATE As String, LEFT_CODE As String, RIGHT_CODE As String
        Dim EXT_LEFT_CODE As String, EXT_RIGHT_CODE As String, Last_FC_MO As String
        Dim BLSTATE As String, INTERCNT As String, CONCNT As String, CONERRCNT As String, LASTFLT As String, LFREASON As String
        Dim FAULT_CODES As String, FAULT_NAME As String, EXTERNAL_FAULT_CODES As String, EXTERNAL_FAULT_NAME As String, REPLACEMENT_UNITS As String, REPLACEMENT_UNITS_NAME As String
        ' STATE  BLSTATE  INTERCNT  CONCNT  CONERRCNT  LASTFLT   LFREASON
        Dim temp_arr As Object
        tempFaultyCode = ""
        MO = ""
        Last_FC_MO = ""
        BTSSWVER = ""
        LEFT_CODE = ""
        RIGHT_CODE = ""
        EXT_LEFT_CODE = ""
        EXT_RIGHT_CODE = ""
        FAULT_CODES = ""
        FAULT_NAME = ""
        EXTERNAL_FAULT_CODES = ""
        EXTERNAL_FAULT_NAME = ""
        REPLACEMENT_UNITS = ""
        REPLACEMENT_UNITS_NAME = ""
        STATE = ""
        BLSTATE = ""
        INTERCNT = ""
        CONCNT = ""
        CONERRCNT = ""
        LASTFLT = ""
        LFREASON = ""

        filepath1 = outPath & "RXMFP_FAULT_CODES_" & CDDtime & ".csv"
        If fso.fileExists(filepath1) Then
            outputFile1 = fso.OpenTextFile(filepath1, 8, 1, 0)
        Else
            outputFile1 = fso.OpenTextFile(filepath1, 2, 1, 0)
            outputFile1.writeLine("BSC" & "," & "MO" & "," & "STATE" & "," & "FAULT CODES" & "," & "告警名称" & "," & "EXTERNAL FAULT CODES" & "," & "外部告警名称" & "," & "REPLACEMENT UNITS" & "," & "替换单元名称" _
                                 & "," & "BLSTATE" & "," & "INTERCNT" & "," & "CONCNT" & "," & "CONERRCNT" & "," & "LASTFLT" & "," & "LFREASON")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If Last_FC_MO <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        outputFile1.close()
                        outputFile1 = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        If Last_FC_MO <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        outputFile1.close()
                        outputFile1 = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "MO                   BTSSWVER" Then
                        If Last_FC_MO <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        templine = getNowLine()
                        'MO                   BTSSWVER          
                        MO = Trim(Mid(templine, 1, 21))
                        Last_FC_MO = Trim(Mid(templine, 1, 21))
                        BTSSWVER = Trim(Mid(templine, 22, 18))
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "STATE  BLSTATE" Then
                        ' STATE  BLSTATE  INTERCNT  CONCNT  CONERRCNT  LASTFLT   LFREASON
                        templine = getNowLine()
                        STATE = Trim(Mid(templine, 1, 7))
                        BLSTATE = Trim(Mid(templine, 8, 9))
                        INTERCNT = Trim(Mid(templine, 17, 10))
                        CONCNT = Trim(Mid(templine, 27, 8))
                        CONERRCNT = Trim(Mid(templine, 35, 11))
                        LASTFLT = Trim(Mid(templine, 46, 10))
                        LFREASON = Trim(Mid(templine, 56, 10))
                    ElseIf UCase(Strings.left(Trim(templine), 11)) = "FAULT CODES" Then
                        LEFT_CODE = UCase(Trim(Mid(templine, 19, 10)))
                        tempFaultyCode = type & UCase(Trim(Mid(templine, 19, 10)))
                        FAULT_CODES = LEFT_CODE
                        templine = getNowLine()
                        RIGHT_CODE = Trim(templine)
                        If InStr(RIGHT_CODE, "  ") Then
                            temp_arr = Split(RIGHT_CODE, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & FAULT_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & RIGHT_CODE
                            FAULT_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        FAULT_CODES = LEFT_CODE & ":" & RIGHT_CODE
                    ElseIf UCase(Strings.left(Trim(templine), 20)) = "EXTERNAL FAULT CODES" Then
                        EXT_LEFT_CODE = UCase(Trim(Mid(templine, 28, 10)))
                        tempFaultyCode = type & "E" & UCase(Trim(Mid(templine, 28, 10)))
                        templine = getNowLine()
                        EXT_RIGHT_CODE = Trim(templine)
                        If InStr(EXT_RIGHT_CODE, "  ") Then
                            temp_arr = Split(EXT_RIGHT_CODE, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    EXTERNAL_FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    EXTERNAL_FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & EXTERNAL_FAULT_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & EXT_RIGHT_CODE
                            EXTERNAL_FAULT_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        EXTERNAL_FAULT_CODES = EXT_LEFT_CODE & ":" & EXT_RIGHT_CODE
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "REPLACEMENT UN" Then
                        templine = Trim(getNowLine())
                        tempFaultyCode = type & "RU"
                        If InStr(templine, "  ") Then
                            temp_arr = Split(templine, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    REPLACEMENT_UNITS_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    REPLACEMENT_UNITS_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & REPLACEMENT_UNITS_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & Trim(templine)
                            REPLACEMENT_UNITS_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        REPLACEMENT_UNITS = Trim(templine)
                        outputFile1.writeLine(BSCNAME & "," & Last_FC_MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                              & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                        MO = ""
                        Last_FC_MO = ""
                        BTSSWVER = ""
                        LEFT_CODE = ""
                        RIGHT_CODE = ""
                        EXT_LEFT_CODE = ""
                        EXT_RIGHT_CODE = ""
                        FAULT_CODES = ""
                        FAULT_NAME = ""
                        EXTERNAL_FAULT_CODES = ""
                        EXTERNAL_FAULT_NAME = ""
                        REPLACEMENT_UNITS = ""
                        REPLACEMENT_UNITS_NAME = ""
                        STATE = ""
                        BLSTATE = ""
                        INTERCNT = ""
                        CONCNT = ""
                        CONERRCNT = ""
                        LASTFLT = ""
                        LFREASON = ""
                    End If
                End If
            End If
        Loop
        outputFile1.close()
        outputFile1 = Nothing
    End Sub
    Private Sub doRXMFP_RXOCON(sender As Object, type As String)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim outputFile1 As Object
        Dim filepath1 As String
        Dim tempFaultyCode As String
        Dim EXT_LEFT_CODE As String, EXT_RIGHT_CODE As String, Last_FC_MO As String
        Dim BLSTATE As String, INTERCNT As String, CONCNT As String, CONERRCNT As String, LASTFLT As String, LFREASON As String
        Dim FAULT_CODES As String, FAULT_NAME As String, EXTERNAL_FAULT_CODES As String, EXTERNAL_FAULT_NAME As String, REPLACEMENT_UNITS As String, REPLACEMENT_UNITS_NAME As String
        Dim MO As String, BTSSWVER As String, STATE As String, LEFT_CODE As String, RIGHT_CODE As String, Last_MO As String
        Dim RU As String, RUREVISION As String, RUSERIALNO As String, RUPOSITION As String, RULOGICALID As String
        Dim temp_arr As Object
        MO = ""
        Last_MO = ""
        BTSSWVER = ""
        RU = ""
        RUREVISION = ""
        RUSERIALNO = ""
        RUPOSITION = ""
        RULOGICALID = ""
        STATE = ""
        LEFT_CODE = ""
        RIGHT_CODE = ""
        tempFaultyCode = ""
        Last_FC_MO = ""
        EXT_LEFT_CODE = ""
        EXT_RIGHT_CODE = ""
        FAULT_CODES = ""
        FAULT_NAME = ""
        EXTERNAL_FAULT_CODES = ""
        EXTERNAL_FAULT_NAME = ""
        REPLACEMENT_UNITS = ""
        REPLACEMENT_UNITS_NAME = ""
        BLSTATE = ""
        INTERCNT = ""
        CONCNT = ""
        CONERRCNT = ""
        LASTFLT = ""
        LFREASON = ""

        filepath = outPath & "RXMFP_RXOCON_" & CDDtime & ".csv"
        filepath1 = outPath & "RXMFP_FAULT_CODES_" & CDDtime & ".csv"

        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeLine("BSC" & "," & "MO" & "," & "BTSSWVER" & "," & "RU" & "," & "RUREVISION" & "," & "RUSERIALNO" & "," & "RUPOSITION" & "," & "RULOGICALID")
        End If
        If fso.fileExists(filepath1) Then
            outputFile1 = fso.OpenTextFile(filepath1, 8, 1, 0)
        Else
            outputFile1 = fso.OpenTextFile(filepath1, 2, 1, 0)
            outputFile1.writeLine("BSC" & "," & "MO" & "," & "STATE" & "," & "FAULT CODES" & "," & "告警名称" & "," & "EXTERNAL FAULT CODES" & "," & "外部告警名称" & "," & "REPLACEMENT UNITS" & "," & "替换单元名称" _
                                 & "," & "BLSTATE" & "," & "INTERCNT" & "," & "CONCNT" & "," & "CONERRCNT" & "," & "LASTFLT" & "," & "LFREASON")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If RU <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                        End If
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        outputFile1.close()
                        outputFile1 = Nothing
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        If RU <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                        End If
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        outputFile1.close()
                        outputFile1 = Nothing
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "MO                   BTSSWVER" Then
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        'MO                   BTSSWVER          
                        'RU  RUREVISION                           RUSERIALNO
                        '    RUPOSITION                           RULOGICALID

                        templine = getNowLine()
                        MO = Trim(Mid(templine, 1, 21))
                        Last_FC_MO = Trim(Mid(templine, 1, 21))
                        BTSSWVER = Trim(Mid(templine, 22, 18))
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "RU  RUREVISION" Then
                        templine = getNowLine()
                        If templine <> "" Then
                            RU = Trim(Mid(templine, 1, 4))
                            RUREVISION = Trim(Mid(templine, 5, 37))
                            RUSERIALNO = Trim(Mid(templine, 42, 32))
                        Else
                            If Last_MO <> MO Then
                                outputFile.writeLine(BSCNAME & "," & MO & "," & BTSSWVER)
                                Last_MO = MO
                            End If
                        End If
                    ElseIf UCase(Strings.left(templine, 14)) = "    RUPOSITION" Then
                        templine = getNowLine()
                        If templine <> "" Then
                            RUPOSITION = Trim(Mid(templine, 5, 36))
                            RULOGICALID = Trim(Mid(templine, 41, 32))
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                            RU = ""
                            RUREVISION = ""
                            RUSERIALNO = ""
                            RUPOSITION = ""
                            RULOGICALID = ""
                        End If
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "STATE  BLSTATE" Then
                        ' STATE  BLSTATE  INTERCNT  CONCNT  CONERRCNT  LASTFLT   LFREASON
                        templine = getNowLine()
                        STATE = Trim(Mid(templine, 1, 7))
                        BLSTATE = Trim(Mid(templine, 8, 9))
                        INTERCNT = Trim(Mid(templine, 17, 10))
                        CONCNT = Trim(Mid(templine, 27, 8))
                        CONERRCNT = Trim(Mid(templine, 35, 11))
                        LASTFLT = Trim(Mid(templine, 46, 10))
                        LFREASON = Trim(Mid(templine, 56, 10))
                    ElseIf UCase(Strings.left(Trim(templine), 11)) = "FAULT CODES" Then
                        LEFT_CODE = UCase(Trim(Mid(templine, 19, 10)))
                        tempFaultyCode = type & UCase(Trim(Mid(templine, 19, 10)))
                        FAULT_CODES = LEFT_CODE
                        templine = getNowLine()
                        RIGHT_CODE = Trim(templine)
                        If InStr(RIGHT_CODE, "  ") Then
                            temp_arr = Split(RIGHT_CODE, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & FAULT_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & RIGHT_CODE
                            FAULT_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        FAULT_CODES = LEFT_CODE & ":" & RIGHT_CODE
                    ElseIf UCase(Strings.left(Trim(templine), 20)) = "EXTERNAL FAULT CODES" Then
                        EXT_LEFT_CODE = UCase(Trim(Mid(templine, 28, 10)))
                        tempFaultyCode = type & "E" & UCase(Trim(Mid(templine, 28, 10)))
                        templine = getNowLine()
                        EXT_RIGHT_CODE = Trim(templine)
                        If InStr(EXT_RIGHT_CODE, "  ") Then
                            temp_arr = Split(EXT_RIGHT_CODE, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    EXTERNAL_FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    EXTERNAL_FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & EXTERNAL_FAULT_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & EXT_RIGHT_CODE
                            EXTERNAL_FAULT_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        EXTERNAL_FAULT_CODES = EXT_LEFT_CODE & ":" & EXT_RIGHT_CODE
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "REPLACEMENT UN" Then
                        templine = Trim(getNowLine())
                        tempFaultyCode = type & "RU"
                        If InStr(templine, "  ") Then
                            temp_arr = Split(templine, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    REPLACEMENT_UNITS_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    REPLACEMENT_UNITS_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & REPLACEMENT_UNITS_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & Trim(templine)
                            REPLACEMENT_UNITS_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        REPLACEMENT_UNITS = Trim(templine)
                        outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                              & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                        MO = ""
                        Last_FC_MO = ""
                        BTSSWVER = ""
                        LEFT_CODE = ""
                        RIGHT_CODE = ""
                        EXT_LEFT_CODE = ""
                        EXT_RIGHT_CODE = ""
                        FAULT_CODES = ""
                        FAULT_NAME = ""
                        EXTERNAL_FAULT_CODES = ""
                        EXTERNAL_FAULT_NAME = ""
                        REPLACEMENT_UNITS = ""
                        REPLACEMENT_UNITS_NAME = ""
                        STATE = ""
                        BLSTATE = ""
                        INTERCNT = ""
                        CONCNT = ""
                        CONERRCNT = ""
                        LASTFLT = ""
                        LFREASON = ""
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
        outputFile1.CLOSE()
        outputFile1 = Nothing
    End Sub
    Private Sub doRXMFP_RXOCF(sender As Object, type As String)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim outputFile1 As Object
        Dim filepath1 As String
        Dim tempFaultyCode As String
        Dim EXT_LEFT_CODE As String, EXT_RIGHT_CODE As String, Last_FC_MO As String
        Dim BLSTATE As String, INTERCNT As String, CONCNT As String, CONERRCNT As String, LASTFLT As String, LFREASON As String
        Dim FAULT_CODES As String, FAULT_NAME As String, EXTERNAL_FAULT_CODES As String, EXTERNAL_FAULT_NAME As String, REPLACEMENT_UNITS As String, REPLACEMENT_UNITS_NAME As String
        Dim MO As String, BTSSWVER As String, STATE As String, LEFT_CODE As String, RIGHT_CODE As String, Last_MO As String
        Dim RU As String, RUREVISION As String, RUSERIALNO As String, RUPOSITION As String, RULOGICALID As String
        Dim BSC_MO_TEMP As String, CDU_TYPE_TEMP As String
        Dim temp_arr As Object
        MO = ""
        Last_MO = ""
        BTSSWVER = ""
        RU = ""
        RUREVISION = ""
        RUSERIALNO = ""
        RUPOSITION = ""
        RULOGICALID = ""
        STATE = ""
        LEFT_CODE = ""
        RIGHT_CODE = ""
        BSC_MO_TEMP = ""
        tempFaultyCode = ""
        Last_FC_MO = ""
        EXT_LEFT_CODE = ""
        EXT_RIGHT_CODE = ""
        FAULT_CODES = ""
        FAULT_NAME = ""
        EXTERNAL_FAULT_CODES = ""
        EXTERNAL_FAULT_NAME = ""
        REPLACEMENT_UNITS = ""
        REPLACEMENT_UNITS_NAME = ""
        BLSTATE = ""
        INTERCNT = ""
        CONCNT = ""
        CONERRCNT = ""
        LASTFLT = ""
        LFREASON = ""

        filepath = outPath & "RXMFP_RXOCF_" & CDDtime & ".csv"
        filepath1 = outPath & "RXMFP_FAULT_CODES_" & CDDtime & ".csv"

        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeLine("BSC" & "," & "MO" & "," & "BTSSWVER" & "," & "RU" & "," & "RUREVISION" & "," & "RUSERIALNO" & "," & "RUPOSITION" & "," & "RULOGICALID")
        End If
        If fso.fileExists(filepath1) Then
            outputFile1 = fso.OpenTextFile(filepath1, 8, 1, 0)
        Else
            outputFile1 = fso.OpenTextFile(filepath1, 2, 1, 0)
            outputFile1.writeLine("BSC" & "," & "MO" & "," & "STATE" & "," & "FAULT CODES" & "," & "告警名称" & "," & "EXTERNAL FAULT CODES" & "," & "外部告警名称" & "," & "REPLACEMENT UNITS" & "," & "替换单元名称" _
                                 & "," & "BLSTATE" & "," & "INTERCNT" & "," & "CONCNT" & "," & "CONERRCNT" & "," & "LASTFLT" & "," & "LFREASON")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If RU <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                        End If
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        outputFile1.close()
                        outputFile1 = Nothing
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        If RU <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                        End If
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        outputFile1.close()
                        outputFile1 = Nothing
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "MO                   BTSSWVER" Then
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        'MO                   BTSSWVER          
                        'RU  RUREVISION                           RUSERIALNO
                        '    RUPOSITION                           RULOGICALID

                        templine = getNowLine()
                        MO = Trim(Mid(templine, 1, 21))
                        Last_FC_MO = Trim(Mid(templine, 1, 21))
                        BSC_MO_TEMP = BSCNAME & "#" & Trim(Replace(MO, "RXOCF", "RXOTG"))
                        BTSSWVER = Trim(Mid(templine, 22, 18))
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "RU  RUREVISION" Then
                        templine = getNowLine()
                        If templine <> "" Then
                            RU = Trim(Mid(templine, 1, 4))
                            RUREVISION = Trim(Mid(templine, 5, 37))
                            RUSERIALNO = Trim(Mid(templine, 42, 32))
                        Else
                            If Last_MO <> MO Then
                                outputFile.writeLine(BSCNAME & "," & MO & "," & BTSSWVER)
                                Last_MO = MO
                            End If
                        End If
                    ElseIf UCase(Strings.left(templine, 14)) = "    RUPOSITION" Then
                        templine = getNowLine()
                        If templine <> "" Then
                            RUPOSITION = Trim(Mid(templine, 5, 36))
                            RULOGICALID = Trim(Mid(templine, 41, 32))
                            CDU_TYPE_TEMP = Trim(Mid(templine, 41, 29))
                            I = I + 1
                            MFP_MO_CDUtype_DIC.Add(I, BSC_MO_TEMP & "@" & CDU_TYPE_TEMP)

                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                            RU = ""
                            RUREVISION = ""
                            RUSERIALNO = ""
                            RUPOSITION = ""
                            RULOGICALID = ""
                        End If
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "STATE  BLSTATE" Then
                        ' STATE  BLSTATE  INTERCNT  CONCNT  CONERRCNT  LASTFLT   LFREASON
                        templine = getNowLine()
                        STATE = Trim(Mid(templine, 1, 7))
                        BLSTATE = Trim(Mid(templine, 8, 9))
                        INTERCNT = Trim(Mid(templine, 17, 10))
                        CONCNT = Trim(Mid(templine, 27, 8))
                        CONERRCNT = Trim(Mid(templine, 35, 11))
                        LASTFLT = Trim(Mid(templine, 46, 10))
                        LFREASON = Trim(Mid(templine, 56, 10))
                    ElseIf UCase(Strings.left(Trim(templine), 11)) = "FAULT CODES" Then
                        LEFT_CODE = UCase(Trim(Mid(templine, 19, 10)))
                        tempFaultyCode = Type & UCase(Trim(Mid(templine, 19, 10)))
                        FAULT_CODES = LEFT_CODE
                        templine = getNowLine()
                        RIGHT_CODE = Trim(templine)
                        If InStr(RIGHT_CODE, "  ") Then
                            temp_arr = Split(RIGHT_CODE, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & FAULT_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & RIGHT_CODE
                            FAULT_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        FAULT_CODES = LEFT_CODE & ":" & RIGHT_CODE
                    ElseIf UCase(Strings.left(Trim(templine), 20)) = "EXTERNAL FAULT CODES" Then
                        EXT_LEFT_CODE = UCase(Trim(Mid(templine, 28, 10)))
                        tempFaultyCode = Type & "E" & UCase(Trim(Mid(templine, 28, 10)))
                        templine = getNowLine()
                        EXT_RIGHT_CODE = Trim(templine)
                        If InStr(EXT_RIGHT_CODE, "  ") Then
                            temp_arr = Split(EXT_RIGHT_CODE, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    EXTERNAL_FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    EXTERNAL_FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & EXTERNAL_FAULT_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & EXT_RIGHT_CODE
                            EXTERNAL_FAULT_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        EXTERNAL_FAULT_CODES = EXT_LEFT_CODE & ":" & EXT_RIGHT_CODE
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "REPLACEMENT UN" Then
                        templine = Trim(getNowLine())
                        tempFaultyCode = Type & "RU"
                        If InStr(templine, "  ") Then
                            temp_arr = Split(templine, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    REPLACEMENT_UNITS_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    REPLACEMENT_UNITS_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & REPLACEMENT_UNITS_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & Trim(templine)
                            REPLACEMENT_UNITS_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        REPLACEMENT_UNITS = Trim(templine)
                        outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                              & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                        MO = ""
                        Last_FC_MO = ""
                        BTSSWVER = ""
                        LEFT_CODE = ""
                        RIGHT_CODE = ""
                        EXT_LEFT_CODE = ""
                        EXT_RIGHT_CODE = ""
                        FAULT_CODES = ""
                        FAULT_NAME = ""
                        EXTERNAL_FAULT_CODES = ""
                        EXTERNAL_FAULT_NAME = ""
                        REPLACEMENT_UNITS = ""
                        REPLACEMENT_UNITS_NAME = ""
                        STATE = ""
                        BLSTATE = ""
                        INTERCNT = ""
                        CONCNT = ""
                        CONERRCNT = ""
                        LASTFLT = ""
                        LFREASON = ""
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
        outputFile1.CLOSE()
        outputFile1 = Nothing
    End Sub
    Private Sub doRXMFP_RXOMCTR(sender As Object, type As String)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim outputFile1 As Object
        Dim filepath1 As String
        Dim tempFaultyCode As String
        Dim EXT_LEFT_CODE As String, EXT_RIGHT_CODE As String, Last_FC_MO As String
        Dim BLSTATE As String, INTERCNT As String, CONCNT As String, CONERRCNT As String, LASTFLT As String, LFREASON As String
        Dim FAULT_CODES As String, FAULT_NAME As String, EXTERNAL_FAULT_CODES As String, EXTERNAL_FAULT_NAME As String, REPLACEMENT_UNITS As String, REPLACEMENT_UNITS_NAME As String
        Dim MO As String, BTSSWVER As String, STATE As String, LEFT_CODE As String, RIGHT_CODE As String, Last_MO As String
        Dim RU As String, RUREVISION As String, RUSERIALNO As String, RUPOSITION As String, RULOGICALID As String
        Dim temp_arr As Object
        MO = ""
        Last_MO = ""
        BTSSWVER = ""
        RU = ""
        RUREVISION = ""
        RUSERIALNO = ""
        RUPOSITION = ""
        RULOGICALID = ""
        STATE = ""
        LEFT_CODE = ""
        RIGHT_CODE = ""
        tempFaultyCode = ""
        Last_FC_MO = ""
        EXT_LEFT_CODE = ""
        EXT_RIGHT_CODE = ""
        FAULT_CODES = ""
        FAULT_NAME = ""
        EXTERNAL_FAULT_CODES = ""
        EXTERNAL_FAULT_NAME = ""
        REPLACEMENT_UNITS = ""
        REPLACEMENT_UNITS_NAME = ""
        BLSTATE = ""
        INTERCNT = ""
        CONCNT = ""
        CONERRCNT = ""
        LASTFLT = ""
        LFREASON = ""

        filepath = outPath & "RXMFP_RXOMCTR_" & CDDtime & ".csv"
        filepath1 = outPath & "RXMFP_FAULT_CODES_" & CDDtime & ".csv"

        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeLine("BSC" & "," & "MO" & "," & "BTSSWVER" & "," & "RU" & "," & "RUREVISION" & "," & "RUSERIALNO" & "," & "RUPOSITION" & "," & "RULOGICALID")
        End If
        If fso.fileExists(filepath1) Then
            outputFile1 = fso.OpenTextFile(filepath1, 8, 1, 0)
        Else
            outputFile1 = fso.OpenTextFile(filepath1, 2, 1, 0)
            outputFile1.writeLine("BSC" & "," & "MO" & "," & "STATE" & "," & "FAULT CODES" & "," & "告警名称" & "," & "EXTERNAL FAULT CODES" & "," & "外部告警名称" & "," & "REPLACEMENT UNITS" & "," & "替换单元名称" _
                                 & "," & "BLSTATE" & "," & "INTERCNT" & "," & "CONCNT" & "," & "CONERRCNT" & "," & "LASTFLT" & "," & "LFREASON")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If RU <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                        End If
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        outputFile1.close()
                        outputFile1 = Nothing
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        If RU <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                        End If
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        outputFile1.close()
                        outputFile1 = Nothing
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "MO                   BTSSWVER" Then
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        'MO                   BTSSWVER          
                        'RU  RUREVISION                           RUSERIALNO
                        '    RUPOSITION                           RULOGICALID

                        templine = getNowLine()
                        MO = Trim(Mid(templine, 1, 21))
                        Last_FC_MO = Trim(Mid(templine, 1, 21))
                        BTSSWVER = Trim(Mid(templine, 22, 18))
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "RU  RUREVISION" Then
                        templine = getNowLine()
                        If templine <> "" Then
                            RU = Trim(Mid(templine, 1, 4))
                            RUREVISION = Trim(Mid(templine, 5, 37))
                            RUSERIALNO = Trim(Mid(templine, 42, 32))
                        Else
                            If Last_MO <> MO Then
                                outputFile.writeLine(BSCNAME & "," & MO & "," & BTSSWVER)
                                Last_MO = MO
                            End If
                        End If
                    ElseIf UCase(Strings.left(templine, 14)) = "    RUPOSITION" Then
                        templine = getNowLine()
                        If templine <> "" Then
                            RUPOSITION = Trim(Mid(templine, 5, 36))
                            RULOGICALID = Trim(Mid(templine, 41, 32))
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                            RU = ""
                            RUREVISION = ""
                            RUSERIALNO = ""
                            RUPOSITION = ""
                            RULOGICALID = ""
                        End If
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "STATE  BLSTATE" Then
                        ' STATE  BLSTATE  INTERCNT  CONCNT  CONERRCNT  LASTFLT   LFREASON
                        templine = getNowLine()
                        STATE = Trim(Mid(templine, 1, 7))
                        BLSTATE = Trim(Mid(templine, 8, 9))
                        INTERCNT = Trim(Mid(templine, 17, 10))
                        CONCNT = Trim(Mid(templine, 27, 8))
                        CONERRCNT = Trim(Mid(templine, 35, 11))
                        LASTFLT = Trim(Mid(templine, 46, 10))
                        LFREASON = Trim(Mid(templine, 56, 10))
                    ElseIf UCase(Strings.left(Trim(templine), 11)) = "FAULT CODES" Then
                        LEFT_CODE = UCase(Trim(Mid(templine, 19, 10)))
                        tempFaultyCode = type & UCase(Trim(Mid(templine, 19, 10)))
                        FAULT_CODES = LEFT_CODE
                        templine = getNowLine()
                        RIGHT_CODE = Trim(templine)
                        If InStr(RIGHT_CODE, "  ") Then
                            temp_arr = Split(RIGHT_CODE, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & FAULT_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & RIGHT_CODE
                            FAULT_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        FAULT_CODES = LEFT_CODE & ":" & RIGHT_CODE
                    ElseIf UCase(Strings.left(Trim(templine), 20)) = "EXTERNAL FAULT CODES" Then
                        EXT_LEFT_CODE = UCase(Trim(Mid(templine, 28, 10)))
                        tempFaultyCode = type & "E" & UCase(Trim(Mid(templine, 28, 10)))
                        templine = getNowLine()
                        EXT_RIGHT_CODE = Trim(templine)
                        If InStr(EXT_RIGHT_CODE, "  ") Then
                            temp_arr = Split(EXT_RIGHT_CODE, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    EXTERNAL_FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    EXTERNAL_FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & EXTERNAL_FAULT_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & EXT_RIGHT_CODE
                            EXTERNAL_FAULT_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        EXTERNAL_FAULT_CODES = EXT_LEFT_CODE & ":" & EXT_RIGHT_CODE
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "REPLACEMENT UN" Then
                        templine = Trim(getNowLine())
                        tempFaultyCode = type & "RU"
                        If InStr(templine, "  ") Then
                            temp_arr = Split(templine, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    REPLACEMENT_UNITS_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    REPLACEMENT_UNITS_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & REPLACEMENT_UNITS_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & Trim(templine)
                            REPLACEMENT_UNITS_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        REPLACEMENT_UNITS = Trim(templine)
                        outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                              & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                        MO = ""
                        Last_FC_MO = ""
                        BTSSWVER = ""
                        LEFT_CODE = ""
                        RIGHT_CODE = ""
                        EXT_LEFT_CODE = ""
                        EXT_RIGHT_CODE = ""
                        FAULT_CODES = ""
                        FAULT_NAME = ""
                        EXTERNAL_FAULT_CODES = ""
                        EXTERNAL_FAULT_NAME = ""
                        REPLACEMENT_UNITS = ""
                        REPLACEMENT_UNITS_NAME = ""
                        STATE = ""
                        BLSTATE = ""
                        INTERCNT = ""
                        CONCNT = ""
                        CONERRCNT = ""
                        LASTFLT = ""
                        LFREASON = ""
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
        outputFile1.CLOSE()
        outputFile1 = Nothing
    End Sub
    Private Sub doRXMFP_RXORX(sender As Object, type As String)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim outputFile1 As Object
        Dim filepath1 As String
        Dim tempFaultyCode As String
        Dim EXT_LEFT_CODE As String, EXT_RIGHT_CODE As String, Last_FC_MO As String
        Dim BLSTATE As String, INTERCNT As String, CONCNT As String, CONERRCNT As String, LASTFLT As String, LFREASON As String
        Dim FAULT_CODES As String, FAULT_NAME As String, EXTERNAL_FAULT_CODES As String, EXTERNAL_FAULT_NAME As String, REPLACEMENT_UNITS As String, REPLACEMENT_UNITS_NAME As String
        Dim MO As String, BTSSWVER As String, STATE As String, LEFT_CODE As String, RIGHT_CODE As String, Last_MO As String
        Dim RU As String, RUREVISION As String, RUSERIALNO As String, RUPOSITION As String, RULOGICALID As String
        Dim temp_arr As Object
        MO = ""
        Last_MO = ""
        BTSSWVER = ""
        RU = ""
        RUREVISION = ""
        RUSERIALNO = ""
        RUPOSITION = ""
        RULOGICALID = ""
        STATE = ""
        LEFT_CODE = ""
        RIGHT_CODE = ""
        tempFaultyCode = ""
        Last_FC_MO = ""
        EXT_LEFT_CODE = ""
        EXT_RIGHT_CODE = ""
        FAULT_CODES = ""
        FAULT_NAME = ""
        EXTERNAL_FAULT_CODES = ""
        EXTERNAL_FAULT_NAME = ""
        REPLACEMENT_UNITS = ""
        REPLACEMENT_UNITS_NAME = ""
        BLSTATE = ""
        INTERCNT = ""
        CONCNT = ""
        CONERRCNT = ""
        LASTFLT = ""
        LFREASON = ""

        filepath = outPath & "RXMFP_RXORX_" & CDDtime & ".csv"
        filepath1 = outPath & "RXMFP_FAULT_CODES_" & CDDtime & ".csv"

        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeLine("BSC" & "," & "MO" & "," & "BTSSWVER" & "," & "RU" & "," & "RUREVISION" & "," & "RUSERIALNO" & "," & "RUPOSITION" & "," & "RULOGICALID")
        End If
        If fso.fileExists(filepath1) Then
            outputFile1 = fso.OpenTextFile(filepath1, 8, 1, 0)
        Else
            outputFile1 = fso.OpenTextFile(filepath1, 2, 1, 0)
            outputFile1.writeLine("BSC" & "," & "MO" & "," & "STATE" & "," & "FAULT CODES" & "," & "告警名称" & "," & "EXTERNAL FAULT CODES" & "," & "外部告警名称" & "," & "REPLACEMENT UNITS" & "," & "替换单元名称" _
                                 & "," & "BLSTATE" & "," & "INTERCNT" & "," & "CONCNT" & "," & "CONERRCNT" & "," & "LASTFLT" & "," & "LFREASON")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If RU <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                        End If
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        outputFile1.close()
                        outputFile1 = Nothing
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        If RU <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                        End If
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        outputFile1.close()
                        outputFile1 = Nothing
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "MO                   BTSSWVER" Then
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        'MO                   BTSSWVER          
                        'RU  RUREVISION                           RUSERIALNO
                        '    RUPOSITION                           RULOGICALID

                        templine = getNowLine()
                        MO = Trim(Mid(templine, 1, 21))
                        Last_FC_MO = Trim(Mid(templine, 1, 21))
                        BTSSWVER = Trim(Mid(templine, 22, 18))
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "RU  RUREVISION" Then
                        templine = getNowLine()
                        If templine <> "" Then
                            RU = Trim(Mid(templine, 1, 4))
                            RUREVISION = Trim(Mid(templine, 5, 37))
                            RUSERIALNO = Trim(Mid(templine, 42, 32))
                        Else
                            If Last_MO <> MO Then
                                outputFile.writeLine(BSCNAME & "," & MO & "," & BTSSWVER)
                                Last_MO = MO
                            End If
                        End If
                    ElseIf UCase(Strings.left(templine, 14)) = "    RUPOSITION" Then
                        templine = getNowLine()
                        If templine <> "" Then
                            RUPOSITION = Trim(Mid(templine, 5, 36))
                            RULOGICALID = Trim(Mid(templine, 41, 32))
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                            RU = ""
                            RUREVISION = ""
                            RUSERIALNO = ""
                            RUPOSITION = ""
                            RULOGICALID = ""
                        End If
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "STATE  BLSTATE" Then
                        ' STATE  BLSTATE  INTERCNT  CONCNT  CONERRCNT  LASTFLT   LFREASON
                        templine = getNowLine()
                        STATE = Trim(Mid(templine, 1, 7))
                        BLSTATE = Trim(Mid(templine, 8, 9))
                        INTERCNT = Trim(Mid(templine, 17, 10))
                        CONCNT = Trim(Mid(templine, 27, 8))
                        CONERRCNT = Trim(Mid(templine, 35, 11))
                        LASTFLT = Trim(Mid(templine, 46, 10))
                        LFREASON = Trim(Mid(templine, 56, 10))
                    ElseIf UCase(Strings.left(Trim(templine), 11)) = "FAULT CODES" Then
                        LEFT_CODE = UCase(Trim(Mid(templine, 19, 10)))
                        tempFaultyCode = Type & UCase(Trim(Mid(templine, 19, 10)))
                        FAULT_CODES = LEFT_CODE
                        templine = getNowLine()
                        RIGHT_CODE = Trim(templine)
                        If InStr(RIGHT_CODE, "  ") Then
                            temp_arr = Split(RIGHT_CODE, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & FAULT_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & RIGHT_CODE
                            FAULT_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        FAULT_CODES = LEFT_CODE & ":" & RIGHT_CODE
                    ElseIf UCase(Strings.left(Trim(templine), 20)) = "EXTERNAL FAULT CODES" Then
                        EXT_LEFT_CODE = UCase(Trim(Mid(templine, 28, 10)))
                        tempFaultyCode = Type & "E" & UCase(Trim(Mid(templine, 28, 10)))
                        templine = getNowLine()
                        EXT_RIGHT_CODE = Trim(templine)
                        If InStr(EXT_RIGHT_CODE, "  ") Then
                            temp_arr = Split(EXT_RIGHT_CODE, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    EXTERNAL_FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    EXTERNAL_FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & EXTERNAL_FAULT_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & EXT_RIGHT_CODE
                            EXTERNAL_FAULT_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        EXTERNAL_FAULT_CODES = EXT_LEFT_CODE & ":" & EXT_RIGHT_CODE
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "REPLACEMENT UN" Then
                        templine = Trim(getNowLine())
                        tempFaultyCode = Type & "RU"
                        If InStr(templine, "  ") Then
                            temp_arr = Split(templine, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    REPLACEMENT_UNITS_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    REPLACEMENT_UNITS_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & REPLACEMENT_UNITS_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & Trim(templine)
                            REPLACEMENT_UNITS_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        REPLACEMENT_UNITS = Trim(templine)
                        outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                              & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                        MO = ""
                        Last_FC_MO = ""
                        BTSSWVER = ""
                        LEFT_CODE = ""
                        RIGHT_CODE = ""
                        EXT_LEFT_CODE = ""
                        EXT_RIGHT_CODE = ""
                        FAULT_CODES = ""
                        FAULT_NAME = ""
                        EXTERNAL_FAULT_CODES = ""
                        EXTERNAL_FAULT_NAME = ""
                        REPLACEMENT_UNITS = ""
                        REPLACEMENT_UNITS_NAME = ""
                        STATE = ""
                        BLSTATE = ""
                        INTERCNT = ""
                        CONCNT = ""
                        CONERRCNT = ""
                        LASTFLT = ""
                        LFREASON = ""
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
        outputFile1.CLOSE()
        outputFile1 = Nothing
    End Sub
    Private Sub doRXMFP_RXOTF(sender As Object, type As String)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim outputFile1 As Object
        Dim filepath1 As String
        Dim tempFaultyCode As String
        Dim EXT_LEFT_CODE As String, EXT_RIGHT_CODE As String, Last_FC_MO As String
        Dim BLSTATE As String, INTERCNT As String, CONCNT As String, CONERRCNT As String, LASTFLT As String, LFREASON As String
        Dim FAULT_CODES As String, FAULT_NAME As String, EXTERNAL_FAULT_CODES As String, EXTERNAL_FAULT_NAME As String, REPLACEMENT_UNITS As String, REPLACEMENT_UNITS_NAME As String
        Dim MO As String, BTSSWVER As String, STATE As String, LEFT_CODE As String, RIGHT_CODE As String, Last_MO As String
        Dim RU As String, RUREVISION As String, RUSERIALNO As String, RUPOSITION As String, RULOGICALID As String
        Dim temp_arr As Object
        MO = ""
        Last_MO = ""
        BTSSWVER = ""
        RU = ""
        RUREVISION = ""
        RUSERIALNO = ""
        RUPOSITION = ""
        RULOGICALID = ""
        STATE = ""
        LEFT_CODE = ""
        RIGHT_CODE = ""
        tempFaultyCode = ""
        Last_FC_MO = ""
        EXT_LEFT_CODE = ""
        EXT_RIGHT_CODE = ""
        FAULT_CODES = ""
        FAULT_NAME = ""
        EXTERNAL_FAULT_CODES = ""
        EXTERNAL_FAULT_NAME = ""
        REPLACEMENT_UNITS = ""
        REPLACEMENT_UNITS_NAME = ""
        BLSTATE = ""
        INTERCNT = ""
        CONCNT = ""
        CONERRCNT = ""
        LASTFLT = ""
        LFREASON = ""

        filepath = outPath & "RXMFP_RXOTF_" & CDDtime & ".csv"
        filepath1 = outPath & "RXMFP_FAULT_CODES_" & CDDtime & ".csv"

        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeLine("BSC" & "," & "MO" & "," & "BTSSWVER" & "," & "RU" & "," & "RUREVISION" & "," & "RUSERIALNO" & "," & "RUPOSITION" & "," & "RULOGICALID")
        End If
        If fso.fileExists(filepath1) Then
            outputFile1 = fso.OpenTextFile(filepath1, 8, 1, 0)
        Else
            outputFile1 = fso.OpenTextFile(filepath1, 2, 1, 0)
            outputFile1.writeLine("BSC" & "," & "MO" & "," & "STATE" & "," & "FAULT CODES" & "," & "告警名称" & "," & "EXTERNAL FAULT CODES" & "," & "外部告警名称" & "," & "REPLACEMENT UNITS" & "," & "替换单元名称" _
                                 & "," & "BLSTATE" & "," & "INTERCNT" & "," & "CONCNT" & "," & "CONERRCNT" & "," & "LASTFLT" & "," & "LFREASON")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If RU <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                        End If
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        outputFile1.close()
                        outputFile1 = Nothing
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        If RU <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                        End If
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        outputFile1.close()
                        outputFile1 = Nothing
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "MO                   BTSSWVER" Then
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        'MO                   BTSSWVER          
                        'RU  RUREVISION                           RUSERIALNO
                        '    RUPOSITION                           RULOGICALID

                        templine = getNowLine()
                        MO = Trim(Mid(templine, 1, 21))
                        Last_FC_MO = Trim(Mid(templine, 1, 21))
                        BTSSWVER = Trim(Mid(templine, 22, 18))
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "RU  RUREVISION" Then
                        templine = getNowLine()
                        If templine <> "" Then
                            RU = Trim(Mid(templine, 1, 4))
                            RUREVISION = Trim(Mid(templine, 5, 37))
                            RUSERIALNO = Trim(Mid(templine, 42, 32))
                        Else
                            If Last_MO <> MO Then
                                outputFile.writeLine(BSCNAME & "," & MO & "," & BTSSWVER)
                                Last_MO = MO
                            End If
                        End If
                    ElseIf UCase(Strings.left(templine, 14)) = "    RUPOSITION" Then
                        templine = getNowLine()
                        If templine <> "" Then
                            RUPOSITION = Trim(Mid(templine, 5, 36))
                            RULOGICALID = Trim(Mid(templine, 41, 32))
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                            RU = ""
                            RUREVISION = ""
                            RUSERIALNO = ""
                            RUPOSITION = ""
                            RULOGICALID = ""
                        End If
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "STATE  BLSTATE" Then
                        ' STATE  BLSTATE  INTERCNT  CONCNT  CONERRCNT  LASTFLT   LFREASON
                        templine = getNowLine()
                        STATE = Trim(Mid(templine, 1, 7))
                        BLSTATE = Trim(Mid(templine, 8, 9))
                        INTERCNT = Trim(Mid(templine, 17, 10))
                        CONCNT = Trim(Mid(templine, 27, 8))
                        CONERRCNT = Trim(Mid(templine, 35, 11))
                        LASTFLT = Trim(Mid(templine, 46, 10))
                        LFREASON = Trim(Mid(templine, 56, 10))
                    ElseIf UCase(Strings.left(Trim(templine), 11)) = "FAULT CODES" Then
                        LEFT_CODE = UCase(Trim(Mid(templine, 19, 10)))
                        tempFaultyCode = type & UCase(Trim(Mid(templine, 19, 10)))
                        FAULT_CODES = LEFT_CODE
                        templine = getNowLine()
                        RIGHT_CODE = Trim(templine)
                        If InStr(RIGHT_CODE, "  ") Then
                            temp_arr = Split(RIGHT_CODE, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & FAULT_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & RIGHT_CODE
                            FAULT_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        FAULT_CODES = LEFT_CODE & ":" & RIGHT_CODE
                    ElseIf UCase(Strings.left(Trim(templine), 20)) = "EXTERNAL FAULT CODES" Then
                        EXT_LEFT_CODE = UCase(Trim(Mid(templine, 28, 10)))
                        tempFaultyCode = type & "E" & UCase(Trim(Mid(templine, 28, 10)))
                        templine = getNowLine()
                        EXT_RIGHT_CODE = Trim(templine)
                        If InStr(EXT_RIGHT_CODE, "  ") Then
                            temp_arr = Split(EXT_RIGHT_CODE, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    EXTERNAL_FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    EXTERNAL_FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & EXTERNAL_FAULT_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & EXT_RIGHT_CODE
                            EXTERNAL_FAULT_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        EXTERNAL_FAULT_CODES = EXT_LEFT_CODE & ":" & EXT_RIGHT_CODE
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "REPLACEMENT UN" Then
                        templine = Trim(getNowLine())
                        tempFaultyCode = type & "RU"
                        If InStr(templine, "  ") Then
                            temp_arr = Split(templine, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    REPLACEMENT_UNITS_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    REPLACEMENT_UNITS_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & REPLACEMENT_UNITS_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & Trim(templine)
                            REPLACEMENT_UNITS_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        REPLACEMENT_UNITS = Trim(templine)
                        outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                              & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                        MO = ""
                        Last_FC_MO = ""
                        BTSSWVER = ""
                        LEFT_CODE = ""
                        RIGHT_CODE = ""
                        EXT_LEFT_CODE = ""
                        EXT_RIGHT_CODE = ""
                        FAULT_CODES = ""
                        FAULT_NAME = ""
                        EXTERNAL_FAULT_CODES = ""
                        EXTERNAL_FAULT_NAME = ""
                        REPLACEMENT_UNITS = ""
                        REPLACEMENT_UNITS_NAME = ""
                        STATE = ""
                        BLSTATE = ""
                        INTERCNT = ""
                        CONCNT = ""
                        CONERRCNT = ""
                        LASTFLT = ""
                        LFREASON = ""
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
        outputFile1.CLOSE()
        outputFile1 = Nothing
    End Sub
    Private Sub doRXMFP_RXOTRX(sender As Object, type As String)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim outputFile1 As Object
        Dim filepath1 As String
        Dim tempFaultyCode As String
        Dim EXT_LEFT_CODE As String, EXT_RIGHT_CODE As String, Last_FC_MO As String
        Dim BLSTATE As String, INTERCNT As String, CONCNT As String, CONERRCNT As String, LASTFLT As String, LFREASON As String
        Dim FAULT_CODES As String, FAULT_NAME As String, EXTERNAL_FAULT_CODES As String, EXTERNAL_FAULT_NAME As String, REPLACEMENT_UNITS As String, REPLACEMENT_UNITS_NAME As String
        Dim MO As String, BTSSWVER As String, STATE As String, LEFT_CODE As String, RIGHT_CODE As String, Last_MO As String
        Dim RU As String, RUREVISION As String, RUSERIALNO As String, RUPOSITION As String, RULOGICALID As String
        Dim temp_arr As Object
        Dim TRX_TYPE_TEMP As String
        TRX_TYPE_TEMP = ""
        MO = ""
        Last_MO = ""
        BTSSWVER = ""
        RU = ""
        RUREVISION = ""
        RUSERIALNO = ""
        RUPOSITION = ""
        RULOGICALID = ""
        STATE = ""
        LEFT_CODE = ""
        RIGHT_CODE = ""
        tempFaultyCode = ""
        Last_FC_MO = ""
        EXT_LEFT_CODE = ""
        EXT_RIGHT_CODE = ""
        FAULT_CODES = ""
        FAULT_NAME = ""
        EXTERNAL_FAULT_CODES = ""
        EXTERNAL_FAULT_NAME = ""
        REPLACEMENT_UNITS = ""
        REPLACEMENT_UNITS_NAME = ""
        BLSTATE = ""
        INTERCNT = ""
        CONCNT = ""
        CONERRCNT = ""
        LASTFLT = ""
        LFREASON = ""

        filepath = outPath & "RXMFP_RXOTRX_" & CDDtime & ".csv"
        filepath1 = outPath & "RXMFP_FAULT_CODES_" & CDDtime & ".csv"

        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeLine("BSC" & "," & "MO" & "," & "BSC_MO" & "," & "BTSSWVER" & "," & "RU" & "," & "RUREVISION" & "," & "RUSERIALNO" & "," & "RUPOSITION" & "," & "RULOGICALID")
        End If
        If fso.fileExists(filepath1) Then
            outputFile1 = fso.OpenTextFile(filepath1, 8, 1, 0)
        Else
            outputFile1 = fso.OpenTextFile(filepath1, 2, 1, 0)
            outputFile1.writeLine("BSC" & "," & "MO" & "," & "STATE" & "," & "FAULT CODES" & "," & "告警名称" & "," & "EXTERNAL FAULT CODES" & "," & "外部告警名称" & "," & "REPLACEMENT UNITS" & "," & "替换单元名称" _
                                 & "," & "BLSTATE" & "," & "INTERCNT" & "," & "CONCNT" & "," & "CONERRCNT" & "," & "LASTFLT" & "," & "LFREASON")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If RU <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BSCNAME & "_" & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                        End If
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        outputFile1.close()
                        outputFile1 = Nothing
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        If RU <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BSCNAME & "_" & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                        End If
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        outputFile1.close()
                        outputFile1 = Nothing
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "MO                   BTSSWVER" Then
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        'MO                   BTSSWVER          
                        'RU  RUREVISION                           RUSERIALNO
                        '    RUPOSITION                           RULOGICALID

                        templine = getNowLine()
                        MO = Trim(Mid(templine, 1, 21))
                        Last_FC_MO = Trim(Mid(templine, 1, 21))
                        BTSSWVER = Trim(Mid(templine, 22, 18))
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "RU  RUREVISION" Then
                        templine = getNowLine()
                        If templine <> "" Then
                            RU = Trim(Mid(templine, 1, 4))
                            RUREVISION = Trim(Mid(templine, 5, 37))
                            RUSERIALNO = Trim(Mid(templine, 42, 32))
                        Else
                            If Last_MO <> MO Then
                                outputFile.writeLine(BSCNAME & "," & MO & "," & BSCNAME & "_" & MO & "," & BTSSWVER)
                                Last_MO = MO
                            End If
                        End If
                    ElseIf UCase(Strings.left(templine, 14)) = "    RUPOSITION" Then
                        templine = getNowLine()
                        If templine <> "" Then
                            RUPOSITION = Trim(Mid(templine, 5, 36))
                            RULOGICALID = Trim(Mid(templine, 41, 32))
                            If Trim(RULOGICALID) = "" Then
                                TRX_TYPE_TEMP = RULOGICALID
                            Else
                                TRX_TYPE_TEMP = Trim(Strings.left(RULOGICALID, Len(RULOGICALID) - 2))
                            End If
                            If TRX_TYPE_DIC.ContainsKey(BSCNAME & "#" & MO) Then
                                TRX_TYPE_DIC(BSCNAME & "#" & MO) = TRX_TYPE_DIC(BSCNAME & "#" & MO) & " " & TRX_TYPE_TEMP
                            Else
                                TRX_TYPE_DIC.Add(BSCNAME & "#" & MO, TRX_TYPE_TEMP)
                            End If
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BSCNAME & "_" & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                            RU = ""
                            RUREVISION = ""
                            RUSERIALNO = ""
                            RUPOSITION = ""
                            RULOGICALID = ""
                        End If
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "STATE  BLSTATE" Then
                        ' STATE  BLSTATE  INTERCNT  CONCNT  CONERRCNT  LASTFLT   LFREASON
                        templine = getNowLine()
                        STATE = Trim(Mid(templine, 1, 7))
                        BLSTATE = Trim(Mid(templine, 8, 9))
                        INTERCNT = Trim(Mid(templine, 17, 10))
                        CONCNT = Trim(Mid(templine, 27, 8))
                        CONERRCNT = Trim(Mid(templine, 35, 11))
                        LASTFLT = Trim(Mid(templine, 46, 10))
                        LFREASON = Trim(Mid(templine, 56, 10))
                    ElseIf UCase(Strings.left(Trim(templine), 11)) = "FAULT CODES" Then
                        LEFT_CODE = UCase(Trim(Mid(templine, 19, 10)))
                        tempFaultyCode = Type & UCase(Trim(Mid(templine, 19, 10)))
                        FAULT_CODES = LEFT_CODE
                        templine = getNowLine()
                        RIGHT_CODE = Trim(templine)
                        If InStr(RIGHT_CODE, "  ") Then
                            temp_arr = Split(RIGHT_CODE, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & FAULT_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & RIGHT_CODE
                            FAULT_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        FAULT_CODES = LEFT_CODE & ":" & RIGHT_CODE
                    ElseIf UCase(Strings.left(Trim(templine), 20)) = "EXTERNAL FAULT CODES" Then
                        EXT_LEFT_CODE = UCase(Trim(Mid(templine, 28, 10)))
                        tempFaultyCode = Type & "E" & UCase(Trim(Mid(templine, 28, 10)))
                        templine = getNowLine()
                        EXT_RIGHT_CODE = Trim(templine)
                        If InStr(EXT_RIGHT_CODE, "  ") Then
                            temp_arr = Split(EXT_RIGHT_CODE, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    EXTERNAL_FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    EXTERNAL_FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & EXTERNAL_FAULT_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & EXT_RIGHT_CODE
                            EXTERNAL_FAULT_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        EXTERNAL_FAULT_CODES = EXT_LEFT_CODE & ":" & EXT_RIGHT_CODE
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "REPLACEMENT UN" Then
                        templine = Trim(getNowLine())
                        tempFaultyCode = Type & "RU"
                        If InStr(templine, "  ") Then
                            temp_arr = Split(templine, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    REPLACEMENT_UNITS_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    REPLACEMENT_UNITS_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & REPLACEMENT_UNITS_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & Trim(templine)
                            REPLACEMENT_UNITS_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        REPLACEMENT_UNITS = Trim(templine)
                        outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                              & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                        MO = ""
                        Last_FC_MO = ""
                        BTSSWVER = ""
                        LEFT_CODE = ""
                        RIGHT_CODE = ""
                        EXT_LEFT_CODE = ""
                        EXT_RIGHT_CODE = ""
                        FAULT_CODES = ""
                        FAULT_NAME = ""
                        EXTERNAL_FAULT_CODES = ""
                        EXTERNAL_FAULT_NAME = ""
                        REPLACEMENT_UNITS = ""
                        REPLACEMENT_UNITS_NAME = ""
                        STATE = ""
                        BLSTATE = ""
                        INTERCNT = ""
                        CONCNT = ""
                        CONERRCNT = ""
                        LASTFLT = ""
                        LFREASON = ""
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
        outputFile1.CLOSE()
        outputFile1 = Nothing
    End Sub
    Private Sub doRXMFP_RXOTS(sender As Object, type As String)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim outputFile1 As Object
        Dim filepath1 As String
        Dim tempFaultyCode As String
        Dim EXT_LEFT_CODE As String, EXT_RIGHT_CODE As String, Last_FC_MO As String
        Dim BLSTATE As String, INTERCNT As String, CONCNT As String, CONERRCNT As String, LASTFLT As String, LFREASON As String
        Dim FAULT_CODES As String, FAULT_NAME As String, EXTERNAL_FAULT_CODES As String, EXTERNAL_FAULT_NAME As String, REPLACEMENT_UNITS As String, REPLACEMENT_UNITS_NAME As String
        Dim MO As String, BTSSWVER As String, STATE As String, LEFT_CODE As String, RIGHT_CODE As String, Last_MO As String
        Dim RU As String, RUREVISION As String, RUSERIALNO As String, RUPOSITION As String, RULOGICALID As String
        Dim temp_arr As Object
        MO = ""
        Last_MO = ""
        BTSSWVER = ""
        RU = ""
        RUREVISION = ""
        RUSERIALNO = ""
        RUPOSITION = ""
        RULOGICALID = ""
        STATE = ""
        LEFT_CODE = ""
        RIGHT_CODE = ""
        tempFaultyCode = ""
        Last_FC_MO = ""
        EXT_LEFT_CODE = ""
        EXT_RIGHT_CODE = ""
        FAULT_CODES = ""
        FAULT_NAME = ""
        EXTERNAL_FAULT_CODES = ""
        EXTERNAL_FAULT_NAME = ""
        REPLACEMENT_UNITS = ""
        REPLACEMENT_UNITS_NAME = ""
        BLSTATE = ""
        INTERCNT = ""
        CONCNT = ""
        CONERRCNT = ""
        LASTFLT = ""
        LFREASON = ""

        filepath = outPath & "RXMFP_RXOTS_" & CDDtime & ".csv"
        filepath1 = outPath & "RXMFP_FAULT_CODES_" & CDDtime & ".csv"

        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeLine("BSC" & "," & "MO" & "," & "BTSSWVER" & "," & "RU" & "," & "RUREVISION" & "," & "RUSERIALNO" & "," & "RUPOSITION" & "," & "RULOGICALID")
        End If
        If fso.fileExists(filepath1) Then
            outputFile1 = fso.OpenTextFile(filepath1, 8, 1, 0)
        Else
            outputFile1 = fso.OpenTextFile(filepath1, 2, 1, 0)
            outputFile1.writeLine("BSC" & "," & "MO" & "," & "STATE" & "," & "FAULT CODES" & "," & "告警名称" & "," & "EXTERNAL FAULT CODES" & "," & "外部告警名称" & "," & "REPLACEMENT UNITS" & "," & "替换单元名称" _
                                 & "," & "BLSTATE" & "," & "INTERCNT" & "," & "CONCNT" & "," & "CONERRCNT" & "," & "LASTFLT" & "," & "LFREASON")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If RU <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                        End If
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        outputFile1.close()
                        outputFile1 = Nothing
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        If RU <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                        End If
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        outputFile1.close()
                        outputFile1 = Nothing
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "MO                   BTSSWVER" Then
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        'MO                   BTSSWVER          
                        'RU  RUREVISION                           RUSERIALNO
                        '    RUPOSITION                           RULOGICALID

                        templine = getNowLine()
                        MO = Trim(Mid(templine, 1, 21))
                        Last_FC_MO = Trim(Mid(templine, 1, 21))
                        BTSSWVER = Trim(Mid(templine, 22, 18))
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "RU  RUREVISION" Then
                        templine = getNowLine()
                        If templine <> "" Then
                            RU = Trim(Mid(templine, 1, 4))
                            RUREVISION = Trim(Mid(templine, 5, 37))
                            RUSERIALNO = Trim(Mid(templine, 42, 32))
                        Else
                            If Last_MO <> MO Then
                                outputFile.writeLine(BSCNAME & "," & MO & "," & BTSSWVER)
                                Last_MO = MO
                            End If
                        End If
                    ElseIf UCase(Strings.left(templine, 14)) = "    RUPOSITION" Then
                        templine = getNowLine()
                        If templine <> "" Then
                            RUPOSITION = Trim(Mid(templine, 5, 36))
                            RULOGICALID = Trim(Mid(templine, 41, 32))
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                            RU = ""
                            RUREVISION = ""
                            RUSERIALNO = ""
                            RUPOSITION = ""
                            RULOGICALID = ""
                        End If
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "STATE  BLSTATE" Then
                        ' STATE  BLSTATE  INTERCNT  CONCNT  CONERRCNT  LASTFLT   LFREASON
                        templine = getNowLine()
                        STATE = Trim(Mid(templine, 1, 7))
                        BLSTATE = Trim(Mid(templine, 8, 9))
                        INTERCNT = Trim(Mid(templine, 17, 10))
                        CONCNT = Trim(Mid(templine, 27, 8))
                        CONERRCNT = Trim(Mid(templine, 35, 11))
                        LASTFLT = Trim(Mid(templine, 46, 10))
                        LFREASON = Trim(Mid(templine, 56, 10))
                    ElseIf UCase(Strings.left(Trim(templine), 11)) = "FAULT CODES" Then
                        LEFT_CODE = UCase(Trim(Mid(templine, 19, 10)))
                        tempFaultyCode = type & UCase(Trim(Mid(templine, 19, 10)))
                        FAULT_CODES = LEFT_CODE
                        templine = getNowLine()
                        RIGHT_CODE = Trim(templine)
                        If InStr(RIGHT_CODE, "  ") Then
                            temp_arr = Split(RIGHT_CODE, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & FAULT_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & RIGHT_CODE
                            FAULT_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        FAULT_CODES = LEFT_CODE & ":" & RIGHT_CODE
                    ElseIf UCase(Strings.left(Trim(templine), 20)) = "EXTERNAL FAULT CODES" Then
                        EXT_LEFT_CODE = UCase(Trim(Mid(templine, 28, 10)))
                        tempFaultyCode = type & "E" & UCase(Trim(Mid(templine, 28, 10)))
                        templine = getNowLine()
                        EXT_RIGHT_CODE = Trim(templine)
                        If InStr(EXT_RIGHT_CODE, "  ") Then
                            temp_arr = Split(EXT_RIGHT_CODE, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    EXTERNAL_FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    EXTERNAL_FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & EXTERNAL_FAULT_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & EXT_RIGHT_CODE
                            EXTERNAL_FAULT_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        EXTERNAL_FAULT_CODES = EXT_LEFT_CODE & ":" & EXT_RIGHT_CODE
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "REPLACEMENT UN" Then
                        templine = Trim(getNowLine())
                        tempFaultyCode = type & "RU"
                        If InStr(templine, "  ") Then
                            temp_arr = Split(templine, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    REPLACEMENT_UNITS_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    REPLACEMENT_UNITS_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & REPLACEMENT_UNITS_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & Trim(templine)
                            REPLACEMENT_UNITS_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        REPLACEMENT_UNITS = Trim(templine)
                        outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                              & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                        MO = ""
                        Last_FC_MO = ""
                        BTSSWVER = ""
                        LEFT_CODE = ""
                        RIGHT_CODE = ""
                        EXT_LEFT_CODE = ""
                        EXT_RIGHT_CODE = ""
                        FAULT_CODES = ""
                        FAULT_NAME = ""
                        EXTERNAL_FAULT_CODES = ""
                        EXTERNAL_FAULT_NAME = ""
                        REPLACEMENT_UNITS = ""
                        REPLACEMENT_UNITS_NAME = ""
                        STATE = ""
                        BLSTATE = ""
                        INTERCNT = ""
                        CONCNT = ""
                        CONERRCNT = ""
                        LASTFLT = ""
                        LFREASON = ""
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
        outputFile1.CLOSE()
        outputFile1 = Nothing
    End Sub
    Private Sub doRXMFP_RXOTX(sender As Object, type As String)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim outputFile1 As Object
        Dim filepath1 As String
        Dim tempFaultyCode As String
        Dim EXT_LEFT_CODE As String, EXT_RIGHT_CODE As String, Last_FC_MO As String
        Dim BLSTATE As String, INTERCNT As String, CONCNT As String, CONERRCNT As String, LASTFLT As String, LFREASON As String
        Dim FAULT_CODES As String, FAULT_NAME As String, EXTERNAL_FAULT_CODES As String, EXTERNAL_FAULT_NAME As String, REPLACEMENT_UNITS As String, REPLACEMENT_UNITS_NAME As String
        Dim MO As String, BTSSWVER As String, STATE As String, LEFT_CODE As String, RIGHT_CODE As String, Last_MO As String
        Dim RU As String, RUREVISION As String, RUSERIALNO As String, RUPOSITION As String, RULOGICALID As String
        Dim temp_arr As Object
        MO = ""
        Last_MO = ""
        BTSSWVER = ""
        RU = ""
        RUREVISION = ""
        RUSERIALNO = ""
        RUPOSITION = ""
        RULOGICALID = ""
        STATE = ""
        LEFT_CODE = ""
        RIGHT_CODE = ""
        tempFaultyCode = ""
        Last_FC_MO = ""
        EXT_LEFT_CODE = ""
        EXT_RIGHT_CODE = ""
        FAULT_CODES = ""
        FAULT_NAME = ""
        EXTERNAL_FAULT_CODES = ""
        EXTERNAL_FAULT_NAME = ""
        REPLACEMENT_UNITS = ""
        REPLACEMENT_UNITS_NAME = ""
        BLSTATE = ""
        INTERCNT = ""
        CONCNT = ""
        CONERRCNT = ""
        LASTFLT = ""
        LFREASON = ""

        filepath = outPath & "RXMFP_RXOTX_" & CDDtime & ".csv"
        filepath1 = outPath & "RXMFP_FAULT_CODES_" & CDDtime & ".csv"

        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeLine("BSC" & "," & "MO" & "," & "BTSSWVER" & "," & "RU" & "," & "RUREVISION" & "," & "RUSERIALNO" & "," & "RUPOSITION" & "," & "RULOGICALID")
        End If
        If fso.fileExists(filepath1) Then
            outputFile1 = fso.OpenTextFile(filepath1, 8, 1, 0)
        Else
            outputFile1 = fso.OpenTextFile(filepath1, 2, 1, 0)
            outputFile1.writeLine("BSC" & "," & "MO" & "," & "STATE" & "," & "FAULT CODES" & "," & "告警名称" & "," & "EXTERNAL FAULT CODES" & "," & "外部告警名称" & "," & "REPLACEMENT UNITS" & "," & "替换单元名称" _
                                 & "," & "BLSTATE" & "," & "INTERCNT" & "," & "CONCNT" & "," & "CONERRCNT" & "," & "LASTFLT" & "," & "LFREASON")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If RU <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                        End If
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        outputFile1.close()
                        outputFile1 = Nothing
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        If RU <> "" Then
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                        End If
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        outputFile1.close()
                        outputFile1 = Nothing
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If Trim(templine) = "MO                   BTSSWVER" Then
                        If FAULT_CODES <> "" Or EXTERNAL_FAULT_CODES <> "" Then
                            outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                                 & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                            MO = ""
                            Last_FC_MO = ""
                            BTSSWVER = ""
                            LEFT_CODE = ""
                            RIGHT_CODE = ""
                            EXT_LEFT_CODE = ""
                            EXT_RIGHT_CODE = ""
                            FAULT_CODES = ""
                            FAULT_NAME = ""
                            EXTERNAL_FAULT_CODES = ""
                            EXTERNAL_FAULT_NAME = ""
                            REPLACEMENT_UNITS = ""
                            REPLACEMENT_UNITS_NAME = ""
                            STATE = ""
                            BLSTATE = ""
                            INTERCNT = ""
                            CONCNT = ""
                            CONERRCNT = ""
                            LASTFLT = ""
                            LFREASON = ""
                        End If
                        'MO                   BTSSWVER          
                        'RU  RUREVISION                           RUSERIALNO
                        '    RUPOSITION                           RULOGICALID

                        templine = getNowLine()
                        MO = Trim(Mid(templine, 1, 21))
                        Last_FC_MO = Trim(Mid(templine, 1, 21))
                        BTSSWVER = Trim(Mid(templine, 22, 18))
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "RU  RUREVISION" Then
                        templine = getNowLine()
                        If templine <> "" Then
                            RU = Trim(Mid(templine, 1, 4))
                            RUREVISION = Trim(Mid(templine, 5, 37))
                            RUSERIALNO = Trim(Mid(templine, 42, 32))
                        Else
                            If Last_MO <> MO Then
                                outputFile.writeLine(BSCNAME & "," & MO & "," & BTSSWVER)
                                Last_MO = MO
                            End If
                        End If
                    ElseIf UCase(Strings.left(templine, 14)) = "    RUPOSITION" Then
                        templine = getNowLine()
                        If templine <> "" Then
                            RUPOSITION = Trim(Mid(templine, 5, 36))
                            RULOGICALID = Trim(Mid(templine, 41, 32))
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & BTSSWVER & "," & RU & "," & RUREVISION & "," & RUSERIALNO & "," & RUPOSITION & "," & RULOGICALID)
                            RU = ""
                            RUREVISION = ""
                            RUSERIALNO = ""
                            RUPOSITION = ""
                            RULOGICALID = ""
                        End If
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "STATE  BLSTATE" Then
                        ' STATE  BLSTATE  INTERCNT  CONCNT  CONERRCNT  LASTFLT   LFREASON
                        templine = getNowLine()
                        STATE = Trim(Mid(templine, 1, 7))
                        BLSTATE = Trim(Mid(templine, 8, 9))
                        INTERCNT = Trim(Mid(templine, 17, 10))
                        CONCNT = Trim(Mid(templine, 27, 8))
                        CONERRCNT = Trim(Mid(templine, 35, 11))
                        LASTFLT = Trim(Mid(templine, 46, 10))
                        LFREASON = Trim(Mid(templine, 56, 10))
                    ElseIf UCase(Strings.left(Trim(templine), 11)) = "FAULT CODES" Then
                        LEFT_CODE = UCase(Trim(Mid(templine, 19, 10)))
                        tempFaultyCode = Type & UCase(Trim(Mid(templine, 19, 10)))
                        FAULT_CODES = LEFT_CODE
                        templine = getNowLine()
                        RIGHT_CODE = Trim(templine)
                        If InStr(RIGHT_CODE, "  ") Then
                            temp_arr = Split(RIGHT_CODE, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & FAULT_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & RIGHT_CODE
                            FAULT_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        FAULT_CODES = LEFT_CODE & ":" & RIGHT_CODE
                    ElseIf UCase(Strings.left(Trim(templine), 20)) = "EXTERNAL FAULT CODES" Then
                        EXT_LEFT_CODE = UCase(Trim(Mid(templine, 28, 10)))
                        tempFaultyCode = Type & "E" & UCase(Trim(Mid(templine, 28, 10)))
                        templine = getNowLine()
                        EXT_RIGHT_CODE = Trim(templine)
                        If InStr(EXT_RIGHT_CODE, "  ") Then
                            temp_arr = Split(EXT_RIGHT_CODE, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    EXTERNAL_FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    EXTERNAL_FAULT_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & EXTERNAL_FAULT_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & EXT_RIGHT_CODE
                            EXTERNAL_FAULT_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        EXTERNAL_FAULT_CODES = EXT_LEFT_CODE & ":" & EXT_RIGHT_CODE
                    ElseIf UCase(Strings.left(Trim(templine), 14)) = "REPLACEMENT UN" Then
                        templine = Trim(getNowLine())
                        tempFaultyCode = Type & "RU"
                        If InStr(templine, "  ") Then
                            temp_arr = Split(templine, "  ")
                            For x = 0 To UBound(temp_arr)
                                If x = 0 Then
                                    REPLACEMENT_UNITS_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x)))
                                Else
                                    REPLACEMENT_UNITS_NAME = "(" & tempFaultyCode & ":" & Trim(temp_arr(x)) & ")" & getFaultyCodeName(tempFaultyCode & Trim(temp_arr(x))) & "/" & REPLACEMENT_UNITS_NAME
                                End If
                            Next
                        Else
                            tempFaultyCode = tempFaultyCode & Trim(templine)
                            REPLACEMENT_UNITS_NAME = getFaultyCodeName(tempFaultyCode)
                        End If
                        REPLACEMENT_UNITS = Trim(templine)
                        outputFile1.writeLine(BSCNAME & "," & MO & "," & STATE & "," & FAULT_CODES & "," & FAULT_NAME & "," & EXTERNAL_FAULT_CODES & "," & EXTERNAL_FAULT_NAME & "," & REPLACEMENT_UNITS & "," & REPLACEMENT_UNITS_NAME _
                                              & "," & BLSTATE & "," & INTERCNT & "," & CONCNT & "," & CONERRCNT & "," & LASTFLT & "," & LFREASON)
                        MO = ""
                        Last_FC_MO = ""
                        BTSSWVER = ""
                        LEFT_CODE = ""
                        RIGHT_CODE = ""
                        EXT_LEFT_CODE = ""
                        EXT_RIGHT_CODE = ""
                        FAULT_CODES = ""
                        FAULT_NAME = ""
                        EXTERNAL_FAULT_CODES = ""
                        EXTERNAL_FAULT_NAME = ""
                        REPLACEMENT_UNITS = ""
                        REPLACEMENT_UNITS_NAME = ""
                        STATE = ""
                        BLSTATE = ""
                        INTERCNT = ""
                        CONCNT = ""
                        CONERRCNT = ""
                        LASTFLT = ""
                        LFREASON = ""
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
        outputFile1.CLOSE()
        outputFile1 = Nothing
    End Sub
    Private Sub doRXMOP_RXOCF(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim MO As String, TEI As String, SIG As String
        Dim SWVERREPL As String, SWVERDLD As String, SWVERACT As String
        Dim BSSWANTED As String, NEGSTATUS As String
        'MO                TEI      SIG
        '                  SWVERREPL      SWVERDLD       SWVERACT
        '                  BSSWANTED     NEGSTATUS       
        MO = ""
        TEI = ""
        SIG = ""
        SWVERREPL = ""
        SWVERDLD = ""
        SWVERACT = ""
        BSSWANTED = ""
        NEGSTATUS = ""
        filepath = outPath & "RXMOP_RXOCF_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "MO" & "," & "TEI" & "," & "SIG" & "," & "SWVERREPL" & "," & "SWVERDLD" & "," & "SWVERACT" & "," & "BSSWANTED" & "," & "NEGSTATUS")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If MO <> "" Then
                            outputFile.writeline(BSCNAME & "," & MO & "," & TEI & "," & SIG & "," & SWVERREPL & "," & SWVERDLD & "," & SWVERACT & "," & BSSWANTED & "," & NEGSTATUS)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If MO <> "" Then
                            outputFile.writeline(BSCNAME & "," & MO & "," & TEI & "," & SIG & "," & SWVERREPL & "," & SWVERDLD & "," & SWVERACT & "," & BSSWANTED & "," & NEGSTATUS)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "MO                TEI      SIG" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'MO                TEI      SIG
                            MO = Trim(Mid(templine, 1, 18))
                            TEI = Trim(Mid(templine, 19, 9))
                            SIG = Trim(Mid(templine, 28, 10))
                        End If
                    ElseIf UCase(Trim(templine)) = "SWVERREPL      SWVERDLD       SWVERACT" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            '                  SWVERREPL      SWVERDLD       SWVERACT
                            '                  BSSWANTED     NEGSTATUS
                            SWVERREPL = Trim(Mid(templine, 19, 15))
                            SWVERDLD = Trim(Mid(templine, 34, 15))
                            SWVERACT = Trim(Mid(templine, 49, 15))
                        End If
                    ElseIf UCase(Trim(templine)) = "BSSWANTED     NEGSTATUS" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            '                  BSSWANTED     NEGSTATUS
                            BSSWANTED = Trim(Mid(templine, 19, 14))
                            NEGSTATUS = Trim(Mid(templine, 33, 15))
                        End If
                        outputFile.writeline(BSCNAME & "," & MO & "," & TEI & "," & SIG & "," & SWVERREPL & "," & SWVERDLD & "," & SWVERACT & "," & BSSWANTED & "," & NEGSTATUS)
                        MO = ""
                        TEI = ""
                        SIG = ""
                        SWVERREPL = ""
                        SWVERDLD = ""
                        SWVERACT = ""
                        BSSWANTED = ""
                        NEGSTATUS = ""
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    private sub doRXMOP_RXOTF(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'MO              TFMODE  SYNCSRC  TFCOMPNEG  TFCOMPPOS  FSOFFSET  TFAUTO
        filepath = outPath & "RXMOP_RXOTF_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "MO" & "," & "TFMODE" & "," & "SYNCSRC" & "," & "TFCOMPNEG" & "," & "TFCOMPPOS" & "," & "FSOFFSET" & "," & "TFAUTO")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "MO              TFMODE  SYNCSRC  TFCOMPNEG  TFCOMPPOS  FSOFFSET  TFAUTO" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 16)) & "," & Trim(Mid(templine, 17, 8)) & "," & Trim(Mid(templine, 25, 9)) & "," & Trim(Mid(templine, 34, 11)) & "," & Trim(Mid(templine, 45, 11)) & "," & Trim(Mid(templine, 56, 10)) & "," & Trim(Mid(templine, 66, 10)))
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    private sub doRXMOP_RXOTG(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim MO As String, RSITE As String, AHOP As String, COMB As String, FHOP As String, MODEL As String
        Dim SWVERREPL As String, SWVERDLD As String, SWVERACT As String, TMODE As String
        Dim CONFMD As String, CONFACT As String, TRACO As String, ABISALLOC As String, CLUSTERID As String, SCGR As String
        Dim DAMRCR As String, CLTGINST As String, CCCHCMD As String
        Dim PTA As String, JBSDL As String, PAL As String, JBPTA As String
        Dim TGFID As String, SIGDEL As String, BSSWANTED As String, PACKALG As String
        'MO                RSITE                               COMB  FHOP  MODEL
        'MO                RSITE                         AHOP  COMB  FHOP  MODEL

        '                  SWVERREPL      SWVERDLD       SWVERACT      TMODE
        '                  CONFMD  CONFACT  TRACO  ABISALLOC  CLUSTERID  SCGR
        '                  DAMRCR  CLTGINST  CCCHCMD
        '                  PTA  JBSDL  PAL  JBPTA
        '                  TGFID         SIGDEL        BSSWANTED    PACKALG
        MO = ""
        RSITE = ""
        AHOP = ""
        COMB = ""
        FHOP = ""
        MODEL = ""
        SWVERREPL = ""
        SWVERDLD = ""
        SWVERACT = ""
        TMODE = ""
        CONFMD = ""
        CONFACT = ""
        TRACO = ""
        ABISALLOC = ""
        CLUSTERID = ""
        SCGR = ""
        DAMRCR = ""
        CLTGINST = ""
        CCCHCMD = ""
        PTA = ""
        JBSDL = ""
        PAL = ""
        JBPTA = ""
        TGFID = ""
        SIGDEL = ""
        BSSWANTED = ""
        PACKALG = ""

        filepath = outPath & "RXMOP_RXOTG_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "MO" & "," & "RSITE" & "," & "AHOP" & "," & "COMB" & "," & "FHOP" & "," & "MODEL" _
                                 & "," & "SWVERREPL" & "," & "SWVERDLD" & "," & "SWVERACT" & "," & "TMODE" _
                                 & "," & "CONFMD" & "," & "CONFACT" & "," & "TRACO" & "," & "ABISALLOC" & "," & "CLUSTERID" & "," & "SCGR" _
                                 & "," & "DAMRCR" & "," & "CLTGINST" & "," & "CCCHCMD" _
                                 & "," & "PTA" & "," & "JBSDL" & "," & "PAL" & "," & "JBPTA" _
                                 & "," & "TGFID" & "," & "SIGDEL" & "," & "BSSWANTED" & "," & "PACKALG")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If MO <> "" Then
                            outputFile.writeline(BSCNAME & "," & MO & "," & RSITE & "," & AHOP & "," & COMB & "," & FHOP & "," & MODEL _
                                                 & "," & SWVERREPL & "," & SWVERDLD & "," & SWVERACT & "," & TMODE _
                                                 & "," & CONFMD & "," & CONFACT & "," & TRACO & "," & ABISALLOC & "," & CLUSTERID & "," & SCGR _
                                                 & "," & DAMRCR & "," & CLTGINST & "," & CCCHCMD _
                                                 & "," & PTA & "," & JBSDL & "," & PAL & "," & JBPTA _
                                                 & "," & TGFID & "," & SIGDEL & "," & BSSWANTED & "," & PACKALG)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If MO <> "" Then
                            outputFile.writeline(BSCNAME & "," & MO & "," & RSITE & "," & AHOP & "," & COMB & "," & FHOP & "," & MODEL _
                                                 & "," & SWVERREPL & "," & SWVERDLD & "," & SWVERACT & "," & TMODE _
                                                 & "," & CONFMD & "," & CONFACT & "," & TRACO & "," & ABISALLOC & "," & CLUSTERID & "," & SCGR _
                                                 & "," & DAMRCR & "," & CLTGINST & "," & CCCHCMD _
                                                 & "," & PTA & "," & JBSDL & "," & PAL & "," & JBPTA _
                                                 & "," & TGFID & "," & SIGDEL & "," & BSSWANTED & "," & PACKALG)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Strings.left(Trim(templine), 23)) = "MO                RSITE" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'MO                RSITE                               COMB  FHOP  MODEL
                            'MO                RSITE                         AHOP  COMB  FHOP  MODEL

                            MO = Trim(Mid(templine, 1, 18))
                            RSITE = Trim(Mid(templine, 19, 30))
                            AHOP = Trim(Mid(templine, 49, 6))
                            COMB = Trim(Mid(templine, 55, 6))
                            FHOP = Trim(Mid(templine, 61, 6))
                            MODEL = Trim(Mid(templine, 67, 5))
                        End If
                    ElseIf UCase(Trim(templine)) = "SWVERREPL      SWVERDLD       SWVERACT      TMODE" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            '                  SWVERREPL      SWVERDLD       SWVERACT      TMODE
                            SWVERREPL = Trim(Mid(templine, 19, 15))
                            SWVERDLD = Trim(Mid(templine, 34, 15))
                            SWVERACT = Trim(Mid(templine, 49, 14))
                            TMODE = Trim(Mid(templine, 63, 7))
                        End If
                    ElseIf UCase(Trim(templine)) = "CONFMD  CONFACT  TRACO  ABISALLOC  CLUSTERID  SCGR" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            '                  CONFMD  CONFACT  TRACO  ABISALLOC  CLUSTERID  SCGR
                            CONFMD = Trim(Mid(templine, 19, 8))
                            CONFACT = Trim(Mid(templine, 27, 9))
                            TRACO = Trim(Mid(templine, 36, 7))
                            ABISALLOC = Trim(Mid(templine, 43, 11))
                            CLUSTERID = Trim(Mid(templine, 54, 11))
                            SCGR = Trim(Mid(templine, 65, 10))
                        End If
                    ElseIf UCase(Trim(templine)) = "DAMRCR  CLTGINST  CCCHCMD" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            '                  DAMRCR  CLTGINST  CCCHCMD
                            DAMRCR = Trim(Mid(templine, 19, 8))
                            CLTGINST = Trim(Mid(templine, 27, 10))
                            CCCHCMD = Trim(Mid(templine, 37, 10))
                        End If
                    ElseIf UCase(Trim(templine)) = "PTA  JBSDL  PAL  JBPTA" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            '                  PTA  JBSDL  PAL  JBPTA
                            PTA = Trim(Mid(templine, 19, 5))
                            JBSDL = Trim(Mid(templine, 24, 7))
                            PAL = Trim(Mid(templine, 31, 5))
                            JBPTA = Trim(Mid(templine, 36, 8))
                        End If
                    ElseIf UCase(Trim(templine)) = "TGFID         SIGDEL        BSSWANTED    PACKALG" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            '                  TGFID         SIGDEL        BSSWANTED    PACKALG
                            TGFID = Trim(Mid(templine, 19, 14))
                            SIGDEL = Trim(Mid(templine, 33, 14))
                            BSSWANTED = Trim(Mid(templine, 47, 13))
                            PACKALG = Trim(Mid(templine, 60, 13))
                        End If
                        outputFile.writeline(BSCNAME & "," & MO & "," & RSITE & "," & AHOP & "," & COMB & "," & FHOP & "," & MODEL _
                                             & "," & SWVERREPL & "," & SWVERDLD & "," & SWVERACT & "," & TMODE _
                                             & "," & CONFMD & "," & CONFACT & "," & TRACO & "," & ABISALLOC & "," & CLUSTERID & "," & SCGR _
                                             & "," & DAMRCR & "," & CLTGINST & "," & CCCHCMD _
                                             & "," & PTA & "," & JBSDL & "," & PAL & "," & JBPTA _
                                             & "," & TGFID & "," & SIGDEL & "," & BSSWANTED & "," & PACKALG)
                        MO = ""
                        RSITE = ""
                        AHOP = ""
                        COMB = ""
                        FHOP = ""
                        MODEL = ""
                        SWVERREPL = ""
                        SWVERDLD = ""
                        SWVERACT = ""
                        TMODE = ""
                        CONFMD = ""
                        CONFACT = ""
                        TRACO = ""
                        ABISALLOC = ""
                        CLUSTERID = ""
                        SCGR = ""
                        DAMRCR = ""
                        CLTGINST = ""
                        CCCHCMD = ""
                        PTA = ""
                        JBSDL = ""
                        PAL = ""
                        JBPTA = ""
                        TGFID = ""
                        SIGDEL = ""
                        BSSWANTED = ""
                        PACKALG = ""
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    private sub doRXMOP_RXORX(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'MO                RXD       BAND     ANTA      ANTB
        filepath = outPath & "RXMOP_RXORX_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "MO" & "," & "RXD" & "," & "BAND" & "," & "ANTA" & "," & "ANTB")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "MO                RXD       BAND     ANTA      ANTB" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 18)) & "," & Trim(Mid(templine, 19, 10)) & "," & Trim(Mid(templine, 29, 9)) & "," & Trim(Mid(templine, 38, 10)) & "," & Trim(Mid(templine, 48, 6)))
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    private sub doRXMOP_RXOTRX(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim MO As String
        Dim DCP1 As String
        Dim DCP2 As String
        Dim SWVERREPL As String
        Dim SWVERDLD As String
        Dim SWVERACT As String
        Dim CELL As String, CHGR As String, TEI As String, SIG As String, MCTRI As String, SC As String
        Dim TRX_CELL_NO_TEMP As Integer, TRX_SITE_TRXSNO_TEMP As Integer, TRX_SITETYPE_TEMP As String, TRX_SITETYPE_TRXNO_TEMP As Integer
        Dim TRX_NO_TEMP As String, LAST_CELL As String

        LAST_CELL = ""
        TRX_CELL_NO_TEMP = 0
        TRX_SITE_TRXSNO_TEMP = 0
        TRX_SITETYPE_TEMP = ""
        TRX_NO_TEMP = ""
        TRX_SITETYPE_TRXNO_TEMP = 0
        MO = ""
        DCP1 = ""
        DCP2 = ""
        CELL = ""
        CHGR = ""
        TEI = ""
        SIG = ""
        MCTRI = ""
        SC = ""
        SWVERREPL = ""
        SWVERDLD = ""
        SWVERACT = ""
        filepath = outPath & "RXMOP_RXOTRX_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "MO" & "," & "BSC_MO" & "," & "CELL" & "," & "CHGR" & "," & "TEI" & "," & "SIG" & "," & "MCTRI" & "," & "DCP1" & "," & "SC" & "," & "SWVERREPL" & "," & "SWVERDLD" & "," & "SWVERACT" & "," & "DCP2")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If DCP2 <> "" Then
                            outputFile.writeLINE("," & SWVERREPL & "," & SWVERDLD & "," & SWVERACT & "," & DCP2)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If DCP2 <> "" Then
                            outputFile.writeLINE("," & SWVERREPL & "," & SWVERDLD & "," & SWVERACT & "," & DCP2)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Strings.left(Trim(templine), 21)) = "MO                CEL" Then
                        templine = getNowLine()
                        'MO                CELL      CHGR    TEI      SIG      MCTRI    DCP1  SC
                        '                  SWVERREPL      SWVERDLD       SWVERACT       DCP2
                        If DCP1 <> "" And DCP2 <> "" Then
                            outputFile.writeLINE("," & SWVERREPL & "," & SWVERDLD & "," & SWVERACT & "," & DCP2)
                            MO = Trim(Mid(templine, 1, 18))
                            CELL = Trim(Mid(templine, 19, 10))
                            CHGR = Trim(Mid(templine, 29, 8))
                            TEI = Trim(Mid(templine, 37, 9))
                            SIG = Trim(Mid(templine, 46, 9))
                            MCTRI = Trim(Mid(templine, 55, 9))
                            DCP1 = Trim(Mid(templine, 64, 6))
                            SC = Trim(Mid(templine, 70, 3))
                        Else
                            MO = Trim(Mid(templine, 1, 18))
                            CELL = Trim(Mid(templine, 19, 10))
                            CHGR = Trim(Mid(templine, 29, 8))
                            TEI = Trim(Mid(templine, 37, 9))
                            SIG = Trim(Mid(templine, 46, 9))
                            MCTRI = Trim(Mid(templine, 55, 9))
                            DCP1 = Trim(Mid(templine, 64, 6))
                            SC = Trim(Mid(templine, 70, 3))
                        End If
                    ElseIf templine = "                  SWVERREPL      SWVERDLD       SWVERACT       DCP2" Then
                        outputFile.write(BSCNAME & "," & MO & "," & BSCNAME & "_" & MO & "," & CELL & "," & CHGR & "," & TEI & "," & SIG & "," & MCTRI & "," & DCP1 & "," & SC)

                        '   TRX_MO_TEMP = Trim(Replace(Strings.left(MO, InStrRev(MO, "-")), "TRX", "TG"))
                        'If TRX_BSC_CELL_DIC.ContainsKey(BSCNAME & "#" & CELL) Then
                        '    TRX_NO_TEMP = TRX_NO_TEMP + 1
                        '    TRX_BSC_CELL_DIC(BSCNAME & "#" & CELL) = TRX_NO_TEMP
                        '    Console.WriteLine(TRX_NO_TEMP)
                        'Else
                        '    TRX_NO_TEMP = 1
                        '    TRX_BSC_CELL_DIC.Add(BSCNAME & "#" & CELL, TRX_NO_TEMP)
                        '    Console.WriteLine(TRX_NO_TEMP)
                        'End If
                        If TRX_BSC_MO_CELL_DIC.ContainsKey(BSCNAME & "#" & CELL) Then
                            TRX_BSC_MO_CELL_DIC(BSCNAME & "#" & CELL) = TRX_BSC_MO_CELL_DIC(BSCNAME & "#" & CELL) & "@" & BSCNAME & "#" & MO
                        Else
                            TRX_BSC_MO_CELL_DIC.Add(BSCNAME & "#" & CELL, BSCNAME & "#" & MO)
                        End If
                        'If TRX_SITE_TRXNO_DIC.ContainsKey(BSCNAME & "#" & Strings.left(CELL, Len(CELL) - 1)) Then
                        '    If LAST_CELL = CELL Then
                        '        TRX_SITE_TRXSNO_TEMP = TRX_SITE_TRXSNO_TEMP + 1
                        '    Else
                        '        TRX_SITE_TRXSNO_TEMP = TRX_SITE_TRXSNO_TEMP + 1
                        '        TRX_CELL_NO_TEMP = TRX_CELL_NO_TEMP + 1
                        '        If TRX_CELL_NO_TEMP = 2 Then
                        '            TRX_SITETYPE_TEMP = "S" & CStr(TRX_BSC_CELL_DIC(BSCNAME & "#" & LAST_CELL))
                        '            TRX_SITETYPE_TRXNO_TEMP = CInt(TRX_BSC_CELL_DIC(BSCNAME & "#" & LAST_CELL))
                        '        ElseIf TRX_CELL_NO_TEMP > 2 Then
                        '            TRX_SITETYPE_TEMP = TRX_SITETYPE_TEMP & "/" & TRX_BSC_CELL_DIC(BSCNAME & "#" & LAST_CELL)
                        '            TRX_SITETYPE_TRXNO_TEMP = TRX_SITETYPE_TRXNO_TEMP + CInt(TRX_BSC_CELL_DIC(BSCNAME & "#" & LAST_CELL))
                        '        End If
                        '    End If
                        '    LAST_CELL = CELL
                        '    TRX_SITE_TRXNO_DIC(BSCNAME & "#" & Strings.left(CELL, Len(CELL) - 1)) = TRX_SITE_TRXSNO_TEMP & "@" & TRX_CELL_NO_TEMP & "@" & TRX_SITETYPE_TEMP & "@" & TRX_SITETYPE_TRXNO_TEMP
                        'Else
                        '    LAST_CELL = CELL
                        '    TRX_SITE_TRXSNO_TEMP = 1
                        '    TRX_CELL_NO_TEMP = 1
                        '    TRX_SITETYPE_TEMP = "S" & CStr(TRX_BSC_CELL_DIC(BSCNAME & "#" & CELL))
                        '    TRX_SITE_TRXNO_DIC.Add(BSCNAME & "#" & Strings.left(CELL, Len(CELL) - 1), TRX_SITE_TRXSNO_TEMP & "@" & TRX_CELL_NO_TEMP & "@" & TRX_SITETYPE_TEMP)
                        'End If
                        templine = getNowLine()
                        SWVERREPL = Trim(Mid(templine, 19, 15))
                        SWVERDLD = Trim(Mid(templine, 34, 15))
                        SWVERACT = Trim(Mid(templine, 49, 15))
                        DCP2 = Trim(Mid(templine, 64, 6))
                    ElseIf Trim(templine) <> "TG TO CHANNEL GROUP CONNECTION DATA" And Trim(templine) <> "RADIO X-CEIVER ADMINISTRATION" Then
                        DCP2 = Trim(DCP2 & " " & Trim(Mid(templine, 64, 6)))
                    End If
                End If
            End If
        Loop
        If Trim(DCP2) <> "" Then
            outputFile.writeLINE("," & SWVERREPL & "," & SWVERDLD & "," & SWVERACT & "," & DCP2)
        End If
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    private sub doRXMOP_RXOTX(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'MO                CELL      CHGR  BAND     ANT      MPWR
        filepath = outPath & "RXMOP_RXOTX_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "MO" & "," & "CELL" & "," & "CHGR" & "," & "BAND" & "," & "ANT" & "," & "MPWR")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Trim(templine)) = "MO                CELL      CHGR  BAND     ANT      MPWR" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 18)) & "," & Trim(Mid(templine, 19, 10)) & "," & Trim(Mid(templine, 29, 6)) & "," & Trim(Mid(templine, 35, 9)) & "," & Trim(Mid(templine, 44, 9)) & "," & Trim(Mid(templine, 53, 5)))
                        End If
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    private sub doRXMOP_RXOMCTR(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim MO As String, MAXPWR As String, MAXTRX As String, PAO As String, TXMPWR As String, IPM As String, MIXEDMODE As String
        Dim MINFREQ As String, MAXFREQ As String, CCMD As String, CCTP As String, EMTP As String, EMDI As String, EMRA As String
        Dim TPO As String, CCID As String
        'MO                MAXPWR   MAXTRX   PAO   TXMPWR   IPM  MIXEDMODE
        '                  MINFREQ  MAXFREQ  CCMD  CCTP  EMTP  EMDI  EMRA

        'G14 
        'MO                MAXPWR   MAXTRX   PAO   TXMPWR   IPM  MIXEDMODE  TPO
        '                  MINFREQ  MAXFREQ
        '                  CCMD  CCID   CCTP  EMTP  EMDI   EMRA
        MO = ""
        MAXPWR = ""
        MAXTRX = ""
        PAO = ""
        TXMPWR = ""
        IPM = ""
        MIXEDMODE = ""
        MINFREQ = ""
        MAXFREQ = ""
        CCMD = ""
        CCTP = ""
        EMTP = ""
        EMDI = ""
        EMRA = ""
        TPO = ""
        CCID = ""
        filepath = outPath & "RXMOP_RXOMCTR_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "MO" & "," & "BSC_MO" & "," & "MAXPWR" & "," & "MAXTRX" & "," & "PAO" & "," & "TXMPWR" & "," & "IPM" & "," & "MIXEDMODE" _
                                 & "," & "MINFREQ" & "," & "MAXFREQ" & "," & "CCMD" & "," & "CCTP" & "," & "EMTP" & "," & "EMDI" & "," & "EMRA" & "," & "TPO" & "," & "CCID")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        If MO <> "" Then
                            outputFile.writeLine(BSCNAME & "," & MO & "," & BSCNAME & "_" & MO & "," & MAXPWR & "," & MAXTRX & "," & PAO & "," & TXMPWR & "," & IPM & "," & MIXEDMODE _
                                                 & "," & MINFREQ & "," & MAXFREQ & "," & CCMD & "," & CCTP & "," & EMTP & "," & EMDI & "," & EMRA & "," & TPO & "," & CCID)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        If MO <> "" Then
                            outputFile.writeLine(BSCNAME & "," & MO & "," & BSCNAME & "_" & MO & "," & MAXPWR & "," & MAXTRX & "," & PAO & "," & TXMPWR & "," & IPM & "," & MIXEDMODE _
                                                 & "," & MINFREQ & "," & MAXFREQ & "," & CCMD & "," & CCTP & "," & EMTP & "," & EMDI & "," & EMRA & "," & TPO & "," & CCID)
                        End If
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Mid(Trim(templine), 1, 54)) = "MO                MAXPWR   MAXTRX   PAO   TXMPWR   IPM" Then
                        If MO <> "" Then
                            outputFile.writeLine(BSCNAME & "," & MO & "," & BSCNAME & "_" & MO & "," & MAXPWR & "," & MAXTRX & "," & PAO & "," & TXMPWR & "," & IPM & "," & MIXEDMODE _
                                                 & "," & MINFREQ & "," & MAXFREQ & "," & CCMD & "," & CCTP & "," & EMTP & "," & EMDI & "," & EMRA & "," & TPO & "," & CCID)
                        End If
                        templine = getNowLine()
                        'MO                MAXPWR   MAXTRX   PAO   TXMPWR   IPM  MIXEDMODE
                        If Not Trim(templine) = "NONE" Then
                            MO = Trim(Mid(templine, 1, 18))
                            MAXPWR = Trim(Mid(templine, 19, 9))
                            MAXTRX = Trim(Mid(templine, 28, 9))
                            PAO = Trim(Mid(templine, 37, 6))
                            TXMPWR = Trim(Mid(templine, 43, 9))
                            IPM = Trim(Mid(templine, 52, 5))
                            MIXEDMODE = Trim(Mid(templine, 57, 11))
                            TPO = Trim(Mid(templine, 68, 11))
                        End If
                    ElseIf UCase(Mid(Trim(templine), 1, 16)) = "MINFREQ  MAXFREQ" Then
                        templine = getNowLine()
                        '                  MINFREQ  MAXFREQ  CCMD  CCTP  EMTP  EMDI  EMRA

                        'G14
                        '                  MINFREQ  MAXFREQ
                        If Not Trim(templine) = "NONE" Then
                            MINFREQ = Trim(Mid(templine, 19, 9))
                            MAXFREQ = Trim(Mid(templine, 28, 9))
                            CCMD = Trim(Mid(templine, 37, 6))
                            CCTP = Trim(Mid(templine, 43, 6))
                            EMTP = Trim(Mid(templine, 49, 6))
                            EMDI = Trim(Mid(templine, 55, 6))
                            EMRA = Trim(Mid(templine, 61, 4))
                        End If
                        If CCMD <> "" Then
                            outputFile.writeLine(BSCNAME & "," & MO & "," & BSCNAME & "_" & MO & "," & MAXPWR & "," & MAXTRX & "," & PAO & "," & TXMPWR & "," & IPM & "," & MIXEDMODE _
                                                 & "," & MINFREQ & "," & MAXFREQ & "," & CCMD & "," & CCTP & "," & EMTP & "," & EMDI & "," & EMRA & "," & TPO & "," & CCID)
                            MO = ""
                            MAXPWR = ""
                            MAXTRX = ""
                            PAO = ""
                            TXMPWR = ""
                            IPM = ""
                            MIXEDMODE = ""
                            MINFREQ = ""
                            MAXFREQ = ""
                            CCMD = ""
                            CCTP = ""
                            EMTP = ""
                            EMDI = ""
                            EMRA = ""
                            TPO = ""
                            CCID = ""
                        End If
                    ElseIf UCase(Mid(Trim(templine), 1, 17)) = "CCMD  CCID   CCTP" Then
                        templine = getNowLine()
                        'g14 
                        '                  CCMD  CCID   CCTP  EMTP  EMDI   EMRA
                        If Not Trim(templine) = "NONE" Then
                            CCMD = Trim(Mid(templine, 19, 6))
                            CCID = Trim(Mid(templine, 25, 7))
                            CCTP = Trim(Mid(templine, 32, 6))
                            EMTP = Trim(Mid(templine, 38, 6))
                            EMDI = Trim(Mid(templine, 44, 7))
                            EMRA = Trim(Mid(templine, 51, 10))
                        End If
                        outputFile.writeLine(BSCNAME & "," & MO & "," & BSCNAME & "_" & MO & "," & MAXPWR & "," & MAXTRX & "," & PAO & "," & TXMPWR & "," & IPM & "," & MIXEDMODE _
                                                 & "," & MINFREQ & "," & MAXFREQ & "," & CCMD & "," & CCTP & "," & EMTP & "," & EMDI & "," & EMRA & "," & TPO & "," & CCID)
                        MO = ""
                        MAXPWR = ""
                        MAXTRX = ""
                        PAO = ""
                        TXMPWR = ""
                        IPM = ""
                        MIXEDMODE = ""
                        MINFREQ = ""
                        MAXFREQ = ""
                        CCMD = ""
                        CCTP = ""
                        EMTP = ""
                        EMDI = ""
                        EMRA = ""
                        TPO = ""
                        CCID = ""
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    private sub doRXMSP_RXOTRX(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        filepath = outPath & "RXMSP_RXOTRX_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "MO" & "," & "STATE" & "," & "BLSTATE" & "," & "BLO" & "," & "BLA" & "," & "LMO" & "," & "BTS" & "," & "CONF")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Strings.left(Trim(templine), 22)) = "MO               STATE" Then
                        templine = getNowLine()
                        'MO               STATE  BLSTATE    BLO   BLA   LMO    BTS   CONF
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 16)) & "," & Trim(Mid(templine, 17, 7)) & "," & Trim(Mid(templine, 24, 11)) & "," & Trim(Mid(templine, 35, 6)) & "," & Trim(Mid(templine, 41, 6)) & "," & Trim(Mid(templine, 47, 7)) & "," & Trim(Mid(templine, 54, 6)) & "," & Trim(Mid(templine, 60, 5)))
                    ElseIf Trim(Mid(templine, 1, 6)) = "RXOTRX" Then
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 16)) & "," & Trim(Mid(templine, 17, 7)) & "," & Trim(Mid(templine, 24, 11)) & "," & Trim(Mid(templine, 35, 6)) & "," & Trim(Mid(templine, 41, 6)) & "," & Trim(Mid(templine, 47, 7)) & "," & Trim(Mid(templine, 54, 6)) & "," & Trim(Mid(templine, 60, 5)))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub
    Private Sub doRXMSP_RXOCF(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        filepath = outPath & "RXMSP_RXOCF_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "MO" & "," & "STATE" & "," & "BLSTATE" & "," & "BLO" & "," & "BLA" & "," & "LMO" & "," & "BTS" & "," & "CONF")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Strings.Left(Trim(templine), 22)) = "MO               STATE" Then
                        templine = getNowLine()
                        'MO               STATE  BLSTATE    BLO   BLA   LMO    BTS   CONF
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 16)) & "," & Trim(Mid(templine, 17, 7)) & "," & Trim(Mid(templine, 24, 11)) & "," & Trim(Mid(templine, 35, 6)) & "," & Trim(Mid(templine, 41, 6)) & "," & Trim(Mid(templine, 47, 7)) & "," & Trim(Mid(templine, 54, 6)) & "," & Trim(Mid(templine, 60, 5)))
                    ElseIf Trim(Mid(templine, 1, 5)) = "RXOCF" Then
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 16)) & "," & Trim(Mid(templine, 17, 7)) & "," & Trim(Mid(templine, 24, 11)) & "," & Trim(Mid(templine, 35, 6)) & "," & Trim(Mid(templine, 41, 6)) & "," & Trim(Mid(templine, 47, 7)) & "," & Trim(Mid(templine, 54, 6)) & "," & Trim(Mid(templine, 60, 5)))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub

    Private Sub doRXMSP_RXOTF(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        filepath = outPath & "RXMSP_RXOTF_" & CDDTime & ".csv"
        'MO               STATE  BLSTATE    BLO   BLA   LMO    BTS   CONF
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "MO" & "," & "STATE" & "," & "BLSTATE" & "," & "BLO" & "," & "BLA" & "," & "LMO" & "," & "BTS" & "," & "CONF")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(templine, sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Strings.Left(Trim(templine), 22)) = "MO               STATE" Then
                        templine = getNowLine()
                        'MO               STATE  BLSTATE    BLO   BLA   LMO    BTS   CONF
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 16)) & "," & Trim(Mid(templine, 17, 7)) & "," & Trim(Mid(templine, 24, 11)) & "," & Trim(Mid(templine, 35, 6)) & "," & Trim(Mid(templine, 41, 6)) & "," & Trim(Mid(templine, 47, 7)) & "," & Trim(Mid(templine, 54, 6)) & "," & Trim(Mid(templine, 60, 5)))
                    ElseIf Trim(Mid(templine, 1, 5)) = "RXOTF" Then
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 16)) & "," & Trim(Mid(templine, 17, 7)) & "," & Trim(Mid(templine, 24, 11)) & "," & Trim(Mid(templine, 35, 6)) & "," & Trim(Mid(templine, 41, 6)) & "," & Trim(Mid(templine, 47, 7)) & "," & Trim(Mid(templine, 54, 6)) & "," & Trim(Mid(templine, 60, 5)))
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing
    End Sub

    Private sub doRXTCP_RXOTG(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim MO As String, CELL As String, CHGR As String
        Dim temp As String

        'MO               CELL                CHGR

        MO = ""
        CELL = ""
        CHGR = ""
        filepath = outPath & "RXTCP_RXOTG_" & CDDtime & ".csv"

        TCP_CELL_MO_DIC("BSC#CELL") = "BSC#MO"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "MO" & "," & "CELL" & "," & "CHGR")
        End If


        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        doWhichMethod(Trim(templine), sender)
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Strings.left(Trim(templine), 21)) = "MO               CELL" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'MO               CELL                CHGR
                            MO = Trim(Mid(templine, 1, 17))
                            CELL = Trim(Mid(templine, 18, 20))
                            CHGR = Trim(Mid(templine, 38, 5))
                            outputFile.writeLINE(BSCNAME & "," & MO & "," & CELL & "," & CHGR)
                            temp = BSCNAME & "#" & CELL
                            If TCP_CELL_MO_DIC.ContainsKey(temp) Then
                                If Not InStr(TCP_CELL_MO_DIC(temp), BSCNAME & "#" & MO) > 0 Then
                                    TCP_CELL_MO_DIC(temp) = TCP_CELL_MO_DIC(temp) & "|" & BSCNAME & "#" & MO
                                End If
                            Else
                                TCP_CELL_MO_DIC.Add(temp, BSCNAME & "#" & MO)
                            End If
                        End If
                    ElseIf Strings.left(templine, 17) = "                 " Then
                        CELL = Trim(Mid(templine, 18, 20))
                        CHGR = Trim(Mid(templine, 38, 5))
                        outputFile.writeLINE(BSCNAME & "," & MO & "," & CELL & "," & CHGR)
                        temp = BSCNAME & "#" & CELL
                        If TCP_CELL_MO_DIC.ContainsKey(temp) Then
                            If Not InStr(TCP_CELL_MO_DIC(temp), BSCNAME & "#" & MO) > 0 Then
                                TCP_CELL_MO_DIC(temp) = TCP_CELL_MO_DIC(temp) & "|" & BSCNAME & "#" & MO
                            End If
                        Else
                            TCP_CELL_MO_DIC.Add(temp, BSCNAME & "#" & MO)
                        End If
                    ElseIf Trim(Mid(templine, 1, 10)) <> "" And InStr(Trim(Mid(templine, 1, 10)), "-") And Trim(templine) <> "TG TO CHANNEL GROUP CONNECTION DATA" And Trim(templine) <> "RADIO X-CEIVER ADMINISTRATION" Then
                        MO = Trim(Mid(templine, 1, 17))
                        CELL = Trim(Mid(templine, 18, 20))
                        CHGR = Trim(Mid(templine, 38, 5))
                        temp = BSCNAME & "#" & Trim(Mid(templine, 18, 20))
                        If TCP_CELL_MO_DIC.ContainsKey(temp) Then
                            If Not InStr(TCP_CELL_MO_DIC(temp), BSCNAME & "#" & MO) > 0 Then
                                TCP_CELL_MO_DIC(temp) = TCP_CELL_MO_DIC(temp) & "|" & BSCNAME & "#" & MO
                            End If
                        Else
                            TCP_CELL_MO_DIC.Add(temp, BSCNAME & "#" & MO)
                        End If
                        outputFile.writeLINE(BSCNAME & "," & MO & "," & CELL & "," & CHGR)
                    End If
                End If
            End If
        Loop
        outputFile.CLOSE()
        outputFile = Nothing

    End Sub
    private sub doSAAEP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        'SAE    BLOCK    CNTRTYP  NI          NIU         NIE         NIR

        filepath = outPath & "SAAEP_" & CDDtime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "SAE" & "," & "BLOCK" & "," & "CNTRTYP" & "," & "NI" & "," & "NIU" & "," & "NIE" & "," & "NIR")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        Exit Sub
                    End If
                Else
                    If UCase(Strings.left(Trim(templine), 20)) = "SAE    BLOCK    CNTR" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'SAE    BLOCK    CNTRTYP  NI          NIU         NIE         NIR
                            outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 7)) & "," & Trim(Mid(templine, 8, 9)) & "," & Trim(Mid(templine, 17, 9)) & "," & Trim(Mid(templine, 26, 12)) & "," & Trim(Mid(templine, 38, 12)) & "," & Trim(Mid(templine, 50, 12)) & "," & Trim(Mid(templine, 62, 10)))
                        End If
                    ElseIf UCase(Trim(templine)) <> "SIZE ALTERATION OF DATA FILES INFORMATION" Then
                        outputFile.writeLINE(BSCNAME & "," & Trim(Mid(templine, 1, 7)) & "," & Trim(Mid(templine, 8, 9)) & "," & Trim(Mid(templine, 17, 9)) & "," & Trim(Mid(templine, 26, 12)) & "," & Trim(Mid(templine, 38, 12)) & "," & Trim(Mid(templine, 50, 12)) & "," & Trim(Mid(templine, 62, 10)))
                    End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
    Private Sub doTPCOP(sender As Object)
        Dim outputFile As Object
        Dim filepath As String
        Dim templine As String
        Dim SDIP As String
        Dim SNT As String
        'SDIP     SNT          MS    CLREF  HP     LP       K  L  M  DIP
        SDIP = ""
        SNT = ""
        filepath = outPath & "TPCOP_" & CDDTime & ".csv"
        If fso.fileExists(filepath) Then
            outputFile = fso.OpenTextFile(filepath, 8, 1, 0)
        Else
            outputFile = fso.OpenTextFile(filepath, 2, 1, 0)
            outputFile.writeline("BSC" & "," & "SDIP" & "," & "BSC_SDIP" & "," & "SNT" & "," & "MS" & "," & "CLREF" & "," & "HP" & "," & "LP" & "," & "K" & "," & "L" & "," & "M" & "," & "DIP")
        End If
        Do While Not EOF(1)
            templine = getNowLine()
            If Trim(templine) <> "" Then
                If isEnd(templine) Then
                    If Strings.Left(templine, 1) = "<" Then
                        outputFile.close()
                        outputFile = Nothing
                        SDIP = ""
                        SNT = ""
                        doWhichMethod(Trim(templine), sender)
                        Exit Sub
                    Else
                        outputFile.close()
                        outputFile = Nothing
                        SDIP = ""
                        SNT = ""
                        Exit Sub
                    End If
                Else
                    If UCase(Strings.Left(Trim(templine), 24)) = "SDIP     SNT          MS" Then
                        templine = getNowLine()
                        If Not Trim(templine) = "NONE" Then
                            'SDIP     SNT          MS    CLREF  HP     LP       K  L  M  DIP
                            If Trim(Mid(templine, 1, 9)) <> "" Then
                                SDIP = Trim(Mid(templine, 1, 9))
                            End If
                            If Trim(Mid(templine, 10, 13)) <> "" Then
                                SNT = Trim(Mid(templine, 10, 13))
                            End If
                            outputFile.writeLINE(BSCNAME & "," & SDIP & "," & BSCNAME & "_" & SDIP & "," & SNT & "," & Trim(Mid(templine, 23, 6)) & "," & Trim(Mid(templine, 29, 7)) & "," & Trim(Mid(templine, 36, 7)) & "," & Trim(Mid(templine, 43, 9)) & "," & Trim(Mid(templine, 52, 3)) & "," & Trim(Mid(templine, 55, 3)) & "," & Trim(Mid(templine, 58, 3)) & "," & Trim(Mid(templine, 61, 13)))
                        End If
                    ElseIf UCase(Trim(templine)) = "SDIP     SDIPOWNER" Then
                        templine = getNowLine()
                    ElseIf UCase(Trim(templine)) <> "SDIP     SNT          MS    CLREF  HP     LP       K  L  M  DIP" And UCase(Trim(templine)) <> "SYNCHRONOUS DIGITAL PATH DATA" And UCase(Trim(templine)) <> "SDIP     SDIPOWNER" Then
                        If Trim(Mid(templine, 1, 9)) <> "" Then
                            SDIP = Trim(Mid(templine, 1, 9))
                        End If
                        If Trim(Mid(templine, 10, 13)) <> "" Then
                            SNT = Trim(Mid(templine, 10, 13))
                        End If
                        outputFile.writeLINE(BSCNAME & "," & SDIP & "," & BSCNAME & "_" & SDIP & "," & SNT & "," & Trim(Mid(templine, 23, 6)) & "," & Trim(Mid(templine, 29, 7)) & "," & Trim(Mid(templine, 36, 7)) & "," & Trim(Mid(templine, 43, 9)) & "," & Trim(Mid(templine, 52, 3)) & "," & Trim(Mid(templine, 55, 3)) & "," & Trim(Mid(templine, 58, 3)) & "," & Trim(Mid(templine, 61, 13)))
                        End If
                End If
            End If
        Loop
        outputFile.close()
        outputFile = Nothing
    End Sub
End Class
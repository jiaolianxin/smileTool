Imports Microsoft.Office.Interop
Imports System.ComponentModel

Module MakeForteFileWithSiteOrSiteDatabase
    Private fileName, outPutPath

    Friend Function getFiles()
        MsgBox("请选择Mcom Site文件(*.txt) 或者 基站信息表文件(*.xls;*.xlsx)!", MsgBoxStyle.Information)
        MainForm.OpenFileDialog1.Title = "选择Mcom Site 和 Mcom Carrier 文件"
        MainForm.OpenFileDialog1.Filter = "McomSite文件(*.txt)|*.txt|基站信息表文件(*.xls;*.xlsx)|*.xls;*.xlsx"
        MainForm.OpenFileDialog1.Multiselect = False
        MainForm.OpenFileDialog1.ShowHelp = True
        MainForm.OpenFileDialog1.RestoreDirectory = True
        MainForm.OpenFileDialog1.ShowDialog()
        fileName = MainForm.OpenFileDialog1.FileName
        MainForm.OpenFileDialog1.Reset()
        Return fileName
    End Function
    Friend Sub doMaker(sender As Object, e As DoWorkEventArgs)
        Dim fso, ext_name, nowLine, nowLineNo, temp, myPath, temp2, temp3, sheetsCount, siteInfoMark, temp4
        Dim siteFile, outSectorFile, outPutFolder, outCHGRFile
        Dim excel_app As New Excel.Application, excel_book As Excel.Workbook
        Dim cell_arr, long_arr, lat_arr, dir_arr, BSC_arr
        Dim MSC, BSC, Vendor, site, Latitude, Longitude, Sector, ID, Master, LAC, CI, Keywords, Azimuth, BCCH_frequency, BSIC, Intracell_HO, Synchronization_group, AMR_HR_Allocation, AMR_HR_Threshold, HR_Allocation, HR_Threshold, TCH_allocation_priority, GPRS_allocation_priority, Remote
        Dim Sector_chgr, Channel_Group, Subcell, Band, Extended, Hopping_method, Contains_BCCH, HSN, DTX, Power_control, Subcell_Signal_Threshold, Subcell_Tx_Power, TRXs, SDCCH_TSs, Fixed_Data_TSs, Dynamic_Data_TSs, Priority
        Dim total_cell
        Dim cell_position As Integer, bsc_position As Integer, lon_position As Integer, lat_position As Integer, dir_position As Integer

        If fileName = "" Then
            Exit Sub
        End If
        sender.reportProgress(0, "准备数据......")
        ext_name = Strings.Split(fileName, ".")(1)
        myPath = Strings.Left(fileName, InStrRev(fileName, "\"))
        fso = CreateObject("Scripting.FileSystemObject")
        nowLine = ""
        nowLineNo = ""
        cell_arr = New ArrayList
        long_arr = New ArrayList
        lat_arr = New ArrayList
        dir_arr = New ArrayList
        BSC_arr = New ArrayList
        temp4 = New ArrayList
        If ext_name = "txt" Then                               ''''''''''''''''''''''''''读取Site文件
            '         siteFile = fso.OpenTextFile(fileName, 1)
            FileClose(1)
            FileOpen(1, fileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
            nowLineNo = 0
            sender.reportProgress(0, "读取Mcom site数据......")
            Try
                Do While Not EOF(1)
                    nowLineNo = nowLineNo + 1
                    nowLine = LineInput(1)
                    temp = Strings.Split(nowLine, "	")
                    If nowLineNo = 1 And temp(2) <> "Longitude" Then
                        MsgBox("请选择有效的Mcom site文件后重试！(特别是检查经纬度是否存在换行符？）")
                        sender.reportProgress(0, "欢迎使用！")
                        FileClose(1)
                        fso = Nothing
                        Exit Sub
                    End If
                    If temp.length < 5 Then
                        MsgBox("请检查" & temp(0) & "的数据是否存在问题？特别是检查经纬度是否存在换行符？修改后再重新运行！")
                        sender.reportProgress(0, "欢迎使用！")
                        FileClose(1)
                        fso = Nothing
                        Exit Sub
                    End If
                    cell_arr.add(temp(0))
                    long_arr.add(temp(2))
                    lat_arr.add(temp(3))
                    dir_arr.add(temp(4))
                    If temp(19) = "" Then
                        BSC_arr.add("BSC01")
                    Else
                        BSC_arr.add(temp(19))
                    End If

                Loop
            Catch ex As Exception
                MsgBox("Mcom site文件格式有误！请检查第" & nowLineNo & "行的数据！")
                cell_arr = Nothing
                long_arr = Nothing
                lat_arr = Nothing
                dir_arr = Nothing
                BSC_arr = Nothing
                Exit Sub
            Finally
                FileClose(1)
            End Try
        ElseIf ext_name = "xls" Or ext_name = "xlsx" Then                      '''''''''''''''''读取基站信息表
            excel_app = New Excel.Application
            excel_book = GetObject(fileName)
            sheetsCount = excel_book.Sheets.Count
            siteInfoMark = False
            total_cell = ""
            sender.reportProgress(0, "读取基站信息表数据......")
            For x = 1 To sheetsCount
                If excel_book.Sheets(x).name = "基站信息" Then
                    total_cell = excel_book.Sheets("基站信息").UsedRange.rows.Count
                    For y = 1 To excel_book.Sheets("基站信息").UsedRange.columns.Count
                        If Trim(excel_book.Sheets("基站信息").cells(1, y).value) = "小区编号" Then
                            cell_position = y
                        ElseIf Trim(excel_book.Sheets("基站信息").cells(1, y).value) = "归属BSC" Then
                            bsc_position = y
                        ElseIf Trim(excel_book.Sheets("基站信息").cells(1, y).value) = "经度" Then
                            lon_position = y
                        ElseIf Trim(excel_book.Sheets("基站信息").cells(1, y).value) = "纬度" Then
                            lat_position = y
                        ElseIf InStr(Trim(excel_book.Sheets("基站信息").cells(1, y).value), "天线方向") > 0 Then
                            dir_position = y
                        End If
                    Next
                    siteInfoMark = True
                    Exit For
                End If
            Next
            sender.reportProgress(5, "读取基站信息表数据......")
            For m = 1 To total_cell
                cell_arr.add(excel_book.Sheets("基站信息").cells(m, cell_position).value)
                BSC_arr.add(excel_book.Sheets("基站信息").cells(m, bsc_position).value)
                long_arr.add(excel_book.Sheets("基站信息").cells(m, lon_position).value)
                lat_arr.add(excel_book.Sheets("基站信息").cells(m, lat_position).value)
                dir_arr.add(excel_book.Sheets("基站信息").cells(m, dir_position).value)
                sender.reportProgress(5 + m / total_cell * 15, "读取基站信息表数据......")
            Next
            If Not siteInfoMark Then
                MsgBox("请选择带有'基站信息'表的基站信息表文件！")
                excel_app.DisplayAlerts = False
                excel_book.Close()
                Exit Sub
            End If
            excel_app.DisplayAlerts = False
            excel_book.Close()
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''sector  和 Channel Group文件制作
        sender.reportProgress(20, "Forte文件制作......")
        If fso.folderExists(myPath & "ForteFile") Then
            outPutPath = myPath & "ForteFile\"
        Else
            outPutFolder = fso.createFolder(myPath & "ForteFile")
            outPutPath = myPath & "ForteFile\"
        End If
        outSectorFile = fso.openTextFile(outPutPath & "Sectors.txt", 2, 1, 0)
        outCHGRFile = fso.openTextFile(outPutPath & "ChannelGroups.txt", 2, 1, 0)
        Try
            For x = 0 To cell_arr.count - 1
                If x = 0 Then
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''Sector 表头
                    MSC = "MSC"
                    BSC = "BSC"
                    Vendor = "Vendor"
                    site = "Site"
                    Latitude = "Latitude"
                    Longitude = "Longitude"
                    Sector = "Sector"
                    ID = "ID"
                    Master = "Master"
                    LAC = "LAC"
                    CI = "CI"
                    Keywords = "Keywords"
                    Azimuth = "Azimuth"
                    BCCH_frequency = "BCCH frequency"
                    BSIC = "BSIC"
                    Intracell_HO = "Intracell HO"
                    Synchronization_group = "Synchronization group"
                    AMR_HR_Allocation = "AMR HR Allocation"
                    AMR_HR_Threshold = "AMR HR Threshold"
                    HR_Allocation = "HR Allocation"
                    HR_Threshold = "HR Threshold"
                    TCH_allocation_priority = "TCH allocation priority"
                    GPRS_allocation_priority = "GPRS allocation priority"
                    Remote = "Remote"
                    outSectorFile.writeLine(MSC & "	" & BSC & "	" & Vendor & "	" & site & "	" & Latitude & "	" & Longitude & "	" & Sector & "	" & ID & "	" & Master & "	" & LAC & "	" & CI & "	" & Keywords & "	" & Azimuth & "	" & BCCH_frequency & "	" & BSIC & "	" & Intracell_HO & "	" & Synchronization_group & "	" & AMR_HR_Allocation & "	" & AMR_HR_Threshold & "	" & HR_Allocation & "	" & HR_Threshold & "	" & TCH_allocation_priority & "	" & GPRS_allocation_priority & "	" & Remote)
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''ChannelGroup 表头
                    Sector_chgr = "Sector"
                    Channel_Group = "Channel Group"
                    Subcell = "Subcell"
                    Band = "Band"
                    Extended = "Extended"
                    Hopping_method = "Hopping method"
                    Contains_BCCH = "Contains BCCH"
                    HSN = "HSN"
                    DTX = "DTX"
                    Power_control = "Power control"
                    Subcell_Signal_Threshold = "Subcell Signal Threshold"
                    Subcell_Tx_Power = "Subcell Tx Power"
                    TRXs = "# TRXs"
                    SDCCH_TSs = "# SDCCH TSs"
                    Fixed_Data_TSs = "# Fixed Data TSs"
                    Dynamic_Data_TSs = "# Dynamic Data TSs"
                    Priority = "Priority"
                    outCHGRFile.write(Sector_chgr & "	" & Channel_Group & "	" & Subcell & "	" & Band & "	" & Extended & "	" & Hopping_method & "	" & Contains_BCCH & "	" & HSN & "	" & DTX & "	" & Power_control & "	" & Subcell_Signal_Threshold & "	" & Subcell_Tx_Power & "	" & TRXs & "	" & SDCCH_TSs & "	" & Fixed_Data_TSs & "	" & Dynamic_Data_TSs & "	" & Priority)
                    temp2 = ""
                    temp3 = ""
                    For M = 1 To 64
                        If temp2 = "" Then
                            temp2 = "	" & "TCH " & M
                        Else
                            temp2 = temp2 & "	" & "TCH " & M
                        End If
                    Next
                    outCHGRFile.write(temp2)
                    For N = 1 To 32
                        If temp3 = "" Then
                            temp3 = "	" & "MAIO " & N
                        Else
                            temp3 = temp3 & "	" & "MAIO " & N
                        End If
                    Next
                    outCHGRFile.write(temp3)
                    outCHGRFile.write(vbCrLf)
                Else
                    ''''''''''''''''''''''''''''''''''''''Sector文件内容
                    MSC = "MSC01"
                    BSC = BSC_arr(x)
                    Vendor = "Ericsson"
                    site = cell_arr(x)
                    Sector = cell_arr(x)
                    ID = ""
                    Master = ""
                    LAC = 65535
                    CI = x
                    Keywords = ""
                    Latitude = lat_arr(x)
                    Longitude = long_arr(x)
                    Azimuth = dir_arr(x)
                    If Azimuth = "360" Then
                        Azimuth = "0"
                    ElseIf InStr(Azimuth, "/") > 0 Then
                        Azimuth = Strings.Split(Azimuth, "/")(0)
                    ElseIf InStr(Azimuth, "\") > 0 Then
                        Azimuth = Strings.Split(Azimuth, "\")(0)
                    ElseIf InStr(Azimuth, "、") > 0 Then
                        Azimuth = Strings.Split(Azimuth, "、")(0)
                    End If
                    Randomize()
                    '''''''''''''''''''''''''''取一定范围内的随机数int((上限-下限+1)*rnd()+下限)
                    BCCH_frequency = Int((124) * Rnd() + 1)
                    BSIC = Int(8 * Rnd()) & Int(8 * Rnd())
                    Intracell_HO = "FALSE"
                    Synchronization_group = BSC_arr(x) & "_RXOTG-" & CI
                    AMR_HR_Allocation = "FALSE"
                    AMR_HR_Threshold = "0"
                    HR_Allocation = "FALSE"
                    HR_Threshold = "0"
                    TCH_allocation_priority = "Random"
                    GPRS_allocation_priority = "No Preference"
                    Remote = "FALSE"
                    outSectorFile.writeLine(MSC & "	" & BSC & "	" & Vendor & "	" & site & "	" & Latitude & "	" & Longitude & "	" & Sector & "	" & ID & "	" & Master & "	" & LAC & "	" & CI & "	" & Keywords & "	" & Azimuth & "	" & BCCH_frequency & "	" & BSIC & "	" & Intracell_HO & "	" & Synchronization_group & "	" & AMR_HR_Allocation & "	" & AMR_HR_Threshold & "	" & HR_Allocation & "	" & HR_Threshold & "	" & TCH_allocation_priority & "	" & GPRS_allocation_priority & "	" & Remote)

                    ''''''''''''''''''''''''''''''''''''''Channel Group文件内容
                    Sector_chgr = cell_arr(x)
                    Channel_Group = 0
                    Subcell = "UL"
                    Band = 900
                    Extended = "PGSM"
                    Hopping_method = "Non hopping"
                    Contains_BCCH = "TRUE"
                    HSN = 0
                    DTX = "Downlink and Uplink"
                    Power_control = "Downlink and Uplink"
                    Subcell_Signal_Threshold = "N/A"
                    Subcell_Tx_Power = "43"
                    TRXs = 1
                    SDCCH_TSs = 1
                    Fixed_Data_TSs = 1
                    Dynamic_Data_TSs = 0
                    Priority = "Normal"
                    outCHGRFile.write(Sector_chgr & "	" & Channel_Group & "	" & Subcell & "	" & Band & "	" & Extended & "	" & Hopping_method & "	" & Contains_BCCH & "	" & HSN & "	" & DTX & "	" & Power_control & "	" & Subcell_Signal_Threshold & "	" & Subcell_Tx_Power & "	" & TRXs & "	" & SDCCH_TSs & "	" & Fixed_Data_TSs & "	" & Dynamic_Data_TSs & "	" & Priority)
                    temp2 = ""
                    For M = 1 To 96
                        If temp2 = "" Then
                            temp2 = "	" & "N/A"
                        Else
                            temp2 = temp2 & "	" & "N/A"
                        End If
                    Next
                    outCHGRFile.write(temp2)
                    outCHGRFile.write(vbCrLf)

                End If
                sender.reportProgress(20 + ((x + 1) / cell_arr.count) * 80, "Forte文件制作......")
            Next
        Catch ex As Exception
            MsgBox("Forte文件制作出错，请检查！" & vbCrLf & ex.ToString)
            Exit Sub
        Finally
            outSectorFile.close()
            outCHGRFile.close()
            excel_app.Quit()
            excel_app = Nothing
            siteFile = Nothing
            outSectorFile = Nothing
            outPutFolder = Nothing
            outCHGRFile = Nothing
            cell_arr = Nothing
            long_arr = Nothing
            lat_arr = Nothing
            dir_arr = Nothing
            BSC_arr = Nothing
            fso = Nothing
            excel_book = Nothing
            excel_app = Nothing
        End Try

        sender.reportProGress(100, "Forte文件制作完成！")
        MsgBox("Forte文件制作完成！")
        showFolder()
    End Sub
    Private Sub showFolder()
        Shell("explorer.exe " & outPutPath, 1)
        outPutPath = Nothing
    End Sub
End Module

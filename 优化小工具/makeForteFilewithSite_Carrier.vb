Imports System.ComponentModel

Module makeForteFilewithSite_Carrier
    Private fileNames, outPutPath

    Friend Function getFiles()
        MsgBox("请选择MCOM Site 和 Carrier两个文件！")
        MainForm.OpenFileDialog1.Title = "选择Mcom Site 和 Mcom Carrier 文件"
        MainForm.OpenFileDialog1.Filter = "Mcom文件(*.txt)|*.txt"
        MainForm.OpenFileDialog1.Multiselect = True
        MainForm.OpenFileDialog1.ShowHelp = True
        MainForm.OpenFileDialog1.RestoreDirectory = True
        MainForm.OpenFileDialog1.ShowDialog()
        fileNames = MainForm.OpenFileDialog1.FileNames
        MainForm.OpenFileDialog1.Reset()
        Return fileNames
    End Function
    Friend Sub makeForteFilewithSiteCarrer(sender As Object, e As DoWorkEventArgs)
        Dim myPath, nowLineNo, temp, site_dic, outPutFolder, temp2, temp3, temp4, temp5, chgr_total
        Dim Cell_arr, BCCH_arr, BSIC_arr, TCH_arr, BSC_arr, LAC_arr, CI_arr, Hop_arr, HSN_arr
        Dim nowLine As String, filetype As String
        Dim fso, inputFile, outSectorFile, outCHGRFile
        Dim MSC, BSC, Vendor, site, Latitude, Longitude, Sector, ID, Master, LAC, CI, Keywords, Azimuth, BCCH_frequency, BSIC, Intracell_HO, Synchronization_group, AMR_HR_Allocation, AMR_HR_Threshold, HR_Allocation, HR_Threshold, TCH_allocation_priority, GPRS_allocation_priority, Remote
        Dim Sector_chgr, Channel_Group, Subcell, Band, Extended, Hopping_method, Contains_BCCH, HSN, DTX, Power_control, Subcell_Signal_Threshold, Subcell_Tx_Power, TRXs, SDCCH_TSs, Fixed_Data_TSs, Dynamic_Data_TSs, Priority
        Dim CHGR_TCH_ARR, CHGR_MAIO_ARR, SPLIT_TCH_ARR

        If fileNames.length < 2 And fileNames.length <> 0 Then
            MsgBox("请同时选择MCOM Site 和 Carrier两个文件！")
            Exit Sub
        ElseIf fileNames.length < 2 And fileNames.length = 0 Then
            Exit Sub
        End If
        sender.reportProgress(0, "读取Mcom文件数据...")
        fso = CreateObject("Scripting.FileSystemObject")
        site_dic = CreateObject("Scripting.Dictionary")
        site_dic.removeall()
        site_dic.CompareMode = vbTextCompare
        filetype = ""
        myPath = ""
        temp2 = ""
        temp3 = ""
        temp4 = ""
        temp5 = ""
        Cell_arr = New ArrayList
        BSC_arr = New ArrayList
        LAC_arr = New ArrayList
        CI_arr = New ArrayList
        BCCH_arr = New ArrayList
        BSIC_arr = New ArrayList
        TCH_arr = New ArrayList
        Hop_arr = New ArrayList
        HSN_arr = New ArrayList
        CHGR_TCH_ARR = New ArrayList
        CHGR_MAIO_ARR = New ArrayList
        SPLIT_TCH_ARR = New ArrayList
        For fileNo = 0 To UBound(fileNames)
            myPath = Strings.Left(fileNames(fileNo), InStrRev(fileNames(fileNo), "\"))
            inputFile = fso.OpenTextFile(fileNames(fileNo), 1)
            filetype = ""
            Do While Not inputFile.AtEndOfStream
                nowLineNo = inputFile.Line
                nowLine = inputFile.readline()
                If nowLine <> "" Then
                    If filetype = "" Then
                        If InStr(nowLine, "BSIC") > 0 Then
                            filetype = "carrier"
                        End If
                        If InStr(nowLine, "Longitude") > 0 Then
                            filetype = "site"
                        End If
                    End If
                    If filetype = "site" Then
                        temp = Strings.Split(nowLine, "	")
                        If InStr(temp(0), "#N/A") > 0 Or InStr(temp(1), "#N/A") > 0 Or InStr(temp(2), "#N/A") > 0 Or InStr(temp(3), "#N/A") > 0 Then
                            MsgBox("Mcom site文件中存在#N/A的内容，请修改后重试！")
                            Exit Sub
                        Else
                            Try
                                site_dic.add(temp(0), temp(2) & "#" & temp(3) & "#" & temp(4))
                            Catch ex As Exception
                                MsgBox(temp(0) & "在Mcom site信息中可能重复，请检查后重试！")
                                site_dic.removeall()
                                inputFile.close()
                                fso = Nothing
                                Exit Sub
                            End Try
                        End If
                    ElseIf filetype = "carrier" Then
                        temp = Strings.Split(nowLine, "	")
                        Cell_arr.add(temp(0))
                        BSC_arr.add(temp(1))
                        If Trim(temp(2)) = "" Then
                            LAC_arr.add("")
                        Else
                            LAC_arr.add(temp(2))
                        End If
                        CI_arr.add(temp(3))
                        If Trim(temp(4)) = "" Then
                            BCCH_arr.add("")
                        Else
                            BCCH_arr.add(temp(4))
                        End If
                        BSIC_arr.add(temp(5))
                        TCH_arr.add(temp(6))
                        Hop_arr.ADD(temp(7))
                        HSN_arr.ADD(temp(8))
                    End If
                End If
            Loop
            inputFile.close()
        Next
        If filetype = "" Then
            MsgBox("未选择有效的MCOM Site 和 Carrier文件！")
            Exit Sub
        End If
        If fso.folderExists(myPath & "ForteFile") Then
            outPutPath = myPath & "ForteFile\"
        Else
            outPutFolder = fso.createFolder(myPath & "ForteFile")
            outPutPath = myPath & "ForteFile\"
        End If
        '''''''''''''''''''Sector.txt文件制作
        sender.reportProgress(20, "正在进行Sector文件制作......")
        outSectorFile = fso.openTextFile(outPutPath & "Sectors.txt", 2, 1, 0)
        For x = 0 To Cell_arr.count - 1
            If x = 0 Then
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
            Else
                MSC = Strings.Left(BSC_arr(x), 2) & "MSC01"
                BSC = BSC_arr(x)
                Vendor = "Ericsson"
                site = Cell_arr(x)
                Sector = Cell_arr(x)
                ID = ""
                Master = ""
                LAC = LAC_arr(x)
                If LAC <> "" Then
                    If LAC = "0" Then
                        LAC = 65535
                    End If
                Else
                    LAC = 65535
                End If

                CI = CI_arr(x)
                If x <> 0 Then
                    If LAC_arr(x) & "#" & CI_arr(x) = LAC_arr(x - 1) & "#" & CI_arr(x - 1) Then
                        CI = CI_arr(x - 1) + 1
                        CI_arr(x) = CI
                    ElseIf LAC_arr(x) & "#" & CI_arr(x) = "0#0" Or LAC_arr(x) & "#" & CI_arr(x) = "65535#0" Then
                        If x > 1 Then
                            LAC = 65535
                            CI = CI_arr(x - 1) + 1
                            CI_arr(x) = CI
                        ElseIf x = 1 Then
                            LAC = 65535
                            CI = 1
                            CI_arr(x) = CI
                        End If
                    End If
                End If
                Keywords = ""
                If site_dic.exists(Cell_arr(x)) Then
                    Latitude = Strings.Split(site_dic(Cell_arr(x)), "#")(1)
                    Longitude = Strings.Split(site_dic(Cell_arr(x)), "#")(0)
                    Azimuth = Strings.Split(site_dic(Cell_arr(x)), "#")(2)
                    If Azimuth = "360" Then
                        Azimuth = "0"
                    ElseIf InStr(Azimuth, "/") > 0 Then
                        Azimuth = Strings.Split(Azimuth, "/")(0)
                    ElseIf InStr(Azimuth, "\") > 0 Then
                        Azimuth = Strings.Split(Azimuth, "\")(0)
                    ElseIf InStr(Azimuth, "、") > 0 Then
                        Azimuth = Strings.Split(Azimuth, "、")(0)
                    End If
                Else
                    Latitude = ""
                    Longitude = ""
                    Azimuth = ""
                End If
                BCCH_frequency = BCCH_arr(x)
                BSIC = BSIC_arr(x)
                Intracell_HO = "FALSE"
                Synchronization_group = BSC_arr(x) & "_RXOTG-" & CI_arr(x)
                AMR_HR_Allocation = "FALSE"
                AMR_HR_Threshold = "0"
                HR_Allocation = "FALSE"
                HR_Threshold = "0"
                TCH_allocation_priority = "Random"
                GPRS_allocation_priority = "No Preference"
                Remote = "FALSE"
                outSectorFile.writeLine(MSC & "	" & BSC & "	" & Vendor & "	" & site & "	" & Latitude & "	" & Longitude & "	" & Sector & "	" & ID & "	" & Master & "	" & LAC & "	" & CI & "	" & Keywords & "	" & Azimuth & "	" & BCCH_frequency & "	" & BSIC & "	" & Intracell_HO & "	" & Synchronization_group & "	" & AMR_HR_Allocation & "	" & AMR_HR_Threshold & "	" & HR_Allocation & "	" & HR_Threshold & "	" & TCH_allocation_priority & "	" & GPRS_allocation_priority & "	" & Remote)
            End If
            sender.reportProgress(20 + ((x + 1) / Cell_arr.count) * 40, "正在进行Sector文件制作......")
        Next
        '''''''''''''''''''''''''''''''''ChannelGroups.txt制作
        sender.reportProgress(60, "正在进行ChannelGroups.txt制作......")
        outCHGRFile = fso.openTextFile(outPutPath & "ChannelGroups.txt", 2, 1, 0)
        For y = 0 To Cell_arr.count - 1
            If y = 0 Then                      '表头
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
                For M = 1 To 64
                    If temp2 = "" Then
                        temp2 = "	" & "TCH " & M
                    Else
                        temp2 = temp2 & "	" & "TCH " & M
                    End If
                    CHGR_TCH_ARR.ADD("	" & "TCH " & M)
                Next
                outCHGRFile.write(temp2)
                For N = 1 To 32
                    If temp3 = "" Then
                        temp3 = "	" & "MAIO " & N
                    Else
                        temp3 = temp3 & "	" & "MAIO " & N
                    End If
                    CHGR_MAIO_ARR.ADD("	" & "MAIO " & N)
                Next
                outCHGRFile.write(temp3)
                outCHGRFile.write(vbCrLf)
            ElseIf y > 0 Then                 '''''''''''''内容
                DTX = "Downlink and Uplink"
                Power_control = "Downlink and Uplink"
                Subcell_Signal_Threshold = "N/A"
                Subcell_Tx_Power = 43
                SDCCH_TSs = 1
                Fixed_Data_TSs = 1
                Dynamic_Data_TSs = 0
                Priority = "Normal"
                Sector_chgr = Cell_arr(y)
                Subcell = ""
                Band = ""
                Extended = ""
                Hopping_method = ""
                Contains_BCCH = ""
                HSN = ""
                TRXs = ""
                If InStr(TCH_arr(y), ",") > 0 Then
                    TCH_arr(y) = Strings.Replace(TCH_arr(y), ",", " ", 1)
                End If
                If InStr(TCH_arr(y), ";") > 0 Then
                    TCH_arr(y) = Strings.Replace(TCH_arr(y), ";", " ", 1)
                End If
                If TCH_arr(y) <> "" Then
                    TCH_arr(y) = Strings.Replace(TCH_arr(y), "  ", " ", 1)
                End If
                TCH_arr(y) = Trim(TCH_arr(y))
                If InStr(TCH_arr(y), "[") > 0 Then                                        '带信道组的情况
                    SPLIT_TCH_ARR = Strings.Split(Trim(TCH_arr(y)), "[")
                    Sector_chgr = Cell_arr(y)
                    If BCCH_arr(y) <> "" Then
                        If BCCH_arr(y) > 100 Then
                            Subcell = "UL"
                            Band = 1800
                            Extended = "N/A"
                        Else
                            Subcell = "UL"
                            Band = 900
                            Extended = "PGSM"
                        End If
                    End If
                    If InStr(Hop_arr(y), "[") > 0 Then
                        If Strings.Split(Hop_arr(y), " ")(1) = "ON" Then
                            Hopping_method = "Base band"
                        ElseIf Strings.Split(Hop_arr(y), " ")(1) = "OFF" Then
                            Hopping_method = "Non hopping"
                        ElseIf Hop_arr(y) = "" Then
                            Hopping_method = "Non hopping"
                        End If
                    Else
                        If InStr(Hop_arr(y), "ON") > 0 Then
                            Hopping_method = "Base band"
                        ElseIf InStr(Hop_arr(y), "OFF") > 0 Then
                            Hopping_method = "Non hopping"
                        ElseIf Hop_arr(y) = "" Then
                            Hopping_method = "Non hopping"
                        End If
                    End If
                    For x = 1 To UBound(SPLIT_TCH_ARR)
                        If InStr(HSN_arr(y), "[") > 0 Then
                            HSN = Strings.Right(Strings.Split(HSN_arr(y), "[")(x), Len(Strings.Split(HSN_arr(y), "[")(x)) - 3)
                        Else
                            If HSN_arr(y) <> "" Then
                                HSN = HSN_arr(y)
                            ElseIf HSN_arr(y) = "" Then
                                HSN = 0
                            End If
                        End If
                        If x = 1 And Strings.Left(SPLIT_TCH_ARR(x), 1) > 0 Then
                            Channel_Group = 0
                            Contains_BCCH = "TRUE"
                            TRXs = 1
                            temp4 = ""
                            For M = 1 To 96
                                temp4 = temp4 & "	" & "N/A"
                            Next
                            outCHGRFile.write(Sector_chgr & "	" & Channel_Group & "	" & Subcell & "	" & Band & "	" & Extended & "	" & Hopping_method & "	" & Contains_BCCH & "	" & HSN & "	" & DTX & "	" & Power_control & "	" & Subcell_Signal_Threshold & "	" & Subcell_Tx_Power & "	" & TRXs & "	" & SDCCH_TSs & "	" & Fixed_Data_TSs & "	" & Dynamic_Data_TSs & "	" & Priority)
                            outCHGRFile.write(temp4)
                            temp4 = ""
                            outCHGRFile.write(vbCrLf)
                            Channel_Group = Strings.Left(SPLIT_TCH_ARR(x), 1)
                            Contains_BCCH = "FALSE"
                            TRXs = UBound(Strings.Split(Strings.Right(Trim(SPLIT_TCH_ARR(x)), Len(Trim(SPLIT_TCH_ARR(x))) - 3), " ")) + 1
                            temp4 = ""
                            For j = 1 To TRXs
                                temp4 = temp4 & "	" & Strings.Split(Strings.Right(Strings.Trim(SPLIT_TCH_ARR(x)), Len(Strings.Trim(SPLIT_TCH_ARR(x))) - 3), " ")(j - 1)
                            Next
                            For M = TRXs + 1 To 96
                                temp4 = temp4 & "	" & "N/A"
                            Next
                            outCHGRFile.write(Sector_chgr & "	" & Channel_Group & "	" & Subcell & "	" & Band & "	" & Extended & "	" & Hopping_method & "	" & Contains_BCCH & "	" & HSN & "	" & DTX & "	" & Power_control & "	" & Subcell_Signal_Threshold & "	" & Subcell_Tx_Power & "	" & TRXs & "	" & SDCCH_TSs & "	" & Fixed_Data_TSs & "	" & Dynamic_Data_TSs & "	" & Priority)
                            outCHGRFile.write(temp4)
                            temp4 = ""
                            outCHGRFile.write(vbCrLf)
                        Else
                            Channel_Group = Strings.Left(SPLIT_TCH_ARR(x), 1)
                            If Channel_Group = 0 Or Channel_Group = "0" Then
                                Contains_BCCH = "TRUE"
                                TRXs = UBound(Strings.Split(Strings.Right(Trim(SPLIT_TCH_ARR(x)), Len(Trim(SPLIT_TCH_ARR(x))) - 3), " ")) + 2
                                For j = 1 To TRXs - 1
                                    temp4 = temp4 & "	" & Strings.Split(Strings.Right(Strings.Trim(SPLIT_TCH_ARR(x)), Len(Strings.Trim(SPLIT_TCH_ARR(x))) - 3), " ")(j - 1)
                                Next
                                For M = TRXs To 96
                                    temp4 = temp4 & "	" & "N/A"
                                Next
                            Else
                                Contains_BCCH = "FALSE"
                                TRXs = UBound(Strings.Split(Strings.Right(Trim(SPLIT_TCH_ARR(x)), Len(Trim(SPLIT_TCH_ARR(x))) - 3), " ")) + 1
                                For j = 1 To TRXs
                                    temp4 = temp4 & "	" & Strings.Split(Strings.Right(Strings.Trim(SPLIT_TCH_ARR(x)), Len(Strings.Trim(SPLIT_TCH_ARR(x))) - 3), " ")(j - 1)
                                Next
                                For M = TRXs + 1 To 96
                                    temp4 = temp4 & "	" & "N/A"
                                Next
                            End If
                            outCHGRFile.write(Sector_chgr & "	" & Channel_Group & "	" & Subcell & "	" & Band & "	" & Extended & "	" & Hopping_method & "	" & Contains_BCCH & "	" & HSN & "	" & DTX & "	" & Power_control & "	" & Subcell_Signal_Threshold & "	" & Subcell_Tx_Power & "	" & TRXs & "	" & SDCCH_TSs & "	" & Fixed_Data_TSs & "	" & Dynamic_Data_TSs & "	" & Priority)
                            outCHGRFile.write(temp4)
                            temp4 = ""
                            outCHGRFile.write(vbCrLf)
                        End If
                    Next
                Else                                                                  '不带信道组的情况
                    If UBound(Strings.Split(Trim(TCH_arr(y)), " ")) <= 10 Or Trim(TCH_arr(y)) = "" Then              'TCH小于11的情况
                        Sector_chgr = Cell_arr(y)
                        Channel_Group = 0
                        If BCCH_arr(y) <> "" Then
                            If CInt(BCCH_arr(y)) > 100 Then

                                Subcell = "UL"
                                Band = 1800
                                Extended = "N/A"
                            Else
                                Subcell = "UL"
                                Band = 900
                                Extended = "PGSM"
                            End If
                        End If
                        If InStr(Hop_arr(y), "[") > 0 Then
                            If Strings.Split(Hop_arr(y), " ")(1) = "ON" Then
                                Hopping_method = "Base band"
                            ElseIf Strings.Split(Hop_arr(y), " ")(1) = "OFF" Then
                                Hopping_method = "Non hopping"
                            ElseIf Hop_arr(y) = "" Then
                                Hopping_method = "Non hopping"
                            End If
                        Else
                            If InStr(Hop_arr(y), "ON") > 0 Then
                                Hopping_method = "Base band"
                            ElseIf InStr(Hop_arr(y), "OFF") > 0 Then
                                Hopping_method = "Non hopping"
                            ElseIf Hop_arr(y) = "" Then
                                Hopping_method = "Non hopping"
                            End If
                        End If
                        Contains_BCCH = "TRUE"
                        If InStr(HSN_arr(y), "[") > 0 Then
                        Else
                            If HSN_arr(y) <> "" Then
                                HSN = HSN_arr(y)
                            ElseIf HSN_arr(y) = "" Then
                                HSN = 0
                            End If
                        End If
                        'If InStr(TCH_arr(y), ",") > 0 Then
                        '    TCH_arr(y) = Strings.Replace(TCH_arr(y), ",", " ", 1)
                        'End If
                        If Len(Trim(TCH_arr(y))) = 0 Then
                            TRXs = 1
                        ElseIf UBound(Strings.Split(Trim(TCH_arr(y)), " ")) >= 0 And Len(Trim(TCH_arr(y))) > 0 Then
                            TRXs = UBound(Strings.Split(Trim(TCH_arr(y)), " ")) + 2
                        End If
                        For Z = 0 To UBound(Strings.Split(Trim(TCH_arr(y)), " "))
                            If TRXs = 1 Then
                                temp4 = ""
                            Else
                                temp4 = temp4 & "	" & Strings.Split(Trim(TCH_arr(y)), " ")(Z)
                            End If
                        Next

                        For M = TRXs To 96
                            temp4 = temp4 & "	" & "N/A"
                        Next
                        outCHGRFile.write(Sector_chgr & "	" & Channel_Group & "	" & Subcell & "	" & Band & "	" & Extended & "	" & Hopping_method & "	" & Contains_BCCH & "	" & HSN & "	" & DTX & "	" & Power_control & "	" & Subcell_Signal_Threshold & "	" & Subcell_Tx_Power & "	" & TRXs & "	" & SDCCH_TSs & "	" & Fixed_Data_TSs & "	" & Dynamic_Data_TSs & "	" & Priority)
                        outCHGRFile.write(temp4)
                        temp4 = ""
                        outCHGRFile.write(vbCrLf)
                    Else                                             'TCH大于11的情况
                        chgr_total = -(Int(-((UBound(Strings.Split(Trim(TCH_arr(y)), " ")) + 1) / 11)))                          '判断信道组的数量

                        For x = 0 To chgr_total - 1
                            Sector_chgr = Cell_arr(y)
                            Channel_Group = x
                            temp4 = ""
                            If BCCH_arr(y) <> "" Then
                                If BCCH_arr(y) > 100 Then
                                    Subcell = "UL"
                                    Band = 1800
                                    Extended = "N/A"
                                Else
                                    Subcell = "UL"
                                    Band = 900
                                    Extended = "PGSM"
                                End If
                            End If
                            If InStr(Hop_arr(y), "[") > 0 Then
                                If Strings.Split(Hop_arr(y), " ")(1) = "ON" Then
                                    Hopping_method = "Base band"
                                ElseIf Strings.Split(Hop_arr(y), " ")(1) = "OFF" Then
                                    Hopping_method = "Non hopping"
                                ElseIf Hop_arr(y) = "" Then
                                    Hopping_method = "Non hopping"
                                End If
                            Else
                                If InStr(Hop_arr(y), "ON") > 0 Then
                                    Hopping_method = "Base band"
                                ElseIf InStr(Hop_arr(y), "OFF") > 0 Then
                                    Hopping_method = "Non hopping"
                                ElseIf Hop_arr(y) = "" Then
                                    Hopping_method = "Non hopping"
                                End If
                            End If
                            If InStr(HSN_arr(y), "[") > 0 Then
                            Else
                                If HSN_arr(y) <> "" Then
                                    HSN = HSN_arr(y)
                                ElseIf HSN_arr(y) = "" Then
                                    HSN = 0
                                End If
                            End If
                            If x = 0 Then                                       '信道组0
                                Contains_BCCH = "TRUE"
                                TRXs = 12
                                For Z = 0 To 10
                                    temp4 = temp4 & "	" & Strings.Split(Trim(TCH_arr(y)), " ")(Z)
                                Next
                                For M = 12 To 96
                                    If temp4 = "" Then
                                        temp4 = "	" & "N/A"
                                    Else
                                        temp4 = temp4 & "	" & "N/A"
                                    End If
                                Next
                            ElseIf x = 1 And (UBound(Strings.Split(Trim(TCH_arr(y)), " ")) + 1 - 11) <= 11 Then        '信道组1
                                Contains_BCCH = "FALSE"
                                TRXs = UBound(Strings.Split(Trim(TCH_arr(y)), " ")) + 1 - 11
                                For Z = 11 To UBound(Strings.Split(Trim(TCH_arr(y)), " "))
                                    temp4 = temp4 & "	" & Strings.Split(Trim(TCH_arr(y)), " ")(Z)
                                Next
                                For M = TRXs + 1 To 96
                                    If temp4 = "" Then
                                        temp4 = "	" & "N/A"
                                    Else
                                        temp4 = temp4 & "	" & "N/A"
                                    End If
                                Next
                            ElseIf x = 1 And (UBound(Strings.Split(Trim(TCH_arr(y)), " ")) + 1 - 11) > 11 Then             '信道组1
                                Contains_BCCH = "TRUE"
                                TRXs = 11
                                For Z = 11 To 21
                                    temp4 = temp4 & "	" & Strings.Split(Trim(TCH_arr(y)), " ")(Z)
                                Next
                                For M = 12 To 96
                                    If temp4 = "" Then
                                        temp4 = "	" & "N/A"
                                    Else
                                        temp4 = temp4 & "	" & "N/A"
                                    End If
                                Next
                            ElseIf x = 2 And (UBound(Strings.Split(Trim(TCH_arr(y)), " ")) + 1 - 22) <= 11 Then         '信道组2
                                Contains_BCCH = "FALSE"
                                TRXs = UBound(Strings.Split(Trim(TCH_arr(y)), " ")) + 1 - 22
                                For Z = 22 To UBound(Strings.Split(Trim(TCH_arr(y)), " "))
                                    temp4 = temp4 & "	" & Strings.Split(Trim(TCH_arr(y)), " ")(Z)
                                Next
                                For M = TRXs + 1 To 96
                                    If temp4 = "" Then
                                        temp4 = "	" & "N/A"
                                    Else
                                        temp4 = temp4 & "	" & "N/A"
                                    End If
                                Next
                                'ElseIf x = 2 And (UBound(Strings.Split(TCH_arr(y), " ")) + 1 - 22) > 11 Then            '信道组2
                                '    Contains_BCCH = "FALSE"
                                '    TRXs = 11
                            End If
                            outCHGRFile.write(Sector_chgr & "	" & Channel_Group & "	" & Subcell & "	" & Band & "	" & Extended & "	" & Hopping_method & "	" & Contains_BCCH & "	" & HSN & "	" & DTX & "	" & Power_control & "	" & Subcell_Signal_Threshold & "	" & Subcell_Tx_Power & "	" & TRXs & "	" & SDCCH_TSs & "	" & Fixed_Data_TSs & "	" & Dynamic_Data_TSs & "	" & Priority)
                            outCHGRFile.write(temp4)
                            temp4 = ""
                            outCHGRFile.write(vbCrLf)
                        Next
                    End If
                End If
            End If
            sender.reportProgress(60 + ((y + 1) / Cell_arr.count) * 40, "正在进行ChannelGroups.txt制作......")
        Next
        sender.reportProgress(100, "完成！")
        outSectorFile.close()
        outCHGRFile.close()
        Cell_arr = Nothing
        BSC_arr = Nothing
        LAC_arr = Nothing
        CI_arr = Nothing
        BCCH_arr = Nothing
        BSIC_arr = Nothing
        TCH_arr = Nothing
        inputFile = Nothing
        outSectorFile = Nothing
        outCHGRFile = Nothing
        outPutFolder = Nothing
        Cell_arr = Nothing
        BSC_arr = Nothing
        LAC_arr = Nothing
        CI_arr = Nothing
        BCCH_arr = Nothing
        BSIC_arr = Nothing
        TCH_arr = Nothing
        Hop_arr = Nothing
        HSN_arr = Nothing
        CHGR_TCH_ARR = Nothing
        CHGR_MAIO_ARR = Nothing
        SPLIT_TCH_ARR = Nothing
        fso = Nothing
        myPath = Nothing
        fileNames = Nothing
        site_dic.removeall()
        site_dic = Nothing
        sender.reportProGress(100, "Forte文件制作完成！")
        MsgBox("Forte文件制作完成！")
        showFolder()
    End Sub
    Friend Sub showFolder()
        Shell("explorer.exe " & outPutPath, 1)
        outPutPath = Nothing
    End Sub
End Module

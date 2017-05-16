Imports System.ComponentModel

Module CreateCelFile
    Private siteFilePath
    Private carrierFilePath
    Private neighborFilePath

    Friend Sub startCreateCelFile(sender As Object, e As DoWorkEventArgs)
        Dim outFolder As String
        Dim FSO As Object, outFile As Object
        Dim MaxNeighborNo As Integer, MaxTCHNo As Integer
        Dim Ncell_TEMP As String, TCH_ARFCN_TEMP As String, tempStr As String
        Dim Cell_carrier_arr As ArrayList, BCCH_ARR As ArrayList, BSIC_ARR As ArrayList, LAC_ARR As ArrayList, CI_ARR As ArrayList, TCH_ARR As ArrayList
        Dim Cell_site_arr As ArrayList, Lat_arr As ArrayList, Lon_arr As ArrayList, ANT_DIRECTION_ARR As ArrayList, ANT_BEAM_WIDTH_ARR As ArrayList, site_name_arr As ArrayList
        Dim Cell_neighbor_arr As ArrayList, Ncell_arr As ArrayList
        Dim openFileNo As Integer
        Dim tempLine As String, nowLineNo As Integer, temp_arr As Object, NCELL_ARR_TEMP As Object, tempIndexNo As Integer
        Dim CELL As String, ARFCN As String, BSIC As String, Lat As String, Lon As String, LAC As String, CI As String, ANT_DIRECTION As String, ANT_BEAM_WIDTH As String
        CELL = ""
        ARFCN = ""
        BSIC = ""
        Lat = ""
        Lon = ""
        LAC = ""
        CI = ""
        ANT_DIRECTION = ""
        ANT_BEAM_WIDTH = ""
        Ncell_TEMP = ""
        TCH_ARFCN_TEMP = ""
        tempStr = ""
        tempLine = ""
        tempIndexNo = 0
        If InStr(e.Argument, "#") Then
            If UBound(Split(e.Argument, "#")) = 2 Then
                siteFilePath = Split(e.Argument, "#")(0)
                carrierFilePath = Split(e.Argument, "#")(1)
                neighborFilePath = Split(e.Argument, "#")(2)
            ElseIf UBound(Split(e.Argument, "#")) = 1 Then
                siteFilePath = Split(e.Argument, "#")(0)
                carrierFilePath = Split(e.Argument, "#")(1)
            End If
        End If
        outFolder = Left(siteFilePath, InStrRev(siteFilePath, "\"))
        Cell_site_arr = New ArrayList
        Lat_arr = New ArrayList
        Lon_arr = New ArrayList
        ANT_DIRECTION_ARR = New ArrayList
        ANT_BEAM_WIDTH_ARR = New ArrayList
        site_name_arr = New ArrayList
        Cell_carrier_arr = New ArrayList
        BCCH_ARR = New ArrayList
        BSIC_ARR = New ArrayList
        TCH_ARR = New ArrayList
        LAC_ARR = New ArrayList
        CI_ARR = New ArrayList
        Cell_neighbor_arr = New ArrayList
        Ncell_arr = New ArrayList
        FSO = CreateObject("Scripting.FileSystemObject")
        openFileNo = FreeFile()
        nowLineNo = 1
        sender.reportProgress(0, "读取Mcom Site数据...")
        Try
            FileOpen(openFileNo, siteFilePath, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
            Do While Not EOF(openFileNo)
                tempLine = LineInput(openFileNo)
                temp_arr = Split(tempLine, "	")
                If nowLineNo = 1 Then
                    If UCase(temp_arr(0)) <> "CELL" Or Strings.Left(UCase(temp_arr(2)), 3) <> "LON" Or Strings.Left(UCase(temp_arr(3)), 3) <> "LAT" Or UCase(temp_arr(4)) <> "DIR" Then
                        MsgBox("Mcom Site文件表头不正确，请重新导入！" & vbCrLf & tempLine, MsgBoxStyle.Critical)
                        nowLineNo = 1
                        Exit Sub
                    End If
                Else
                    Cell_site_arr.Add(temp_arr(0))
                    Lat_arr.Add(temp_arr(3))
                    Lon_arr.Add(temp_arr(2))
                    ANT_DIRECTION_ARR.Add(temp_arr(4))
                    If temp_arr(4) = "360" Then
                        ANT_BEAM_WIDTH_ARR.Add("360")
                    Else
                        ANT_BEAM_WIDTH_ARR.Add("65")
                    End If
                    site_name_arr.Add(temp_arr(12))
                End If
                nowLineNo = nowLineNo + 1
            Loop
        Catch ex As Exception
            MsgBox("Mcom site 数据有误，请检查！", MsgBoxStyle.Critical)
            FSO = Nothing
            outFile = Nothing
            Cell_site_arr = Nothing
            Lat_arr = Nothing
            Lon_arr = Nothing
            ANT_DIRECTION_ARR = Nothing
            ANT_BEAM_WIDTH_ARR = Nothing
            site_name_arr = Nothing
            Cell_carrier_arr = Nothing
            BCCH_ARR = Nothing
            BSIC_ARR = Nothing
            TCH_ARR = Nothing
            LAC_ARR = Nothing
            CI_ARR = Nothing
            Cell_neighbor_arr = Nothing
            Ncell_arr = Nothing
            Exit Sub
        Finally
            nowLineNo = 1
            FileClose(openFileNo)
        End Try
        sender.reportProgress(25, "读取Mcom Site数据完毕...")
        sender.reportProgress(25, "读取Mcom Carrier数据...")
        openFileNo = FreeFile()
        nowLineNo = 1
        Try
            FileOpen(openFileNo, carrierFilePath, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
            Do While Not EOF(openFileNo)
                tempLine = LineInput(openFileNo)
                temp_arr = Split(tempLine, "	")
                If nowLineNo = 1 Then
                    If UCase(temp_arr(0)) <> "CELL" Or UCase(temp_arr(2)) <> "LAI" Or UCase(temp_arr(3)) <> "CI" Or UCase(temp_arr(4)) <> "BCCH" Or UCase(temp_arr(5)) <> "BSIC" Then
                        MsgBox("Mcom Carrier文件不正确，请重新导入！", MsgBoxStyle.Critical)
                        nowLineNo = 1
                        Exit Sub
                    End If
                Else
                    Cell_carrier_arr.Add(temp_arr(0))
                    LAC_ARR.Add(temp_arr(2))
                    CI_ARR.Add(temp_arr(3))
                    BCCH_ARR.Add(temp_arr(4))
                    BSIC_ARR.Add(temp_arr(5))
                    If InStr(temp_arr(6), "[") Then
                        If InStr(Trim(temp_arr(6)), " ") Then
                            For s = 0 To UBound(Split(temp_arr(6), " "))
                                If InStr(Split(temp_arr(6), " ")(s), "[") = 0 Then
                                    If tempStr = "" Then
                                        tempStr = Split(temp_arr(6), " ")(s)
                                    Else
                                        tempStr = tempStr & " " & Split(temp_arr(6), " ")(s)
                                    End If
                                End If
                            Next
                            TCH_ARR.Add(tempStr)
                            If UBound(Split(tempStr, " ")) + 1 > MaxTCHNo Then
                                MaxTCHNo = UBound(Split(tempStr, " ")) + 1
                            End If
                            tempStr = ""
                        Else
                            TCH_ARR.Add(Right(Trim(temp_arr(6)), Len(Trim(temp_arr(6))) - 4))
                        End If
                    Else
                        TCH_ARR.Add(temp_arr(6))
                        If UBound(Split(temp_arr(6), " ")) + 1 > MaxTCHNo Then
                            MaxTCHNo = UBound(Split(tempStr, " ")) + 1
                        End If
                        tempStr = ""
                    End If
                End If
                nowLineNo = nowLineNo + 1
            Loop
        Catch ex As Exception
            MsgBox("Mcom Carrier 数据有误，请检查！" & ex.ToString(), MsgBoxStyle.Critical)
            FSO = Nothing
            outFile = Nothing
            Cell_site_arr = Nothing
            Lat_arr = Nothing
            Lon_arr = Nothing
            ANT_DIRECTION_ARR = Nothing
            ANT_BEAM_WIDTH_ARR = Nothing
            site_name_arr = Nothing
            Cell_carrier_arr = Nothing
            BCCH_ARR = Nothing
            BSIC_ARR = Nothing
            TCH_ARR = Nothing
            LAC_ARR = Nothing
            CI_ARR = Nothing
            Cell_neighbor_arr = Nothing
            Ncell_arr = Nothing
            Exit Sub
        Finally
            nowLineNo = 1
            FileClose(openFileNo)
        End Try
        sender.reportProgress(50, "读取Mcom Carrier数据完毕...")
        If neighborFilePath <> "" Then
            sender.reportProgress(50, "读取Mcom Neighbor数据...")
            openFileNo = FreeFile()
            nowLineNo = 1
            Try
                FileOpen(openFileNo, neighborFilePath, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
                Do While Not EOF(openFileNo)
                    tempLine = LineInput(openFileNo)
                    temp_arr = Split(tempLine, "	")
                    If nowLineNo = 1 Then
                        If UCase(temp_arr(0)) <> "CELL" Or UCase(temp_arr(1)) <> "NCELL" Then
                            MsgBox("Mcom Neighbor文件不正确，请重新导入！", MsgBoxStyle.Critical)
                            nowLineNo = 1
                            Exit Sub
                        End If
                    Else
                        Cell_neighbor_arr.Add(temp_arr(0))
                        Ncell_arr.Add(temp_arr(1))
                        If InStr(Trim(temp_arr(1)), " ") Then
                            If UBound(Split(Trim(temp_arr(1)), " ")) + 1 > MaxNeighborNo Then
                                MaxNeighborNo = UBound(Split(Trim(temp_arr(1)), " ")) + 1
                            End If
                        Else
                            If MaxNeighborNo = 0 Then
                                MaxNeighborNo = 1
                            End If
                        End If
                    End If
                    nowLineNo = nowLineNo + 1
                Loop
            Catch ex As Exception
                MsgBox("Mcom Neighbor 数据有误，请检查！", MsgBoxStyle.Critical)
                FSO = Nothing
                outFile = Nothing
                Cell_site_arr = Nothing
                Lat_arr = Nothing
                Lon_arr = Nothing
                ANT_DIRECTION_ARR = Nothing
                ANT_BEAM_WIDTH_ARR = Nothing
                site_name_arr = Nothing
                Cell_carrier_arr = Nothing
                BCCH_ARR = Nothing
                BSIC_ARR = Nothing
                TCH_ARR = Nothing
                LAC_ARR = Nothing
                CI_ARR = Nothing
                Cell_neighbor_arr = Nothing
                Ncell_arr = Nothing
                Exit Sub
            Finally
                nowLineNo = 1

                FileClose(openFileNo)
            End Try

        End If
        sender.reportProgress(75, "制作.CELL文件..." & "0%")
        If Not FSO.folderExists(outFolder & "TEMS cell\") Then
            FSO.createFolder(outFolder & "TEMS cell\")
        End If
        outFile = FSO.OpenTextFile(outFolder & "TEMS cell\TEMS_Cell_" & Format(Now, "yyyyMMdd") & ".cel", 2, 1, 0)
        '2 TEMS_-_Cell_names
        'CELL	ARFCN	BSIC	Lat	Lon	LAC	CI	ANT_DIRECTION	ANT_BEAM_WIDTH
        For x = 1 To MaxNeighborNo
            If x = 1 Then
                Ncell_TEMP = "LAC_N_" & x & "	" & "CI_N_" & x
            Else
                Ncell_TEMP = Trim(Ncell_TEMP & "	" & "LAC_N_" & x & "	" & "CI_N_" & x)
            End If
        Next
        For Y = 1 To MaxTCHNo
            If Y = 1 Then
                TCH_ARFCN_TEMP = "TCH_ARFCN_" & Y
            Else
                TCH_ARFCN_TEMP = Trim(TCH_ARFCN_TEMP & "	" & "TCH_ARFCN_" & Y)
            End If
        Next
        outFile.writeLine("2 TEMS_-_Cell_names")
        outFile.writeLine("CELL" & "	" & "ARFCN" & "	" & "BSIC" & "	" & "Lat" & "	" & "Lon" & "	" & "LAC" & "	" & "CI" & "	" & "ANT_DIRECTION" & "	" & "ANT_BEAM_WIDTH" & "	" & TCH_ARFCN_TEMP & "	" & Ncell_TEMP)
        TCH_ARFCN_TEMP = ""
        Ncell_TEMP = ""
        Try
            For W = 0 To Cell_carrier_arr.Count - 1
                CELL = Cell_carrier_arr(W)
                ARFCN = BCCH_ARR(W)
                BSIC = BSIC_ARR(W)
                LAC = LAC_ARR(W)
                CI = CI_ARR(W)
                If InStr(TCH_ARR(W), " ") <> 0 Then
                    For p = 0 To MaxTCHNo - 1
                        If p = 0 Then
                            If Split(TCH_ARR(W), " ")(p) <> BCCH_ARR(W) Then
                                TCH_ARFCN_TEMP = Split(TCH_ARR(W), " ")(p)
                            End If
                        ElseIf p <> 0 And p <= UBound(Split(TCH_ARR(W), " ")) Then
                            If Split(TCH_ARR(W), " ")(p) <> BCCH_ARR(W) Then
                                If TCH_ARFCN_TEMP = "" Then
                                    TCH_ARFCN_TEMP = Split(TCH_ARR(W), " ")(p)
                                Else
                                    TCH_ARFCN_TEMP = TCH_ARFCN_TEMP & "	" & Split(TCH_ARR(W), " ")(p)
                                End If
                            End If
                        ElseIf p <> 0 And p > UBound(Split(TCH_ARR(W), " ")) Then
                            TCH_ARFCN_TEMP = TCH_ARFCN_TEMP & "	" & ""
                        End If
                    Next
                    If UBound(Split(TCH_ARFCN_TEMP, "	")) < MaxTCHNo - 1 Then
                        TCH_ARFCN_TEMP = TCH_ARFCN_TEMP & "	" & ""
                    End If
                Else
                    For Y = 1 To MaxTCHNo
                        If Y = 1 Then
                            TCH_ARFCN_TEMP = TCH_ARR(W)
                        Else
                            TCH_ARFCN_TEMP = TCH_ARFCN_TEMP & "	" & ""
                        End If
                    Next
                End If

                If Cell_neighbor_arr.Contains(CELL) Then
                    tempIndexNo = Cell_neighbor_arr.IndexOf(CELL)
                    If InStr(Trim(Ncell_arr(tempIndexNo)), " ") Then
                        NCELL_ARR_TEMP = Split(Trim(Ncell_arr(tempIndexNo)), " ")
                        For U = 0 To UBound(NCELL_ARR_TEMP)
                            If U = 0 Then
                                If Cell_carrier_arr.Contains(NCELL_ARR_TEMP(U)) Then
                                    tempIndexNo = Cell_carrier_arr.IndexOf(NCELL_ARR_TEMP(U))
                                    Ncell_TEMP = LAC_ARR(tempIndexNo) & "	" & CI_ARR(tempIndexNo)
                                Else
                                    Ncell_TEMP = NCELL_ARR_TEMP(U) & "	" & NCELL_ARR_TEMP(U)
                                End If
                            Else
                                If Cell_carrier_arr.Contains(NCELL_ARR_TEMP(U)) Then
                                    tempIndexNo = Cell_carrier_arr.IndexOf(NCELL_ARR_TEMP(U))
                                    Ncell_TEMP = Ncell_TEMP & "	" & LAC_ARR(tempIndexNo) & "	" & CI_ARR(tempIndexNo)
                                Else
                                    Ncell_TEMP = Ncell_TEMP & "	" & NCELL_ARR_TEMP(U) & "	" & NCELL_ARR_TEMP(U)
                                End If
                            End If
                        Next
                    Else
                        Ncell_TEMP = ""
                    End If
                Else
                    For x = 1 To MaxNeighborNo
                        If x = 1 Then
                            Ncell_TEMP = "" & "	" & ""
                        Else
                            Ncell_TEMP = Ncell_TEMP & "	" & "" & "	" & ""
                        End If
                    Next
                End If
                If Cell_site_arr.Contains(CELL) Then
                    tempIndexNo = Cell_site_arr.IndexOf(CELL)
                    Lat = Lat_arr(tempIndexNo)
                    Lon = Lon_arr(tempIndexNo)
                    ANT_DIRECTION = ANT_DIRECTION_ARR(tempIndexNo)
                    ANT_BEAM_WIDTH = ANT_BEAM_WIDTH_ARR(tempIndexNo)
                    CELL = CELL & "_" & site_name_arr(tempIndexNo)
                Else
                    Lat = ""
                    Lon = ""
                    ANT_DIRECTION = ""
                    ANT_BEAM_WIDTH = ""
                End If
                outFile.writeLine(CELL & "	" & ARFCN & "	" & BSIC & "	" & Lat & "	" & Lon & "	" & LAC & "	" & CI & "	" & ANT_DIRECTION & "	" & ANT_BEAM_WIDTH & "	" & TCH_ARFCN_TEMP & "	" & Ncell_TEMP)
                Ncell_TEMP = ""
                sender.reportProgress(((W + 1) / Cell_carrier_arr.Count) * 100, "制作.CELL文件...")
            Next
        Catch ex As Exception
            MsgBox(".CEL文件制作失败！" & ex.ToString, MsgBoxStyle.Critical)
            Exit Sub
        Finally
            outFile.close()
            FSO = Nothing
            outFile = Nothing
            Cell_site_arr = Nothing
            Lat_arr = Nothing
            Lon_arr = Nothing
            ANT_DIRECTION_ARR = Nothing
            ANT_BEAM_WIDTH_ARR = Nothing
            site_name_arr = Nothing
            Cell_carrier_arr = Nothing
            BCCH_ARR = Nothing
            BSIC_ARR = Nothing
            TCH_ARR = Nothing
            LAC_ARR = Nothing
            CI_ARR = Nothing
            Cell_neighbor_arr = Nothing
            Ncell_arr = Nothing
        End Try
        MsgBox(".CEL文件制作完成！" & vbCrLf & ".CEL文件生成路径为：" & outFolder & "TEMS cell", MsgBoxStyle.Information)
        Shell("explorer.exe " & outFolder & "TEMS cell", 1)
    End Sub
End Module

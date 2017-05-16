Imports System.ComponentModel

Module CheckCoBCCH_BSIC
    Private siteFilePath As String, carrierFilePath As String
    Private site_position_dic As New Dictionary(Of String, String)
    Private carrier_cell_arr As ArrayList, bcch_bsic_arr As ArrayList
    Private Sub importSite(sender As Object, e As DoWorkEventArgs, mainForm As MainForm)
        Dim fileNo As Integer
        Dim templine As String, temp_arr As Object
        Dim nowLineNo As Integer = 0
        fileNo = FreeFile()
        Try
            FileOpen(fileNo, siteFilePath, OpenMode.Input, OpenAccess.Read, OpenShare.LockReadWrite)
            Do While Not EOF(fileNo)
                nowLineNo = nowLineNo + 1
                templine = LineInput(fileNo)
                temp_arr = Split(templine, "	")
                If nowLineNo = 1 Then
                    If Strings.Left(UCase(temp_arr(2)), 3) <> "LON" Or Strings.Left(UCase(temp_arr(3)), 3) <> "LAT" Or UCase(temp_arr(0)) <> "CELL" Then
                        MsgBox("请选择标准的Mcom Site文件！", MsgBoxStyle.Critical)
                        Exit Sub
                    End If
                End If
                If site_position_dic.ContainsKey(temp_arr(0)) Then
                    '               MsgBox("Mcom site文件中" & temp_arr(0) & "存在重复信息，第" & nowLineNo & "行的数据不使用！")
                    sender.reportProgress(0, "导入Site数据...@=>Mcom site文件中" & temp_arr(0) & "存在重复信息，第" & nowLineNo & "行的数据不使用！" & vbCrLf)
                Else
                    site_position_dic.Add(temp_arr(0), temp_arr(2) & "#" & temp_arr(3))
                End If
            Loop
        Catch ex As Exception
            MsgBox("Mcom site文件有误，请检查第" & nowLineNo & "行的数据！", MsgBoxStyle.Critical)
        Finally
            FileClose(fileNo)
        End Try
    End Sub
    Private Sub importCarrier(sender As Object, e As DoWorkEventArgs, mainForm As MainForm)
        Dim fileNo As Integer
        Dim templine As String, temp_arr As Object
        Dim nowLineNo As Integer = 0
        fileNo = FreeFile()
        Try
            FileOpen(fileNo, carrierFilePath, OpenMode.Input, OpenAccess.Read, OpenShare.LockReadWrite)
            Do While Not EOF(fileNo)
                nowLineNo = nowLineNo + 1
                templine = LineInput(fileNo)
                temp_arr = Split(templine, "	")
                If nowLineNo = 1 Then
                    If UCase(temp_arr(4)) <> "BCCH" Or UCase(temp_arr(5)) <> "BSIC" Or UCase(temp_arr(0)) <> "CELL" Then
                        MsgBox("请选择标准的Mcom Carrier文件！")
                        Exit Sub
                    End If
                End If
                If carrier_cell_arr.Contains(temp_arr(0)) Then
                    '      MsgBox("Mcom Carrier文件中" & temp_arr(0) & "存在重复信息，第" & nowLineNo & "行的数据不使用！")
                    sender.reportProgress(10, "导入Carrier数据...@=>Mcom Carrier文件中" & temp_arr(0) & "存在重复信息，第" & nowLineNo & "行的数据不使用！" & vbCrLf)
                Else
                    carrier_cell_arr.Add(temp_arr(0))
                    bcch_bsic_arr.Add(temp_arr(4) & "#" & temp_arr(5))
                End If
            Loop
        Catch ex As Exception
            MsgBox("Mcom Carrier文件有误，请检查第" & nowLineNo & "行的数据！")
        Finally
            FileClose(fileNo)
        End Try
    End Sub
    Friend Sub Check_CoBcchBsic(sender As Object, e As DoWorkEventArgs, mainForm As MainForm)
        Dim fso As Object
        Dim outPutFile As Object
        Dim outPath As String
        Dim Scell_Ncell_dic As New Dictionary(Of String, String)
        Dim distance As String
        Dim x As Integer, y As Integer
        carrier_cell_arr = New ArrayList
        bcch_bsic_arr = New ArrayList
        site_position_dic.Clear()
        If InStr(e.Argument, "#") Then
            siteFilePath = Split(e.Argument, "#")(0)
            carrierFilePath = Split(e.Argument, "#")(1)
            If siteFilePath = "" Or carrierFilePath = "" Then
                MsgBox("文件有误，请重新选择的Mcom Site 和Mcom Carrier文件！", MsgBoxStyle.Critical)
            End If
        End If
        outPath = Left(siteFilePath, InStrRev(siteFilePath, "\")) & "CheckResultsOfCo_BcchBsic_" & Format(Now, "yyyyMMdd") & ".csv"
        fso = CreateObject("Scripting.FileSystemObject")
        Try
            outPutFile = fso.openTextFile(outPath, 2, 1, 0)
            outPutFile.writeline("源小区" & "," & "目标小区" & "," & "BCCH#BSIC" & "," & "距离（KM）")
        Catch ex As Exception
            MsgBox("输出的结果文件处于打开状态，请先将其关闭，然后重试！", MsgBoxStyle.Critical)
            Exit Sub

        End Try
        sender.reportProgress(0, "导入Site数据...")
        importSite(sender, e, mainForm)
        If site_position_dic.Count = 0 Then
            Exit Sub
        End If
        sender.reportProgress(10, "导入Carrier数据...10%@=>Site数据导入完毕！" & vbCrLf)
        importCarrier(sender, e, mainForm)
        If carrier_cell_arr.Count = 0 Then
            Exit Sub
        End If
        sender.reportProgress(20, "计算同BCCH同BSIC距离...20%@=>Carrier数据导入完毕！" & vbCrLf)
        Try
            For x = 1 To carrier_cell_arr.Count - 1
                For y = 1 To carrier_cell_arr.Count - 1
                    If bcch_bsic_arr(x) = bcch_bsic_arr(y) And carrier_cell_arr(x) <> carrier_cell_arr(y) Then
                        If Not Scell_Ncell_dic.ContainsKey(carrier_cell_arr(x) & "#" & carrier_cell_arr(y)) And Not Scell_Ncell_dic.ContainsKey(carrier_cell_arr(y) & "#" & carrier_cell_arr(x)) Then
                            If site_position_dic.ContainsKey(carrier_cell_arr(x)) And site_position_dic.ContainsKey(carrier_cell_arr(y)) Then
                                Try
                                    distance = CacuDistance(Split(site_position_dic(carrier_cell_arr(x)), "#")(0), Split(site_position_dic(carrier_cell_arr(x)), "#")(1), Split(site_position_dic(carrier_cell_arr(y)), "#")(0), Split(site_position_dic(carrier_cell_arr(y)), "#")(1))
                                Catch ex As Exception
                                    distance = "经纬度信息异常！"
                                    sender.reportProgress(20 + x / (carrier_cell_arr.Count - 1) * 80, "计算同BCCH同BSIC距离..." & Mid(20 + x / (carrier_cell_arr.Count - 1) * 80, 1, 2) & "%@=>" & carrier_cell_arr(x) & "或" & carrier_cell_arr(y) & "的数据异常，请检查！" & vbCrLf)
                                End Try
                                outPutFile.writeline(carrier_cell_arr(x) & "," & carrier_cell_arr(y) & "," & bcch_bsic_arr(y) & "," & distance)
                            Else
                                outPutFile.writeline(carrier_cell_arr(x) & "," & carrier_cell_arr(y) & "," & bcch_bsic_arr(y) & "," & "源小区或目标小区无经纬度！")
                                sender.reportProgress(20 + x / (carrier_cell_arr.Count - 1) * 80, "计算同BCCH同BSIC距离..." & Mid(20 + x / (carrier_cell_arr.Count - 1) * 80, 1, 2) & "%@=>" & carrier_cell_arr(x) & "或" & carrier_cell_arr(y) & "无经纬度信息，请检查！" & vbCrLf)
                            End If
                            Scell_Ncell_dic.Add(carrier_cell_arr(x) & "#" & carrier_cell_arr(y), y)
                        End If
                    End If
                Next
                sender.reportProgress(20 + x / (carrier_cell_arr.Count - 1) * 80, "计算同BCCH同BSIC距离..." & Mid(20 + x / (carrier_cell_arr.Count - 1) * 80, 1, 2) & "%")
            Next
        Catch ex As Exception
            MsgBox("出错了，与" & carrier_cell_arr(x) & "" & carrier_cell_arr(y) & "的数据有关，请检查！", MsgBoxStyle.Critical)
        Finally
            carrier_cell_arr.Clear()
            carrier_cell_arr = Nothing
            bcch_bsic_arr.Clear()
            bcch_bsic_arr = Nothing
            Scell_Ncell_dic.Clear()
            Scell_Ncell_dic = Nothing
            outPutFile.close()
            fso = Nothing
        End Try
        sender.reportProgress(100, "计算同BCCH同BSIC距离...@=>计算同BCCH同BSIC距离完毕！结果生成路径为：" & outPath & vbCrLf)
        MsgBox("计算完毕，共耗时：" & mainForm.totalEvaTime & "秒！", MsgBoxStyle.Information)
        Shell("explorer.exe " & outPath, 1)
    End Sub
    Private Function CacuDistance(lon1 As Double, lat1 As Double, lon2 As Double, lat2 As Double)
        CacuDistance = 111.12 * Math.Acos((Math.Sin(lat1 * 3.14159265358979 / 180) * Math.Sin(lat2 * 3.14159265358979 / 180) + Math.Cos(lat1 * 3.14159265358979 / 180) * Math.Cos(lat2 * 3.14159265358979 / 180) * Math.Cos((lon2 - lon1) * 3.14159265358979 / 180))) * 180 / 3.14159265358979
    End Function
End Module

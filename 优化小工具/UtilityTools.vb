Imports 优化小工具.MainForm
Imports Microsoft.Office.Interop
Imports System.ComponentModel

Module UtilityTools
    '   Friend BSC_SITE_BCCH_DIC As Dictionary(Of String, String)
    Friend Structure McomCarrierData
        Dim cell_arr As ArrayList
        Dim site_arr As ArrayList
        Dim BSC As String
        Dim LAC As String
        Dim CI As String
        Dim BCCH As String
        Dim BSIC As String
        Dim TCH() As Object
        Dim NoOfFreq As Integer
        Dim DCH As String
        Dim Cellr() As Object
        Dim NoOfCellr As Integer
        Dim NoOfCO As Integer
        Dim NoOfCA As Integer
        Dim BSC_CELL_BCCH_ARR As ArrayList
        Dim CELL_BCCH_BSIC_DIC As Dictionary(Of String, String)
        Dim CELL_BCCH_DIC As Dictionary(Of String, String)
        Dim CELL_TRX_DIC As Dictionary(Of String, String)
        'Dim CELL_BCCH_DIC As Dictionary(Of String, String)
        'Dim CELL_TRX_DIC As Dictionary(Of String, String)
        Dim SITE_CELL_DIC As Dictionary(Of String, String)
        Dim BSC_CELL_DIC As Dictionary(Of String, String)
    End Structure
    Friend Structure McomNeighborData
        Dim cell_arr As ArrayList
        Dim ncell_arr As ArrayList
    End Structure
    Friend timeRemainForTool As Boolean = False
    Friend remainDays As Integer
    Friend Function selectFile(def_title As String, def_filter As String, def_multiselect As Boolean, def_showHelp As Boolean, def_restoreDirectory As Boolean)
        MainForm.OpenFileDialog1.Title = def_title
        MainForm.OpenFileDialog1.Filter = def_filter
        MainForm.OpenFileDialog1.Multiselect = def_multiselect
        MainForm.OpenFileDialog1.ShowHelp = def_showHelp
        MainForm.OpenFileDialog1.RestoreDirectory = def_restoreDirectory
        MainForm.OpenFileDialog1.ShowDialog()
        If def_multiselect Then
            selectFile = MainForm.OpenFileDialog1.FileNames
        Else
            selectFile = MainForm.OpenFileDialog1.FileName
        End If
        MainForm.OpenFileDialog1.Reset()
    End Function
    Friend Function getDistance(lon1 As Double, lat1 As Double, lon2 As Double, lat2 As Double)
        getDistance = 111.12 * Math.Acos((Math.Sin(lat1 * 3.14159265358979 / 180) * Math.Sin(lat2 * 3.14159265358979 / 180) + Math.Cos(lat1 * 3.14159265358979 / 180) * Math.Cos(lat2 * 3.14159265358979 / 180) * Math.Cos((lon2 - lon1) * 3.14159265358979 / 180))) * 180 / 3.14159265358979
    End Function
    Friend Sub getInitTime(fileName As String)
        Dim fso As Object
        Dim outPutFile As Object
        Dim fileNo As Integer
        Dim keyFilePath As String
        Dim existsKeyStr As Double
        Dim fileCreateTime As Double
        Dim nowTime As Double
        Dim LicenseDays As Double
        Dim keyFileCreateTime As Double
        Dim usedDays As Double

        keyFilePath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\KeyForMacroTool"
        'keyFilePath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\KeyForMyTool"

        fileCreateTime = (CDbl(Format(FileDateTime(fileName), "yyyy")) * 365 + CDbl(Format(FileDateTime(fileName), "MM")) * 30 + CDbl(Format(FileDateTime(fileName), "dd"))) * 250 - 100
        nowTime = (CDbl(Format(Now, "yyyy")) * 365 + CDbl(Format(Now, "MM")) * 30 + CDbl(Format(Now, "dd"))) * 250 - 100

        'fileCreateTime = CDbl(Format(FileDateTime(fileName), "yyyyMMdd")) * 8 - 100
        'nowTime = CDbl(Format(Now, "yyyyMMdd")) * 8 - 100
        usedDays = (nowTime - fileCreateTime) / 250

        LicenseDays = 16
        remainDays = LicenseDays - usedDays
        fso = CreateObject("Scripting.FileSystemObject")

        Try
            If fso.FileExists(keyFilePath) Then
                keyFileCreateTime = (CDbl(Format(FileDateTime(keyFilePath), "yyyy")) * 365 + CDbl(Format(FileDateTime(keyFilePath), "MM")) * 30 + CDbl(Format(FileDateTime(keyFilePath), "dd"))) * 250 - 100
                'keyFileCreateTime = CDbl(Format(FileDateTime(keyFilePath), "yyyyMMdd")) * 8 - 100
                If keyFileCreateTime > nowTime Then
                    MsgBox("系统时间异常！将无法使用部分功能！", MsgBoxStyle.Critical)
                    timeRemainForTool = False
                    Exit Sub
                End If
                fileNo = FreeFile()
                FileOpen(fileNo, keyFilePath, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
                existsKeyStr = CDbl(LineInput(fileNo))
                FileClose(fileNo)
                If existsKeyStr < nowTime Then
                    outPutFile = fso.OpenTextFile(keyFilePath, 2, 1, 0)
                    outPutFile.WriteLine(nowTime)
                    outPutFile.close()
                    If usedDays > LicenseDays Then
                        timeRemainForTool = False
                    ElseIf usedDays <= LicenseDays And usedDays >= 0 Then
                        timeRemainForTool = True
                    End If
                ElseIf existsKeyStr > nowTime Then
                    timeRemainForTool = False
                ElseIf existsKeyStr = nowTime And usedDays <= LicenseDays And usedDays >= 0 Then
                    timeRemainForTool = True
                End If
            Else
                If fileCreateTime <= nowTime Then
                    outPutFile = fso.OpenTextFile(keyFilePath, 2, 1, 0)
                    outPutFile.WriteLine(nowTime)
                    outPutFile.close()
                    If usedDays > LicenseDays Then
                        timeRemainForTool = False
                    ElseIf usedDays <= LicenseDays Then
                        timeRemainForTool = True
                    End If
                Else
                    MsgBox("系统时间异常！将无法使用部分功能！", MsgBoxStyle.Critical)
                    timeRemainForTool = False
                End If
            End If
        Catch ex As Exception
            FileClose(fileNo)
            outPutFile = Nothing
            fso = Nothing
        Finally
            fso = Nothing
        End Try
    End Sub
    Friend Sub selectWorkSheet(fileName As String)
        Dim excel_app As New Excel.Application
        Dim excel_workbook As Excel.Workbook
        excel_workbook = excel_app.Workbooks.Open(fileName)
        WorkSheetsSelection.WorkSheetsGroup.Items.Clear()
        WorkSheetsSelection.Show()
        For x = 1 To excel_workbook.Sheets.Count
            WorkSheetsSelection.WorkSheetsGroup.Items.Add(excel_app.Worksheets(x).Name)
        Next
        excel_workbook.Close()
        excel_workbook = Nothing
        excel_app.Quit()
        excel_app = Nothing
    End Sub
    Friend Function importSiteFileMethod(filePath As String) As ArrayList
        Dim openFileNo As Integer
        Dim tempLine As String
        Dim nowReadLine As Integer
        Dim temp_arr As Object
        Dim siteInformation_arr As ArrayList
        openFileNo = FreeFile()
        siteInformation_arr = New ArrayList
        Try
            FileOpen(openFileNo, filePath, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        Catch ex As Exception
            MsgBox("Site文件打开失败！" & vbCrLf & ex.ToString(), MsgBoxStyle.Critical)
            FileClose(openFileNo)
        End Try
        Try
            nowReadLine = 1
            Do While Not EOF(openFileNo)
                tempLine = LineInput(openFileNo)
                temp_arr = Split(tempLine, "	")
                If nowReadLine = 1 Then
                    If UCase(temp_arr(0)) <> "CELL" Or Strings.Left(UCase(temp_arr(2)), 3) <> "LON" Or Strings.Left(UCase(temp_arr(3)), 3) <> "LAT" Or UCase(temp_arr(4)) <> "DIR" Then
                        MsgBox("Mcom Site文件表头不正确，请重新导入！" & vbCrLf & tempLine, MsgBoxStyle.Critical)
                        nowReadLine = 1
                        Return siteInformation_arr
                        Exit Function
                    Else
                        siteInformation_arr.Add(tempLine)
                    End If
                Else
                    siteInformation_arr.Add(tempLine)
                End If
                nowReadLine = nowReadLine + 1
            Loop
        Catch ex As Exception
            MsgBox("Site文件导入失败！请检查第" & nowReadLine & "行的数据！！", MsgBoxStyle.Critical)
        Finally
            FileClose(openFileNo)
            temp_arr = Nothing
        End Try
        Return siteInformation_arr
    End Function
    Friend Function importCarrierFileMethod(filePath As String)
        Dim openFileNo As Integer
        Dim tempLine As String
        Dim nowReadLine As Integer
        Dim temp_arr As Object
        Dim carrierData As McomCarrierData

        carrierData = New McomCarrierData
        carrierData.cell_arr = New ArrayList
        carrierData.site_arr = New ArrayList
        carrierData.BSC_CELL_BCCH_ARR = New ArrayList
        carrierData.CELL_BCCH_BSIC_DIC = New Dictionary(Of String, String)
        carrierData.CELL_BCCH_DIC = New Dictionary(Of String, String)
        carrierData.CELL_TRX_DIC = New Dictionary(Of String, String)
        '    BSC_SITE_BCCH_DIC = New Dictionary(Of String, String)
        carrierData.SITE_CELL_DIC = New Dictionary(Of String, String)
        carrierData.BSC_CELL_DIC = New Dictionary(Of String, String)
        openFileNo = FreeFile()
        Try
            FileOpen(openFileNo, filePath, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        Catch ex As Exception
            MsgBox("Carrier文件打开失败！" & vbCrLf & ex.ToString(), MsgBoxStyle.Critical)
            FileClose(openFileNo)
        End Try
        Try
            nowReadLine = 1
            Do While Not EOF(openFileNo)
                tempLine = LineInput(openFileNo)
                temp_arr = Split(tempLine, "	")
                If nowReadLine = 1 Then
                    If UCase(temp_arr(0)) <> "CELL" Or UCase(temp_arr(4)) <> "BCCH" Or UCase(temp_arr(5)) <> "BSIC" Or UCase(temp_arr(6)) <> "TCH" Then
                        MsgBox("Mcom Carrier文件不正确，请重新导入！", MsgBoxStyle.Critical)
                        nowReadLine = 1
                        Return carrierData
                        Exit Function
                        'Else
                        '    siteInformation_arr.Add(tempLine)
                    End If
                Else
                    If Not carrierData.cell_arr.Contains(temp_arr(0)) Then
                        carrierData.cell_arr.Add(temp_arr(0))
                    End If
                    If Not carrierData.BSC_CELL_BCCH_ARR.Contains(temp_arr(1) & "#" & temp_arr(0) & "#" & temp_arr(4)) Then
                        carrierData.BSC_CELL_BCCH_ARR.Add(temp_arr(1) & "#" & temp_arr(0) & "#" & temp_arr(4))
                    End If
                    If Not carrierData.site_arr.Contains(Strings.Left(temp_arr(0), Len(temp_arr(0)) - 1)) Then
                        carrierData.site_arr.Add(Strings.Left(temp_arr(0), Len(temp_arr(0)) - 1))
                    End If
                    If carrierData.CELL_BCCH_BSIC_DIC.ContainsKey(temp_arr(0)) Then
                        MsgBox("Carrier表存在同名小区:" & temp_arr(0) & "，请检查！检查结果将保留在：" & carrierData.BSC_CELL_DIC(temp_arr(0)) & "下的小区", MsgBoxStyle.Critical)
                    Else
                        carrierData.CELL_BCCH_BSIC_DIC.Add(temp_arr(0), temp_arr(4) & "#" & temp_arr(5))
                        carrierData.CELL_BCCH_DIC.Add(temp_arr(0), temp_arr(4))
                        carrierData.CELL_TRX_DIC.Add(temp_arr(0), Trim(temp_arr(4) & " " & temp_arr(6)))
                        carrierData.BSC_CELL_DIC.Add(temp_arr(0), temp_arr(1))
                    End If
                    If carrierData.SITE_CELL_DIC.ContainsKey(Strings.Left(temp_arr(0), Len(temp_arr(0)) - 1)) Then
                        'BSC_SITE_BCCH_DIC(temp_arr(1) & "#" & Strings.Left(temp_arr(0), Len(temp_arr(0)) - 1)) = BSC_SITE_BCCH_DIC(temp_arr(1) & "#" & Strings.Left(temp_arr(0), Len(temp_arr(0)) - 1)) & "$" & temp_arr(0) & "@" & temp_arr(4)
                        'BSC_SITE_TRX_DIC(temp_arr(1) & "#" & Strings.Left(temp_arr(0), Len(temp_arr(0)) - 1)) = BSC_SITE_TRX_DIC(temp_arr(1) & "#" & Strings.Left(temp_arr(0), Len(temp_arr(0)) - 1)) & "$" & temp_arr(0) & "@" & Trim(temp_arr(4) & " " & temp_arr(6))
                        carrierData.SITE_CELL_DIC(Strings.Left(temp_arr(0), Len(temp_arr(0)) - 1)) = carrierData.SITE_CELL_DIC(Strings.Left(temp_arr(0), Len(temp_arr(0)) - 1)) & "$" & temp_arr(0)
                    Else
                        'BSC_SITE_BCCH_DIC.Add(temp_arr(1) & "#" & Strings.Left(temp_arr(0), Len(temp_arr(0)) - 1), temp_arr(0) & "@" & temp_arr(4))
                        'BSC_SITE_TRX_DIC.Add(temp_arr(1) & "#" & Strings.Left(temp_arr(0), Len(temp_arr(0)) - 1), temp_arr(0) & "@" & Trim(temp_arr(4) & " " & temp_arr(6)))
                        carrierData.SITE_CELL_DIC.Add(Strings.Left(temp_arr(0), Len(temp_arr(0)) - 1), temp_arr(0))
                    End If
                End If
                nowReadLine = nowReadLine + 1
            Loop
            Return carrierData
        Catch ex As Exception
            MsgBox("Carrier文件导入失败！" & vbCrLf & ex.ToString(), MsgBoxStyle.Critical)
            Return carrierData
        Finally
            FileClose(openFileNo)
            temp_arr = Nothing
        End Try
    End Function
    Friend Function importNeighborFileMethod(filePath As String)
        Dim openFileNo As Integer
        Dim tempLine As String
        Dim nowReadLine As Integer
        Dim temp_arr As Object
        Dim neighborData As McomNeighborData
        neighborData = New McomNeighborData
        neighborData.cell_arr = New ArrayList
        neighborData.ncell_arr = New ArrayList
        openFileNo = FreeFile()
        Try
            FileOpen(openFileNo, filePath, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        Catch ex As Exception
            MsgBox("Neighbor文件打开失败！" & vbCrLf & ex.ToString(), MsgBoxStyle.Critical)
            FileClose(openFileNo)
        End Try
        Try
            nowReadLine = 1
            Do While Not EOF(openFileNo)
                tempLine = LineInput(openFileNo)
                temp_arr = Split(tempLine, "	")
                If nowReadLine = 1 Then
                    If UCase(temp_arr(0)) <> "CELL" Or UCase(temp_arr(1)) <> "NCELL" Then
                        MsgBox("Mcom Neighbor文件不正确，请重新导入！请按如下格式：(CELL	NCELL为标准Mcom Neighbor文件表头)" & vbCrLf & "CELL	NCELL" & vbCrLf & "GA0001B	GA0001A" & vbCrLf & "GA0001B	GA0001C" & vbCrLf & "或者：" & vbCrLf & "CELL	NCELL" & vbCrLf & "GA0001B	GA0001A GA0001C", MsgBoxStyle.Critical)
                        nowReadLine = 1
                        Return neighborData
                        Exit Function
                        'Else
                        '    siteInformation_arr.Add(tempLine)
                    End If
                Else
                    If Not neighborData.cell_arr.Contains(temp_arr(0) & "#" & temp_arr(1)) Then
                        neighborData.cell_arr.Add(temp_arr(0) & "#" & temp_arr(1))
                    Else
                        'neighborData.cell_arr(temp_arr(0) & "#" & temp_arr(1)) = neighborData.cell_arr(temp_arr(0) & "#" & temp_arr(1)) &" "&
                    End If
                End If
                nowReadLine = nowReadLine + 1
            Loop
            Return neighborData
        Catch ex As Exception
            MsgBox("Neighbor文件导入失败！" & vbCrLf & ex.ToString(), MsgBoxStyle.Critical)
            Return neighborData
        Finally
            FileClose(openFileNo)
            temp_arr = Nothing
        End Try
    End Function
   
    'Friend Function loadCarrierFile() As ArrayList
    '    Dim cellCarrierData() As McomCarrierData
    '    '    Dim cellInfo_arr As ArrayList
    '    '   cellCarrierData = New McomCarrierData
    '    '   cellInfo_arr = New ArrayList
    '    cellCarrierData(0).BCCH = "22"
    '    'cellInfo_arr.Add(cellCarrierData)

    '    Return cellCarrierData
    'End Function
    Friend Sub csv2xlsx(filePath As String)
        Dim excel_app As Excel.Application
        Dim excel_workbook As Excel.Workbook
        Dim originalFile_workbook As Excel.Workbook
        Dim fileName As String
        Dim fso As Object
        fso = CreateObject("Scripting.fileSystemObject")
        fileName = Strings.Right(filePath, Len(filePath) - InStrRev(filePath, "\"))
        excel_app = New Excel.Application
        excel_app.DisplayAlerts = False
        excel_workbook = excel_app.Workbooks.Add
        '      excel_workbook.Sheets.Delete()
        Try
            originalFile_workbook = excel_app.Workbooks.Open(filePath) : excel_workbook.Sheets.Add()
            excel_workbook.ActiveSheet.Name = Strings.Left(Split(fileName, ".")(0), Len(Split(fileName, ".")(0)) - 7)
            originalFile_workbook.ActiveSheet.UsedRange.Copy(excel_workbook.ActiveSheet.range("A1"))
            originalFile_workbook.Close(False)
            excel_workbook.SaveAs(Split(filePath, ".")(0) & ".xlsx")
            fso.DeleteFile(filePath)
            excel_app.DisplayAlerts = True
        Catch ex As Exception
            MsgBox("CSV转XLXS文件失败！" & vbCrLf & ex.ToString, MsgBoxStyle.Critical)
            Exit Sub
        Finally
            excel_workbook.Close(False)
            excel_app.Quit()
            fso = Nothing
        End Try

    End Sub

    '''''计算目标小区相对于源小区的角度
    Friend Function caculateRelativeAngle(lon1 As Double, lat1 As Double, lon2 As Double, lat2 As Double)
        Dim dx, dy, angle, angle_end, pi As Double
        pi = 3.141592654
        dx = lon2 - lon1
        dy = lat2 - lat1
        If (lon2 = lon1) Then
            angle = pi / 2
        ElseIf (lat2 = lat1) Then
            angle = 0
        ElseIf (lat2 < lat1) Then
            angle = 3 * pi / 2
        Else
        End If

        If ((lon2 > lon1) And (lat2 > lat1)) Then
            angle = Math.Atan(dx / dy)
        ElseIf ((lon2 > lon1) And (lat2 < lat1)) Then
            angle = pi / 2 + Math.Atan(-dy / dx)
        ElseIf ((lon2 < lon1) And (lat2 < lat1)) Then
            angle = pi + Math.Atan(dx / dy)
        ElseIf ((lon2 < lon1) And (lat2 > lat1)) Then
            angle = 3 * pi / 2 + Math.Atan(dy / -dx)
        Else
        End If

        angle_end = angle * 180 / pi
        caculateRelativeAngle = angle_end
    End Function
    ''''''判定是目标小区是否在源小区的正打范围内
    Friend Function isCoverBySourceCell(S_azimuth As Double, RelativeAzimuth As Double, RangeAzimuth As Double)
        '' RangeAzimuth 为判定源小区的正打的范围(角度)
        If If(Math.Abs(RelativeAzimuth - S_azimuth) < RangeAzimuth / 2, 1, 0) + If((360 + S_azimuth - RelativeAzimuth) < RangeAzimuth / 2, 1, 0) >= 1 Then
            isCoverBySourceCell = True
        Else
            isCoverBySourceCell = False
        End If
    End Function
    ''''''判定源小区与目标小区是否为对打小区
    Friend Function isAgainstCells(S_lon As Double, S_lat As Double, S_azimuth As Double, T_lon As Double, T_lat As Double, T_azimuth As Double, RangeAzimuth As Double)
        '' RangeAzimuth 为判定源小区的正打的范围(角度)
        If isCoverBySourceCell(S_azimuth, caculateRelativeAngle(S_lon, S_lat, T_lon, T_lat), RangeAzimuth) And isCoverBySourceCell(T_azimuth, caculateRelativeAngle(T_lon, T_lat, S_lon, S_lat), RangeAzimuth) Then
            isAgainstCells = 1 '是一对，对打小区
        ElseIf isCoverBySourceCell(S_azimuth, caculateRelativeAngle(S_lon, S_lat, T_lon, T_lat), RangeAzimuth) And Not isCoverBySourceCell(T_azimuth, caculateRelativeAngle(T_lon, T_lat, S_lon, S_lat), RangeAzimuth) Then
            isAgainstCells = 2 '是目标小区位于源小区正向，且共同覆盖同一方向
        ElseIf Not isCoverBySourceCell(S_azimuth, caculateRelativeAngle(S_lon, S_lat, T_lon, T_lat), RangeAzimuth) And isCoverBySourceCell(T_azimuth, caculateRelativeAngle(T_lon, T_lat, S_lon, S_lat), RangeAzimuth) Then
            isAgainstCells = 3 '是目标小区位于源小区背向，且共同覆盖同一方向
        Else
            isAgainstCells = 4 '目前小区为源小区的背向小区
        End If
    End Function

    Friend Sub myWriteLine(fileName As String, content As String)
        Dim fso As Object
        Dim outWrite As Object
        fso = CreateObject("Scripting.FileSystemObject")
        Try
            If fso.fileExists(fileName) Then
                outWrite = fso.openTextFile(fileName, 8, 1, 0)
            Else
                outWrite = fso.openTextFile(fileName, 2, 1, 0)
            End If
        Catch ex As Exception
            MsgBox("请检查输出文件，是否处于打开状态！", MsgBoxStyle.Critical)
            Exit Sub
            outWrite = Nothing
            fso = Nothing
        End Try
        Try
            outWrite.writeLine(content)
        Catch ex As Exception
        Finally
            outWrite.close()
            outWrite = Nothing
            fso = Nothing
        End Try

    End Sub
    Friend Sub myWrite(fileName As String, content As String)
        Dim fso As Object
        Dim outWrite As Object
        fso = CreateObject("Scripting.FileSystemObject")
        Try
            If fso.fileExists(fileName) Then
                outWrite = fso.openTextFile(fileName, 8, 1, 0)
            Else
                outWrite = fso.openTextFile(fileName, 2, 1, 0)
            End If
        Catch ex As Exception
            MsgBox("请检查输出文件，是否处于打开状态！", MsgBoxStyle.Critical)
            Exit Sub
            outWrite = Nothing
            fso = Nothing
        End Try
        Try
            outWrite.write(content)
        Catch ex As Exception
        Finally
            outWrite.close()
            outWrite = Nothing
            fso = Nothing
        End Try

    End Sub
End Module

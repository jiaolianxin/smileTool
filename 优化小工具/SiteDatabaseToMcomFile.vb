Imports Microsoft.Office.Interop
Imports System.ComponentModel
Imports 优化小工具.MainForm

Module SiteDatabaseToMcomFile
    Friend Sub startConvertSiteDbToMcom(fileName As String, sheetName As String, sender As Object, e As DoWorkEventArgs)
        Dim excel_app As New Excel.Application
        Dim excel_workbook As Excel.Workbook
        Dim excel_worksheet As New Excel.Worksheet
        Dim totalLine As Integer, totalColumn As Integer
        Dim enbcell As ArrayList, eNodeBname As ArrayList, cluster As ArrayList, eNodeBID As ArrayList, cellName As ArrayList, cellID As ArrayList, PCIgroup As ArrayList, subCellId As ArrayList
        Dim ENBCELL_pos As Integer, LOGIC_NAME_pos As Integer, ENODEB_NAME_pos As Integer, CLUSTER_pos As Integer, eNB_id_pos As Integer, CELL_NAME_pos As Integer, CELL_ID_pos As Integer, PCIgroup_pos As Integer, subCellId_pos As Integer
        Dim logicName As ArrayList, myCellName As ArrayList, nodeBid_cellID As ArrayList
        Dim long_arr As ArrayList, lat_arr As ArrayList, Dir_arr As ArrayList, Height_arr As ArrayList, Dtile_arr As ArrayList
        Dim long_pos As Integer, lat_pos As Integer, dir_pos As Integer, height_pos As Integer, dtile_pos As Integer
        Dim fso As Object
        Dim outSiteFile As Object
        Dim outCarrierFile As Object
        Dim outFolder As String
        Dim x As Integer, m As Integer, rowsNo As Integer, startRowNo As Integer
        rowsNo = 1
        outFolder = Strings.Left(fileName, InStrRev(fileName, "\"))
        long_arr = New ArrayList
        lat_arr = New ArrayList
        Dir_arr = New ArrayList
        Height_arr = New ArrayList
        Dtile_arr = New ArrayList
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
        totalLine = 0
        totalColumn = 0
        excel_app.Workbooks.Open(fileName)
        excel_workbook = excel_app.ActiveWorkbook
        Try
            excel_worksheet = excel_workbook.Sheets(sheetName)
            totalLine = excel_worksheet.UsedRange.Rows.Count
            totalColumn = excel_worksheet.UsedRange.Columns.Count
            sender.reportProgress(0, "读取Site database数据......")
            For rowno = 1 To 2
                For y = 1 To 45
                    If UCase(Trim(excel_worksheet.Cells(rowno, y).value)) = "ENBCELL" Then
                        startRowNo = rowno
                        ENBCELL_pos = y
                    ElseIf UCase(Trim(excel_worksheet.Cells(rowno, y).value)) = "LOGIC NAME" Then
                        LOGIC_NAME_pos = y
                    ElseIf UCase(Trim(excel_worksheet.Cells(rowno, y).value)) = "ENODEB NAME" Then
                        ENODEB_NAME_pos = y
                    ElseIf UCase(Trim(excel_worksheet.Cells(rowno, y).value)) = "CLUSTER" Then
                        CLUSTER_pos = y
                    ElseIf Trim(excel_worksheet.Cells(rowno, y).value) = "新eNB id(十进制)" Then
                        eNB_id_pos = y
                    ElseIf UCase(Trim(excel_worksheet.Cells(rowno, y).value)) = "CELL NAME" Then
                        CELL_NAME_pos = y
                    ElseIf UCase(Trim(excel_worksheet.Cells(rowno, y).value)) = "CELL ID" Then
                        CELL_ID_pos = y
                    ElseIf InStr(Trim(excel_worksheet.Cells(rowno, y).value), "CellIdGroup") > 0 Then
                        PCIgroup_pos = y
                    ElseIf InStr(Trim(excel_worksheet.Cells(rowno, y).value), "SubCellId") > 0 Then
                        subCellId_pos = y
                    ElseIf InStr(UCase(Trim(excel_worksheet.Cells(rowno, y).value)), "LONGITUDE") > 0 Then
                        long_pos = y
                    ElseIf InStr(UCase(Trim(excel_worksheet.Cells(rowno, y).value)), "LATITUDE") > 0 Then
                        lat_pos = y
                    ElseIf InStr(UCase(Trim(excel_worksheet.Cells(rowno, y).value)), "AZIMUT") > 0 Then
                        dir_pos = y
                    ElseIf InStr(UCase(Trim(excel_worksheet.Cells(rowno, y).value)), "天线高度") > 0 Then
                        height_pos = y
                    ElseIf InStr(UCase(Trim(excel_worksheet.Cells(rowno, y).value)), "OTAL ANTENNATILT") > 0 Then
                        dtile_pos = y
                    End If
                    sender.reportProgress((x * y / (2 * 45)) * 10, "读取Site database数据......" & Mid((x * y / (2 * 45)) * 10, 1, 2) & "%")
                Next
            Next
            For m = 1 To totalLine
                enbcell.Add(excel_worksheet.Cells(m, ENBCELL_pos).value)
                'logicName.Add(excel_worksheet.Cells(m, LOGIC_NAME_pos).value)
                eNodeBname.Add(excel_worksheet.Cells(m, ENODEB_NAME_pos).value)
                cluster.Add(excel_worksheet.Cells(m, CLUSTER_pos).value)
                eNodeBID.Add(excel_worksheet.Cells(m, eNB_id_pos).value)
                cellName.Add(excel_worksheet.Cells(m, CELL_NAME_pos).value)
                cellID.Add(excel_worksheet.Cells(m, CELL_ID_pos).value)
                PCIgroup.Add(excel_worksheet.Cells(m, PCIgroup_pos).value)
                subCellId.Add(excel_worksheet.Cells(m, subCellId_pos).value)
                nodeBid_cellID.Add(excel_worksheet.Cells(m, eNB_id_pos).value & "-" & excel_worksheet.Cells(m, CELL_ID_pos).value)
                long_arr.Add(excel_worksheet.Cells(m, long_pos).value)
                lat_arr.Add(excel_worksheet.Cells(m, lat_pos).value)
                Dir_arr.Add(excel_worksheet.Cells(m, dir_pos).value)
                Height_arr.Add(excel_worksheet.Cells(m, height_pos).value)
                Dtile_arr.Add(excel_worksheet.Cells(m, dtile_pos).value)
                'myCellName.Add(Strings.Split(excel_worksheet.Cells(m, LOGIC_NAME_pos).value, "-")(0) & "_" & Strings.Right(excel_worksheet.Cells(m, CELL_ID_pos).value, 1))
                sender.reportProgress(10 + (m / totalLine) * 90, "读取Site database数据......" & Mid(10 + (m / totalLine) * 90, 1, 2) & "%")
            Next
        Catch ex As Exception
            MsgBox("SiteDataBase文件读取错误！请检查" & m & "行的数据后重试！" & vbCrLf & ex.ToString, MsgBoxStyle.Critical)
            Exit Sub
        Finally
            excel_workbook.Close()
            excel_workbook = Nothing
            excel_app.Quit()
            excel_app = Nothing
        End Try
        sender.reportProgress(100, "读取Site database数据完毕......")

        '''''''''''''''''''''''''''''''''''''''''''制作Mcom site文件
        sender.reportProgress(0, "制作Mcom文件......")
        fso = CreateObject("scripting.FileSystemObject")
        If Not fso.folderExists(outFolder & "Mcom file\") Then
            fso.createFolder(outFolder & "Mcom file\")
        End If
        outSiteFile = fso.openTextFile(outFolder & "Mcom file\Mcom Site_" & Format(FileDateTime(fileName), "yyMMdd") & ".txt", 2, 1, 0)
        outCarrierFile = fso.openTextFile(outFolder & "Mcom file\Mcom Carrier_" & Format(FileDateTime(fileName), "yyMMdd") & ".txt", 2, 1, 0)
        outSiteFile.writeLine("Cell" & "	" & "Site" & "	" & "Longitude" & "	" & "Latitude" & "	" & "Dir" & "	" & "Height" & "	" & "Tilt" & "	" & "Ground_Height" & "	" & "Cell_Type" & "	" & "Ant_Type" & "	" & "Ant_BW" & "	" & "Note" & "	" & "SiteName" & "	" & "Flag1" & "	" & "Flag2" & "	" & "Flag3" & "	" & "Flag4" & "	" & "Flag5" & "	" & "CI" & "	" & "BSC" & "	" & "Ant_size")
        outCarrierFile.writeLine("CELL" & "	" & "BSC" & "	" & "LAI" & "	" & "CI" & "	" & "BCCH" & "	" & "BSIC" & "	" & "TCH" & "	" & "HOP" & "	" & "HSN" & "	" & "FONT_SIZE")
        Try
            For x = startRowNo To nodeBid_cellID.Count - 1
                If InStr(UCase(Height_arr(x)), "M") > 0 Then
                    Height_arr(x) = Replace(UCase(Height_arr(x)), "M", "")
                End If
                outSiteFile.writeLine(nodeBid_cellID(x) & "	" & eNodeBID(x) & "	" & long_arr(x) & "	" & lat_arr(x) & "	" & Dir_arr(x) & "	" & Height_arr(x) & "	" & Dtile_arr(x) & "	" & "" & "	" & "" & "	" & "" & "	" & "" & "	" & "" & "	" & cellName(x) & "	" & PCIgroup(x) & "	" & subCellId(x) & "	" & "" & "	" & "" & "	" & "" & "	" & "" & "	" & cluster(x) & "	" & "5")
                outCarrierFile.writeLine(nodeBid_cellID(x) & "	" & cluster(x) & "	" & "" & "	" & "" & "	" & "37900" & "	" & PCIgroup(x) & "	" & "" & "	" & "" & "	" & "" & "	" & "5")
                sender.reportProgress(x / (nodeBid_cellID.Count - 1) * 100, "制作Mcom文件......" & Mid(x / (nodeBid_cellID.Count - 1) * 100, 1, 2) & "%")
            Next
            sender.reportProgress(100, "制作Mcom文件完成......100%")
        Catch ex As Exception
            MsgBox("Mcom文件制作失败！请检查: " & eNodeBID(x) & "的数据是否有误？", MsgBoxStyle.Critical)
            Exit Sub
        Finally
            outSiteFile.close()
            outCarrierFile.close()
            outSiteFile = Nothing
            outCarrierFile = Nothing
            fso = Nothing
        End Try
        MsgBox("Mcom文件制作完成，耗时：" & MainForm.totalEvaTime & "秒！", MsgBoxStyle.Information)
        Shell("explorer.exe " & outFolder, AppWinStyle.MaximizedFocus)

    End Sub
End Module

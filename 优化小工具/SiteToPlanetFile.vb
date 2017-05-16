Imports System.ComponentModel
Imports Microsoft.Office.Interop

Module SiteToPlanetFile
    Private cell_arr As ArrayList
    Private antenna_id_arr As ArrayList
    Private sector_id_arr As ArrayList
    Private site_arr As ArrayList
    Private long_arr As ArrayList
    Private lat_arr As ArrayList
    Private dir_arr As ArrayList
    Private height_arr As ArrayList
    Private tilt_arr As ArrayList
    Private siteName_arr As ArrayList
    Private SitesSheet_arr As Object
    Private AntennasSheet_arr As Object
    Private SectorsSheet_arr As Object
    Private Sector_AntennasSheet_arr As Object

    Private Sub importSite(sender As Object, e As DoWorkEventArgs, siteFilePath As String)
        Dim openFileNo As Integer
        Dim tempLine As String
        Dim temp_arr As Object
        Dim nowLineNo As Integer = 1
        Dim cell_id As String
        cell_arr = New ArrayList
        antenna_id_arr = New ArrayList
        sector_id_arr = New ArrayList
        site_arr = New ArrayList
        long_arr = New ArrayList
        lat_arr = New ArrayList
        dir_arr = New ArrayList
        height_arr = New ArrayList
        tilt_arr = New ArrayList
        siteName_arr = New ArrayList
        SitesSheet_arr = Nothing

        cell_arr.Clear()
        antenna_id_arr.Clear()
        sector_id_arr.Clear()
        site_arr.Clear()
        long_arr.Clear()
        lat_arr.Clear()
        dir_arr.Clear()
        height_arr.Clear()
        tilt_arr.Clear()
        siteName_arr.Clear()
        ReDim SitesSheet_arr(0 To 65534, 0 To 6)
        ReDim AntennasSheet_arr(0 To 65534, 0 To 24)
        ReDim SectorsSheet_arr(0 To 65534, 0 To 13)
        ReDim Sector_AntennasSheet_arr(0 To 65534, 0 To 7)
        openFileNo = FreeFile()
        Try
            FileOpen(openFileNo, siteFilePath, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
            Do While Not EOF(openFileNo)
                tempLine = LineInput(openFileNo)
                temp_arr = Split(tempLine, "	")
                If nowLineNo = 1 Then
                    If UCase(temp_arr(0)) <> "CELL" Or Strings.Left(UCase(temp_arr(2)), 3) <> "LON" Or Strings.Left(UCase(temp_arr(3)), 3) <> "LAT" Or UCase(temp_arr(4)) <> "DIR" Then
                        MsgBox("Mcom Site文件不正确，请重新导入！", MsgBoxStyle.Critical)
                        nowLineNo = 1
                        Exit Sub
                    Else
                        '''''''''Sites表的数据
                        SitesSheet_arr(nowLineNo - 1, 0) = "Site ID"
                        SitesSheet_arr(nowLineNo - 1, 1) = "Site UID"
                        SitesSheet_arr(nowLineNo - 1, 2) = "Longitude"
                        SitesSheet_arr(nowLineNo - 1, 3) = "Latitude"
                        SitesSheet_arr(nowLineNo - 1, 4) = "Site Name"
                        SitesSheet_arr(nowLineNo - 1, 5) = "Site Name 2"
                        SitesSheet_arr(nowLineNo - 1, 6) = "Candidate Priority"
                        '''''''''Antennas表的数据
                        AntennasSheet_arr(nowLineNo - 1, 0) = "Site ID"
                        AntennasSheet_arr(nowLineNo - 1, 1) = "Antenna ID"
                        AntennasSheet_arr(nowLineNo - 1, 2) = "Longitude"
                        AntennasSheet_arr(nowLineNo - 1, 3) = "Latitude"
                        AntennasSheet_arr(nowLineNo - 1, 4) = "Antenna File"
                        AntennasSheet_arr(nowLineNo - 1, 5) = "Height (m)"
                        AntennasSheet_arr(nowLineNo - 1, 6) = "Azimuth"
                        AntennasSheet_arr(nowLineNo - 1, 7) = "Mechanical Tilt"
                        AntennasSheet_arr(nowLineNo - 1, 8) = "Twist"
                        AntennasSheet_arr(nowLineNo - 1, 9) = "Donor Antenna"
                        AntennasSheet_arr(nowLineNo - 1, 10) = "Terrain Height (m)"
                        AntennasSheet_arr(nowLineNo - 1, 11) = "Sectors"
                        AntennasSheet_arr(nowLineNo - 1, 12) = "Number of Corrections"
                        AntennasSheet_arr(nowLineNo - 1, 13) = "Correction (dB) at -180 degrees"
                        AntennasSheet_arr(nowLineNo - 1, 14) = "Correction (dB) at -150 degrees"
                        AntennasSheet_arr(nowLineNo - 1, 15) = "Correction (dB) at -120 degrees"
                        AntennasSheet_arr(nowLineNo - 1, 16) = "Correction (dB) at -90 degrees"
                        AntennasSheet_arr(nowLineNo - 1, 17) = "Correction (dB) at -60 degrees"
                        AntennasSheet_arr(nowLineNo - 1, 18) = "Correction (dB) at -30 degrees"
                        AntennasSheet_arr(nowLineNo - 1, 19) = "Correction (dB) at 0 degrees"
                        AntennasSheet_arr(nowLineNo - 1, 20) = "Correction (dB) at 30 degrees"
                        AntennasSheet_arr(nowLineNo - 1, 21) = "Correction (dB) at 60 degrees"
                        AntennasSheet_arr(nowLineNo - 1, 22) = "Correction (dB) at 90 degrees"
                        AntennasSheet_arr(nowLineNo - 1, 23) = "Correction (dB) at 120 degrees"
                        AntennasSheet_arr(nowLineNo - 1, 24) = "Correction (dB) at 150 degrees"

                        ''''''Sectors表的数据
                        SectorsSheet_arr(nowLineNo - 1, 0) = "Site ID"
                        SectorsSheet_arr(nowLineNo - 1, 1) = "BTS Name"
                        SectorsSheet_arr(nowLineNo - 1, 2) = "Sector ID"
                        SectorsSheet_arr(nowLineNo - 1, 3) = "Sector UID"
                        SectorsSheet_arr(nowLineNo - 1, 4) = "Cell ID"
                        SectorsSheet_arr(nowLineNo - 1, 5) = "Technology"
                        SectorsSheet_arr(nowLineNo - 1, 6) = "Band Name"
                        SectorsSheet_arr(nowLineNo - 1, 7) = "Antenna Algorithm"
                        SectorsSheet_arr(nowLineNo - 1, 8) = "Propagation Model"
                        SectorsSheet_arr(nowLineNo - 1, 9) = "Distance (km)"
                        SectorsSheet_arr(nowLineNo - 1, 10) = "Radials"
                        SectorsSheet_arr(nowLineNo - 1, 11) = "Prediction Mode"
                        SectorsSheet_arr(nowLineNo - 1, 12) = "Interpolation Distance (m)"
                        SectorsSheet_arr(nowLineNo - 1, 13) = "Display Scheme"

                        ''''''''Sector_Antennas表的数据
                        Sector_AntennasSheet_arr(nowLineNo - 1, 0) = "Site ID"
                        Sector_AntennasSheet_arr(nowLineNo - 1, 1) = "Sector ID"
                        Sector_AntennasSheet_arr(nowLineNo - 1, 2) = "Antenna ID"
                        Sector_AntennasSheet_arr(nowLineNo - 1, 3) = "Horizontal Beamwidth"
                        Sector_AntennasSheet_arr(nowLineNo - 1, 4) = "Vertical Beamwidth"
                        Sector_AntennasSheet_arr(nowLineNo - 1, 5) = "Link Configuration ID"
                        Sector_AntennasSheet_arr(nowLineNo - 1, 6) = "Cable Length (m)"
                        Sector_AntennasSheet_arr(nowLineNo - 1, 7) = "Power Split (%)"

                    End If
                Else

                    cell_arr.Add(temp_arr(0))
                    cell_id = Right(temp_arr(0), 1)
                    antenna_id_arr.Add(CInt(Right(temp_arr(0), 1)))
                    sector_id_arr.Add("'" & CStr(Right(temp_arr(0), 1)))
                    site_arr.Add(temp_arr(1))
                    lat_arr.Add(temp_arr(3))
                    long_arr.Add(temp_arr(2))
                    dir_arr.Add(temp_arr(4))
                    height_arr.Add(temp_arr(5))
                    tilt_arr.Add(temp_arr(6))
                    siteName_arr.Add(temp_arr(12))
                    '''''''''Sites表的数据
                    SitesSheet_arr(nowLineNo - 1, 0) = CStr(site_arr(nowLineNo - 2))
                    SitesSheet_arr(nowLineNo - 1, 1) = ""
                    SitesSheet_arr(nowLineNo - 1, 2) = long_arr(nowLineNo - 2)
                    SitesSheet_arr(nowLineNo - 1, 3) = lat_arr(nowLineNo - 2)
                    SitesSheet_arr(nowLineNo - 1, 4) = CStr(siteName_arr(nowLineNo - 2))
                    SitesSheet_arr(nowLineNo - 1, 5) = ""
                    SitesSheet_arr(nowLineNo - 1, 6) = CInt(1)
                    '''''''''Antennas表的数据
                    AntennasSheet_arr(nowLineNo - 1, 0) = CStr(site_arr(nowLineNo - 2))
                    AntennasSheet_arr(nowLineNo - 1, 1) = CInt(antenna_id_arr(nowLineNo - 2))
                    AntennasSheet_arr(nowLineNo - 1, 2) = long_arr(nowLineNo - 2)
                    AntennasSheet_arr(nowLineNo - 1, 3) = lat_arr(nowLineNo - 2)
                    AntennasSheet_arr(nowLineNo - 1, 4) = ""
                    AntennasSheet_arr(nowLineNo - 1, 5) = height_arr(nowLineNo - 2)
                    AntennasSheet_arr(nowLineNo - 1, 6) = dir_arr(nowLineNo - 2)
                    AntennasSheet_arr(nowLineNo - 1, 7) = tilt_arr(nowLineNo - 2)
                    AntennasSheet_arr(nowLineNo - 1, 8) = 0
                    AntennasSheet_arr(nowLineNo - 1, 9) = CStr("FALSE")
                    AntennasSheet_arr(nowLineNo - 1, 10) = ""
                    AntennasSheet_arr(nowLineNo - 1, 11) = CStr(sector_id_arr(nowLineNo - 2))
                    AntennasSheet_arr(nowLineNo - 1, 12) = 12
                    AntennasSheet_arr(nowLineNo - 1, 13) = 0
                    AntennasSheet_arr(nowLineNo - 1, 14) = 0
                    AntennasSheet_arr(nowLineNo - 1, 15) = 0
                    AntennasSheet_arr(nowLineNo - 1, 16) = 0
                    AntennasSheet_arr(nowLineNo - 1, 17) = 0
                    AntennasSheet_arr(nowLineNo - 1, 18) = 0
                    AntennasSheet_arr(nowLineNo - 1, 19) = 0
                    AntennasSheet_arr(nowLineNo - 1, 20) = 0
                    AntennasSheet_arr(nowLineNo - 1, 21) = 0
                    AntennasSheet_arr(nowLineNo - 1, 22) = 0
                    AntennasSheet_arr(nowLineNo - 1, 23) = 0
                    AntennasSheet_arr(nowLineNo - 1, 24) = 0

                    ''''''Sectors表的数据
                    SectorsSheet_arr(nowLineNo - 1, 0) = CStr(site_arr(nowLineNo - 2))
                    SectorsSheet_arr(nowLineNo - 1, 1) = "LTE TDD"
                    SectorsSheet_arr(nowLineNo - 1, 2) = CStr(sector_id_arr(nowLineNo - 2))
                    SectorsSheet_arr(nowLineNo - 1, 3) = ""
                    SectorsSheet_arr(nowLineNo - 1, 4) = CStr(sector_id_arr(nowLineNo - 2))
                    SectorsSheet_arr(nowLineNo - 1, 5) = "LTE TDD"
                    SectorsSheet_arr(nowLineNo - 1, 6) = "LteTddBand_D"
                    SectorsSheet_arr(nowLineNo - 1, 7) = "N/A"
                    SectorsSheet_arr(nowLineNo - 1, 8) = ""
                    SectorsSheet_arr(nowLineNo - 1, 9) = 2
                    SectorsSheet_arr(nowLineNo - 1, 10) = 360
                    SectorsSheet_arr(nowLineNo - 1, 11) = "Modeled"
                    SectorsSheet_arr(nowLineNo - 1, 12) = 200
                    SectorsSheet_arr(nowLineNo - 1, 13) = "N/A"

                    ''''''''Sector_Antennas表的数据
                    Sector_AntennasSheet_arr(nowLineNo - 1, 0) = CStr(site_arr(nowLineNo - 2))
                    Sector_AntennasSheet_arr(nowLineNo - 1, 1) = CStr(sector_id_arr(nowLineNo - 2))
                    Sector_AntennasSheet_arr(nowLineNo - 1, 2) = CInt(antenna_id_arr(nowLineNo - 2))
                    Sector_AntennasSheet_arr(nowLineNo - 1, 3) = 67.75289
                    Sector_AntennasSheet_arr(nowLineNo - 1, 4) = 5.41106367
                    Sector_AntennasSheet_arr(nowLineNo - 1, 5) = "Default"
                    Sector_AntennasSheet_arr(nowLineNo - 1, 6) = 4
                    Sector_AntennasSheet_arr(nowLineNo - 1, 7) = 100


                End If
                nowLineNo = nowLineNo + 1
            Loop
        Catch ex As Exception
            MsgBox("Mcom site 数据在第" & nowLineNo - 1 & "行有误，请检查！" & vbCrLf & "暂不支持小区最后一位为‘字母’", MsgBoxStyle.Critical)
            Exit Sub
        Finally
            nowLineNo = 1
            FileClose(openFileNo)
        End Try
    End Sub
    Friend Sub startConvertSiteToPlanetFile(sender As Object, e As DoWorkEventArgs, fileName As String)
        Dim excel_app As Excel.Application
        Dim excel_workbook As Excel.Workbook
        Dim outFolder As String
        Dim outFileName As String
        Dim totalRowsNo As Integer

        sender.reportProgress(0, "正在导入Site文件...")
        importSite(sender, e, fileName)
        sender.reportProgress(100, "导入Site文件完成...")
        sender.reportProgress(0, "正在生成Planet sites文件...")
        excel_app = New Excel.Application
        '  excel_app.Visible = True
        outFolder = Left(fileName, InStrRev(fileName, "\"))
        outFileName = outFolder & "Planet sites_" & Format(FileDateTime(fileName), "MMdd") & ".xlsx"
        excel_workbook = excel_app.Workbooks.Add
        excel_workbook.Sheets.Add()
        excel_workbook.Sheets(1).name = "Sites"
        excel_workbook.Sheets(2).name = "Antennas"
        excel_workbook.Sheets(3).name = "Sectors"
        excel_workbook.Sheets(4).name = "Sector_Antennas"

        excel_workbook.Sheets("Sites").RANGE("A1:G" & cell_arr.Count + 1) = SitesSheet_arr
        excel_workbook.Sheets("Sites").Range("A1:G" & cell_arr.Count + 1).RemoveDuplicates(Columns:=1, Header:=1)
        excel_workbook.Sheets("Antennas").RANGE("A1:Y" & cell_arr.Count + 1) = AntennasSheet_arr
        excel_workbook.Sheets("Sectors").RANGE("A1:N" & cell_arr.Count + 1) = SectorsSheet_arr
        excel_workbook.Sheets("Sector_Antennas").RANGE("A1:H" & cell_arr.Count + 1) = Sector_AntennasSheet_arr
        totalRowsNo = excel_workbook.Sheets("Antennas").UsedRange.Rows.Count
        sender.reportProgress(70, "正在生成Planet sites文件...")

        For i = 2 To totalRowsNo
            If excel_workbook.Sheets("Antennas").cells(i, 7).value = 360 Then
                excel_workbook.Sheets("Antennas").cells(i, 7).Font.Color = -16776961
            End If
            sender.reportProgress(70 + i / totalRowsNo * 20, "正在生成Planet sites文件...")
        Next
        excel_app.DisplayAlerts = False
        excel_workbook.SaveAs(outFileName)
        excel_workbook.Close()
        excel_app.DisplayAlerts = True
        excel_app.Quit()
        sender.reportProgress(100, "生成Planet sites文件成功...")

        MsgBox("转换完成，文件保存路径为：" & outFileName, MsgBoxStyle.Information)
        Shell("explorer.exe " & outFileName, 1)
    End Sub
End Module

Public Class ConsistencyCheck
    Private carrierData As McomCarrierData
    Private neighborData As McomNeighborData

    Protected Friend Sub loadCarrierData(fileName As String)
        carrierData = New McomCarrierData
        carrierData = importCarrierFileMethod(fileName)
    End Sub
    Protected Friend Sub loadNeighborData(fileName As String)
        neighborData = New McomNeighborData
        neighborData = importNeighborFileMethod(fileName)
    End Sub
    Protected Friend Sub coSiteFreqCheck(outFolder As String)
        Dim fso As Object
        Dim outFile As Object
        Dim outFileName As String
        Dim nowLine As Integer
        Dim x As Integer
        Dim y As Integer
        Dim m As Integer
        Dim n As Integer
        Dim site_temp As String
        Dim cell_temp1 As String
        Dim cell_temp2 As String
        Dim freq_temp1 As String
        Dim freq_temp2 As String
        Dim freq_arr_temp1 As Object
        Dim freq_arr_temp2 As Object
        site_temp = ""
        cell_temp1 = ""
        cell_temp2 = ""
        fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.folderExists(outFolder) Then
            fso.createFolder(outFolder)
        End If
        outFileName = outFolder & "CoSiteFreqCheck_" & Format(Now, "yyMMdd") & ".csv"
        outFile = fso.openTextFile(outFileName, 2, 1, 0)
        outFile.writeLine("BSC" & "," & "CELL" & "," & "CellOfCoSite" & "," & "Freq1" & "," & "Freq2" & "," & "FreqSep")
        Try
            For nowLine = 0 To carrierData.site_arr.Count - 1
                site_temp = carrierData.site_arr(nowLine)
                If carrierData.SITE_CELL_DIC.ContainsKey(site_temp) Then
                    For x = 0 To UBound(Split(carrierData.SITE_CELL_DIC(site_temp), "$"))
                        cell_temp1 = Split(carrierData.SITE_CELL_DIC(site_temp), "$")(x)
                        freq_temp1 = Trim(carrierData.CELL_TRX_DIC(cell_temp1))
                        If freq_temp1 <> "" Then
                            For y = x + 1 To UBound(Split(carrierData.SITE_CELL_DIC(site_temp), "$"))
                                cell_temp2 = Split(carrierData.SITE_CELL_DIC(site_temp), "$")(y)
                                freq_temp2 = Trim(carrierData.CELL_TRX_DIC(cell_temp2))
                                If freq_temp2 <> "" Then
                                    freq_arr_temp1 = Split(freq_temp1, " ")
                                    freq_arr_temp2 = Split(freq_temp2, " ")
                                    For m = 0 To UBound(freq_arr_temp1)
                                        For n = 0 To UBound(freq_arr_temp2)
                                            If InStr(freq_arr_temp1(m), "]") = 0 And InStr(freq_arr_temp2(n), "]") = 0 Then
                                                If Math.Abs(freq_arr_temp1(m) - freq_arr_temp2(n)) <= 1 Then
                                                    outFile.writeLine(carrierData.BSC_CELL_DIC(cell_temp1) & "," & cell_temp1 & "," & cell_temp2 & "," & freq_arr_temp1(m) & "," & freq_arr_temp2(n) & "," & Math.Abs(freq_arr_temp1(m) - freq_arr_temp2(n)))
                                                End If
                                            End If
                                        Next
                                    Next
                                End If
                            Next
                        End If
                    Next
                End If
            Next
            MsgBox("站内同邻频检查完毕！", MsgBoxStyle.Information)
            Shell("explorer.exe " & outFolder, 1)
        Catch ex As Exception
            MsgBox("站内同邻频检查失败！请检查" & Split(cell_temp1, "#")(1) & "," & Split(cell_temp2, "#")(1) & "的频点信息！", MsgBoxStyle.Critical)
            Exit Sub
        Finally
            outFile.close()
            fso = Nothing
        End Try
    End Sub
    Protected Friend Sub neighborCellCoBCCHCheck(outFolder As String)
        Dim fso As Object
        Dim outFile As Object
        Dim outFile1 As Object
        Dim outFile2 As Object
        '        Dim outFile3 As Object
        Dim outFileName As String   'Neighbor CoCarrier
        Dim outFileName1 As String  'Neighbor CoBCCH
        Dim outFileName2 As String ' Neighbor'count and Co or Adj count
        '       Dim outFileName3 As String '邻小区同BCCH同BISC
        Dim nowLine As Integer
        Dim x As Integer
        'Dim y As Integer
        Dim m As Integer
        Dim n As Integer
        Dim cell_temp1 As String
        Dim cell_temp2 As String
        Dim freq_temp1 As String
        Dim freq_temp2 As String
        Dim freq_arr_temp1 As Object
        Dim freq_arr_temp2 As Object
        Dim ncell_arr_temp As Object
        Dim coMark As Boolean  '是否存在同频
        Dim adjMark As Boolean  '是否存在邻频
        Dim coNo As Integer   '同频个数
        Dim adjNo As Integer   '邻频个数
        Dim lastCell As String  '上一个源小区
        Dim neighborNo As Integer  '只有一个邻区或竖向邻区检查时的邻区个数
        Dim coAndAdjCount As ArrayList
        coAndAdjCount = New ArrayList
        cell_temp1 = ""
        cell_temp2 = ""
        ncell_arr_temp = ""
        neighborNo = 0
        lastCell = ""
        fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.folderExists(outFolder) Then
            fso.createFolder(outFolder)
        End If
        outFileName = outFolder & "NeigCellCoCarrierFreqCheck_" & Format(Now, "yyMMdd") & ".csv"
        outFileName1 = outFolder & "NeigCellCoBCCHFreqCheck_" & Format(Now, "yyMMdd") & ".csv"
        outFileName2 = outFolder & "NeigCellCoAndAdjCount_" & Format(Now, "yyMMdd") & ".csv"
        '      outFileName3 = outFolder & "CoBCCHCoBSICOfNeighbors_" & Format(Now, "yyMMdd") & ".csv"
        outFile = fso.openTextFile(outFileName, 2, 1, 0)
        outFile1 = fso.openTextFile(outFileName1, 2, 1, 0)
        outFile2 = fso.openTextFile(outFileName2, 2, 1, 0)
        '    outFile3 = fso.openTextFile(outFileName3, 2, 1, 0)
        outFile.writeLine("S_BSC" & "," & "CELL" & "," & "N_BSC" & "," & "NCELL" & "," & "Freq1" & "," & "Freq2" & "," & "FreqSep")
        outFile1.writeLine("S_BSC" & "," & "CELL" & "," & "N_BSC" & "," & "NCELL" & "," & "BCCH1" & "," & "BCCH2" & "," & "FreqSep")
        outFile2.writeLine("S_BSC" & "," & "CELL" & "," & "NoOfNeighborCell" & "," & "NoOfCoNeighbor" & "," & "NoOfAdjNeighbor")
        '      outFile3.writeLine("S_BSC" & "," & "CELL" & "," & "Ncell1_BSC" & "," & "NCELL1" & "," & "Ncell2_BSC" & "," & "NCELL2" & "," & "BCCH" & "," & "BSIC")

        Try
            For nowLine = 0 To neighborData.cell_arr.Count - 1
                cell_temp1 = ""
                cell_temp2 = ""
                freq_temp1 = ""
                freq_temp2 = ""
                cell_temp1 = Split(Trim(neighborData.cell_arr(nowLine)), "#")(0)
                cell_temp2 = Split(Trim(neighborData.cell_arr(nowLine)), "#")(1)
                If carrierData.CELL_TRX_DIC.ContainsKey(cell_temp1) Then
                    freq_temp1 = Trim(carrierData.CELL_TRX_DIC(cell_temp1))
                    If InStr(cell_temp2, " ") > 0 Then
                        coNo = 0
                        adjNo = 0
                        If freq_temp1 <> "" Then
                            freq_arr_temp1 = Split(freq_temp1, " ")
                            If cell_temp2 <> "" Then
                                ncell_arr_temp = Split(cell_temp2, " ")
                                For x = 0 To UBound(ncell_arr_temp)
                                    coMark = False
                                    adjMark = False
                                    If carrierData.CELL_TRX_DIC.ContainsKey(ncell_arr_temp(x)) Then
                                        freq_temp2 = Trim(carrierData.CELL_TRX_DIC(ncell_arr_temp(x)))
                                        If freq_temp2 <> "" Then
                                            freq_arr_temp2 = Split(freq_temp2, " ")
                                            For m = 0 To UBound(freq_arr_temp1)
                                                For n = 0 To UBound(freq_arr_temp2)
                                                    If InStr(freq_arr_temp1(m), "]") = 0 And InStr(freq_arr_temp2(n), "]") = 0 Then
                                                        If Math.Abs(freq_arr_temp1(m) - freq_arr_temp2(n)) <= 1 Then
                                                            If Math.Abs(freq_arr_temp1(m) - freq_arr_temp2(n)) = 0 Then
                                                                coMark = True
                                                            Else
                                                                adjMark = True
                                                            End If
                                                            outFile.writeLine(carrierData.BSC_CELL_DIC(cell_temp1) & "," & cell_temp1 & "," & carrierData.BSC_CELL_DIC(ncell_arr_temp(x)) & "," & ncell_arr_temp(x) & "," & freq_arr_temp1(m) & "," & freq_arr_temp2(n) & "," & Math.Abs(freq_arr_temp1(m) - freq_arr_temp2(n)))
                                                            If carrierData.BSC_CELL_BCCH_ARR.Contains(carrierData.BSC_CELL_DIC(cell_temp1) & "#" & cell_temp1 & "#" & freq_arr_temp1(m)) And carrierData.BSC_CELL_BCCH_ARR.Contains(carrierData.BSC_CELL_DIC(ncell_arr_temp(x)) & "#" & ncell_arr_temp(x) & "#" & freq_arr_temp2(n)) Then
                                                                outFile1.writeLine(carrierData.BSC_CELL_DIC(cell_temp1) & "," & cell_temp1 & "," & carrierData.BSC_CELL_DIC(ncell_arr_temp(x)) & "," & ncell_arr_temp(x) & "," & freq_arr_temp1(m) & "," & freq_arr_temp2(n) & "," & Math.Abs(freq_arr_temp1(m) - freq_arr_temp2(n)))
                                                            End If
                                                        End If
                                                    End If
                                                Next
                                            Next
                                        End If
                                    End If
                                    If coMark Then
                                        coNo = coNo + 1
                                    End If
                                    If adjMark Then
                                        adjNo = adjNo + 1
                                    End If
                                Next
                                coAndAdjCount.Add(carrierData.BSC_CELL_DIC(cell_temp1) & "," & cell_temp1 & "," & UBound(ncell_arr_temp) + 1 & "," & coNo & "," & adjNo)
                                coNo = 0
                                adjNo = 0
                            End If
                        End If
                    Else
                        If lastCell = cell_temp1 Then
                            neighborNo = neighborNo + 1
                        End If
                        If carrierData.CELL_TRX_DIC.ContainsKey(cell_temp2) Then
                            coNo = 0
                            adjNo = 0
                            freq_temp2 = Trim(carrierData.CELL_TRX_DIC(cell_temp2))
                            If freq_temp1 <> "" And freq_temp2 <> "" Then
                                freq_arr_temp1 = Split(freq_temp1, " ")
                                freq_arr_temp2 = Split(freq_temp2, " ")
                                For m = 0 To UBound(freq_arr_temp1)
                                    For n = 0 To UBound(freq_arr_temp2)
                                        If InStr(freq_arr_temp1(m), "]") = 0 And InStr(freq_arr_temp2(n), "]") = 0 Then
                                            If Math.Abs(freq_arr_temp1(m) - freq_arr_temp2(n)) <= 1 Then
                                                If Math.Abs(freq_arr_temp1(m) - freq_arr_temp2(n)) = 0 Then
                                                    coMark = True
                                                Else
                                                    adjMark = True
                                                End If
                                                outFile.writeLine(carrierData.BSC_CELL_DIC(cell_temp1) & "," & cell_temp1 & "," & carrierData.BSC_CELL_DIC(cell_temp2) & "," & cell_temp2 & "," & freq_arr_temp1(m) & "," & freq_arr_temp2(n) & "," & Math.Abs(freq_arr_temp1(m) - freq_arr_temp2(n)))
                                                If carrierData.BSC_CELL_BCCH_ARR.Contains(carrierData.BSC_CELL_DIC(cell_temp1) & "#" & cell_temp1 & "#" & freq_arr_temp1(m)) And carrierData.BSC_CELL_BCCH_ARR.Contains(carrierData.BSC_CELL_DIC(cell_temp2) & "#" & cell_temp2 & "#" & freq_arr_temp2(m)) Then
                                                    outFile1.writeLine(carrierData.BSC_CELL_DIC(cell_temp1) & "," & cell_temp1 & "," & carrierData.BSC_CELL_DIC(cell_temp2) & "," & cell_temp2 & "," & freq_arr_temp1(m) & "," & freq_arr_temp2(n) & "," & Math.Abs(freq_arr_temp1(m) - freq_arr_temp2(n)))
                                                End If
                                            End If
                                        End If
                                    Next
                                Next
                            End If
                        End If
                        If coMark And lastCell = cell_temp1 Then
                            coNo = coNo + 1
                        ElseIf coMark And lastCell <> cell_temp1 And coNo >= 1 And lastCell <> "" Then
                            coNo = coNo - 1
                        Else
                            coNo = coNo + 1
                        End If
                        If adjMark And lastCell = cell_temp1 Then
                            adjNo = adjNo + 1
                        ElseIf adjMark And lastCell <> cell_temp1 And adjNo >= 1 And lastCell <> "" Then
                            adjNo = adjNo - 1
                        Else
                            adjNo = adjNo + 1
                        End If
                        If lastCell <> cell_temp1 And lastCell <> "" Then
                            coAndAdjCount.Add(carrierData.BSC_CELL_DIC(lastCell) & "," & lastCell & "," & neighborNo & "," & coNo & "," & adjNo)
                            lastCell = cell_temp1
                            coNo = 1
                            adjNo = 1
                            neighborNo = 1
                        End If
                    End If
                End If
            Next
            For i = 0 To coAndAdjCount.Count - 1
                outFile2.writeLine(coAndAdjCount(i))
            Next
            MsgBox("邻区同邻频检查完毕！", MsgBoxStyle.Information)
            Shell("explorer.exe " & outFolder, 1)
        Catch ex As Exception
            MsgBox("邻区同邻频检查失败！请检查" & cell_temp1 & "," & ncell_arr_temp(x) & "的频点信息！", MsgBoxStyle.Critical)
            Exit Sub
        Finally
            outFile.close()
            outFile1.CLOSE()
            outFile2.CLOSE()
            'outFile3.CLOSE()
            fso = Nothing
        End Try
    End Sub
End Class
Imports System.ComponentModel
Imports System.Threading

Module FormatCDDlog_backgroundworker
    Private statusText As String
    Private filenames, outPutPath, totalTime
    Private BSCname As String, fso As Object

    Private Sub ChangeStatus(status As String)
        MainForm.statusBar.Text = status
        MainForm.StatusStrip1.Refresh()
    End Sub
    Friend Function getFiles()
        MsgBox("请选择需要修改的CDDlog文件！")
        MainForm.OpenFileDialog1.Title = "选择需要调整格式的CDDlog"
        MainForm.OpenFileDialog1.Filter = "CDDlog文件(*.log)|*.log|CDD文本文件(*.txt)|*.txt"
        MainForm.OpenFileDialog1.Multiselect = True
        MainForm.OpenFileDialog1.ShowHelp = True
        MainForm.OpenFileDialog1.RestoreDirectory = True
        MainForm.OpenFileDialog1.ShowDialog()
        filenames = MainForm.OpenFileDialog1.FileNames
        MainForm.OpenFileDialog1.Reset()
        Return filenames
        '       Return filenames
    End Function
    Friend Sub doWrite(filePath As String)
        Dim nowLine As Integer
        Dim nowOpenFileNo As Integer
        Dim outputFile As Object
        Dim temp As String, temp1 As String
        nowOpenFileNo = FreeFile()
        FileOpen(nowOpenFileNo, Split(filePath, "@")(0), OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        outputFile = fso.OpenTextFile(Split(filePath, "@")(1), 2, 1, 0)
        '  nowWriteFileNo = FreeFile()
        '    FileOpen(nowWriteFileNo, Split(filePath, "@")(1), OpenMode.Output, OpenAccess.Write, OpenShare.LockWrite)
        nowLine = 0
        Do While Not EOF(nowOpenFileNo)
            nowLine = nowLine + 1
            temp1 = LineInput(nowOpenFileNo)
            If nowLine = 1 Then
                temp = "*** Connected to " & UCase(BSCname) & " ***"
                'WriteLine(nowWriteFileNo, Replace(temp, "\""", ""))
                'WriteLine(nowWriteFileNo, temp1)
                outputFile.writeline(temp)
                outputFile.writeline(temp1)
            Else
                If InStr(temp1, "<") > 0 Then
                    '     WriteLine(nowWriteFileNo, vbCrLf & "END" & vbCrLf & vbCrLf & "<" & Split(temp1, "<")(1))
                    outputFile.writeline(vbCrLf & "END" & vbCrLf & vbCrLf & "<" & Split(temp1, "<")(1))
                Else
                    ' WriteLine(nowWriteFileNo, temp1)
                    outputFile.writeline(temp1)
                End If
            End If
        Loop
        'WriteLine(nowWriteFileNo, "END")
        'WriteLine(nowWriteFileNo, "*** Disconnected from " & UCase(BSCname) & " ***")
        outputFile.writeline("END")
        outputFile.writeline("*** Disconnected from " & UCase(BSCname) & " ***")
        outputFile.Close()
        FileClose(nowOpenFileNo)
        '  FileClose(nowWriteFileNo)
    End Sub
    Friend Sub doFormatWork(sender As Object, e As DoWorkEventArgs)
        Dim myPath, outputFolder, logname
        Dim inputFile
        Dim nowProcessPercentage As Integer
        Dim myThread As Thread

        nowProcessPercentage = 0
        If filenames.length = 0 Then
            Exit Sub
        Else
            sender.reportProgress(0, "数据准备...")
            fso = CreateObject("Scripting.FileSystemObject")
            For nowFileNo = 0 To UBound(filenames)
                myPath = Strings.Left(filenames(nowFileNo), InStrRev(filenames(nowFileNo), "\"))
                BSCname = UCase(Split(Split(filenames(nowFileNo), "\")(UBound(Split(filenames(nowFileNo), "\"))), ".")(0))
                myPath = myPath & "newCDD_" & Format(Now, "MMdd")
                If fso.folderExists(myPath) Then
                    outPutPath = myPath & "\"
                Else
                    outputFolder = fso.createFolder(myPath)
                    outPutPath = myPath & "\"
                End If
                logname = UCase(Split(filenames(nowFileNo), "\")(UBound(Split(filenames(nowFileNo), "\"))))
                sender.ReportProgress((nowFileNo + 1) / (UBound(filenames) + 1) * 100, "正在格式化：" & logname)

                '                sender.ReportProgress(nowProcessPercentage, "正在格式化：" & logname)
                '  outputFile = fso.OpenTextFile(outPutPath & logname, 2, 1, 0)
                myThread = New Thread(AddressOf doWrite)
                myThread.Start(filenames(nowFileNo) & "@" & outPutPath & logname)
                myThread.Join()

                'doWrite(filenames(nowFileNo))
                sender.ReportProgress((nowFileNo + 1) / (UBound(filenames) + 1) * 100, "正在格式化：" & logname)
                '               sender.ReportProgress(nowProcessPercentage, "正在格式化：" & logname)
            Next
            sender.ReportProgress(100, "完成！")
            '  outputFile = Nothing
            inputFile = Nothing
            fso = Nothing
            filenames = Nothing
        End If
        MsgBox("CDDlog调整完毕！" & MainForm.totalEvaTime)
        MainForm.totalEvaTime = 0
    End Sub
End Module

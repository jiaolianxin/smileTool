Public Class CDDCommand

    Private Sub Start_button_Click(sender As Object, e As EventArgs) Handles start_button.Click
        Dim savePath As String
        Dim defalutName As String

        If Not CellCommand.Checked And Not CellCommand.Checked And Not MSCCommand.Checked And Not NDCommand.Checked And Not Cell_MOCommand.Checked Then
            MsgBox("请先选择所要生成的CDD指令！")
            Exit Sub
        End If
        MsgBox("请选择保存CDD指令生成文件的文件夹!")
        CommandSave.Description = "请选择保存CDD指令生成文件的文件夹："
        CommandSave.ShowDialog()
        If Now.Month < 10 And Now.Day < 10 Then
            defalutName = "0" & Now.Month & "0" & Now.Day
        ElseIf Now.Month < 10 And Now.Day >= 10 Then
            defalutName = "0" & Now.Month & Now.Day
        ElseIf Now.Month >= 10 And Now.Day < 10 Then
            defalutName = Now.Month & "0" & Now.Day
        Else
            defalutName = Now.Month & Now.Day
        End If

        savePath = CommandSave.SelectedPath
        If savePath = "" Then
            Exit Sub
        Else
            savePath = savePath & "\"
            If CellCommand.Checked Then
                CreateCMDCommand.cReateCellCommand(savePath & "Cell_BSC_Command_" & defalutName & ".txt")
            End If
            If MOCommand.Checked Then
                CreateCMDCommand.cReateMOCommand(savePath & "MOCommand_" & defalutName & ".txt")
            End If
            If MSCCommand.Checked Then
                CreateCMDCommand.cReateMSCCommand(savePath & "MSCCommand_" & defalutName & ".txt")
            End If
            If NDCommand.Checked Then
                CreateCMDCommand.cReateNDCommand(savePath & "NDCommand_" & defalutName & ".txt")
            End If
            If Cell_MOCommand.Checked Then
                CreateCMDCommand.cReateCell_MOCommand(savePath & "Cell_MOCommand_" & defalutName & ".txt")
            End If

            If Not CellCommand.Checked And Not CellCommand.Checked And Not MSCCommand.Checked And Not NDCommand.Checked And Not Cell_MOCommand.Checked Then
                MsgBox("请先选择所要生成的CDD指令！")
            Else
                Shell("explorer.exe " & savePath, 1)
                Close()
            End If
        End If
    End Sub

End Class
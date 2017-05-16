Imports Microsoft.VisualBasic.Compatibility
Imports System.ComponentModel

Module VBAMacroCrack
    Friend Sub crack97_2003VBAmacro(fileName As String, sender As Object, e As DoWorkEventArgs)
        Dim GetData As New VB6.FixedLengthString(5)
        FileOpen(1, fileName, OpenMode.Binary, OpenAccess.ReadWrite, OpenShare.Shared)
        Dim CMGs As Long
        Dim DPBo As Long
        sender.reportProgress(0, "破解中")
        For i = 1 To LOF(1)
            FileGet(1, GetData.Value, i, 0)
            If GetData.Value = "CMG=""" Then
                CMGs = i
            End If
            If GetData.Value = "[Host" Then
                DPBo = i - 2
                Exit For
            End If
            If i > 1000 And i Mod 1000 = 0 Then
                sender.reportProgress(i / LOF(1) * 100, "破解中")
            End If
        Next

        If CMGs = 0 Then
            MsgBox("请先对VBA编码设置一个保护密码...", MsgBoxStyle.Critical, "提示")
            FileClose(1)
            Exit Sub
        End If

        Dim St As New VB6.FixedLengthString(2)
        Dim s20 As New VB6.FixedLengthString(1)

        '取得一个0D0A十六进制字串
        FileGet(1, St.Value, CMGs - 2)
        '取得一个20十六制字串
        FileGet(1, s20.Value, DPBo + 16)
        '替换加密部份机码
        For i = CMGs To DPBo Step 2
            FilePut(1, St.Value, i)
        Next

        '加入不配对符号
        If (DPBo - CMGs) Mod 2 <> 0 Then
            FilePut(1, s20.Value, DPBo + 1)
        End If
        sender.reportProgress(100, "破解成功！")
        MsgBox("文件解密成功......", MsgBoxStyle.Information, "提示")
        FileClose(1)
        Shell("explorer.exe " & fileName, 1)
    End Sub
End Module

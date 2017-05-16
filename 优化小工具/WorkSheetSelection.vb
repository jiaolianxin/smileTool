Public Class WorkSheetsSelection
    Private Sub StartImport_Click(sender As Object, e As EventArgs) Handles StartImport.Click
        If MainForm.isSiteDatabaseToMcom Then
            '  优化小工具.SiteDatabase转Mcom文件.PerformClick()
            MainForm.BGWSiteDBtoMcom.RunWorkerAsync(WorkSheetsGroup.Text)
        End If
        Me.Close()
    End Sub
End Class
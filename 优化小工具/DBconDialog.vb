Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Data.SqlClient


Public Class DBconDialog

    Private Sub StartCon_Click(sender As Object, e As EventArgs) Handles StartCon.Click
        test()
    End Sub
    Private Sub test()
        Dim dbConStr As String
        Dim dbCon As New OleDbConnection
        Dim sqlStr As String
        Dim cmd As New OleDbCommand

        dbConStr = "Provider=Microsoft.ACE.OLEDB.12.0;Jet OLEDB:Database Password=;Data Source=d:\test.accdb;"
        'sqlStr = "drop table test"
        sqlStr = "CREATE TABLE test33 (LogDate text(8),BSC text(10),CELL text(10))"
        Try
            dbCon = New OleDbConnection(dbConStr)
            dbCon.Open()
            cmd = New OleDb.OleDbCommand(sqlStr, dbCon)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "提示")
        Finally
            cmd.Dispose()
            dbCon.Close()
            dbCon.Dispose()
        End Try
    End Sub
End Class
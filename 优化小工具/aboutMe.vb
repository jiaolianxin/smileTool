Imports System.Diagnostics.Process
Public Class aboutMe

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Close()
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles EmailLinkLabel.LinkClicked
        Dim Process = New Process()
        Process.StartInfo.FileName = "mailto:" & EmailLinkLabel.Text
        Process.Start()
    End Sub

    Private Sub aboutMe_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LabelVersion.Text = String.Format("版本： {0}", My.Application.Info.Version.ToString)
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click

    End Sub
End Class
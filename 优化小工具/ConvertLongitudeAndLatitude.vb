Public Class ConvertLongitudeAndLatitude


    Private Sub longDu_ValueChanged(sender As Object, e As EventArgs) Handles longDu.ValueChanged
        Dim newLong As String
        Dim newLat As String
        If longDu.Value = 180 Then
            longFen.Value = 0
            longMiao.Value = 0
            longFen.ReadOnly = True
            longMiao.ReadOnly = True
        Else
            longFen.ReadOnly = False
            longMiao.ReadOnly = False
        End If
        newLong = longDu.Value + longFen.Value / 60 + longMiao.Value / 3600
        newLat = latDu.Value + latFen.Value / 60 + latMiao.Value / 3600
        If InStr(newLong, ".") > 0 Then
            If Len(Split(newLong, ".")(1)) < 5 Then
                newLong = newLong & "00000"
            End If
            If Mid(newLong, InStr(newLong, ".") + 6, 1) >= 5 Then
                newLong = Strings.Left(newLong, InStr(newLong, ".") + 5) + 0.00001
            Else
                newLong = Strings.Left(newLong, InStr(newLong, ".") + 5)
            End If
            newLongitude.Text = newLong
        Else
            newLongitude.Text = newLong & ".00000"
        End If
        If InStr(newLat, ".") > 0 Then
            If Len(Split(newLat, ".")(1)) < 5 Then
                newLat = newLat & "00000"
            End If
            If Mid(newLat, InStr(newLat, ".") + 6, 1) >= 5 Then
                newLat = Strings.Left(newLat, InStr(newLat, ".") + 5) + 0.00001
            Else
                newLat = Strings.Left(newLat, InStr(newLat, ".") + 5)
            End If
            newLatitude.Text = newLat
        Else
            newLatitude.Text = newLat & ".00000"
        End If
    End Sub
    Private Sub longFen_ValueChanged(sender As Object, e As EventArgs) Handles longFen.ValueChanged
        Dim newLong As String
        Dim newLat As String
        If longDu.Value = 180 Then
            longFen.Value = 0
            longMiao.Value = 0
            longFen.ReadOnly = True
            longMiao.ReadOnly = True
        Else
            longFen.ReadOnly = False
            longMiao.ReadOnly = False
        End If
        If longFen.Value = 60 Then
            longFen.Value = 0
            longDu.Value = longDu.Value + 1
        End If
        newLong = longDu.Value + longFen.Value / 60 + longMiao.Value / 3600
        newLat = latDu.Value + latFen.Value / 60 + latMiao.Value / 3600
        If InStr(newLong, ".") > 0 Then
            If Len(Split(newLong, ".")(1)) < 5 Then
                newLong = newLong & "00000"
            End If
            If Mid(newLong, InStr(newLong, ".") + 6, 1) >= 5 Then
                newLong = Strings.Left(newLong, InStr(newLong, ".") + 5) + 0.00001
            Else
                newLong = Strings.Left(newLong, InStr(newLong, ".") + 5)
            End If
            newLongitude.Text = newLong
        Else
            newLongitude.Text = newLong & ".00000"
        End If
        If InStr(newLat, ".") > 0 Then
            If Len(Split(newLat, ".")(1)) < 5 Then
                newLat = newLat & "00000"
            End If
            If Mid(newLat, InStr(newLat, ".") + 6, 1) >= 5 Then
                newLat = Strings.Left(newLat, InStr(newLat, ".") + 5) + 0.00001
            Else
                newLat = Strings.Left(newLat, InStr(newLat, ".") + 5)
            End If
            newLatitude.Text = newLat
        Else
            newLatitude.Text = newLat & ".00000"
        End If
    End Sub
    Private Sub longMiao_ValueChanged(sender As Object, e As EventArgs) Handles longMiao.ValueChanged
        Dim newLong As String
        Dim newLat As String
        If longDu.Value = 180 Then
            longFen.Value = 0
            longMiao.Value = 0
            longFen.ReadOnly = True
            longMiao.ReadOnly = True
        Else
            longFen.ReadOnly = False
            longMiao.ReadOnly = False
        End If
        If longMiao.Value = 60 Then
            longMiao.Value = 0
            longFen.Value = longFen.Value + 1
        End If
        newLong = longDu.Value + longFen.Value / 60 + longMiao.Value / 3600
        newLat = latDu.Value + latFen.Value / 60 + latMiao.Value / 3600
        If InStr(newLong, ".") > 0 Then
            If Len(Split(newLong, ".")(1)) < 5 Then
                newLong = newLong & "00000"
            End If
            If Mid(newLong, InStr(newLong, ".") + 6, 1) >= 5 Then
                newLong = Strings.Left(newLong, InStr(newLong, ".") + 5) + 0.00001
            Else
                newLong = Strings.Left(newLong, InStr(newLong, ".") + 5)
            End If
            newLongitude.Text = newLong
        Else
            newLongitude.Text = newLong & ".00000"
        End If
        If InStr(newLat, ".") > 0 Then
            If Len(Split(newLat, ".")(1)) < 5 Then
                newLat = newLat & "00000"
            End If
            If Mid(newLat, InStr(newLat, ".") + 6, 1) >= 5 Then
                newLat = Strings.Left(newLat, InStr(newLat, ".") + 5) + 0.00001
            Else
                newLat = Strings.Left(newLat, InStr(newLat, ".") + 5)
            End If
            newLatitude.Text = newLat
        Else
            newLatitude.Text = newLat & ".00000"
        End If
    End Sub
    Private Sub latDu_ValueChanged(sender As Object, e As EventArgs) Handles latDu.ValueChanged
        Dim newLong As String
        Dim newLat As String
        If latDu.Value = 90 Then
            latFen.Value = 0
            latMiao.Value = 0
            latFen.ReadOnly = True
            latMiao.ReadOnly = True
        Else
            latFen.ReadOnly = False
            latMiao.ReadOnly = False
        End If
        newLong = longDu.Value + longFen.Value / 60 + longMiao.Value / 3600
        newLat = latDu.Value + latFen.Value / 60 + latMiao.Value / 3600
        If InStr(newLong, ".") > 0 Then
            If Len(Split(newLong, ".")(1)) < 5 Then
                newLong = newLong & "00000"
            End If
            If Mid(newLong, InStr(newLong, ".") + 6, 1) >= 5 Then
                newLong = Strings.Left(newLong, InStr(newLong, ".") + 5) + 0.00001
            Else
                newLong = Strings.Left(newLong, InStr(newLong, ".") + 5)
            End If
            newLongitude.Text = newLong
        Else
            newLongitude.Text = newLong & ".00000"
        End If
        If InStr(newLat, ".") > 0 Then
            If Len(Split(newLat, ".")(1)) < 5 Then
                newLat = newLat & "00000"
            End If
            If Mid(newLat, InStr(newLat, ".") + 6, 1) >= 5 Then
                newLat = Strings.Left(newLat, InStr(newLat, ".") + 5) + 0.00001
            Else
                newLat = Strings.Left(newLat, InStr(newLat, ".") + 5)
            End If
            newLatitude.Text = newLat
        Else
            newLatitude.Text = newLat & ".00000"
        End If
    End Sub
    Private Sub latFen_ValueChanged(sender As Object, e As EventArgs) Handles latFen.ValueChanged
        Dim newLong As String
        Dim newLat As String
        If latDu.Value = 90 Then
            latFen.Value = 0
            latMiao.Value = 0
            latFen.ReadOnly = True
            latMiao.ReadOnly = True
        Else
            latFen.ReadOnly = False
            latMiao.ReadOnly = False
        End If
        If latFen.Value = 60 Then
            latFen.Value = 0
            latDu.Value = latDu.Value + 1
        End If
        newLong = longDu.Value + longFen.Value / 60 + longMiao.Value / 3600
        newLat = latDu.Value + latFen.Value / 60 + latMiao.Value / 3600
        If InStr(newLong, ".") > 0 Then
            If Len(Split(newLong, ".")(1)) < 5 Then
                newLong = newLong & "00000"
            End If
            If Mid(newLong, InStr(newLong, ".") + 6, 1) >= 5 Then
                newLong = Strings.Left(newLong, InStr(newLong, ".") + 5) + 0.00001
            Else
                newLong = Strings.Left(newLong, InStr(newLong, ".") + 5)
            End If
            newLongitude.Text = newLong
        Else
            newLongitude.Text = newLong & ".00000"
        End If
        If InStr(newLat, ".") > 0 Then
            If Len(Split(newLat, ".")(1)) < 5 Then
                newLat = newLat & "00000"
            End If
            If Mid(newLat, InStr(newLat, ".") + 6, 1) >= 5 Then
                newLat = Strings.Left(newLat, InStr(newLat, ".") + 5) + 0.00001
            Else
                newLat = Strings.Left(newLat, InStr(newLat, ".") + 5)
            End If
            newLatitude.Text = newLat
        Else
            newLatitude.Text = newLat & ".00000"
        End If

    End Sub
    Private Sub latMiao_ValueChanged(sender As Object, e As EventArgs) Handles latMiao.ValueChanged
        Dim newLong As String
        Dim newLat As String
        If latDu.Value = 90 Then
            latFen.Value = 0
            latMiao.Value = 0
            latFen.ReadOnly = True
            latMiao.ReadOnly = True
        Else
            latFen.ReadOnly = False
            latMiao.ReadOnly = False
        End If
        If latMiao.Value = 60 Then
            latMiao.Value = 0
            latFen.Value = latFen.Value + 1
        End If
        newLong = longDu.Value + longFen.Value / 60 + longMiao.Value / 3600
        newLat = latDu.Value + latFen.Value / 60 + latMiao.Value / 3600
        If InStr(newLong, ".") > 0 Then
            If Len(Split(newLong, ".")(1)) < 5 Then
                newLong = newLong & "00000"
            End If
            If Mid(newLong, InStr(newLong, ".") + 6, 1) >= 5 Then
                newLong = Strings.Left(newLong, InStr(newLong, ".") + 5) + 0.00001
            Else
                newLong = Strings.Left(newLong, InStr(newLong, ".") + 5)
            End If
            newLongitude.Text = newLong
        Else
            newLongitude.Text = newLong & ".00000"
        End If
        If InStr(newLat, ".") > 0 Then
            If Len(Split(newLat, ".")(1)) < 5 Then
                newLat = newLat & "00000"
            End If
            If Mid(newLat, InStr(newLat, ".") + 6, 1) >= 5 Then
                newLat = Strings.Left(newLat, InStr(newLat, ".") + 5) + 0.00001
            Else
                newLat = Strings.Left(newLat, InStr(newLat, ".") + 5)
            End If
            newLatitude.Text = newLat
        Else
            newLatitude.Text = newLat & ".00000"
        End If
    End Sub


End Class
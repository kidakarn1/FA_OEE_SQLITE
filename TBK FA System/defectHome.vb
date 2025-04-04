Public Class defectHome
    Public Shared dtType As String = "NO DATA"
    Private Sub btn_back_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Working_Pro.ResetRed()
        Working_Pro.Enabled = True
        Me.Hide()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs)
        'dtType = "NC"
        'Dim dfSelecttype = New defectSelecttype()
        'dfSelecttype.show()
        'Me.Close()
        MsgBox("ไม่สามารถ ใส่ NC ได้ เนื่องจาก เมนูปิดแล้ว")
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        dtType = "NG"
        Dim dfSelecttype1 = New defectSelecttype()
        dfSelecttype1.show()
        Me.Close()
    End Sub
    Private Sub btnAdjustnc_Click(sender As Object, e As EventArgs)
        'dtType = "NC"
        'Dim dfDetailnc = New defectDetailnc()
        'dfDetailnc.show()
        'Me.Close()
        MsgBox("ไม่สามารถ ปรับ NC ได้ เนื่องจาก เมนูปิดแล้ว")
    End Sub
    Private Sub btnAdjustng_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        dtType = "NG"
        Dim dfDetailng = New defectDetailng()
        dfDetailng.show()
        Me.Close()
    End Sub

    Private Sub defectHome_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Working_Pro.TowerLamp(8, 0)
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)

    End Sub
End Class

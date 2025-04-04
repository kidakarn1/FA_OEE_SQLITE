Public Class TestForm
    Private Sub TestForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WebView21.Source = New Uri("http://192.168.161.219/productionHrProgress/?line_cd=K1A003&shift=P")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim dfHome As New defectHome()
        dfHome.Show()
        Me.Enabled = False
    End Sub
End Class
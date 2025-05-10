Public Class PopUpLossDelayStart
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Working_Pro.Start_Production()
        Me.Close()
    End Sub
End Class
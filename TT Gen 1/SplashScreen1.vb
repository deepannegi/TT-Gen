Public NotInheritable Class SplashScreen1

    Private Sub SplashScreen1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Timer2.Start()
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        ProgressBar1.PerformStep()
        If ProgressBar1.Value = 100 Then
            Timer2.Stop()
            Me.Hide()
            frmMain.Show()
        End If
    End Sub

End Class

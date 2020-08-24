Public Class logo

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'Timer1.Enabled = True
        'Timer2.Enabled = True
        ProgressBar1.Increment(1)
        If ProgressBar1.Value = 100 Then
            login.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        ' Me.Opacity = Me.Opacity - 0.1
    End Sub

    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        ' If Me.Opacity = 0 Then
        'login.Show()
        'End If
    End Sub

    Private Sub logo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
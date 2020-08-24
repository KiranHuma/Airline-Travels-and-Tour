Public Class editbnk

    Private Sub editbnk_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If (DateTime.Now.Hour < 12) Then
            Button1.Visible = False
            Button2.Visible = False
            Button3.Visible = True
            Button4.Visible = False
            lblgrtng.Text = "Good Morning"
            'Label12.Text = Convert.ToString(DateTime.Now)

        ElseIf (DateTime.Now.Hour < 17) Then
            Button1.Visible = False
            Button2.Visible = True
            Button4.Visible = False
            Button3.Visible = False
            lblgrtng.Text = "Good Afternoon"
            'Label12.Text = Convert.ToString(DateTime.Now)
        ElseIf (DateTime.Now.Hour < 19) Then
            Button1.Visible = False
            Button2.Visible = False
            Button3.Visible = False
            Button4.Visible = True
            lblgrtng.Text = "Good Evening"
            'Label12.Text = Convert.ToString(DateTime.Now)
        Else
            Button1.Visible = True
            Button2.Visible = False
            Button3.Visible = False
            Button4.Visible = False
            lblgrtng.Visible = False
            lblgrtng.Text = "Good Night"
            'Label12.Text = Convert.ToString(DateTime.Now)
        End If
        Dim t As Timer = New Timer()
        t.Interval = 5000
        AddHandler t.Tick, AddressOf HandleTimerTick
        t.Start()
    End Sub

    Private Sub HandleTimerTick()
        Me.Close()
    End Sub


End Class
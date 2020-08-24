Public Class tprft
    Dim con As New SqlClient.SqlConnection                      'for sql
    Dim cmd As New SqlClient.SqlCommand

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Profit.ShowDialog()
        Catch ex As Exception
            MsgBox("Form not loaded sucessfully!!!! Try again")
        End Try
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Me.Dispose()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            frmuRcrds.ShowDialog()
        Catch ex As Exception
            MsgBox("Form not loaded sucessfully!!!! Try again")
        End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Try
            frmGrndrecord.ShowDialog()
        Catch ex As Exception
            MsgBox("Form not loaded sucessfully!!!!profit must be in numbers")
        End Try
    End Sub


    Private Sub tprft_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        
        'Acces DataBase Connectivity and for MS Access 2003 PROVIDER=Microsoft.Jet.OLEDB.4.0
        Try
            con.ConnectionString = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
            cmd.Connection = con
            'MessageBox.Show(con.State.ToString())
        Catch ex As Exception
            MsgBox("DataBase not connected due to the reason because " & ex.Message)
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Profitvissa.ShowDialog()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Profitvisit.ShowDialog()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Pexpensis.ShowDialog()
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        Try
            Profit.ShowDialog()
        Catch ex As Exception
            MsgBox("Form not loaded sucessfully!!!! Try again")
        End Try
    End Sub

    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3.Click
        Profitvissa.ShowDialog()
    End Sub

    Private Sub PictureBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox4.Click
        Profitvisit.ShowDialog()
    End Sub

    Private Sub PictureBox7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox7.Click
        Pexpensis.ShowDialog()
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        frmuRcrds.ShowDialog()
    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub
End Class
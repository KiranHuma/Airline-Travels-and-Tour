Imports System.Data.SqlClient
Public Class userreg
    Dim con As New SqlClient.SqlConnection                      'for sql
    Dim cmd As New SqlClient.SqlCommand
    Private Sub dbaccessconnection()
        'Acces DataBase Connectivity and for MS Access 2003 PROVIDER=Microsoft.Jet.OLEDB.4.0
        Try
            con.ConnectionString = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
            cmd.Connection = con
            'MessageBox.Show(con.State.ToString())
        Catch ex As Exception
            MsgBox("DataBase not connected due to the reason because " & ex.Message)
        End Try
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
        login.Show()
    End Sub

    Private Sub userreg_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("Male")
        ComboBox1.Items.Add("Female")
        ComboBox1.Text = "select"
    End Sub
    Private Sub insert()
        dbaccessconnection()
        con.Open()
        cmd.CommandText = "insert into auth(login,password,name,age,sex,contactno)values('" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox6.Text & "','" & TextBox7.Text & "','" & ComboBox1.Text & "','" & TextBox9.Text & "')"
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Or TextBox9.Text = "" Or ComboBox1.Text = "" Then
            MessageBox.Show("Please complete the required fields..", "Authentication Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If Label2.Text = "Match" Then
            Try
                insert()
                MsgBox(" Account created successfully!!", MsgBoxStyle.Information, "Record Added = ")
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox5.Text = ""
                TextBox6.Text = ""
                TextBox7.Text = ""
                ComboBox1.Text = ""
                TextBox9.Text = ""
                Label2.Text = ""

            Catch ex As Exception
                MsgBox(" Failed to creat account!!", MsgBoxStyle.Critical)
            End Try
        Else
            MessageBox.Show("Password not match... Try again", "Authentication Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox4.Text = ""
            TextBox5.Text = ""
            Label2.Text = ""
        End If
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        If TextBox4.Text = TextBox5.Text Then
            Label2.Text = "Match"
            Label2.ForeColor = Color.Green

        Else
            Label2.Text = "Not match"
            Label2.ForeColor = Color.Red
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        ComboBox1.Text = ""
        TextBox9.Text = ""
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()
    End Sub

    Private Sub RectangleShape5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RectangleShape5.Click

    End Sub
End Class
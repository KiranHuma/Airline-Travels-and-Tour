Imports System.Data.SqlClient
Public Class login
    Dim con As New SqlClient.SqlConnection
    Dim cmd As New SqlClient.SqlCommand
    Private Sub dbaccessconnection()
        Try
            con.ConnectionString = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
            cmd.Connection = con
            'MessageBox.Show(con.State.ToString())
        Catch ex As Exception
            MsgBox("DataBase not connected due to the reason because " & ex.Message)
        End Try
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ComboBox1.Text = "USER" Then
            Dim username As String
            username = TextBox1.Text
            Dim pwd As String
            pwd = TextBox2.Text
            Dim conn As New SqlConnection
            conn.ConnectionString = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
            conn.Open()
            Dim cm As New SqlCommand
            cm.CommandText = "SELECT * FROM auth where login = '" & username & "' And password = '" & pwd & "'"
            cm.Connection = conn
            Dim dr As SqlDataReader
            dr = cm.ExecuteReader
            If dr.HasRows Then
                MessageBox.Show("succsessfully login", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                choose.Button10.Enabled = False
                choose.Button8.Enabled = False
                choose.Button2.Enabled = False
                choose.Button6.Enabled = False
                choose.Button7.Enabled = False
                choose.ShowDialog()
                TextBox1.Text = ""
                TextBox2.Text = ""

                Me.Hide()
                dr.Close()
            Else
                Beep()
                MessageBox.Show("Your username Or password is not match", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Label6.ForeColor = Color.Red
                Label6.Text = " Not succsessfully login "
                Label6.Visible = True
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox1.Focus()
            End If
            conn.Close()
        ElseIf ComboBox1.Text = "ADMIN" Then
            If TextBox1.Text = "adnan " Or TextBox2.Text = "123" Then
                Beep()
                Beep()
                MessageBox.Show("You are successfully logged", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                choose.ShowDialog()
                TextBox1.Text = ""
                TextBox2.Text = ""
                Me.Hide()

            Else
                Beep()
                MessageBox.Show("Your username Or password is not match", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox1.Text = ""
                TextBox2.Text = ""
                Label6.ForeColor = Color.Red
                Label6.Text = " Not succsessfully login "
                Label6.Visible = True
                TextBox1.Focus()

            End If

        Else
            MessageBox.Show("Select your choice", "ADMIN or USER", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ComboBox1.Items.Add("USER")
        ComboBox1.Items.Add("ADMIN")
        ComboBox1.Text = "SELECT"

    End Sub




    '/////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub login()
        'not call anywhere but it is good and running sussessfully 

        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim rd As SqlDataReader

        con.ConnectionString = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
        cmd.Connection = con
        con.Open()
        cmd.CommandText = "Select login,password from auth where login= '" & TextBox1.Text & "' And password ='" & TextBox2.Text & "'"
        rd = cmd.ExecuteReader
        If rd.HasRows Then
            choose.Show()
        Else
            MsgBox(" invalid password or username")

        End If
    End Sub
    Private Sub insert()
        dbaccessconnection()
        con.Open()
        cmd.CommandText = "insert into auth(login,password)values('" & TextBox1.Text & "','" & TextBox2.Text & "')"
        cmd.ExecuteNonQuery()
        con.Close()
        'not call anywhere but it is good and running sussessfully 
    End Sub

    '////////////////////////////////////old login code////////////////////////////////
    'Private Function isauthenticated(ByVal username As String, ByVal password As String) As Boolean
    ' Dim isusernamecrrect As Boolean = False
    ' If username = "username" Then
    '  isusernamecrrect = True
    ' End If
    ' Dim ispasswordcorrect As Boolean = False
    ' If password = "password" Then
    '  ispasswordcorrect = True
    ' End If
    'if username and password are both wrong
    ' If isusernamepasswordincorrect(isusernamecrrect, ispasswordcorrect) Then
    '  Return False
    ' End If
    'if username is wrong and password is correct
    ' If isusernameincorrect(isusernamecrrect) Then
    ' Return False
    ' End If
    'if password is wrong and username is correct 
    ' If ispasswordincorrect(ispasswordcorrect) Then
    ' Return False
    ' End If
    ' Return True

    'End Function
    ' Private Function isusernamepasswordincorrect(ByVal isusernamecrrect As Boolean, ByVal ispasswordcorrect As Boolean) As Boolean
    ' If isusernamecrrect = False And ispasswordcorrect = False Then
    ' MessageBox.Show("both user name and password are wrong ", "Error message", MessageBoxButtons.OK, MessageBoxIcon.Error)

    ' Return True
    ' End If
    ' Return False
    ' End Function
    ' Private Function isusernameincorrect(ByVal isusernamecrrect As Boolean) As Boolean
    'If isusernamecrrect = False Then
    ' MessageBox.Show("Sorry username is wrong  ", "Error message", MessageBoxButtons.OK, MessageBoxIcon.Error)

    ' Return True
    ' End If
    ' Return False
    ' End Function
    ' Private Function ispasswordincorrect(ByVal ispasswordcorrect As Boolean) As Boolean
    'If ispasswordcorrect = False Then
    ' MessageBox.Show("Sorry Password is wrong  ", "Error message", MessageBoxButtons.OK, MessageBoxIcon.Error)

    ' Return True

    'End If
    ' Return False
    ' End Function

    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
    ' Dim us As String = TextBox1.Text
    'Dim ps As String = TextBox2.Text
    'If isauthenticated(us, ps) Then
    ' Me.Hide()

    ' Dim form As New choose
    ' choose.Show()
    'End If
    ' End Sub

    ' Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
    ' choose.Button8.Visible = False
    ' choose.Button10.Visible = False
    'choose.Show()
    ' choose.PictureBox9.Visible = False
    ' choose.PictureBox8.Visible = False
    'Ticket.txtcommissin.Visible = False
    ' End Sub
    ' Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
    '  Me.Close()
    'End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

   

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        userreg.Show()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        End

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged

        If CheckBox1.Checked Then
            TextBox2.PasswordChar = ControlChars.NullChar
            ' TextBox2.PasswordChar = ""
        Else
            TextBox2.PasswordChar = "*"
        End If
    End Sub

    
End Class
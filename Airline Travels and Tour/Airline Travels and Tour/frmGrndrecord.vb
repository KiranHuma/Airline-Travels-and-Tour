Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.Data.DataTable
Imports System.Data.SqlClient
Public Class frmGrndrecord
    Dim rdr As SqlDataReader
    Dim provider As String
    Dim dataFile As String
    Dim connString As String
    Dim myConnection As SqlConnection = New SqlConnection
    Dim ds As DataSet = New DataSet         
    Dim da As SqlDataAdapter
    Dim tables As DataTableCollection = ds.Tables
    Dim source1 As New BindingSource()
    Dim source2 As New BindingSource()
    Dim con As New SqlClient.SqlConnection
    Dim cmd As New SqlClient.SqlCommand

    Dim dt As New DataTable
    Dim cs As String = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
    Private Sub dbaccessconnection()
        'Acces DataBase Connectivity and for MS Access 2003 PROVIDER=Microsoft.Jet.OLEDB.4.0
        Try
            con.ConnectionString = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
            cmd.Connection = con
            'MessageBox.Show(con.State.ToString())
        Catch ex As Exception
            MsgBox("DataBase not connected due to the reason because " & ex.Message)
            Me.Dispose()
        End Try
    End Sub
    Private Sub frmGrndrecord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Tsum()
        Usum()
        visum()
        Vsum()
        Esalrysum()
        checkprft()
    End Sub
    Private Sub Tsum()
        'search data through textboxs
        Try
            dbaccessconnection()
            con.Open()
            Dim num As New Integer
            cmd.CommandText = "SELECT SUM(comissin) from ticket "
            If (IsDBNull(cmd.ExecuteScalar)) Then
                num = 0
                Label1.Text = num.ToString
            Else
                'num = cmd.ExecuteScalar()
                'txtticket.Text = num + 1
                num = cmd.ExecuteScalar
                Label1.Text = num.ToString
            End If
            con.Close()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
            Me.Dispose()
        End Try
    End Sub
    Private Sub Vsum()
        'search data through textboxs
        Try
            dbaccessconnection()
            con.Open()
            Dim num As New Integer
            cmd.CommandText = "SELECT SUM(vprft) from vissa "
            If (IsDBNull(cmd.ExecuteScalar)) Then
                num = 0
                Label9.Text = num.ToString
            Else
                'num = cmd.ExecuteScalar()
                'txtticket.Text = num + 1
                num = cmd.ExecuteScalar
                Label9.Text = num.ToString
            End If
            con.Close()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
            Me.Dispose()
        End Try
    End Sub
    Private Sub Esalrysum()
        'search data through textboxs
        Try
            dbaccessconnection()
            con.Open()
            Dim num As New Integer
            cmd.CommandText = "SELECT SUM(salry+bill+pckgecash+prsnluse+rent) from expensis"
            If (IsDBNull(cmd.ExecuteScalar)) Then
                num = 0
                Label10.Text = num.ToString
            Else
                'num = cmd.ExecuteScalar()
                'txtticket.Text = num + 1
                num = cmd.ExecuteScalar
                Label10.Text = num.ToString
            End If
            con.Close()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
            Me.Dispose()
        End Try
    End Sub
    Private Sub Usum()
        'search data through textboxs
        Try
            dbaccessconnection()
            con.Open()
            Dim num As New Integer
            cmd.CommandText = "SELECT Sum(uprft) FROM umrahh "
            If (IsDBNull(cmd.ExecuteScalar)) Then
                num = 0
                Label2.Text = num.ToString
            Else
                'num = cmd.ExecuteScalar()
                'txtticket.Text = num + 1
                num = cmd.ExecuteScalar
                Label2.Text = num.ToString
            End If
            con.Close()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
            Me.Dispose()
        End Try
    End Sub
    Private Sub visum()
        'search data through textboxs
        Try
            dbaccessconnection()
            con.Open()
            Dim num As New Integer
            cmd.CommandText = "SELECT Sum(proft) FROM visit"
            If (IsDBNull(cmd.ExecuteScalar)) Then
                num = 0
                Label7.Text = num.ToString
            Else
                'num = cmd.ExecuteScalar()
                'txtticket.Text = num + 1
                num = cmd.ExecuteScalar
                Label7.Text = num.ToString
            End If
            con.Close()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
            Me.Dispose()
        End Try
    End Sub
    Private Sub fill()
        dbaccessconnection()
        con.Open()
        cmd.CommandText = "SELECT SUM(comissin) from ticket='" & Label1.Text & "'"
        Label1.Text = cmd.ExecuteScalar
        'cmd.CommandText = "select Fdte from fee1 where fCnic=" & Label2.Text & ""
        'Label2.Text = cmd.ExecuteScalar
        ' cmd.CommandText = "select fmth from fee1 where fCnic=" & Label7.Text & ""
        ' Label7.Text = cmd.ExecuteScalar
        ' cmd.CommandText = "select fyr from fee1 where fCnic=" & Label9.Text & ""
        ' Label9.Text = cmd.ExecuteScalar
        ' cmd.CommandText = "select fduefee from fee1 where fCnic=" & Label10.Text & ""
        ' Label10.Text = cmd.ExecuteScalar
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
    Private Sub checkprft()
        Dim x As Integer
        Dim y As Integer
        Dim z As Integer
        Dim xx As Integer
        Dim yy As Integer
        x = Label1.Text
        y = Label2.Text
        z = Label7.Text
        xx = Label9.Text
        yy = Label10.Text
        If x > 0 Then
            'MsgBox("You have profit which is =" & Label3.Text)
            Label4.Text = " &Profit from Tickets"
            ' Label12.Text = Label3.Text
            Label4.BackColor = System.Drawing.Color.Green
        Else
            If x < 0 Then
                'MsgBox("You have loss which is " & Label3.Text)
                Label4.Text = "&Loss from Tickets "
                'Label12.Text = Label3.Text
                Label4.BackColor = System.Drawing.Color.Red
            End If
        End If
        If y > 0 Then
            'MsgBox("You have profit which is =" & Label3.Text)
            Label5.Text = " &Profit from Umrahs"
            ' Label12.Text = Label3.Text
            Label5.BackColor = System.Drawing.Color.Green
        Else
            If y < 0 Then
                'MsgBox("You have loss which is " & Label3.Text)
                Label5.Text = "&Loss from Umrahs"
                'Label12.Text = Label3.Text
                Label5.BackColor = System.Drawing.Color.Red
            End If
        End If
        If z > 0 Then
            'MsgBox("You have profit which is =" & Label3.Text)
            Label6.Text = " &Profit from Visits"
            ' Label12.Text = Label3.Text
            Label6.BackColor = System.Drawing.Color.Green
        Else
            If z < 0 Then
                'MsgBox("You have loss which is " & Label3.Text)
                Label6.Text = "&Loss from Visits"
                'Label12.Text = Label3.Text
                Label6.BackColor = System.Drawing.Color.Red
            End If
        End If
        If xx > 0 Then
            'MsgBox("You have profit which is =" & Label3.Text)
            Label8.Text = " &Profit from Visa"
            ' Label12.Text = Label3.Text
            Label8.BackColor = System.Drawing.Color.Green
        Else
            If xx < 0 Then
                'MsgBox("You have loss which is " & Label3.Text)
                Label8.Text = "&Loss from Visa"
                'Label12.Text = Label3.Text
                Label8.BackColor = System.Drawing.Color.Red
            End If
        End If
        If yy > 0 Then
            'MsgBox("You have profit which is =" & Label3.Text)
            Label11.Text = " &Expensis"
            ' Label12.Text = Label3.Text
            Label11.BackColor = System.Drawing.Color.Green
        Else
            If yy < 0 Then
                'MsgBox("You have loss which is " & Label3.Text)
                Label11.Text = "&Expensis are more"
                'Label12.Text = Label3.Text
                Label11.BackColor = System.Drawing.Color.Red
            End If
        End If
    End Sub
    Private Sub profit()
        Try
            Dim sum As Integer
            ' Dim minus As Integer
            Dim a As Integer
            Dim b As Integer
            Dim c As Integer
            Dim d As Integer
            Dim e As Integer
            a = Label1.Text
            b = Label2.Text
            c = Label7.Text
            d = Label9.Text
            e = Label10.Text
            sum = b + a + c + d - e
            Label3.Text = sum
            If sum > 0 Then
                MsgBox("You have profit which is =" & Label3.Text)
                Label13.Text = "You getting profit"
                ' Label12.Text = Label3.Text
                Label13.ForeColor = System.Drawing.Color.Green
                Label13.BackColor = System.Drawing.Color.White
            Else
                If sum < 0 Then
                    'MsgBox("You have loss which is " & Label3.Text)
                    Label13.Text = "You getting loss "
                    'Label12.Text = Label3.Text
                    Label13.ForeColor = System.Drawing.Color.White
                    Label13.BackColor = System.Drawing.Color.Red
                End If
            End If
            'MessageBox.Show(minus)
        Catch ex As Exception
            MsgBox("Not loading try again")
            Me.Dispose()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        profit()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Dispose()
    End Sub
End Class
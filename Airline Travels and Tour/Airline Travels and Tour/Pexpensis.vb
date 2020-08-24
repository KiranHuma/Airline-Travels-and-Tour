Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.Data.DataTable
Imports System.Data.SqlClient
Public Class Pexpensis
    Dim provider As String
    Dim dataFile As String
    Dim connString As String
    Dim myConnection As SqlConnection = New SqlConnection
    Dim ds As DataSet = New DataSet
    Dim da As SqlDataAdapter
    Dim tables As DataTableCollection = ds.Tables
    Dim source1 As New BindingSource()
    Dim dt As New DataTable
    Dim con As New SqlClient.SqlConnection
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
    Private Sub FillCombo()
        ' Dim myConnToAccess As OleDbConnection
        Dim myConnToAccess As SqlConnection
        Dim ds As DataSet
        'Dim da As OleDbDataAdapter
        Dim da As SqlDataAdapter
        Dim tables As DataTableCollection
        '  myConnToAccess = New OleDbConnection("provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb")
        myConnToAccess = New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
        myConnToAccess.Open()
        ds = New DataSet
        tables = ds.Tables
        'da = New OleDbDataAdapter("SELECT nme from expensis", myConnToAccess)
        da = New SqlDataAdapter("SELECT nme from expensis ", myConnToAccess)
        da.Fill(ds, "expensis")
        Dim view1 As New DataView(tables(0))
        With ComboBox1
            .DataSource = ds.Tables("expensis")
            .DisplayMember = "nme"
            .ValueMember = "nme"
            .SelectedIndex = -1
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.ListItems
        End With
    End Sub
    Private Sub gridfill()
        'dbaccessconnection()
        Me.Label1.Text = Format(Now, "dd-MMM-yyyy")
        Timer1.Enabled = True
        'provider = "provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb"
        provider = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        'da = New OleDbDataAdapter("Select [entry], [nme], [dte], [salry] ,[bill],[pckgecash],[prsnluse], [rent] from expensis", myConnection)
        da = New SqlDataAdapter("Select [entry],[nme],[bill],[pckgecash],[prsnluse], [rent],[salry] from expensis", myConnection)
        da.Fill(ds, "expensis")
        Dim view1 As New DataView(tables(0))
        source1.DataSource = view1
        DataGridView1.DataSource = view1
        DataGridView1.Refresh()
    End Sub
    Private Sub vprofit()
        Dim Tsum As Integer = 0
        Dim Tsum1 As Integer = 0
        Dim Tsum2 As Integer = 0
        Dim Tsum3 As Integer = 0
        Dim Tsum4 As Integer = 0
        For i As Integer = 0 To DataGridView1.Rows.Count() - 1 Step +1
            Tsum = Tsum + DataGridView1.Rows(i).Cells(2).Value
            Tsum1 = Tsum1 + DataGridView1.Rows(i).Cells(3).Value
            Tsum2 = Tsum2 + DataGridView1.Rows(i).Cells(4).Value
            Tsum3 = Tsum3 + DataGridView1.Rows(i).Cells(5).Value
            Tsum4 = Tsum4 + DataGridView1.Rows(i).Cells(6).Value

        Next
        Ebill.Text = Tsum.ToString()
        Ecash.Text = Tsum1.ToString()
        Puse.Text = Tsum2.ToString()
        prent.Text = Tsum3.ToString()
        Esalary.Text = Tsum4.ToString()

    End Sub
    Private Sub profit()
        Try
            Dim sum As Integer
            Dim sum1 As Integer
            Dim sum2 As Integer
            Dim a As Integer
            Dim b As Integer
            Dim c As Integer
            Dim d As Integer
            Dim f As Integer
            Dim g As Integer
            a = Ebill.Text
            b = Ecash.Text
            c = Puse.Text
            g = prent.Text
            f = Esalary.Text
            sum1 = a + b + c
            sum2 = d + f + g
            sum = sum1 + sum2
            texpensis.Text = sum
            'MessageBox.Show(minus)
        Catch ex As Exception
            MsgBox("please add the grand totals from umrah and ticket records")

        End Try
    End Sub
   

    Private Sub texpensis_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles texpensis.MouseClick
        profit()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Dispose()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        source1.Filter = "[nme] = '" & ComboBox1.Text & "'"
        DataGridView1.Refresh()
        ComboBox1.Text = ""
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        timer.Text = Date.Now.ToString(" hh:mm:ss")
    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Pexpensis_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.Label1.Text = Format(Now, "dd-MMM-yyyy")
        'Call CenterToScreen()
        'Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        'Me.WindowState = FormWindowState.Maximized
        gridfill()
        FillCombo()
        vprofit()
        dbaccessconnection()
        DataGridView1.Sort(DataGridView1.Columns("entry"), System.ComponentModel.ListSortDirection.Ascending)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        frmGrndrecord.Label10.Text = texpensis.Text
    End Sub

    Private Sub texpensis_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles texpensis.TextChanged

    End Sub
End Class
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.Data.DataTable
Imports System.Data.SqlClient
Public Class frmuRcrds
    Dim provider As String
    Dim dataFile As String
    Dim connString As String
    ' Dim myConnection As OleDbConnection = New OleDbConnection
    Dim myConnection As SqlConnection = New SqlConnection
    Dim ds As DataSet = New DataSet
    ' Dim da As OleDbDataAdapter
    Dim da As SqlDataAdapter
    Dim tables As DataTableCollection = ds.Tables
    Dim source1 As New BindingSource()
    Dim dt As New DataTable



    ' Dim con As New OleDb.OleDbConnection
    'Dim cmd As New OleDb.OleDbCommand
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
    Private Sub gridfill()
        Timer1.Enabled = True
        Me.Label1.Text = Format(Now, "dd-MMM-yyyy")
        ' provider = "provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb"
        provider = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        'entry,unme,udte,upasprt,utrvlngdte,uexdte,uorgnlcst,usale,uprft,uduratin,upkge,urefrnce)
        'da = New OleDbDataAdapter("Select [entry],[uorgnlcst],[usale],[uprft]from umrah ", myConnection)
        da = New SqlDataAdapter("Select [entry], [udte],[uorgnlcst],[usale],[uprft]from umrahh ", myConnection)
        da.Fill(ds, "umrahh ")
        Dim view1 As New DataView(tables(0))
        source1.DataSource = view1
        DataGridView1.DataSource = view1
        DataGridView1.Refresh()
    End Sub
    Private Sub Combofill()
        Dim myConnToAccess As SqlConnection
        Dim ds As DataSet
        Dim da As SqlDataAdapter
        Dim tables As DataTableCollection
        myConnToAccess = New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
        myConnToAccess.Open()
        ds = New DataSet
        tables = ds.Tables
        da = New SqlDataAdapter("SELECT entry from umrahh ", myConnToAccess)
        da.Fill(ds, "umrahh ")
        Dim view1 As New DataView(tables(0))
        With ComboBox1
            .DataSource = ds.Tables("umrahh ")
            .DisplayMember = "entry"
            .ValueMember = "entry"
            .SelectedIndex = 0
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.ListItems
        End With
    End Sub
   

    Private Sub frmuRcrds_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        gridfill()
        ' Call CenterToScreen()
        'Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        'Me.WindowState = FormWindowState.Maximized
        Combofill()
        urofit()
        txtucnt.Text = DataGridView1.RowCount
        Me.Label1.Text = Format(Now, "dd-MMM-yyyy")

        DataGridView1.Sort(DataGridView1.Columns("entry"), System.ComponentModel.ListSortDirection.Ascending)

        
    End Sub
    Private Sub urofit()
        Dim Tsum As Integer = 0
        For i As Integer = 0 To DataGridView1.Rows.Count() - 1 Step +1
            Tsum = Tsum + DataGridView1.Rows(i).Cells(4).Value
        Next
        txtuprofit.Text = Tsum.ToString()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        source1.Filter = "[entry] = '" & ComboBox1.Text & "'"
        DataGridView1.Refresh()
        ComboBox1.Text = ""
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        timer.Text = Date.Now.ToString(" hh:mm:ss")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Dispose()
    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        frmGrndrecord.Label2.Text = txtuprofit.Text
    End Sub

    Private Sub frmuRcrds_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseClick
        Call CenterToScreen()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub frmuRcrds_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDoubleClick
        Call CenterToScreen()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim cn As New SqlConnection
        Dim ds As New DataSet
        Dim dt As New DataTable
        Dim dfrom As DateTime = DateTimePicker1.Value
        Dim dto As DateTime = DateTimePicker2.Value
        cn.ConnectionString = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
        cn.Open()
        Dim str As String = "select * from umrahh where udte >= '" & Format(dfrom, "MM-dd-yyyy") & "' and udte <='" & Format(dto, "MM-dd-yyyy") & "'"
        Dim da As SqlDataAdapter = New SqlDataAdapter(str, cn)
        da.Fill(dt)
        DataGridView1.DataSource = dt
    End Sub

    Private Sub GroupBox2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox2.Enter

    End Sub
End Class
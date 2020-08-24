Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.Data.DataTable
Imports System.Data.SqlClient
Public Class Profit
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
    Private Sub gridfill()
        Timer1.Enabled = True
        Me.Label1.Text = Format(Now, "dd-MMM-yyyy")
        'provider = "provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb"
        provider = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        'vientry,nme,dte,psprtno,trvelingdte,expirdte,vipresent,viduration,refrnce
        da = New SqlDataAdapter("Select sum(comissin)from ticket group by ticketnmbr", myConnection)
        da.Fill(ds, "ticket")
        Dim view1 As New DataView(tables(0))
        source1.DataSource = View1
        DataGridView1.DataSource = view1
        DataGridView1.Refresh()
    End Sub
    Private Sub gridviewfill()


        Timer1.Enabled = True
        provider = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        da = New SqlDataAdapter("Select [ticketnmbr], [nme], [comissin] from ticket", myConnection)
        da.Fill(ds, "ticket")
        Dim view1 As New DataView(tables(0))
        source1.DataSource = view1
        DataGridView1.DataSource = view1
        DataGridView1.Refresh()
    End Sub
    Private Sub FillCombo()
        Dim myConnToAccess As SqlConnection
        Dim ds As DataSet
        Dim da As SqlDataAdapter
        Dim tables As DataTableCollection
        myConnToAccess = New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
        myConnToAccess.Open()
        ds = New DataSet
        tables = ds.Tables
        da = New SqlDataAdapter("SELECT ticketnmbr from ticket", myConnToAccess)
        da.Fill(ds, "ticket")
        Dim view1 As New DataView(tables(0))
        With ComboBox1
            .DataSource = ds.Tables("ticket")
            .DisplayMember = "ticketnmbr"
            .ValueMember = "ticketnmbr"
            .SelectedIndex = 0
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.ListItems
        End With
    End Sub
    Private Sub Tprofit()
        Dim Tsum As Integer = 0
        For i As Integer = 0 To DataGridView1.Rows.Count() - 1 Step +1
            Tsum = Tsum + DataGridView1.Rows(i).Cells(2).Value
        Next
        Tprftticket.Text = Tsum.ToString()

    End Sub
   Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        source1.Filter = "[nme] = '" & ComboBox1.Text & "'"
        DataGridView1.Refresh()
        ComboBox1.Text = ""
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        timer.Text = Date.Now.ToString(" hh:mm:ss")
    End Sub
    Private Sub Profit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        gridviewfill()
        FillCombo()
        Tprofit()
        Ttickeet.Text = DataGridView1.RowCount
        Me.Label1.Text = Format(Now, "dd-MMM-yyyy")

        'Call CenterToScreen()
        'Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        'Me.WindowState = FormWindowState.Maximized

        DataGridView1.Sort(DataGridView1.Columns("ticketnmbr"), System.ComponentModel.ListSortDirection.Ascending)

    End Sub
  

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Dispose()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        frmGrndrecord.Label1.Text = Tprftticket.Text
    End Sub

    Private Sub GroupBox2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub Profit_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseClick
        Call CenterToScreen()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub Profit_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDoubleClick
        Call CenterToScreen()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.WindowState = FormWindowState.Normal
    End Sub
End Class
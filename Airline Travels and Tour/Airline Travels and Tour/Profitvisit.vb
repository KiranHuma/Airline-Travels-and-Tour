Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.Data.DataTable
Imports System.Data.SqlClient
Public Class Profitvisit
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
        da = New SqlDataAdapter("Select [vientry], [nme],[proft]from visit ", myConnection)
        da.Fill(ds, "visit ")
        Dim view1 As New DataView(tables(0))
        source1.DataSource = view1
        DataGridView1.DataSource = view1
        DataGridView1.Refresh()
    End Sub
    Private Sub FillCombo()
        'Dim myConnToAccess As OleDbConnection
        Dim myConnToAccess As SqlConnection
        Dim ds As DataSet
        ' Dim da As OleDbDataAdapter
        Dim da As SqlDataAdapter
        Dim tables As DataTableCollection
        ' myConnToAccess = New OleDbConnection("provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb")
        myConnToAccess = New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
        myConnToAccess.Open()
        ds = New DataSet
        tables = ds.Tables
        da = New SqlDataAdapter("SELECT nme from visit", myConnToAccess)
        da.Fill(ds, "visit")
        Dim view1 As New DataView(tables(0))
        With ComboBox1
            .DataSource = ds.Tables("visit")
            .DisplayMember = "nme"
            .ValueMember = "nme"
            .SelectedIndex = 0
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.ListItems
        End With
    End Sub
    Private Sub vprofit()
        Dim Tsum As Integer = 0
        For i As Integer = 0 To DataGridView1.Rows.Count() - 1 Step +1
            Tsum = Tsum + DataGridView1.Rows(i).Cells(2).Value
        Next
        Vprft.Text = Tsum.ToString()

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
    Private Sub listview()
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim itemcoll(100) As String
        Me.ListView1.View = View.Details
        Me.ListView1.GridLines = True
        Dim conn As New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
        Dim strQ As String = String.Empty
        strQ = "SELECT * FROM visit"
        cmd = New SqlCommand(strQ, conn)
        da = New SqlDataAdapter(cmd)
        ds = New DataSet
        da.Fill(ds, "visit")
        Dim i As Integer = 0
        Dim j As Integer = 0
        ' adding the columns in ListView
        For i = 0 To ds.Tables(0).Columns.Count - 1
            Me.ListView1.Columns.Add(ds.Tables(0).Columns(i).ColumnName.ToString())
        Next
        'Now adding the Items in Listview
        For i = 0 To ds.Tables(0).Rows.Count - 1
            For j = 0 To ds.Tables(0).Columns.Count - 1
                itemcoll(j) = ds.Tables(0).Rows(i)(j).ToString()
            Next
            Dim lvi As New ListViewItem(itemcoll)
            Me.ListView1.Items.Add(lvi)
            Me.ListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)
        Next
        Me.ListView1.View = View.Details
        Me.ListView1.GridLines = True
      
    End Sub
    Private Sub Profitvisit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'totlvisit.Text = DataGridView1.RowCount
        Me.Label1.Text = Format(Now, "dd-MMM-yyyy")
        listview()
        'Call CenterToScreen()
        ' Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        ' Me.WindowState = FormWindowState.Maximized
        gridfill()
        FillCombo()
        vprofit()
        count1()
        totlvisit.Text = DataGridView1.RowCount
        dbaccessconnection()

        DataGridView1.Sort(DataGridView1.Columns("vientry"), System.ComponentModel.ListSortDirection.Ascending)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        frmGrndrecord.Label7.Text = Vprft.Text
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        
        searchlstview()
        ' movls()
    End Sub
    Private Sub movoneitm()
        Dim lvi As New ListViewItem
        If ListView1.SelectedItems.Count > 0 Then
            For i As Integer = 0 To ListView1.SelectedItems.Count - 1
                lvi = ListView1.SelectedItems(i)

                Dim lvi2 As New ListViewItem
                lvi2 = CType(lvi.Clone, ListViewItem)
                ListView2.Items.Add(lvi2)
            Next
        End If


    End Sub

    Private Sub movwholerow()
        
    End Sub
    Private Sub searchlstview()

        ListView1.Focus()

        For i = 0 To ListView1.Items.Count - 1
            ' If ListView1.Items(i).SubItems(1).Text = TextBox1.Text.ToLower Then //for whole word search
            If ListView1.Items(i).SubItems(1).Text.ToLower.Contains(TextBox1.Text.ToLower) Then
                ListView1.Items(i).Selected = True
                ListView1.EnsureVisible(i)
            End If
        Next
    End Sub
    Private Sub count1()
        Try
            dbaccessconnection()
            con.Open()
            Dim num As New Integer
            cmd.CommandText = "Select count(*) from vissa "
            If (IsDBNull(cmd.ExecuteScalar)) Then
                num = 0
                totlvisit.Text = num.ToString
            Else
                'num = cmd.ExecuteScalar()
                'txtticket.Text = num + 1
                num = cmd.ExecuteScalar
                totlvisit.Text = num.ToString
            End If
            con.Close()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
            Me.Dispose()
        End Try

    End Sub
    Private Sub search()
        Dim SqlQuery As String = "SELECT * FROM visit WHERE nme like ' % " & TextBox1.Text & " % ' "
        Dim SqlCommand As New SqlCommand
        Dim SqlAdapter As New SqlDataAdapter
        Dim TABLE As New DataTable
        'MsgBox("trigger")
        Dim conn As New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
        With SqlCommand
            .CommandText = SqlQuery
            .Connection = conn

        End With

        With SqlAdapter
            .SelectCommand = SqlCommand
            .Fill(TABLE)
        End With

        ListView1.Items.Clear()
        For i = 0 To TABLE.Rows.Count - 1
            Dim li As New ListViewItem
            li = ListView1.Items.Add(TABLE.Rows(i)("vientry").ToString())
            li.SubItems.Add(TABLE.Rows(i)("nme").ToString())
            li.SubItems.Add(TABLE.Rows(i)("dte").ToString())
            li.SubItems.Add(TABLE.Rows(i)("psprtno").ToString())
            li.SubItems.Add(TABLE.Rows(i)("trvelingdte").ToString())
            li.SubItems.Add(TABLE.Rows(i)("expirdte").ToString())
            li.SubItems.Add(TABLE.Rows(i)("vipresent").ToString())
            li.SubItems.Add(TABLE.Rows(i)("orgnlcst").ToString())
            li.SubItems.Add(TABLE.Rows(i)("viduration").ToString())
            li.SubItems.Add(TABLE.Rows(i)("refrnce").ToString())
        Next
        ' vientry,nme,dte,psprtno,trvelingdte,expirdte,vipresent,orgnlcst,sale,proft,viduration,refrnce
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        movoneitm()
    End Sub
End Class
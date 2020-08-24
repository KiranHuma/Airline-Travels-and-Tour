Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.Data.DataTable
Imports System.Data.SqlClient
Public Class Profitvissa
    Dim provider As String
    Dim dataFile As String
    Dim connString As String
    ' Dim myConnection As OleDbConnection = New OleDbConnection
    Dim myConnection As SqlConnection = New SqlConnection
    Dim ds As DataSet = New DataSet
    'Dim da As OleDbDataAdapter
    Dim da As SqlDataAdapter
    Dim tables As DataTableCollection = ds.Tables
    Dim source1 As New BindingSource()
    Dim source2 As New BindingSource()
    Dim dv As DataView


    Dim dt As New DataTable
    Dim con As New SqlClient.SqlConnection
    Dim cmd As New SqlClient.SqlCommand

    Private Sub dbaccessconnection()
        'Acces DataBase Connectivity and for MS Access 2003 PROVIDER=Microsoft.Jet.OLEDB.4.0
        Try
            con.ConnectionString = "Data Source=ADMINRG-HKIL2V5;Initial Catalog=airlinee;Integrated Security=True"
            cmd.Connection = con
            'MessageBox.Show(con.State.ToString())
        Catch ex As Exception
            MsgBox("DataBase not connected due to the reason because " & ex.Message)
        End Try
    End Sub

    Private Sub Profitvissa_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dbaccessconnection()
        Label12.Visible = False
        ' gridfill()

        Fillvwakala()
        Fillname()
        Fillvnmber()
        Fillvcntry()
        Fillvwrk()
        count1()
        count2()
        count3()
        ' totlvisit.Text = DataGridView1.RowCount
        ' TextBox3.Text = DataGridView1.RowCount
        'TextBox1.Text = DataGridView1.RowCount
        ' DataGridView1.Sort(DataGridView1.Columns("visanmbr"), System.ComponentModel.ListSortDirection.Ascending)
    End Sub
    Private Sub Tprofit()
        'search data through textboxs
        Try
            dbaccessconnection()
            con.Open()
            Dim num As New Integer
            cmd.CommandText = "SELECT SUM(comissin) from vissa "
            If (IsDBNull(cmd.ExecuteScalar)) Then
                num = 0
                Viprofit.Text = num.ToString
            Else
                'num = cmd.ExecuteScalar()
                'txtticket.Text = num + 1
                num = cmd.ExecuteScalar
                Viprofit.Text = num.ToString
            End If
            con.Close()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
            Me.Dispose()
        End Try
    End Sub
    Private Sub Tsale()
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
    Private Sub count1()
        Try
            dbaccessconnection()
            con.Open()
            Dim num As New Integer
            cmd.CommandText = "Select count(*) from vissa "
            If (IsDBNull(cmd.ExecuteScalar)) Then
                num = 0
                TextBox1.Text = num.ToString
            Else
                'num = cmd.ExecuteScalar()
                'txtticket.Text = num + 1
                num = cmd.ExecuteScalar
                TextBox1.Text = num.ToString
            End If
            con.Close()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
            Me.Dispose()
        End Try
      
    End Sub
    Private Sub count2()
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
    Private Sub count3()
        Try
            dbaccessconnection()
            con.Open()
            Dim num As New Integer
            cmd.CommandText = "Select count(*) from vissa "
            If (IsDBNull(cmd.ExecuteScalar)) Then
                num = 0
                TextBox3.Text = num.ToString
            Else
                'num = cmd.ExecuteScalar()
                'txtticket.Text = num + 1
                num = cmd.ExecuteScalar
                TextBox3.Text = num.ToString
            End If
            con.Close()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
            Me.Dispose()
        End Try

    End Sub
    Private Sub Viprofit()
        Dim Tsum As Integer = 0
        For i As Integer = 0 To DataGridView1.Rows.Count() - 1 Step +1
            Tsum = Tsum + DataGridView1.Rows(i).Cells(7).Value
        Next
        Viprft.Text = Tsum.ToString()

    End Sub
    Private Sub Visale()
        Dim Tsum As Integer = 0
        For i As Integer = 0 To DataGridView1.Rows.Count() - 1 Step +1
            Tsum = Tsum + DataGridView1.Rows(i).Cells(5).Value
        Next
        TextBox2.Text = Tsum.ToString()

    End Sub
    Private Sub Vipurchse()
        Dim Tsum As Integer = 0
        For i As Integer = 0 To DataGridView1.Rows.Count() - 1 Step +1
            Tsum = Tsum + DataGridView1.Rows(i).Cells(6).Value
        Next
        TextBox4.Text = Tsum.ToString()

    End Sub
    Private Sub Fillname()
        ' Dim myConnToAccess As OleDbConnection
        Dim myConnToAccess As SqlConnection
        Dim ds As DataSet
        'Dim da As OleDbDataAdapter
        Dim da As SqlDataAdapter
        Dim tables As DataTableCollection
        ' myConnToAccess = New OleDbConnection("provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb")
        myConnToAccess = New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
        myConnToAccess.Open()
        ds = New DataSet
        tables = ds.Tables
        'da = New OleDbDataAdapter("SELECT vnme from vissa ", myConnToAccess)
        da = New SqlDataAdapter("SELECT vnme from vissa ", myConnToAccess)
        da.Fill(ds, "vissa ")
        Dim view1 As New DataView(tables(0))
        With ComboBox1
            .DataSource = ds.Tables("vissa ")
            .DisplayMember = "vnme"
            .ValueMember = "vnme"
            .SelectedIndex = -1
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.ListItems
        End With
    End Sub
    Private Sub Fillvnmber()
        ' Dim myConnToAccess As OleDbConnection
        Dim myConnToAccess As SqlConnection
        Dim ds As DataSet
        'Dim da As OleDbDataAdapter
        Dim da As SqlDataAdapter
        Dim tables As DataTableCollection
        ' myConnToAccess = New OleDbConnection("provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb")
        myConnToAccess = New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
        myConnToAccess.Open()
        ds = New DataSet
        tables = ds.Tables
        'da = New OleDbDataAdapter("SELECT vnme from vissa ", myConnToAccess)
        da = New SqlDataAdapter("SELECT visanmbr from vissa ", myConnToAccess)
        da.Fill(ds, "vissa ")
        Dim view1 As New DataView(tables(0))
        With Combonmbr
            .DataSource = ds.Tables("vissa ")
            .DisplayMember = "visanmbr"
            .ValueMember = "visanmbr"
            .SelectedIndex = -1
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.ListItems
        End With
    End Sub
    Private Sub Fillvcntry()
        ' Dim myConnToAccess As OleDbConnection
        Dim myConnToAccess As SqlConnection
        Dim ds As DataSet
        'Dim da As OleDbDataAdapter
        Dim da As SqlDataAdapter
        Dim tables As DataTableCollection
        ' myConnToAccess = New OleDbConnection("provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb")
        myConnToAccess = New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
        myConnToAccess.Open()
        ds = New DataSet
        tables = ds.Tables
        'da = New OleDbDataAdapter("SELECT vnme from vissa ", myConnToAccess)
        da = New SqlDataAdapter("SELECT vcntry from vissa ", myConnToAccess)
        da.Fill(ds, "vissa ")
        Dim view1 As New DataView(tables(0))
        With Combocntry
            .DataSource = ds.Tables("vissa ")
            .DisplayMember = "vcntry"
            .ValueMember = "vcntry"
            .SelectedIndex = -1
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.ListItems
        End With
    End Sub
    Private Sub Fillvwrk()
        ' Dim myConnToAccess As OleDbConnection
        Dim myConnToAccess As SqlConnection
        Dim ds As DataSet
        'Dim da As OleDbDataAdapter
        Dim da As SqlDataAdapter
        Dim tables As DataTableCollection
        ' myConnToAccess = New OleDbConnection("provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb")
        myConnToAccess = New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
        myConnToAccess.Open()
        ds = New DataSet
        tables = ds.Tables
        'da = New OleDbDataAdapter("SELECT vnme from vissa ", myConnToAccess)
        da = New SqlDataAdapter("SELECT vwrk from vissa ", myConnToAccess)
        da.Fill(ds, "vissa ")
        Dim view1 As New DataView(tables(0))
        With Combowrk
            .DataSource = ds.Tables("vissa ")
            .DisplayMember = "vwrk"
            .ValueMember = "vwrk"
            .SelectedIndex = -1
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.ListItems
        End With
    End Sub
    Private Sub Fillvwakala()
        ' Dim myConnToAccess As OleDbConnection
        Dim myConnToAccess As SqlConnection
        Dim ds As DataSet
        'Dim da As OleDbDataAdapter
        Dim da As SqlDataAdapter
        Dim tables As DataTableCollection
        ' myConnToAccess = New OleDbConnection("provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb")
        myConnToAccess = New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
        myConnToAccess.Open()
        ds = New DataSet
        tables = ds.Tables
        'da = New OleDbDataAdapter("SELECT vnme from vissa ", myConnToAccess)
        da = New SqlDataAdapter("SELECT vwakala from vissa ", myConnToAccess)
        da.Fill(ds, "vissa ")
        Dim view1 As New DataView(tables(0))
        With Combowakala
            .DataSource = ds.Tables("vissa ")
            .DisplayMember = "vwakala"
            .ValueMember = "vwakala"
            .SelectedIndex = -1
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.ListItems
        End With
    End Sub
    Private Sub gridfill1()
        Dim ds1 As DataSet = New DataSet
        'Dim da As OleDbDataAdapter
        Dim da1 As SqlDataAdapter
        Dim tables1 As DataTableCollection = ds.Tables

        Dim source2 As New BindingSource()


        Dim dt As New DataTable
        Timer1.Enabled = True
        Me.Label1.Text = Format(Now, "dd-MMM-yyyy")
        ' provider = "provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb"
        provider = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        'vissa( visanmbr,vnme,vdte,vcntry,vwrk,vprce,vorgnlcst,vprft,vwakala
        ' da = New OleDbDataAdapter("Select [visanmbr], [vnme], [vdte],[vcntry],[vwrk],[vprce],[vorgnlcst],[vprft],[vwakala]from vissa ", myConnection)
        da1 = New SqlDataAdapter("Select visanmbr,vnme,vdte,vcntry,vwrk,vprce,vorgnlcst,vprft,vwakala from vissa ", myConnection)
        da1.Fill(ds1, "vissa")
        Dim view2 As New DataView(tables1(0))
        source2.DataSource = view2
        DataGridView1.DataSource = view2
        DataGridView1.Refresh()
    End Sub

    Private Sub gridfill()
        Try
            Timer1.Enabled = True
            Me.Label1.Text = Format(Now, "dd-MMM-yyyy")
            ' provider = "provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb"
            provider = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            'vissa( visanmbr,vnme,vdte,vcntry,vwrk,vprce,vorgnlcst,vprft,vwakala
            ' da = New OleDbDataAdapter("Select [visanmbr], [vnme], [vdte],[vcntry],[vwrk],[vprce],[vorgnlcst],[vprft],[vwakala]from vissa ", myConnection)
            da = New SqlDataAdapter("Select visanmbr,vnme,vdte,vcntry,vwrk,vprce,vorgnlcst,vprft,vwakala from vissa ", myConnection)
            da.Fill(ds, "vissa")
            Dim view1 As New DataView(tables(0))
            source1.DataSource = view1
            DataGridView1.DataSource = view1
            DataGridView1.Refresh()
        Catch ex As Exception

            MsgBox("DataBase not connected due to the reason because " & ex.Message)
        End Try

    End Sub

    Private Sub getdata()
        Dim con As New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
        con.Open()
        Dim da As New SqlDataAdapter("Select visanmbr,vnme,vdte,vcntry,vwrk,vprce,vorgnlcst,vprft,vwakala from vissa", con)
        Dim dt As New DataTable
        da.Fill(dt)
        'Dim view1 As New DataView(tables1(0))
        source2.DataSource = dt
        DataGridView1.DataSource = dt
        DataGridView1.Refresh()
        con.Close()
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        source1.Filter = "[vnme] = '" & ComboBox1.Text & "'"
        DataGridView1.Refresh()
        ComboBox1.Text = ""
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Dispose()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        frmGrndrecord.Label9.Text = Viprft.Text
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        DataGridView1.DataSource = Nothing
        GroupBox1.Visible = True
        GroupBox1.Enabled = True
        GroupBox2.Visible = False
        GroupBox2.Enabled = False
        GroupBox3.Visible = False
        GroupBox3.Enabled = False
        GroupBox4.Visible = False
        GroupBox4.Enabled = False
        ComboBox1.Visible = True
        ComboBox1.Enabled = True
        Combonmbr.Visible = False
        Combonmbr.Enabled = False
        Combocntry.Enabled = False
        Combocntry.Visible = False
        Combowrk.Enabled = False
        Combowrk.Visible = False
        Combowakala.Enabled = False
        Combowakala.Visible = False


        
        bydte.Visible = False
        Label12.Visible = False
        DateTimePicker1.Visible = False
        DateTimePicker1.Enabled = False
        DateTimePicker2.Visible = False
        DateTimePicker2.Enabled = False


    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        source1.Filter = "[vnme] = '" & ComboBox1.Text & "'"
        source2.Filter = "[vnme] = '" & ComboBox1.Text & "'"
        DataGridView1.Refresh()
        ComboBox1.Text = ""

    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        GroupBox1.Visible = True
        GroupBox1.Enabled = True
        GroupBox2.Visible = False
        GroupBox2.Enabled = False
        GroupBox3.Visible = False
        GroupBox3.Enabled = False
        ComboBox1.Visible = False
        ComboBox1.Enabled = False
        Combonmbr.Visible = True
        Combonmbr.Enabled = True
        Combocntry.Enabled = False
        Combocntry.Visible = False
        Combowrk.Enabled = False
        Combowrk.Visible = False
        Combowakala.Enabled = False
        Combowakala.Visible = False
        GroupBox4.Visible = False
        GroupBox4.Enabled = False


        bydte.Enabled = False
        bydte.Visible = False
        Label12.Visible = False
        DateTimePicker1.Visible = False
        DateTimePicker1.Enabled = False
        DateTimePicker2.Visible = False
        DateTimePicker2.Enabled = False
    End Sub

    Private Sub btnnmbr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        source1.Filter = "[visanmbr] = '" & Combonmbr.Text & "'"
        source2.Filter = "[visanmbr] = '" & Combonmbr.Text & "'"
        DataGridView1.Refresh()
        ComboBox1.Text = ""
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        DataGridView1.DataSource = Nothing

        GroupBox1.Visible = True
        GroupBox1.Enabled = True
        GroupBox4.Visible = False
        GroupBox4.Enabled = False
        GroupBox2.Visible = False
        GroupBox2.Enabled = False
        GroupBox3.Visible = False
        GroupBox3.Enabled = False
        ComboBox1.Visible = False
        ComboBox1.Enabled = False
        Combonmbr.Visible = False
        Combonmbr.Enabled = False
        Combocntry.Enabled = True
        Combocntry.Visible = True
        Combowrk.Enabled = False
        Combowrk.Visible = False
        Combowakala.Enabled = False
        Combowakala.Visible = False

        
        bydte.Enabled = False
        bydte.Visible = False
        Label12.Visible = False
        DateTimePicker1.Visible = False
        DateTimePicker1.Enabled = False
        DateTimePicker2.Visible = False
        DateTimePicker2.Enabled = False
    End Sub

    Private Sub btncntry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        source1.Filter = "[vcntry] = '" & Combocntry.Text & "'"
        source2.Filter = "[vcntry] = '" & Combocntry.Text & "'"
        DataGridView1.Refresh()
        ComboBox1.Text = ""
    End Sub

    Private Sub btnwrk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        source1.Filter = "[vwrk] = '" & Combowrk.Text & "'"
        source2.Filter = "[vwrk] = '" & Combowrk.Text & "'"
        DataGridView1.Refresh()
        ComboBox1.Text = ""
    End Sub

    Private Sub btnwakala_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        source1.Filter = "[vwakala] = '" & Combowakala.Text & "'"
        source2.Filter = "[vwakala] = '" & Combowakala.Text & "'"
        DataGridView1.Refresh()
        ComboBox1.Text = ""
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        DataGridView1.DataSource = Nothing

        GroupBox1.Visible = True
        GroupBox1.Enabled = True
        GroupBox2.Visible = False
        GroupBox2.Enabled = False
        GroupBox3.Visible = False
        GroupBox3.Enabled = False
        ComboBox1.Visible = False
        ComboBox1.Enabled = False
        Combonmbr.Visible = False
        Combonmbr.Enabled = False
        Combocntry.Enabled = False
        Combocntry.Visible = False
        Combowrk.Enabled = True
        Combowrk.Visible = True
        GroupBox4.Visible = False
        GroupBox4.Enabled = False
        Combowakala.Enabled = False
        Combowakala.Visible = False


        Label12.Visible = False
        bydte.Enabled = False
        bydte.Visible = False
        DateTimePicker1.Visible = False
        DateTimePicker1.Enabled = False
        DateTimePicker2.Visible = False
        DateTimePicker2.Enabled = False
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        DataGridView1.DataSource = Nothing

        GroupBox1.Visible = True
        GroupBox1.Enabled = True
        GroupBox2.Visible = False
        GroupBox2.Enabled = False
        GroupBox3.Visible = False
        GroupBox3.Enabled = False
        ComboBox1.Visible = False
        ComboBox1.Enabled = False
        Combonmbr.Visible = False
        Combonmbr.Enabled = False
        Combocntry.Enabled = False
        Combocntry.Visible = False
        Combowrk.Enabled = False
        Combowrk.Visible = False
        GroupBox4.Visible = False
        GroupBox4.Enabled = False
        Combowakala.Enabled = True
        Combowakala.Visible = True

        
        bydte.Enabled = False
        bydte.Visible = False
        Label12.Visible = False
        DateTimePicker1.Visible = False
        DateTimePicker1.Enabled = False
        DateTimePicker2.Visible = False
        DateTimePicker2.Enabled = False
    End Sub

    Private Sub reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' ds.Tables.Clear()
        Try
            con = New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
            con.Open()
            cmd = New SqlCommand("Select visanmbr,vnme,vdte,vcntry,vwrk,vprce,vorgnlcst,vprft,vwakala from vissa where vnme='" & ComboBox1.Text & "' order by vnme", con)
            Dim myDA As SqlDataAdapter = New SqlDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "vissa")
            DataGridView1.DataSource = myDataSet.Tables("vissa").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Me.DataGridView1.DataSource = Nothing 'clear out the datasource for the Grid view
        ' da.Fill(ds, "vissa ") ' refill the table adapter from the dataset table
        ' Me.DataGridView1.DataSource = Me.source1 'reset the datasource from the binding source
        ' Me.DataGridView1.Refresh() 'should redraw with the new data
        ' DataSet1.Table2.Clear() ' clear related dataset table

        ' TableAdapter1.Fill(DataSet1.Table1) ' refill datatable and datagridview

        'TableAdapter2.Fill(Me.DataSet1.Table2)


    End Sub

    Private Sub bydte_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim cn As New SqlConnection
        Dim ds As New DataSet
        Dim dt As New DataTable
        Dim dfrom As DateTime = DateTimePicker1.Value
        Dim dto As DateTime = DateTimePicker2.Value
        cn.ConnectionString = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
        cn.Open()
        Dim str As String = "select * from vissa  where vdte>= '" & Format(dfrom, "MM-dd-yyyy") & "' and vdte <='" & Format(dto, "MM-dd-yyyy") & "'"
        Dim da As SqlDataAdapter = New SqlDataAdapter(str, cn)
        da.Fill(dt)
        DataGridView1.DataSource = dt
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        GroupBox1.Visible = True
        GroupBox1.Enabled = True
        GroupBox2.Visible = False
        GroupBox2.Enabled = False
        GroupBox3.Visible = False
        GroupBox3.Enabled = False
        ComboBox1.Visible = False
        ComboBox1.Enabled = False
        Combonmbr.Visible = False
        Combonmbr.Enabled = False
        Combocntry.Enabled = False
        Combocntry.Visible = False
        Combowrk.Enabled = False
        Combowrk.Visible = False
        GroupBox4.Visible = False
        GroupBox4.Enabled = False
        Combowakala.Enabled = False
        Combowakala.Visible = False

        bydte.Enabled = True
        bydte.Visible = True

        Label12.Visible = True
        DateTimePicker1.Visible = True
        DateTimePicker1.Enabled = True
        DateTimePicker2.Visible = True
        DateTimePicker2.Enabled = True
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView1.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView1.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ComboBox1.SelectedIndex = -1

        DataGridView1.DataSource = Nothing
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            con = New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
            con.Open()
            cmd = New SqlCommand("Select visanmbr,vnme,vdte,vcntry,vwrk,vprce,vorgnlcst,vprft,vwakala from vissa where vnme='" & ComboBox1.Text & "' order by vnme", con)
            Dim myDA As SqlDataAdapter = New SqlDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "vissa")
            DataGridView1.DataSource = myDataSet.Tables("vissa").DefaultView
            con.Close()
        Catch ex As Exception
            MsgBox("DataBase not connected due to the reason because " & ex.Message)
        End Try
    End Sub

    Private Sub Label15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label15.Click
        Try
            con = New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
            con.Open()
            cmd = New SqlCommand("Select visanmbr,vnme,vdte,vcntry,vwrk,vprce,vorgnlcst,vprft,vwakala from vissa where vnme='" & ComboBox1.Text & "' order by vnme", con)
            Dim myDA As SqlDataAdapter = New SqlDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "vissa")
            DataGridView1.DataSource = myDataSet.Tables("vissa").DefaultView
            con.Close()
        Catch ex As Exception
            MsgBox("DataBase not connected due to the reason because " & ex.Message)
        End Try
    End Sub

    Private Sub Label14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label14.Click
        ComboBox1.SelectedIndex = -1

        DataGridView1.DataSource = Nothing
    End Sub

    Private Sub Combowakala_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combowakala.SelectedIndexChanged
        Try
            con = New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
            con.Open()
            cmd = New SqlCommand("Select visanmbr,vnme,vdte,vcntry,vwrk,vprce,vorgnlcst,vprft,vwakala from vissa where vwakala='" & Combowakala.Text & "' order by vwakala ", con)
            Dim myDA As SqlDataAdapter = New SqlDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "vissa")
            DataGridView1.DataSource = myDataSet.Tables("vissa").DefaultView
            con.Close()
        Catch ex As Exception
            MsgBox("DataBase not connected due to the reason because " & ex.Message)
        End Try
    End Sub

    Private Sub Combocntry_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combocntry.SelectedIndexChanged
        Try
            con = New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
            con.Open()
            cmd = New SqlCommand("Select visanmbr,vnme,vdte,vcntry,vwrk,vprce,vorgnlcst,vprft,vwakala from vissa where vcntry='" & Combocntry.Text & "' order by vcntry", con)
            Dim myDA As SqlDataAdapter = New SqlDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "vissa")
            DataGridView1.DataSource = myDataSet.Tables("vissa").DefaultView
            con.Close()
        Catch ex As Exception
            MsgBox("DataBase not connected due to the reason because " & ex.Message)
        End Try
    End Sub


    Private Sub Combowrk_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combowrk.SelectedIndexChanged
        Try
            con = New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
            con.Open()
            cmd = New SqlCommand("Select visanmbr,vnme,vdte,vcntry,vwrk,vprce,vorgnlcst,vprft,vwakala from vissa where vwrk='" & Combowrk.Text & "'order by vwrk", con)
            Dim myDA As SqlDataAdapter = New SqlDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "vissa")
            DataGridView1.DataSource = myDataSet.Tables("vissa").DefaultView
            con.Close()
        Catch ex As Exception
            MsgBox("DataBase not connected due to the reason because " & ex.Message)
        End Try
    End Sub

End Class
'Option Explicit On   ' for excel
Imports System.Data 'for access and sql same
Imports System.Data.OleDb 'for access and sql same
Imports System.Data.Odbc 'for access and sql same
Imports System.Data.DataTable 'for access and sql same
Imports System.Data.SqlClient 'for sql
'Imports System.Configuration   ' for excel
Imports Excel = Microsoft.Office.Interop.Excel ' for excel
'Imports System.Drawing.Printing  ' for excel
'Imports System.IO                   ' FOR FILE ACCESS.
Imports ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat ' for excel

Public Class Ticket

    Private bitmap As Bitmap 'for print grid
    Dim rdr As SqlDataReader
    Dim colColors As Collection = New Collection 'for color of listbox
    Dim provider As String  'for access and sql same
    Dim dataFile As String  'for access and sql same
    Dim connString As String   'for access and sql same
    ' Dim myConnection As OleDbConnection = New OleDbConnection   'for access replace it  Dim myConnection As SqlConnection = New SqlConnection
    Dim myConnection As SqlConnection = New SqlConnection
    Dim ds As DataSet = New DataSet            'for access and sql same
    ' Dim da As OleDbDataAdapter                'for access replace it with Dim da As SqlDataAdapter
    Dim da As SqlDataAdapter
    Dim tables As DataTableCollection = ds.Tables  'for access and sql same
    Dim source1 As New BindingSource()                    'for access and sql same
    Dim source2 As New BindingSource()
    Dim con As New SqlClient.SqlConnection                      'for sql
    Dim cmd As New SqlClient.SqlCommand                        'for sql

    Dim dt As New DataTable
    Dim cs As String = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
    ' Dim con As New OleDb.OleDbConnection 'for access
    'Dim cmd As New OleDb.OleDbCommand  'for access
    'replace
    ' Dim con As New SqlClient.SqlConnection            
    ' Dim cmd As New SqlClient.SqlCommand
    '(((((((((((Private Sub dbaccessconnection()
    'Acces DataBase Connectivity and for MS Access 2003 PROVIDER=Microsoft.Jet.OLEDB.4.0
    ' Try
    ' con.ConnectionString = "provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb"
    ' cmd.Connection = con
    'MessageBox.Show("connection created")
    ' Catch ex As Exception
    'MsgBox("DataBase not connected due to the reason because " & ex.Message)
    ' End Try
    'End Sub))))))))))))))))))) brackets offcourse removed
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
    Private Sub listview()
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim itemcoll(100) As String
        Me.ListView1.View = View.Details
        Me.ListView1.GridLines = True
        Dim conn As New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
        Dim strQ As String = String.Empty
        strQ = "Select ticketnmbr, nme as [Name], sector as [Sector] ,dte as [Date],pnr as [PNR],issuedte as [Issue Date], expiredte as [Expire Date], mobileno as [Mobile No.],flight as [Flight],pasprtnmbr as [Passport Number], comissin as [Comission], basicfair as [Basic Fair], deprttme as [Departure Time] ,arrtme as [Arrival Time],totalamount as [Total Amount],status as [Status]from ticket"
        cmd = New SqlCommand(strQ, conn)
        da = New SqlDataAdapter(cmd)
        ds = New DataSet
        da.Fill(ds, "ticket")
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
    Private Sub Ticket_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadall()
        FillCombo1()
        FillCombo2()

        
        'Dim checkbx As New DataGridViewCheckBoxColumn
        'checkbx.HeaderText = "select"
        'DataGridView1.Columns.Add(checkbx)
        ' DataGridView1.Columns(2).DefaultCellStyle.ForeColor = Color.Blue


        'Call CenterToScreen()
        'Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        'Me.WindowState = FormWindowState.Maximized
        'DataGridView1.BackColor = Color.Transparent
        'TransparencyKey = BackColor
        'txtticket.BackColor = Me.BackColor
        dbaccessconnection()
        ComboBox1.Text = "----Select Or Type Name---"
        Me.Label23.Text = Format(Now, "dd-MMM-yyyy")
        Timer1.Enabled = True
        GroupBox1.Enabled = False
        Timer1.Start()

    End Sub


    Private Sub loadall()
        gridviewfill()
        FillCombo()
        ' Me.PopulateCombobox()
        ' Me.DataGridView1.AllowUserToAddRows = False
       ' DataGridView1.Sort(DataGridView1.Columns("ticketnmbr"), System.ComponentModel.ListSortDirection.Ascending)
        Me.Refresh()
    End Sub

    Private Sub PopulateCombobox()
        Dim constr As String = ("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
        Using conn As New SqlConnection(constr)
            Using cmd As New SqlCommand("SELECT [nme]  from ticket", conn)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    ' Create a new row
                    Dim dr As DataRow = dt.NewRow()
                    dr("nme") = ""
                    dt.Rows.InsertAt(dr, 0)
                    Me.ComboBox1.DisplayMember = "nme"
                    Me.ComboBox1.ValueMember = "nme"

                    Me.ComboBox1.DataSource = dt

                End Using
            End Using
        End Using
    End Sub
    Private Sub FillCombo()
        Try
            ' Dim myConnToAccess As OleDbConnection for access
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
            da = New SqlDataAdapter("SELECT nme from ticket", myConnToAccess)
            '    da = New OleDbDataAdapter("SELECT ticketnmbr from ticket", myConnToAccess)
            da.Fill(ds, "ticket")
            Dim view1 As New DataView(tables(0))
            With ComboBox1
                .DataSource = ds.Tables("ticket")
                .DisplayMember = "nme"
                .ValueMember = "nme"
                .SelectedIndex = -1
                '  .SelectedIndex = 0 this give error when no entry in form occur to solve this error i use aove line 0 is replace by -1

                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .AutoCompleteSource = AutoCompleteSource.ListItems
            End With
        Catch ex As Exception
            MessageBox.Show("Errpr while loading combobox", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.Dispose()
        End Try
    End Sub
    Private Sub FillCombo1()
        Try
            ' Dim myConnToAccess As OleDbConnection for access
            Dim myConnToAccess As SqlConnection
            Dim ds1 As DataSet
            ' Dim da As OleDbDataAdapter
            Dim da1 As SqlDataAdapter
            Dim tables As DataTableCollection
            ' myConnToAccess = New OleDbConnection("provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb")
            myConnToAccess = New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
            myConnToAccess.Open()
            ds1 = New DataSet
            tables = ds1.Tables
            da1 = New SqlDataAdapter("SELECT mobileno from ticket", myConnToAccess)
            '    da = New OleDbDataAdapter("SELECT ticketnmbr from ticket", myConnToAccess)
            da1.Fill(ds1, "ticket")
            Dim view1 As New DataView(tables(0))
            With txtmobile
                .DataSource = ds.Tables("ticket")
                .DisplayMember = "mobileno"
                .ValueMember = "mobileno"
                .SelectedIndex = -1
                '  .SelectedIndex = 0 this give error when no entry in form occur to solve this error i use aove line 0 is replace by -1

                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .AutoCompleteSource = AutoCompleteSource.ListItems
            End With
        Catch ex As Exception
            MessageBox.Show("Errpr while loading combobox2", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.Dispose()
        End Try
    End Sub
    Private Sub FillCombo2()
        Try
            ' Dim myConnToAccess As OleDbConnection for access
            Dim myConnToAccess As SqlConnection
            Dim ds2 As DataSet
            ' Dim da As OleDbDataAdapter
            Dim da2 As SqlDataAdapter
            Dim tables As DataTableCollection
            ' myConnToAccess = New OleDbConnection("provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb")
            myConnToAccess = New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
            myConnToAccess.Open()
            ds2 = New DataSet
            tables = ds2.Tables
            da2 = New SqlDataAdapter("SELECT pasprtnmbr from ticket", myConnToAccess)
            '    da = New OleDbDataAdapter("SELECT ticketnmbr from ticket", myConnToAccess)
            da2.Fill(ds2, "ticket")
            Dim view1 As New DataView(tables(0))
            With txtpasprt
                .DataSource = ds.Tables("ticket")
                .DisplayMember = "pasprtnmbr"
                .ValueMember = "pasprtnmbr"
                .SelectedIndex = -1
                '  .SelectedIndex = 0 this give error when no entry in form occur to solve this error i use aove line 0 is replace by -1

                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .AutoCompleteSource = AutoCompleteSource.ListItems
            End With
        Catch ex As Exception
            MessageBox.Show("Errpr while loading combobox3", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.Dispose()
        End Try
    End Sub

    Private Sub gridviewfill()
        Try
            'provider = "provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb"   'fro access
            provider = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            'for access da = New OleDbDataAdapter("Select [ticketnmbr], [nme], [sector] ,[dte],[pnr],[issuedte], [expiredte], [mobileno] ,[flight],[pasprtnmbr], [comissin], [basicfair], [deprttme] ,[arrtme],[totalamount],[status]  from ticket", myConnection)
            da = New SqlDataAdapter("Select ticketnmbr, nme , sector as [Sector] ,dte as [Date],pnr as [PNR],issuedte as [Issue Date], expiredte as [Expire Date], mobileno,flight as [Flight],pasprtnmbr, comissin as [Comission], basicfair as [Basic Fair], deprttme as [Departure Time] ,arrtme as [Arrival Time],totalamount as [Total Amount],status as [Status]from ticket", myConnection)
            da.Fill(ds, "ticket")
            Dim view1 As New DataView(tables(0))
            source1.DataSource = view1
            DataGridView1.DataSource = view1
            DataGridView1.Refresh()
        Catch ex As Exception
            ' MsgBox("not loaded " & ex.Message)
            MessageBox.Show(" Not loaded successfully", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.Dispose()
        End Try

    End Sub
    Private Sub datechk()
        Dim dteissue As Date
        Dim dteex As Date
        dteissue = DateTimePicker2.Value
        dteex = DateTimePicker3.Value
        If dteex < dteissue Then
            MessageBox.Show("Expire date must be greater than isssue date!", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
    Private Sub timeechk()
        Dim arrtime As DateTime = DateTime.Parse("1:42:21 PM")
        Dim deptime As DateTime = DateTime.Parse("1:42:21 PM")
        arrtime = txtarrtime.Value
        deptime = txtdeptme.Value
        If arrtime < deptime Then
            MessageBox.Show("arrival time must be greater than depature time!", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub insert()

        dbaccessconnection()
        con.Open()
        cmd.CommandText = "insert into ticket(ticketnmbr,nme,sector,dte,pnr,issuedte,expiredte,mobileno,flight,pasprtnmbr,comissin,basicfair,deprttme,arrtme,totalamount,status)values('" & txtticket.Text & "','" & txtnme.Text & "','" & txtsectr.Text & "','" & DateTimePicker1.Value & "','" & txtpnr.Text & "','" & DateTimePicker2.Value & "','" & DateTimePicker3.Value & "','" & txtmobile.Text & "','" & txtflight.Text & "','" & txtpasprt.Text & "','" & txtcommissin.Text & "','" & txtbasicfair.Text & "','" & txtdeptme.Value & "','" & txtarrtime.Value & "','" & txttotal.Text & "','" & txtstatus.Text & "')"
        cmd.ExecuteReader()
        con.Close()

    End Sub
    Private Sub del()
        dbaccessconnection()
        con.Open()
        cmd.CommandText = "delete from ticket where ticketnmbr=" & txtticket.Text & ""
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
    Private Sub edit()
        dbaccessconnection()
        con.Open()
        cmd.CommandText = ("UPDATE ticket SET nme='" & txtnme.Text & "',  sector= '" & txtsectr.Text & "',dte= '" & DateTimePicker1.Value & "',pnr = '" & txtpnr.Text & "',issuedte = '" & DateTimePicker2.Value & "',expiredte = '" & DateTimePicker3.Value & "',mobileno = '" & txtmobile.Text & "',flight = '" & txtflight.Text & "',pasprtnmbr = '" & txtpasprt.Text & "',comissin = '" & txtcommissin.Text & "',basicfair= '" & txtbasicfair.Text & "',deprttme= '" & txtdeptme.Value & "',arrtme= '" & txtarrtime.Value & "',totalamount= '" & txttotal.Text & "',status= '" & txtstatus.Text & "'  where ticketnmbr = " & txtticket.Text & "")
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
    Private Sub dataalredy()
        dbaccessconnection()
        con.Open()
        Dim ct As String = "select ticketnmbr from ticket where ticketnmbr='" & txtticket.Text & "'"
        cmd = New SqlCommand(ct)
        cmd.Connection = con
        rdr = cmd.ExecuteReader()
        If rdr.Read Then
            MessageBox.Show("hostel name already exists", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtticket.Text = ""
            txtticket.Focus()
            If Not rdr Is Nothing Then
                rdr.Close()
            End If

            Exit Sub
        End If
    End Sub

    Private Sub Btnsve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btnsve.Click


        If Len(Trim(txtticket.Text)) = 0 Then
            MessageBox.Show("Please enter ticket number", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtticket.Focus()
            Exit Sub
        End If
        If Len(Trim(txtnme.Text)) = 0 Then
            MessageBox.Show("Please enter name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtnme.Focus()
            Exit Sub
        End If
        If Len(Trim(txtsectr.Text)) = 0 Then
            MessageBox.Show("Please enter sector.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtsectr.Focus()
            Exit Sub
        End If
        Try
            MessageBox.Show("Are you sure to add data", "Data Adding", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            insert()
            getdata()
            FillCombo()
            FillCombo1()
            FillCombo2()
            Label25.Text = "'" & txtticket.Text & "' Ticket details saved successfully!"
            Label25.ForeColor = System.Drawing.Color.DarkGreen
            Btnsve.Enabled = False
            Button9.Enabled = True
            GroupBox1.Enabled = False
        Catch ex As Exception
            Label25.Text = "Error while saving '" & txtticket.Text & "' ticket details"
            Label25.ForeColor = System.Drawing.Color.Red
            MsgBox("DataBase not connected due to the reason because " & ex.Message)
            'MessageBox.Show("Data already exist, you again select Ticket Details and Try other entry", "Data Invalid, Application is closing", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Dispose()
        End Try

        clear()

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'timer.Text = Date.Now.ToString(" hh:mm:ss")
        timer.Text = TimeOfDay

    End Sub

    Private Sub Btndel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btndel.Click

        Try
            If Not txtnme.Text = "" Then
                MessageBox.Show("Are you sure to delete data", "Data Deleting", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                del()
                getdata()
                FillCombo()
                FillCombo1()
                FillCombo2()
                Label25.Text = "'" & txtticket.Text & "' Ticket details removed successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen
                GroupBox1.Enabled = False
            Else

                MessageBox.Show("Select rows from grid to remove", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                TabPage1.Visible = True
                TabPage2.Visible = False
                TabPage1.Visible = False
                TabPage2.Visible = True
                TabControl1.SelectedTab = TabPage2
            End If
        Catch ex As Exception
            ' MessageBox.Show("Data is not remove succesfully", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Label25.Text = "Error while removing '" & txtticket.Text & "' ticket details"
            Label25.ForeColor = System.Drawing.Color.Red
            Me.Dispose()
        End Try
        clear()
    End Sub
    Private Sub clear()
        txtticket.Text = ""
        txtarrtime.Text = ""
        txtstatus.Text = ""
        DateTimePicker1.Text = ""
        DateTimePicker2.Text = ""
        DateTimePicker3.Text = ""
        txtbasicfair.Text = ""
        txtcommissin.Text = ""
        txtflight.Text = ""
        txtdeptme.Text = ""
        txtmobile.Text = ""
        txtnme.Text = ""
        txtpasprt.Text = ""
        txtpnr.Text = ""
        txtsectr.Text = ""
        txttotal.Text = ""
    End Sub

    Private Sub btnupdte_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnupdte.Click
        Try
            If Not txtnme.Text = "" Then
                MessageBox.Show("Are you sure to update data", "Data Updating", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                edit()
                getdata()
                FillCombo()
                FillCombo1()
                FillCombo2()
                Label25.Enabled = True
                Label25.Text = "'" & txtticket.Text & "' Ticket details updated successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen
                GroupBox1.Enabled = False
                Me.Refresh()
            Else

                MessageBox.Show("Select rows from grid to edit", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                TabPage1.Visible = True
                TabPage2.Visible = False
                TabPage1.Visible = False
                TabPage2.Visible = True
                TabControl1.SelectedTab = TabPage2
            End If
        Catch ex As Exception
            ' MessageBox.Show("Data not updated successfully", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Label25.Text = "Error while updating '" & txtticket.Text & "' ticket details"
            Label25.ForeColor = System.Drawing.Color.Red
            Me.Dispose()
        End Try
        Me.Refresh()
        clear()

    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try

            If Not txtnme.Text = "" Then
                MessageBox.Show("Are you sure to print data", "Data Adding", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                Dim report As New CRticket
                Dim objText As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(1).ReportObjects("Text3")
                Dim objText1 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text7")
                Dim objText2 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text8")
                Dim objText3 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text9")
                Dim objText4 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text10")
                Dim objText5 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text11")
                Dim objText6 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text12")
                Dim objText7 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text13")
                Dim objText8 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text14")
                Dim objText9 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text15")
                Dim objText10 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text17")
                Dim objText11 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text18")
                Dim objText12 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text19")
                Dim objText13 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text20")
                Dim objText14 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text21")
                Dim objText15 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text2")
                objText.Text = Me.txtticket.Text
                objText1.Text = Me.txtnme.Text
                objText2.Text = Me.txtsectr.Text
                objText3.Text = Me.DateTimePicker1.Text
                objText4.Text = Me.txtpnr.Text
                objText5.Text = Me.DateTimePicker2.Text
                objText6.Text = Me.DateTimePicker3.Text
                objText7.Text = Me.txtmobile.Text
                objText8.Text = Me.txtpasprt.Text
                objText9.Text = Me.txtflight.Text
                objText10.Text = Me.txtbasicfair.Text
                objText11.Text = Me.txtdeptme.Text
                objText12.Text = Me.txtarrtime.Text
                objText13.Text = Me.txttotal.Text
                objText14.Text = Me.txtstatus.Text
                objText15.Text = Me.txtcommissin.Text
                Rprtticket.CrystalReportViewer1.ReportSource = report
                Rprtticket.Show()
                Label25.Text = "'" & txtticket.Text & "' Ticket details printed successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen
            Else
                MessageBox.Show("Select value from gridview to print", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                TabPage1.Visible = True
                TabPage2.Visible = False
                TabPage1.Visible = False
                TabPage2.Visible = True
                TabControl1.SelectedTab = TabPage2
            End If
        Catch ex As Exception
            ' MessageBox.Show("Reports are not loding properly,try again", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Label25.Text = "Error while printing'" & txtticket.Text & "' ticket details"
            Label25.ForeColor = System.Drawing.Color.Red
            Me.Dispose()
        End Try

    End Sub
    Private Sub btnsearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsearch.Click
        Try
            source1.Filter = "[nme] = '" & ComboBox1.Text & "'"
            source2.Filter = "[nme] = '" & ComboBox1.Text & "'"
            DataGridView1.Refresh()
            ComboBox1.Text = ""
        Catch ex As Exception
            MessageBox.Show("Error while searching,try again", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.Dispose()
        End Try
    End Sub
    '////////////////////not used now////////////
    Private Sub done_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' Me.Refresh()
        ' Tfrmload()
        If txtnme.Text = "" Then
            MessageBox.Show("Are you sure to save changes", "Data Saving", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            Panel2.Enabled = False
            Panel2.Enabled = False
            Dim oForm As Ticket
            oForm = New Ticket()
            oForm.ShowDialog()
            oForm = Nothing
        ElseIf Panel2.Enabled = True Then
            txtticket.Enabled = False
        Else
            MessageBox.Show("Not Choose Any Operation?????", "Changes not saved properly", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        clear()
    End Sub
    '/////////////////////////
    Private Sub Btnadd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btnadd.Click
        Label25.Text = " Add new record!"
        Label25.ForeColor = System.Drawing.Color.DarkCyan
        Try

            GroupBox1.Enabled = True
            Btnsve.Enabled = True
            Button9.Enabled = True
            Btndel.Enabled = False
            btnupdte.Enabled = False
            Button1.Enabled = False
            
            clear()
            txtboxid()
        Catch ex As Exception
            MessageBox.Show("Something while adding,Close application and try again", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.Dispose()
        End Try
    End Sub
    Private Sub txtboxid()
        Try
            dbaccessconnection()
            con.Open()
            Dim num As New Integer
            cmd.CommandText = "SELECT MAX(ticketnmbr) FROM ticket "
            If (IsDBNull(cmd.ExecuteScalar)) Then
                num = 1
                txtticket.Text = num.ToString
            Else
                'num = cmd.ExecuteScalar()
                'txtticket.Text = num + 1
                num = cmd.ExecuteScalar + 1
                txtticket.Text = num.ToString
            End If
            con.Close()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
            Me.Dispose()
        End Try
    End Sub
    '///////////////////////not call anywhere but useful////////////////////
    Private Sub idd()
        '///////////////////for maximu value in textbox ////////////////////
        con.Open()
        Dim num As New Integer
        cmd.CommandText = "SELECT MAX(ticketnmbr) FROM ticket "
        num = cmd.ExecuteScalar()
        txtticket.Text = num + 1
        con.Close()
        '///////////////for maximu value in gridview////////////////
        Dim MaxVal As Double = 0
        For Each row As DataGridViewRow In DataGridView1.Rows

            If row.Cells(0).Value > MaxVal Then MaxVal = row.Cells(0).Value 'Maximum value of first column
        Next
        txtticket.Text = MaxVal
    End Sub
    '///////////////////////not call anywhere but useful////////////////////


    Private Sub txtticket_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtticket.KeyPress
        'If IsNumeric(txtticket.Text + e.KeyChar) = False Then e.Handled = True
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub


    Private Sub txtticket_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtticket.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtticket, "Enter Ticket ID in numbers")
    End Sub

    Private Sub txtnme_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtnme.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtnme, "Enter name in letters, numbers ,.,#")
    End Sub

    Private Sub txtsectr_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtsectr.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtsectr, "Enter sector in letters, numbers ,.,#")
    End Sub

    Private Sub DateTimePicker1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DateTimePicker1.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(DateTimePicker1, "Select date from calender")
    End Sub

    Private Sub txtpnr_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtpnr.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtpnr, "enter pnr in letters ,digits,.,#")
    End Sub

    Private Sub DateTimePicker2_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker2.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(DateTimePicker2, "Select date from calender")
    End Sub
    Private Sub DateTimePicker3_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker3.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(DateTimePicker3, "Select date from calender")
    End Sub

    Private Sub txtmobile_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtmobile, "Enter number only in digits")

    End Sub

    Private Sub txtflight_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtflight.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtflight, "enter Flight either in letters ,digits,.,#")
    End Sub

    Private Sub txtpasprt_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtpasprt, "Enter passport number either in letters ,digits,.,#")
    End Sub

    Private Sub txtcommissin_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtcommissin.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtcommissin, "Enter Comission  in digits")
    End Sub

    Private Sub txtbasicfair_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtbasicfair.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtbasicfair, "Enter Basic Fair  in digits")
    End Sub

    Private Sub txtdeptme_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtdeptme.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtdeptme, "Enter Departure time in 00:00AM or PM")
    End Sub
    Private Sub txtarrtime_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtarrtime.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtarrtime, "Enter Arrival time in like 00:00AM or PM")
    End Sub

    Private Sub txttotal_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txttotal.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txttotal, "Enter Total Amount in format like 00Rs")
    End Sub

    Private Sub txtstatus_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtstatus.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtstatus, "Select one of menu")
    End Sub
    '/////////////////clear button not used/////////////////
    Private Sub Btnclear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        clear()
    End Sub
    Private Sub DataGridView1_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        Try
            txtticket.Enabled = False
            Panel2.Enabled = True
            GroupBox1.Enabled = True
            Btnsve.Enabled = False
            Btndel.Enabled = True
            btnupdte.Enabled = True
            Button1.Enabled = True
            Me.txtticket.Text = DataGridView1.CurrentRow.Cells(0).Value.ToString
            Me.txtnme.Text = DataGridView1.CurrentRow.Cells(1).Value.ToString
            Me.txtsectr.Text = DataGridView1.CurrentRow.Cells(2).Value.ToString
            Me.DateTimePicker1.Text = DataGridView1.CurrentRow.Cells(3).Value.ToString
            Me.txtpnr.Text = DataGridView1.CurrentRow.Cells(4).Value.ToString
            Me.DateTimePicker2.Text = DataGridView1.CurrentRow.Cells(5).Value.ToString
            Me.DateTimePicker3.Text = DataGridView1.CurrentRow.Cells(6).Value.ToString
            Me.txtmobile.Text = DataGridView1.CurrentRow.Cells(7).Value.ToString
            Me.txtflight.Text = DataGridView1.CurrentRow.Cells(8).Value.ToString
            Me.txtpasprt.Text = DataGridView1.CurrentRow.Cells(9).Value.ToString
            Me.txtcommissin.Text = DataGridView1.CurrentRow.Cells(10).Value.ToString
            Me.txtbasicfair.Text = DataGridView1.CurrentRow.Cells(11).Value.ToString
            Me.txtdeptme.Text = DataGridView1.CurrentRow.Cells(12).Value.ToString
            Me.txtarrtime.Text = DataGridView1.CurrentRow.Cells(13).Value.ToString
            Me.txttotal.Text = DataGridView1.CurrentRow.Cells(14).Value.ToString
            Me.txtstatus.Text = DataGridView1.CurrentRow.Cells(15).Value.ToString
            Me.txtstatus.Text = DataGridView1.CurrentRow.Cells(15).Value.ToString
            editticket.txtstatus.Text = DataGridView1.CurrentRow.Cells(15).Value.ToString
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Dispose()
        End Try
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If MsgBox("Are you sure want to exit now?", MsgBoxStyle.YesNo, "Closing warning") = MsgBoxResult.Yes Then ' If you select yes in the MsgBox then it will close the window
            Me.Close() ' Close the window
            Me.Dispose()
        Else
            ' Will not close the application
        End If
    End Sub

    Private Sub txtmobile_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtcommissin_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtcommissin.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtbasicfair_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtbasicfair.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txttotal_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txttotal.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub DateTimePicker3_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker3.Validated
        datechk()
    End Sub
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call CenterToScreen()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.WindowState = FormWindowState.Minimized
    End Sub
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        txtticket.Text = Val(txtticket.Text) + 1
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        txtticket.Text = Val(txtticket.Text) - 1
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        If Len(Trim(txtticket.Text)) = 0 Then
            MessageBox.Show("Please enter ticket number", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtticket.Focus()
            Exit Sub
        End If
        If Len(Trim(txtnme.Text)) = 0 Then
            MessageBox.Show("Please enter name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtnme.Focus()
            Exit Sub
        End If
        If Len(Trim(txtsectr.Text)) = 0 Then
            MessageBox.Show("Please enter sector.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtsectr.Focus()
            Exit Sub
        End If
        Try
            MessageBox.Show("Are you sure to add data", "Data Adding", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            insert()
            Label25.Text = "'" & txtticket.Text & "' Ticket details saved successfully!"
            Label25.ForeColor = System.Drawing.Color.DarkGreen
            Button9.Enabled = False
            GroupBox1.Enabled = False
        Catch ex As Exception
            Label25.Text = "Error while saving '" & txtticket.Text & "' ticket details"
            Label25.ForeColor = System.Drawing.Color.Red
            MessageBox.Show("Data already exist, you again select Ticket Details and Try other entry", "Data Invalid, Application is closing", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Dispose()
        End Try
        clear()
        Me.Refresh()
    End Sub

    Private Sub listboxfill()
        Try
            Dim SqlSb As New SqlConnectionStringBuilder()
            SqlSb.DataSource = "MEERHAMZA"
            SqlSb.InitialCatalog = "airlinee"
            SqlSb.IntegratedSecurity = True
            Using SqlConn As SqlConnection = New SqlConnection(SqlSb.ConnectionString)
                SqlConn.Open()

                Dim cmd As SqlCommand = SqlConn.CreateCommand()
                cmd.CommandText = "SELECT [nme],[ticketnmbr], [sector] ,[dte],[pnr],[issuedte], [expiredte], [mobileno] ,[flight],[pasprtnmbr], [deprttme] ,[arrtme],[totalamount],[status]  FROM [ticket]"
                Using reader As SqlDataReader = cmd.ExecuteReader
                    While (reader.Read())
                        'for access da = New OleDbDataAdapter("Select [ticketnmbr], [nme], [sector] ,[dte],[pnr],[issuedte], [expiredte], [mobileno] ,[flight],[pasprtnmbr], [comissin], [basicfair], [deprttme] ,[arrtme],[totalamount],[status]  from ticket", myConnection)
                        Me.ListBox1.Items.Add("  ..Ticket Number..  ")
                        Me.ListBox1.Items.Add(reader("ticketnmbr"))
                        Me.ListBox1.Items.Add("  ..Name..  ")
                        Me.ListBox1.Items.Add(reader("nme"))
                        Me.ListBox1.Items.Add("  ..Sector..  ")
                        Me.ListBox1.Items.Add(reader("sector"))
                        Me.ListBox1.Items.Add("  ..Date..  ")
                        Me.ListBox1.Items.Add(reader("dte"))
                        Me.ListBox1.Items.Add("  ..Pnr..  ")
                        Me.ListBox1.Items.Add(reader("pnr"))
                        Me.ListBox1.Items.Add("  ..Issue Date..  ")
                        Me.ListBox1.Items.Add(reader("issuedte"))
                        Me.ListBox1.Items.Add("  ..Expire Date..  ")
                        Me.ListBox1.Items.Add(reader("expiredte"))
                        Me.ListBox1.Items.Add("  ..Mobile Number..  ")
                        Me.ListBox1.Items.Add(reader("mobileno"))
                        Me.ListBox1.Items.Add("  ..Flight..  ")
                        Me.ListBox1.Items.Add(reader("flight"))
                        Me.ListBox1.Items.Add("  ..Passport Number..  ")
                        Me.ListBox1.Items.Add(reader("pasprtnmbr"))
                        Me.ListBox1.Items.Add("  ..Depature Time..  ")
                        Me.ListBox1.Items.Add(reader("deprttme"))
                        Me.ListBox1.Items.Add("  ..Arrival Time..  ")
                        Me.ListBox1.Items.Add(reader("arrtme"))
                        Me.ListBox1.Items.Add("  ..Total Amount..  ")
                        Me.ListBox1.Items.Add(reader("totalamount"))
                        Me.ListBox1.Items.Add("  ..Status..  ")
                        Me.ListBox1.Items.Add(reader("status"))
                        Me.ListBox1.Items.Add("```````````````")
                        Me.ListBox1.Items.Add("               ")

                    End While
                End Using
                SqlConn.Close()
            End Using
        Catch ex As Exception
            ' MsgBox("Not Loading", MsgBoxStyle.OkOnly, "Error", & ex.Message)
            MsgBox("DataBase not connected due to the reason because " & ex.Message)
        End Try
    End Sub
    Private Sub listboxfill2()
        Try
            Dim SqlSb As New SqlConnectionStringBuilder()
            SqlSb.DataSource = "MEERHAMZA"
            SqlSb.InitialCatalog = "airlinee"
            SqlSb.IntegratedSecurity = True
            Using SqlConn As SqlConnection = New SqlConnection(SqlSb.ConnectionString)
                SqlConn.Open()

                Dim cmd As SqlCommand = SqlConn.CreateCommand()
                cmd.CommandText = "SELECT [nme] FROM [ticket]"
                Using reader As SqlDataReader = cmd.ExecuteReader
                    While (reader.Read())
                        'for access da = New OleDbDataAdapter("Select [ticketnmbr], [nme], [sector] ,[dte],[pnr],[issuedte], [expiredte], [mobileno] ,[flight],[pasprtnmbr], [comissin], [basicfair], [deprttme] ,[arrtme],[totalamount],[status]  from ticket", myConnection)



                        Me.ListBox1.Items.Add(reader("nme"))

                    End While
                End Using
                SqlConn.Close()
            End Using
        Catch ex As Exception
            ' MsgBox("Not Loading", MsgBoxStyle.OkOnly, "Error", & ex.Message)
            MsgBox("DataBase of listbox2 is not connected due to the reason because " & ex.Message)
        End Try
    End Sub
    Private Sub loaddata_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles loaddata.Click
        ListBox1.Items.Clear()
        listboxfill2()
        listview()
    End Sub
    Private Sub profit()
        If Not txtbasicfair.Text = "" Then
            Dim minus As Integer
            Dim a As Integer
            Dim b As Integer
            a = txttotal.Text
            b = txtbasicfair.Text
            minus = b - a
            txtcommissin.Text = minus
            'MessageBox.Show(minus)
        Else
            MsgBox("Orignal cost and sale must not be empty ")

        End If
    End Sub



    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        TabPage1.Visible = True
        Label25.Text = ""
        TabControl1.SelectedTab = TabPage1

    End Sub
    Private Sub getdata()
        Dim con As New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
        con.Open()
        Dim da As New SqlDataAdapter("Select ticketnmbr , nme , sector as [Sector] ,dte as [Date],pnr as [PNR],issuedte as [Issue Date], expiredte as [Expire Date], mobileno,flight as [Flight],pasprtnmbr, comissin as [Comission], basicfair as [Basic Fair], deprttme as [Departure Time] ,arrtme as [Arrival Time],totalamount as [Total Amount],status as [Status]from ticket", con)
        Dim dt As New DataTable
        da.Fill(dt)
        'Dim view1 As New DataView(tables1(0))
        source2.DataSource = dt
        DataGridView1.DataSource = dt
        DataGridView1.Refresh()
    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        e.Graphics.DrawImage(bitmap, 0, 0)
    End Sub
    Private Sub printgrid()
        Dim height As Integer = DataGridView1.Height
        DataGridView1.Height = DataGridView1.RowCount * DataGridView1.RowTemplate.Height
        bitmap = New Bitmap(Me.DataGridView1.Width, Me.DataGridView1.Height)
        DataGridView1.DrawToBitmap(bitmap, New Rectangle(0, 0, Me.DataGridView1.Width, Me.DataGridView1.Height))

        'Resize DataGridView back to original height.
        DataGridView1.Height = height

        'Show the Print Preview Dialog.
        PrintPreviewDialog1.Document = PrintDocument1

        PrintPreviewDialog1.PrintPreviewControl.Zoom = 1


        'PrintDocument1.PrinterSettings.DefaultPageSettings.Landscape = True
        PrintDocument1.DefaultPageSettings.Landscape = True
        PrintPreviewDialog1.ShowDialog()


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged()
        ' Dim constr As String = ("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
        'Using conn As New SqlConnection(constr)
        'Using cmd As New SqlCommand("SELECT  [ticketnmbr], [nme], [sector] ,[dte],[pnr],[issuedte], [expiredte], [mobileno] ,[flight],[pasprtnmbr], [comissin], [basicfair], [deprttme] ,[arrtme],[totalamount],[status]   FROM ticket WHERE nme   = @nme  OR @nme  = ''", conn)
        'Using da As New SqlDataAdapter(cmd)
        'cmd.Parameters.AddWithValue("@nme", Me.ComboBox1.SelectedValue)
        'Dim dt As New DataTable()
        ' da.Fill(dt)
        ' Me.DataGridView1.DataSource = dt
        ' End Using
        '. End Using
        ' End Using
    End Sub

    Private Sub Button5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try

            If Not DataGridView1.CurrentRow.IsNewRow Then
                'Query string
                dbaccessconnection()
                con.Open()
                cmd.CommandText = "delete  from ticket  where ticketnmbr='" & DataGridView1.CurrentRow.Cells(0).Value & "'"
                cmd.ExecuteNonQuery()
                DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
                MessageBox.Show("Record Deleted")
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Private Sub DeleteSelecedRows()
        Dim ObjConnection As New SqlConnection()
        Dim i As Integer
        Dim mResult
        mResult = MsgBox("Want you really delete the selected records?", _
        vbYesNo + vbQuestion, "Removal confirmation")
        If mResult = vbNo Then
            Exit Sub
        End If
        ObjConnection.ConnectionString = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
        Dim ObjCommand As New SqlCommand()
        ObjCommand.Connection = ObjConnection
        For i = Me.DataGridView1.SelectedRows.Count - 1 To 0 Step -1
            ObjCommand.CommandText = "delete from ticket where ticketnmbr='" & DataGridView1.SelectedRows(i).Cells("ticketnmbr").Value & "'"
            ObjConnection.Open()
            ObjCommand.ExecuteNonQuery()
            ObjConnection.Close()

            Me.DataGridView1.Rows.Remove(Me.DataGridView1.SelectedRows(i))
        Next

    End Sub

    Private Sub txtcommissin_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtcommissin.TextChanged
        '  txtcommissin.Text = Val(txttotal.Text) + Val(txtbasicfair.Text)
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        If MsgBox("Are you sure want to exit now?", MsgBoxStyle.YesNo, "Closing warning") = MsgBoxResult.Yes Then ' If you select yes in the MsgBox then it will close the window
            Me.Close()
            Me.Dispose() ' Close the window
        Else
            ' Will not close the application
        End If
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call CenterToScreen()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call CenterToScreen()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Call CenterToScreen()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub Label25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label25.Click

    End Sub

    Private Sub Label17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label17.Click

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        source1.Filter = "[nme] = '" & ComboBox1.Text & "'"
        DataGridView1.Refresh()
        ComboBox1.Text = ""
    End Sub

    Private Sub excel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles excel.Click
        If DataGridView1.RowCount = Nothing Then
            MessageBox.Show("Sorry nothing to export into excel sheet.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application
        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True

            rowsTotal = DataGridView1.RowCount
            colsTotal = DataGridView1.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView1.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView1.Rows(I).Cells(j).Value.ToString()
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 12

                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'RELEASE ALLOACTED RESOURCES
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub


    Private Sub txtbasicfair_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtbasicfair.Validating
        profit()
    End Sub


    Private Sub Button2_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ListBox1.Items.Clear()
        For Each str As String In ListBox1.SelectedItems
            ListBox1.Items.Add(str)
        Next str
        ListBox1.Items.Clear()
        For x As Integer = 0 To 2
            ListBox1.Items.Add(ListBox1.Items(x).ToString)
        Next x
    End Sub

    Private Sub Button6_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        DeleteSelecedRows()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        '///////////// For ListBox to richtextbox
        ' Label28.Text = ListBox1.SelectedItem.ToString()
        ' RichTextBox1.Clear()
        ' For Each Item As Object In ListBox1.SelectedItems
        'RichTextBox1.AppendText(Item.ToString + Environment.NewLine)
        'Next
        ' Me.ListBox1.SetSelected(2, True)
    End Sub

    Private Sub Button14_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        searchlstview()
        movoneitm()
        Label27.Text = "Search Values are: " & ListView2.Items.Count
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

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        ListView1.Items.Clear()
        ListView2.Items.Clear()
        ListView1.SelectedIndices.Clear()
        ListBox1.Items.Clear()
        TextBox1.Text = ""

    End Sub

    Private Sub TextBox1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox1.MouseClick
        TextBox1.Text = ""
    End Sub

End Class

Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.Data.DataTable
Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel '
Imports ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat
Public Class frmexpensis
    Private bitmap As Bitmap
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
    Dim source2 As New BindingSource()
    Dim dt As New DataTable
    Dim con As New SqlClient.SqlConnection
    Dim cmd As New SqlClient.SqlCommand
    '/////////////////////////////////////////////
    ' Private Sub dbaccessconnection()
    'Acces DataBase Connectivity and for MS Access 2003 PROVIDER=Microsoft.Jet.OLEDB.4.0
    '  Try
    '  con.ConnectionString = "provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb"
    ' cmd.Connection = con
    ' MessageBox.Show("connection created")
    'Catch ex As Exception
    '  MsgBox("DataBase not connected due to the reason because " & ex.Message)
    ' End Try
    'End Sub
    '////////////////////////////////////
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

    Private Sub frmexpensis_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If (DateTime.Now.Hour < 12) Then
            PictureBox5.Visible = False
            PictureBox18.Visible = False
            PictureBox19.Visible = True
            PictureBox17.Visible = False
            lblgrting.Text = "Good Morning"
            'Label12.Text = Convert.ToString(DateTime.Now)

        ElseIf (DateTime.Now.Hour < 17) Then
            PictureBox5.Visible = False
            PictureBox18.Visible = True
            PictureBox19.Visible = False
            PictureBox17.Visible = False
            lblgrting.Text = "Good Afternoon"
            ' Label12.Text = Convert.ToString(DateTime.Now)
        ElseIf (DateTime.Now.Hour < 19) Then
            PictureBox17.Visible = False
            PictureBox18.Visible = False
            PictureBox19.Visible = False
            PictureBox5.Visible = True
            lblgrting.Text = "Good Evening"
        Else
            PictureBox17.Visible = True
            PictureBox18.Visible = False
            PictureBox19.Visible = False
            PictureBox5.Visible = False
            lblgrting.Text = "Good Night"
            ' Label12.Text = Convert.ToString(DateTime.Now)
        End If
        Label8.Text = Date.Today.ToString("dddd")
        'Call CenterToScreen()
        'Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        'Me.WindowState = FormWindowState.Maximized
        gridfill()
        dbaccessconnection()
        GroupBox1.Enabled = False
        FillCombo()
        FillCombo2()

        ComboBox1.Text = "----Select Or Type Name---"
        ' DataGridView1.Sort(DataGridView1.Columns("entry"), System.ComponentModel.ListSortDirection.Ascending)
        Me.Label23.Text = Format(Now, "dd-MMM-yyyy")
        Timer1.Enabled = True
        Timer1.Start()
    End Sub
    Private Sub FillCombo()
        Try
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
        Catch ex As Exception
            MsgBox("Search have problem!!!!")
        End Try
    End Sub
    Private Sub FillCombo2()
        Try
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
            With txtexpnme
                .DataSource = ds.Tables("expensis")
                .DisplayMember = "nme"
                .ValueMember = "nme"
                .SelectedIndex = -1
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .AutoCompleteSource = AutoCompleteSource.ListItems
            End With
        Catch ex As Exception
            MsgBox("Search have problem!!!!")
        End Try
    End Sub
    Private Sub gridfill()
        Try
            'dbaccessconnection()
            'provider = "provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb"
            provider = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            'da = New OleDbDataAdapter("Select [entry], [nme], [dte], [salry] ,[bill],[pckgecash],[prsnluse], [rent] from expensis", myConnection)
            da = New SqlDataAdapter("Select [entry], [nme], [dte], [salry] ,[bill],[pckgecash],[prsnluse], [rent] from expensis", myConnection)
            da.Fill(ds, "expensis")
            Dim view1 As New DataView(tables(0))
            source1.DataSource = view1
            DataGridView1.DataSource = view1
            DataGridView1.Refresh()
        Catch ex As Exception
            MsgBox("DataBase not connected due to the reason because " & ex.Message)
            MessageBox.Show(" Not loaded successfully", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub
    Private Sub getdata()

        Dim con As New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
        con.Open()
        Dim da As New SqlDataAdapter("Select [entry], [nme], [dte], [salry] ,[bill],[pckgecash],[prsnluse], [rent] from expensis", con)
        Dim dt As New DataTable
        da.Fill(dt)
        'Dim view1 As New DataView(tables1(0))
        source2.DataSource = dt
        DataGridView1.DataSource = dt
        DataGridView1.Refresh()
    End Sub
    Private Sub insert()
        dbaccessconnection()
        con.Open()
        cmd.CommandText = "insert into expensis(entry,nme,dte,salry,bill,pckgecash,prsnluse,rent)values('" & txtexentry.Text & "','" & txtexpnme.Text & "','" & txtexpdte.Value & "','" & txtexpsalry.Text & "','" & txtexpbl.Text & "','" & txtexpacge.Text & "','" & txtexpprsnl.Text & "','" & txteprnt.Text & "')"
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
    Private Sub del()
        dbaccessconnection()
        con.Open()
        cmd.CommandText = "delete from expensis where entry=" & txtexentry.Text & ""
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
    Private Sub edit()
        dbaccessconnection()
        con.Open()
        cmd.CommandText = ("UPDATE expensis SET  nme = '" & txtexpnme.Text & "',dte='" & txtexpdte.Value & "', salry= '" & txtexpsalry.Text & "',bill= '" & txtexpbl.Text & "',pckgecash= '" & txtexpacge.Text & "',prsnluse= '" & txtexpprsnl.Text & "',rent= '" & txteprnt.Text & "'  where entry=" & txtexentry.Text & "")
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub

    Private Sub Btnsve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btnsve.Click
        If Len(Trim(txtexentry.Text)) = 0 Then
            MessageBox.Show("Please enter Entry number", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtexentry.Focus()
            Exit Sub
        End If
        If Len(Trim(txtexpnme.Text)) = 0 Then
            MessageBox.Show("Please enter Name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtexpnme.Focus()
            Exit Sub
        End If
        Me.Refresh()
        Try
            If Not txtexentry.Text = "" Then
                MessageBox.Show("Are you sure to add data", "Data Adding", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                insert()
                getdata()
                FillCombo()
                Label25.Text = "'" & txtexentry.Text & "' Expensis details saved successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen
                Btnsve.Enabled = False
                Btnadd.Enabled = True
                Btndel.Enabled = True
                btnupdte.Enabled = True
                Button2.Enabled = True
                Button9.Enabled = True
                Panel2.Enabled = False

            Else

                MessageBox.Show("please fill all above textboxes", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Label25.Text = "Problem while saving '" & txtexentry.Text & "' Expensis details"
                Label25.ForeColor = System.Drawing.Color.Red
                
            End If
        Catch ex As Exception
            Label25.Text = "Error while saving '" & txtexentry.Text & "' Expensis details"
            Label25.ForeColor = System.Drawing.Color.Red
            MessageBox.Show("Data already exist, you again select Expenesis Details and Try other entry", "Data Invalid, Application is closing", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Dispose()
        End Try
        clear()
    End Sub

    Private Sub Btndel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btndel.Click
        gridclick()
        Try
            If Not txtexentry.Text = "" Then
                MessageBox.Show("Are you sure to delete data", "Data Deleting", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                del()
                FillCombo()
                getdata()
                Label25.Text = "'" & txtexentry.Text & "'Expensis details removed successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen
                Panel2.Enabled = False
                Btnadd.Enabled = True

            Else
                MessageBox.Show("Select rows from grid to remove", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Label25.Text = "Error while removing '" & txtexentry.Text & "' Expensis details"
                Label25.ForeColor = System.Drawing.Color.Red

                GroupBox1.Visible = False
                GroupBox1.Enabled = False
                GroupBox5.Visible = False
                GroupBox5.Enabled = False
                GroupBox3.Visible = True
                GroupBox3.Enabled = True
                Panel1.Visible = False
                Panel1.Enabled = False
            End If
        Catch ex As Exception
            'MessageBox.Show("Data is not remove succesfully", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Label25.Text = "Error while removing '" & txtexentry.Text & "'Expensis details"
            Label25.ForeColor = System.Drawing.Color.Red
            Me.Dispose()
        End Try
        Me.Refresh()
        clear()
    End Sub

    Private Sub btnupdte_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnupdte.Click
        Try
            If Not txtexentry.Text = "" Then
                MessageBox.Show("Are you sure to update data", "Data Updating", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                edit()
                FillCombo()
                getdata()

                Label25.Enabled = True
                Label25.Text = "'" & txtexentry.Text & "' Expensis details updated successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen
                Panel2.Enabled = False
                Btnadd.Enabled = True
                Me.Refresh()
            Else
                MessageBox.Show("Select rows from grid to edit", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Label25.Text = "Error while updating '" & txtexentry.Text & "' Expensis details"
                Label25.ForeColor = System.Drawing.Color.Red

                GroupBox1.Visible = False
                GroupBox1.Enabled = False
                GroupBox5.Visible = False
                GroupBox5.Enabled = False
                GroupBox3.Visible = True
                GroupBox3.Enabled = True
                Panel1.Visible = False
                Panel1.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show("Data not updated successfully", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Label25.Text = "Error while updating '" & txtexentry.Text & "' Expensis details"
            Label25.ForeColor = System.Drawing.Color.Red
            Me.Dispose()
        End Try
        clear()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        timer.Text = TimeOfDay
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        GroupBox1.Visible = True
        GroupBox1.Enabled = True
        Panel1.Visible = True
        Panel1.Enabled = True
        GroupBox5.Visible = False
        GroupBox5.Enabled = False
        GroupBox3.Visible = False
        GroupBox3.Enabled = False
        gridclick()
    End Sub
    Private Sub gridclick()
        Try
            Label25.Text = "Edit Values in above textboxes"
            GroupBox1.Enabled = True
            txtexentry.Enabled = False
            Btnsve.Enabled = False
            Btndel.Enabled = True
            btnupdte.Enabled = True
            Me.txtexentry.Text = DataGridView1.CurrentRow.Cells(0).Value.ToString
            Me.txtexpnme.Text = DataGridView1.CurrentRow.Cells(1).Value.ToString
            Me.txtexpdte.Text = DataGridView1.CurrentRow.Cells(2).Value.ToString
            Me.txtexpsalry.Text = DataGridView1.CurrentRow.Cells(3).Value.ToString
            Me.txtexpbl.Text = DataGridView1.CurrentRow.Cells(4).Value.ToString
            Me.txtexpacge.Text = DataGridView1.CurrentRow.Cells(5).Value.ToString
            Me.txtexpprsnl.Text = DataGridView1.CurrentRow.Cells(6).Value.ToString
            Me.txteprnt.Text = DataGridView1.CurrentRow.Cells(7).Value.ToString
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub clear()
        txtexentry.Text = ""
        txtexpnme.Text = ""
        txtexpdte.Text = ""
        txtexpsalry.Text = ""
        txtexpbl.Text = ""
        txtexpacge.Text = ""
        txtexpprsnl.Text = ""
        txteprnt.Text = ""

    End Sub

    Private Sub Btnadd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btnadd.Click
        Label25.Text = " Add new record!"
        Label25.ForeColor = System.Drawing.Color.DarkCyan
        Try
            Btnsve.Enabled = True
            Button9.Enabled = True
            Btndel.Enabled = False
            btnupdte.Enabled = False
            Button2.Enabled = False
            Panel2.Enabled = True

            clear()
            txtboxid()
        Catch ex As Exception
            MessageBox.Show("Something is going wrong,Close application and try again", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.Dispose()
        End Try
    End Sub
    Private Sub txtboxid()
        Try
            dbaccessconnection()
            con.Open()
            Dim num As New Integer

            cmd.CommandText = "SELECT MAX(entry) FROM expensis"
            If (IsDBNull(cmd.ExecuteScalar)) Then
                num = 1
                txtexentry.Text = num.ToString
            Else
                num = cmd.ExecuteScalar + 1
                txtexentry.Text = num.ToString
            End If
            con.Close()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If Not txtexpnme.Text = "" Then
                MessageBox.Show("Are you sure to print data", "Data Adding", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                Dim report As New CRexpensis
                Dim objText As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(1).ReportObjects("Text1")
                Dim objText1 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text2")
                Dim objText2 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text3")
                Dim objText3 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text4")
                Dim objText4 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text8")
                Dim objText5 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text5")
                Dim objText6 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text6")
                Dim objText7 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text7")

                objText.Text = Me.txtexentry.Text
                objText1.Text = Me.txtexpnme.Text
                objText2.Text = Me.txtexpdte.Text
                objText3.Text = Me.txtexpsalry.Text
                objText4.Text = Me.txtexpbl.Text
                objText5.Text = Me.txteprnt.Text
                objText6.Text = Me.txtexpacge.Text
                objText7.Text = Me.txtexpprsnl.Text
                RprtExpnsis.CrystalReportViewer1.ReportSource = report
                RprtExpnsis.Show()
                Label25.Text = "'" & txtexentry.Text & "' expensis details printed successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen


            Else
                MessageBox.Show("Select value from gridview to print", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Label25.Text = "Problem while Printing '" & txtexentry.Text & "' Expensis details"
                Label25.ForeColor = System.Drawing.Color.Red
                GroupBox1.Visible = False
                GroupBox1.Enabled = False
                GroupBox5.Visible = False
                GroupBox5.Enabled = False
                GroupBox3.Visible = True
                GroupBox3.Enabled = True
                Panel1.Visible = False
                Panel1.Enabled = False

            End If
        Catch ex As Exception
            MessageBox.Show("Reports are not loding properly,try again", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Label25.Text = "Error while printing'" & txtexentry.Text & "' Expensis details"
            Label25.ForeColor = System.Drawing.Color.Red

            Me.Dispose()
        End Try
    End Sub
    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        If MsgBox("Are you sure want to exit now?", MsgBoxStyle.YesNo, "Closing warning") = MsgBoxResult.Yes Then ' If you select yes in the MsgBox then it will close the window
            Me.Close()
            Me.Dispose() ' Close the window
        Else
            ' Will not close the application
        End If
    End Sub
    Private Sub listboxfill()
        Try
            Dim SqlSb As New SqlConnectionStringBuilder()
            SqlSb.DataSource = "MEERHAMZA"
            SqlSb.InitialCatalog = "airlinee"
            SqlSb.IntegratedSecurity = True
            Using SqlConn As SqlConnection = New SqlConnection(SqlSb.ConnectionString)
                SqlConn.Open()
                ' visit( vientry,nme,dte
                ' vientry,nme,dte,psprtno,trvelingdte,expirdte,vipresent,orgnlcst,sale,proft,viduration,refrnce
                Dim cmd As SqlCommand = SqlConn.CreateCommand()
                cmd.CommandText = "Select [nme] FROM [expensis]"
                Using reader As SqlDataReader = cmd.ExecuteReader
                    While (reader.Read())

                        Me.ListBox1.Items.Add(reader("nme"))
                        
                    End While
                End Using
                SqlConn.Close()
            End Using
        Catch ex As Exception
            MsgBox("DataBase not connected due to the reason because " & ex.Message)
            ' MsgBox("Not Loading", MsgBoxStyle.OkOnly, "Error")

        End Try
    End Sub
    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        GroupBox1.Visible = True
        GroupBox1.Enabled = True
        GroupBox5.Visible = False
        GroupBox5.Enabled = False
        GroupBox3.Visible = False
        GroupBox3.Enabled = False
        Panel1.Visible = True
        Panel1.Enabled = True
        Label25.Text = "Add new"
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        If Len(Trim(txtexentry.Text)) = 0 Then
            MessageBox.Show("Please enter Entry number", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtexentry.Focus()
            Exit Sub
        End If
        If Len(Trim(txtexpnme.Text)) = 0 Then
            MessageBox.Show("Please enter Name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtexpnme.Focus()
            Exit Sub
        End If
        Me.Refresh()
        Try

            MessageBox.Show("Are you sure to add data", "Data Adding", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            insert()
            Label25.Text = "'" & txtexentry.Text & "'expensis details saved successfully!"
            Label25.ForeColor = System.Drawing.Color.DarkGreen

            Btnsve.Enabled = False
            Button9.Enabled = False
            Panel2.Enabled = False
        Catch ex As Exception
            ' MessageBox.Show("Data is already exist", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Label25.Text = "Error while saving '" & txtexentry.Text & "'expensis details"
            Label25.ForeColor = System.Drawing.Color.Red

            MessageBox.Show("Data already exist, you again select Expensis  Details and Try other entry", "Data Invalid, Application is closing", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Dispose()
        End Try
        clear()
        Me.Refresh()
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Call CenterToScreen()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub txtexpsalry_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtexpsalry.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtexpbl_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtexpbl.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txteprnt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txteprnt.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub
    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        e.Graphics.DrawImage(bitmap, 0, 0)
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
    Private Sub DeleteSelecedRows()
        Dim ObjConnection As New SqlConnection()
        Dim i As Integer
        Dim mResult
        mResult = MsgBox("Records Deleted", _
        vbOK + vbQuestion, "Removal confirmation")
        If mResult = vbNo Then
            Exit Sub
        End If
        ObjConnection.ConnectionString = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
        Dim ObjCommand As New SqlCommand()
        ObjCommand.Connection = ObjConnection
        For i = Me.DataGridView1.SelectedRows.Count - 1 To 0 Step -1
            ObjCommand.CommandText = "delete from expensis where entry='" & DataGridView1.SelectedRows(i).Cells("entry").Value & "'"
            ObjConnection.Open()
            ObjCommand.ExecuteNonQuery()
            ObjConnection.Close()

            Me.DataGridView1.Rows.Remove(Me.DataGridView1.SelectedRows(i))
        Next

    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        GroupBox1.Enabled = True
        gridclick()
        Btnsve.Enabled = False
        GroupBox1.Visible = True
        GroupBox1.Enabled = True
        GroupBox5.Visible = False
        GroupBox5.Enabled = False
        GroupBox3.Visible = False
        GroupBox3.Enabled = False
        Panel1.Visible = True
        Panel1.Enabled = True
    End Sub

    Private Sub RemoveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveToolStripMenuItem.Click
        DeleteSelecedRows()
        FillCombo()
        FillCombo2()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        DeleteSelecedRows()
        FillCombo()
    End Sub
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        txtexentry.Text = Val(txtexentry.Text) - 1
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        txtexentry.Text = Val(txtexentry.Text) + 1
    End Sub

    Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        GroupBox1.Visible = False
        GroupBox1.Enabled = False
        GroupBox5.Visible = False
        GroupBox5.Enabled = False
        GroupBox3.Visible = True
        GroupBox3.Enabled = True
        Panel1.Visible = False
        Panel1.Enabled = False
    End Sub

    Private Sub LinkLabel4_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        GroupBox1.Visible = False
        GroupBox1.Enabled = False
        GroupBox5.Visible = True
        GroupBox5.Enabled = True
        GroupBox3.Visible = False
        GroupBox3.Enabled = False
        Panel1.Visible = False
        Panel1.Enabled = False
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        GroupBox1.Visible = False
        GroupBox1.Enabled = False
        GroupBox5.Visible = False
        GroupBox5.Enabled = False
        GroupBox3.Visible = False
        GroupBox3.Enabled = False
        Panel1.Visible = False
        Panel1.Enabled = False
    End Sub

    Private Sub btnsearch_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsearch.Click
        source1.Filter = "[nme] = '" & ComboBox1.Text & "'"
        source2.Filter = "[nme] = '" & ComboBox1.Text & "'"
        DataGridView1.Refresh()
        ComboBox1.Text = ""
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
        strQ = "Select entry as [Entry],nme as [Name],dte as [Date],salry as [Salary],bill as [Bill],pckgrcash as [Package],prsnluse as [Personal Use],rent as [Rent]from expensis"
        cmd = New SqlCommand(strQ, conn)
        da = New SqlDataAdapter(cmd)
        ds = New DataSet
        da.Fill(ds, "expensis")
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


    Private Sub Label27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label27.Click

    End Sub

    Private Sub PictureBox24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox24.Click
        searchlstview()
        movoneitm()
    End Sub

    Private Sub PictureBox15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox15.Click
        Try

            ListBox1.Items.Clear()
            listboxfill()
            listview()
            Label26.Text = "Umrah details loaded successfully!"
            Label26.ForeColor = System.Drawing.Color.DarkGreen

            Label27.Text = "Click to Hide Data"
            Label27.ForeColor = System.Drawing.Color.Black
        Catch ex As Exception
            Label26.Text = "Umrah details not loaded successfully!"
            Label26.ForeColor = System.Drawing.Color.Red
        End Try
    End Sub

    Private Sub PictureBox14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox14.Click
        Try
            ListBox1.Items.Clear()
            ListView1.Items.Clear()
            ListView2.Items.Clear()
            TextBox1.Text = ""
            Label27.Text = "Umrah details Unloaded successfully!"
            Label27.ForeColor = System.Drawing.Color.DarkGreen
            Label26.Text = "Click to view data!"
            Label26.ForeColor = System.Drawing.Color.Black
        Catch ex As Exception
            Label27.Text = "Umrah details not Unloaded successfully!"
            Label27.ForeColor = System.Drawing.Color.Red
        End Try
    End Sub

    Private Sub GroupBox5_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox5.Enter

    End Sub
End Class
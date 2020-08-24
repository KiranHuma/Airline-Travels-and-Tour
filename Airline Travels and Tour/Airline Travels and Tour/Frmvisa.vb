Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.Data.DataTable
Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel '
Imports ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat
Public Class Frmvisa
    Private bitmap As Bitmap

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
    Private Sub insert()
        dbaccessconnection()
        con.Open()
        cmd.CommandText = "insert into vissa( visanmbr,vnme,vdte,vcntry,vwrk,vprce,vorgnlcst,vprft,vwakala)values('" & txtvnmbr.Text & "','" & txtvnme.Text & "','" & txtvdte.Value & "','" & txtvcntry.Text & "','" & txtvwrk.Text & "','" & txtvprce.Text & "','" & txtvorgnlcst.Text & "','" & txtvprft.Text & "','" & txtvwakla.Text & "')"
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
    Private Sub del()
        dbaccessconnection()
        con.Open()
        cmd.CommandText = "delete from vissa where visanmbr=" & txtvnmbr.Text & ""
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
    Private Sub edit()
        dbaccessconnection()
        con.Open()
        cmd.CommandText = ("UPDATE vissa SET vnme = '" & txtvnme.Text & "',vdte= '" & txtvdte.Value & "',vcntry= '" & txtvcntry.Text & "',vwrk= '" & txtvwrk.Text & "',vprce= '" & txtvprce.Text & "',vorgnlcst='" & txtvorgnlcst.Text & "',vprft='" & txtvprft.Text & "',vwakala='" & txtvwakla.Text & "'  where visanmbr=" & txtvnmbr.Text & "")
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub


    Private Sub Btnsve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btnsve.Click
        If Len(Trim(txtvnmbr.Text)) = 0 Then
            MessageBox.Show("Please enter ticket number", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtvnmbr.Focus()
            Exit Sub
        End If
        If Len(Trim(txtvnme.Text)) = 0 Then
            MessageBox.Show("Please enter name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtvnme.Focus()
            Exit Sub
        End If
        If Len(Trim(txtvcntry.Text)) = 0 Then
            MessageBox.Show("Please enter sector.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtvcntry.Focus()
            Exit Sub
        End If
        Try
            MessageBox.Show("Are you sure to add data", "Data Adding", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            insert()
            getdata()
            FillCombo()
            FillCombo1()
            FillCombo2()
            Label25.Text = "'" & txtvnmbr.Text & "' visa details saved successfully!"
            Label25.ForeColor = System.Drawing.Color.DarkGreen
            Btnsve.Enabled = False
            Button9.Enabled = True
            GroupBox1.Enabled = False
            Catch ex As Exception
            Label25.Text = "Error while saving '" & txtvnmbr.Text & "' visa details"
            Label25.ForeColor = System.Drawing.Color.Red
            MsgBox("DataBase not connected due to the reason because " & ex.Message)
            ' MessageBox.Show("Data already exist, you again select visa Details and Try other entry", "Data Invalid, Application is closing", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Dispose()
            End Try
        clear()
    End Sub


    Private Sub btnupdte_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnupdte.Click
        Try
            If Not txtvnme.Text = "" Then
                MessageBox.Show("Are you sure to update data", "Data Updating", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                edit()
                getdata()
                FillCombo()
                FillCombo1()
                FillCombo2()
                Label25.Enabled = True
                Label25.Text = "'" & txtvnmbr.Text & "' Ticket details updated successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen
                GroupBox1.Enabled = False
               
            Else

                MessageBox.Show("Select rows from grid to edit", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                TabPage1.Visible = True
                TabPage2.Visible = False
                TabPage1.Visible = False
                TabPage2.Visible = True
                TabControl1.SelectedTab = TabPage2
            End If
        Catch ex As Exception
            MessageBox.Show("Data not updated successfully", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ' MsgBox("DataBase not connected due to the reason because " & ex.Message)
            Label25.Text = "Error while updating '" & txtvnmbr.Text & "' visa details"
            Label25.ForeColor = System.Drawing.Color.Red
            Me.Dispose()
        End Try
        Me.Refresh()
       
        clear()
    End Sub

    Private Sub Btndel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btndel.Click
        Try
            If Not txtvnme.Text = "" Then
                MessageBox.Show("Are you sure to delete data", "Data Deleting", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                del()
                getdata()
                FillCombo()
                FillCombo1()
                FillCombo2()
                Label25.Text = "'" & txtvnme.Text & "' visa details removed successfully!"
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
            Label25.Text = "Error while removing '" & txtvnmbr.Text & "' visa details"
            Label25.ForeColor = System.Drawing.Color.Red

            Me.Dispose()
        End Try
        clear()
    End Sub

    Private Sub Frmvisa_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dbaccessconnection()
        gridfill()
        FillCombo()
        FillCombo1()
        FillCombo2()
        ComboBox1.Text = "----Select Or Type Name---"
        Me.Label23.Text = Format(Now, "dd-MMM-yyyy")
        Timer1.Enabled = True
        GroupBox1.Enabled = False
        Timer1.Start()
        ' DataGridView1.Sort(DataGridView1.Columns("visanmbr"), System.ComponentModel.ListSortDirection.Ascending)


    End Sub
    Private Sub profit()
        Try
            Dim minus As Integer
            Dim a As Integer
            Dim b As Integer
            a = txtvorgnlcst.Text
            b = txtvprce.Text
            minus = b - a
            txtvprft.Text = minus
            'MessageBox.Show(minus)
        Catch ex As Exception
            MsgBox("Orignal cost and sale must not be empty")
        End Try
    End Sub
    Private Sub FillCombo()
        Try
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
        Catch ex As Exception
            MessageBox.Show("Errpr while loading combobox", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.Dispose()
        End Try
    End Sub
    Private Sub FillCombo1()
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
            da2 = New SqlDataAdapter("SELECT vcntry  from vissa", myConnToAccess)
            '    da = New OleDbDataAdapter("SELECT ticketnmbr from ticket", myConnToAccess)
            da2.Fill(ds2, "vissa")
            Dim view1 As New DataView(tables(0))
            With txtvcntry
                .DataSource = ds.Tables("vissa")
                .DisplayMember = "vcntry "
                .ValueMember = "vcntry"
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
            da2 = New SqlDataAdapter("SELECT vnme from vissa", myConnToAccess)
            '    da = New OleDbDataAdapter("SELECT ticketnmbr from ticket", myConnToAccess)
            da2.Fill(ds2, "vissa")
            Dim view1 As New DataView(tables(0))
            With txtvnme
                .DataSource = ds.Tables("vissa")
                .DisplayMember = "vnme"
                .ValueMember = "vnme"
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
    Private Sub gridfill()
        Try
            provider = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            da = New SqlDataAdapter("Select visanmbr, [vnme] ,vdte as [Visa Date],vcntry,vwrk as [work],vprce as [Price],vorgnlcst  as [Original Cost],vprft as [Profit],vwakala as [Wakala] from vissa ", myConnection)
            da.Fill(ds, "vissa")
            Dim view1 As New DataView(tables(0))
            source1.DataSource = view1
            DataGridView1.DataSource = view1
            DataGridView1.Refresh()
        Catch ex As Exception
            ' MsgBox("not loaded " & ex.Message)
            MessageBox.Show(" All records Not loaded successfully", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.Dispose()
        End Try
    End Sub

    Private Sub btnsearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsearch.Click
        Try
            source1.Filter = "[vnme] = '" & ComboBox1.Text & "'"
            source2.Filter = "[vnme] = '" & ComboBox1.Text & "'"
            DataGridView1.Refresh()
            ComboBox1.Text = ""
        Catch ex As Exception
            MessageBox.Show("Error while searching,try again", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.Dispose()
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        timer.Text = TimeOfDay
    End Sub


    Private Sub clear()
        txtvnmbr.Text = ""
        txtvnme.Text = ""
        txtvdte.Text = ""
        txtvcntry.Text = ""
        txtvwrk.Text = ""
        txtvprce.Text = ""
        txtvorgnlcst.Text = ""
        txtvprft.Text = ""
        txtvwakla.Text = ""
    End Sub

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
            txtboxidd()
        Catch ex As Exception
            MessageBox.Show("Error while adding,Close application and try again", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.Dispose()
        End Try

    End Sub
    Private Sub txtboxidd()
        Try
            dbaccessconnection()
            con.Open()
            Dim num As New Integer
            cmd.CommandText = "SELECT MAX(visanmbr)FROM vissa "
            If (IsDBNull(cmd.ExecuteScalar)) Then
                num = 1
                txtvnmbr.Text = num.ToString
            Else
                'num = cmd.ExecuteScalar()
                'txtticket.Text = num + 1
                num = cmd.ExecuteScalar + 1
                txtvnmbr.Text = num.ToString
            End If
            con.Close()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
            Me.Dispose()
        End Try
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try

            If Not txtvnme.Text = "" Then
                MessageBox.Show("Are you sure to print data", "Data Adding", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                Dim report As New CRvissa
                Dim objText As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(1).ReportObjects("Text1")
                Dim objText1 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text2")
                Dim objText2 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text3")
                Dim objText3 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text4")
                Dim objText4 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text5")
                Dim objText5 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text6")
                Dim objText7 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text7")
                Dim objText8 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text8")

                objText.Text = Me.txtvnmbr.Text
                objText1.Text = Me.txtvnme.Text
                objText2.Text = Me.txtvdte.Text
                objText3.Text = Me.txtvcntry.Text
                objText4.Text = Me.txtvprce.Text
                objText5.Text = Me.txtvorgnlcst.Text
                objText7.Text = Me.txtvwakla.Text
                objText8.Text = Me.txtvwrk.Text
                Rprtvissa.CrystalReportViewer1.ReportSource = report
                Rprtvissa.Show()
                Label25.Text = "'" & txtvnme.Text & "' visa details printed successfully!"
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
            Label25.Text = "Error while printing'" & txtvnmbr.Text & "' visa details"
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

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Call CenterToScreen()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub txtvorgnlcst_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtvorgnlcst.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtvprce_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtvprce.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub
    Private Sub DataGridView1_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        Try
            txtvnmbr.Enabled = False
            Panel2.Enabled = True
            GroupBox1.Enabled = True
            Btnsve.Enabled = False
            Btndel.Enabled = True
            btnupdte.Enabled = True
            Button1.Enabled = True
            Me.Refresh()
            Me.txtvnmbr.Text = DataGridView1.CurrentRow.Cells(0).Value.ToString
            Me.txtvnme.Text = DataGridView1.CurrentRow.Cells(1).Value.ToString
            Me.txtvdte.Text = DataGridView1.CurrentRow.Cells(2).Value.ToString
            Me.txtvcntry.Text = DataGridView1.CurrentRow.Cells(3).Value.ToString
            Me.txtvwrk.Text = DataGridView1.CurrentRow.Cells(4).Value.ToString
            Me.txtvprce.Text = DataGridView1.CurrentRow.Cells(5).Value.ToString
            Me.txtvorgnlcst.Text = DataGridView1.CurrentRow.Cells(6).Value.ToString
            Me.txtvprft.Text = DataGridView1.CurrentRow.Cells(7).Value.ToString
            Me.txtvwakla.Text = DataGridView1.CurrentRow.Cells(8).Value.ToString
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Dispose()
        End Try
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        If Len(Trim(txtvnmbr.Text)) = 0 Then
            MessageBox.Show("Please enter ticket number", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtvnmbr.Focus()
            Exit Sub
        End If
        If Len(Trim(txtvnme.Text)) = 0 Then
            MessageBox.Show("Please enter name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtvnme.Focus()
            Exit Sub
        End If
        If Len(Trim(txtvcntry.Text)) = 0 Then
            MessageBox.Show("Please enter sector.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtvcntry.Focus()
            Exit Sub
        End If
        Try
            MessageBox.Show("Are you sure to add data", "Data Adding", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            insert()
            Label25.Text = "'" & txtvnmbr.Text & "' visa details saved successfully!"
            Label25.ForeColor = System.Drawing.Color.DarkGreen
            Button9.Enabled = False
            GroupBox1.Enabled = False
        Catch ex As Exception
            Label25.Text = "Error while saving '" & txtvnmbr.Text & "' ticket details"
            Label25.ForeColor = System.Drawing.Color.Red
            MessageBox.Show("Data already exist, you again select visa Details and Try other entry", "Data Invalid, Application is closing", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
                cmd.CommandText = "Select [vnme]from [vissa] "
                Using reader As SqlDataReader = cmd.ExecuteReader
                    While (reader.Read())
                        'for access da = New OleDbDataAdapter("Select [ticketnmbr], [nme], [sector] ,[dte],[pnr],[issuedte], [expiredte], [mobileno] ,[flight],[pasprtnmbr], [comissin], [basicfair], [deprttme] ,[arrtme],[totalamount],[status]  from ticket", myConnection)



                        Me.ListBox1.Items.Add(reader("vnme"))
                       

                    End While
                End Using
                SqlConn.Close()
            End Using
        Catch ex As Exception
            ' MsgBox("Not Loading", MsgBoxStyle.OkOnly, "Error", & ex.Message)
            MsgBox("DataBase not connected due to the reason because " & ex.Message)
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        TabPage1.Visible = True
        Label25.Text = ""
        TabControl1.SelectedTab = TabPage1
    End Sub
    Private Sub getdata()
        Dim con As New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
        con.Open()
        Dim da As New SqlDataAdapter("Select visanmbr,vnme,vdte as [Visa Date],vcntry,vwrk as [work],vprce as [Price],vorgnlcst  as [Original Cost],vprft as [Profit],vwakala as [Wakala] from vissa", con)
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

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try

            If Not DataGridView1.CurrentRow.IsNewRow Then
                'Query string
                dbaccessconnection()
                con.Open()
                cmd.CommandText = "delete  from vissa  where visanmbr='" & DataGridView1.CurrentRow.Cells(0).Value & "'"
                cmd.ExecuteNonQuery()
                DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
                MessageBox.Show("Record Deleted")
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub txtvprce_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtvprce.Validating
        profit()
    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

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

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        txtvnmbr.Text = Val(txtvnmbr.Text) - 1
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        txtvnmbr.Text = Val(txtvnmbr.Text) + 1
    End Sub
    Private Sub loaddata_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ListBox1.Items.Clear()
        listboxfill()
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
            ObjCommand.CommandText = "delete  from vissa  where visanmbr='" & DataGridView1.SelectedRows(i).Cells("visanmbr").Value & "'"
            ObjConnection.Open()
            ObjCommand.ExecuteNonQuery()
            ObjConnection.Close()

            Me.DataGridView1.Rows.Remove(Me.DataGridView1.SelectedRows(i))
        Next

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        DeleteSelecedRows()
    End Sub
    Private Sub TextBox1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox1.MouseClick
        TextBox1.Text = ""
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
        strQ = "Select visanmbr as [Visa Number],vnme as [Visa Name] ,vdte as [Visa Date],vcntry as [Country],vwrk as [Work],vorgnlcst  as [Original Cost],vwakala as [Wakala] from vissa"
        cmd = New SqlCommand(strQ, conn)
        da = New SqlDataAdapter(cmd)
        ds = New DataSet
        da.Fill(ds, "vissa")
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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        searchlstview()
        movoneitm()
    End Sub

    Private Sub loaddata_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles loaddata.Click
        ListBox1.Items.Clear()
        listboxfill()
        listview()
    End Sub

    Private Sub Button11_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        ListView1.Items.Clear()
        ListView2.Items.Clear()
        ListBox1.Items.Clear()
        TextBox1.Text = ""
    End Sub

    Private Sub GroupBox4_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox4.Enter

    End Sub
End Class
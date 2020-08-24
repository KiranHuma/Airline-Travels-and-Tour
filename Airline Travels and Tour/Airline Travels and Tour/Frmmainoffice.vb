Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.Data.DataTable
Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel '
Imports ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat
Public Class Frmmainoffice
    Private bitmap As Bitmap
    Dim provider As String
    Dim dataFile As String
    Dim connString As String
    'Dim myConnection As OleDbConnection = New OleDbConnection
    Dim myConnection As SqlConnection = New SqlConnection
    Dim ds As DataSet = New DataSet
    'Dim da As OleDbDataAdapter
    Dim da As SqlDataAdapter
    Dim tables As DataTableCollection = ds.Tables
    Dim source1 As New BindingSource()
    Dim source2 As New BindingSource()


    Dim dt As New DataTable
    ' Dim con As New OleDb.OleDbConnection
    'Dim cmd As New OleDb.OleDbCommand
    Dim con As New SqlClient.SqlConnection
    Dim cmd As New SqlClient.SqlCommand
    '////////////////////////////////////////////////////////////////////
    ' Private Sub dbaccessconnection()
    ''Acces DataBase Connectivity and for MS Access 2003 PROVIDER=Microsoft.Jet.OLEDB.4.0
    ' Try
    ' con.ConnectionString = "provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb"
    'cmd.Connection = con
    ' MessageBox.Show("connection created")
    '  Catch ex As Exception
    '  MsgBox("DataBase not connected due to the reason because " & ex.Message)
    ' End Try
    'End Sub
    '///////////////////////////////////////////////////////////////
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
        cmd.CommandText = "insert into mainoffice( trnsctinnmbr,nme,dte,cash,bank,accntnmbr,amnt)values('" & txtofftrnsctin.Text & "','" & txtoffnme.Text & "','" & txtoffdte.Value & "','" & txtoffcash.Text & "','" & txtoffbnk.Text & "','" & txtacntnmbr.Text & "','" & txtamnt.Text & "')"
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
    Private Sub del()
        dbaccessconnection()
        con.Open()
        cmd.CommandText = "delete from mainoffice where trnsctinnmbr=" & txtofftrnsctin.Text & ""
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
    Private Sub edit()
        dbaccessconnection()
        con.Open()
        cmd.CommandText = ("UPDATE mainoffice SET  nme = '" & txtoffnme.Text & "', dte= '" & txtoffdte.Value & "',cash= '" & txtoffcash.Text & "',bank= '" & txtoffbnk.Text & "',accntnmbr= '" & txtacntnmbr.Text & "',amnt='" & txtamnt.Text & "'  where trnsctinnmbr=" & txtofftrnsctin.Text & "")
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
    Private Sub Frmmainoffice_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If (DateTime.Now.Hour < 12) Then
            PictureBox18.Visible = False
            PictureBox5.Visible = False
            PictureBox19.Visible = True
            PictureBox17.Visible = False
            lblgrting.Text = "Good Morning"
            'Label12.Text = Convert.ToString(DateTime.Now)

        ElseIf (DateTime.Now.Hour < 17) Then
            PictureBox18.Visible = True
            PictureBox5.Visible = False
            PictureBox19.Visible = False
            PictureBox17.Visible = False
            lblgrting.Text = "Good Afternoon"
            ' Label12.Text = Convert.ToString(DateTime.Now)
        ElseIf (DateTime.Now.Hour < 19) Then
            PictureBox18.Visible = False
            PictureBox5.Visible = True
            PictureBox19.Visible = False
            PictureBox17.Visible = False
            lblgrting.Text = "Good Evening"
        Else
            PictureBox18.Visible = False
            PictureBox5.Visible = False
            PictureBox19.Visible = False
            PictureBox17.Visible = True
            ' lblgrting.Text = "Good Night"
            ' Label12.Text = Convert.ToString(DateTime.Now)
        End If
        Label24.Text = Date.Today.ToString("dddd")

        Panel1.Enabled = False
        FillCombo()
        FillCombo1()
        ComboBox1.Text = "----Select Or Type Name---"
        ' DataGridView1.Sort(DataGridView1.Columns("entry"), System.ComponentModel.ListSortDirection.Ascending)
        Me.Label23.Text = Format(Now, "dd-MMM-yyyy")
        Timer1.Enabled = True
        Timer1.Start()
        dbaccessconnection()
        gridfill()
        ' DataGridView1.Sort(DataGridView1.Columns("trnsctinnmbr"), System.ComponentModel.ListSortDirection.Ascending)
       
    End Sub
    Private Sub FillCombo()
        Try
            ' Dim myConnToAccess As OleDbConnection
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
            'da = New OleDbDataAdapter("SELECT nme from mainoffice ", myConnToAccess)
            da = New SqlDataAdapter("SELECT nme from mainoffice ", myConnToAccess)
            da.Fill(ds, "mainoffice")
            Dim view1 As New DataView(tables(0))
            With ComboBox1
                .DataSource = ds.Tables("mainoffice")
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
    Private Sub FillCombo1()
        Try
            ' Dim myConnToAccess As OleDbConnection
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
            'da = New OleDbDataAdapter("SELECT nme from mainoffice ", myConnToAccess)
            da = New SqlDataAdapter("SELECT nme from mainoffice ", myConnToAccess)
            da.Fill(ds, "mainoffice")
            Dim view1 As New DataView(tables(0))
            With txtoffnme
                .DataSource = ds.Tables("mainoffice")
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
            'provider = "provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb"
            provider = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            '( bnknme,Transctinnmbr,dte,accntnmbr,amount,accnthldr
            ' da = New OleDbDataAdapter("Select [trnsctinnmbr], [nme], [dte],[cash],[bank],[accntnmbr],[amnt]from mainoffice ", myConnection)
            da = New SqlDataAdapter("Select [trnsctinnmbr], [nme], [dte],[cash],[bank],[accntnmbr],[amnt]from mainoffice ", myConnection)
            da.Fill(ds, "mainoffice ")
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
        Dim da As New SqlDataAdapter("Select [trnsctinnmbr], [nme], [dte],[cash],[bank],[accntnmbr],[amnt]from mainoffice ", con)
        Dim dt As New DataTable
        da.Fill(dt)
        'Dim view1 As New DataView(tables1(0))
        source2.DataSource = dt
        DataGridView1.DataSource = dt
        DataGridView1.Refresh()
    End Sub
    Private Sub btnsearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsearch.Click
      
    End Sub
    Private Sub clear()
        txtofftrnsctin.Text = ""
        txtoffnme.Text = ""
        txtoffdte.Text = ""
        txtoffcash.Text = ""
        txtoffbnk.Text = ""
        txtacntnmbr.Text = ""
        txtamnt.Text = ""
    End Sub
    Private Sub PictureBox10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Panel2.Enabled = True
        txtofftrnsctin.Enabled = True
        clear()
    End Sub

    Private Sub Label17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label17.Click
        Panel1.Visible = True
        Panel1.Enabled = True
        PictureBox8.Visible = False
        Panel4.Visible = False
        Panel4.Enabled = False
        Panel5.Visible = False
        Panel5.Enabled = False
        Panel8.Enabled = False
        Panel8.Visible = False
        Panel7.Enabled = True
        Panel7.Visible = True
        Panel6.Enabled = False
        Panel6.Visible = False




    End Sub

    Private Sub Label18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label18.Click
        Panel1.Visible = False
        Panel4.Visible = True
        Panel4.Enabled = True
        Panel1.Visible = False
        PictureBox8.Visible = False
        Panel1.Enabled = False
        Panel5.Visible = True
        Panel5.Enabled = True
        Panel8.Enabled = False
        Panel8.Visible = False
        Panel7.Enabled = False
        Panel7.Visible = False
        Panel6.Enabled = False
        Panel6.Visible = False

    End Sub
    Private Sub Label20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label20.Click
        Panel1.Visible = False
        Panel4.Visible = False
        Panel4.Enabled = False
        Panel1.Visible = False
        Panel1.Enabled = False
        PictureBox8.Visible = False
        Panel5.Visible = False
        Panel5.Enabled = False
        Panel8.Enabled = True
        Panel8.Visible = True
        Panel7.Enabled = False
        Panel7.Visible = False
        Panel6.Enabled = True
        Panel6.Visible = True
    End Sub

    Private Sub EmptyRecycleBinToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmptyRecycleBinToolStripMenuItem.Click
        PictureBox12.Visible = False
        PictureBox13.Visible = True
    End Sub

    Private Sub Label10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label10.Click
        PictureBox12.Visible = True
        PictureBox13.Visible = False
        DeleteSelecedRows()
        FillCombo()
        FillCombo1()
        PictureBox9.Visible = False
        PictureBox10.Visible = True
        PictureBox16.Visible = False
     
    End Sub

    Private Sub Label12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label12.Click
        PictureBox2.Visible = False
        PictureBox4.Visible = True
    End Sub



    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

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
                cmd.CommandText = "Select [nme]from [mainoffice]"
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


    Private Sub Btnadd_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btnadd.Click
        Label25.Text = " Add new record!"
        Label25.ForeColor = System.Drawing.Color.DarkCyan
        PictureBox9.Visible = False
        PictureBox10.Visible = True
        PictureBox16.Visible = False
        Try
            Panel1.Enabled = True
            Btnsve.Enabled = True
            Button9.Enabled = True
            Btndel.Enabled = False
            btnupdte.Enabled = False
            Button1.Enabled = False
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
            'visit( vientry,nme,dte,
            cmd.CommandText = "SELECT MAX(trnsctinnmbr) from mainoffice"
            If (IsDBNull(cmd.ExecuteScalar)) Then
                num = 1
                txtofftrnsctin.Text = num.ToString
            Else
                'num = cmd.ExecuteScalar()
                'txtticket.Text = num + 1
                num = cmd.ExecuteScalar + 1
                txtofftrnsctin.Text = num.ToString
            End If
            con.Close()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        End Try
    End Sub

    Private Sub btnupdte_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnupdte.Click
        Try
            If Not txtofftrnsctin.Text = "" Then
                MessageBox.Show("Are you sure to update data", "Data Updating", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                edit()
                FillCombo()
                getdata()
                Label25.Enabled = True
                Label25.Text = "'" & txtofftrnsctin.Text & "' office details updated successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen
                PictureBox9.Visible = False
                PictureBox10.Visible = True
                PictureBox16.Visible = False
                Panel1.Enabled = False
                Panel1.Visible = False
                Btnadd.Enabled = True
                Me.Refresh()
            Else
                MessageBox.Show("Select rows from grid to edit", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Label25.Text = "Error while updating '" & txtofftrnsctin.Text & "' Office details"
                Label25.ForeColor = System.Drawing.Color.Red
                PictureBox10.Visible = False
                PictureBox16.Visible = False
                PictureBox9.Visible = True
                PictureBox8.Visible = False
                Panel4.Visible = True
                Panel4.Enabled = True
                Panel5.Visible = True
                Panel5.Enabled = True
                Panel7.Visible = False
                Panel7.Enabled = False
                Panel1.Enabled = False
                Panel1.Visible = False
                Label18.Focus()
            End If
        Catch ex As Exception
            MessageBox.Show("Data not updated successfully", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Label25.Text = "Error while updating '" & txtofftrnsctin.Text & "' office details"
            Label25.ForeColor = System.Drawing.Color.Red
            PictureBox10.Visible = False
            PictureBox16.Visible = False
            PictureBox9.Visible = True
            Me.Dispose()
        End Try

        Me.Refresh()
        clear()
    End Sub

    Private Sub Btndel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btndel.Click
        Try
            If Not txtofftrnsctin.Text = "" Then
                MessageBox.Show("Are you sure to delete data", "Data Deleting", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                del()
                FillCombo()
                getdata()
                Label25.Text = "'" & txtofftrnsctin.Text & "'office  details removed successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen
                PictureBox9.Visible = False
                PictureBox10.Visible = True
                PictureBox16.Visible = False
                Panel1.Enabled = False
                Panel1.Visible = False
                Btnadd.Enabled = True


            Else
                MessageBox.Show("Select rows from grid to remove", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Label25.Text = "Error while removing '" & txtofftrnsctin.Text & "' office details"
                Label25.ForeColor = System.Drawing.Color.Red
                PictureBox10.Visible = False
                PictureBox16.Visible = False
                PictureBox9.Visible = True
                Panel4.Visible = True
                Panel4.Enabled = True
                Panel5.Visible = True
                Panel5.Enabled = True
                Panel7.Visible = False
                Panel7.Enabled = False
                Panel1.Enabled = False
                Panel1.Visible = False
                Panel4.Visible = True
                Label18.Focus()


            End If
        Catch ex As Exception
            'MessageBox.Show("Data is not remove succesfully", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Label25.Text = "Error while removing '" & txtofftrnsctin.Text & "' office details"
            Label25.ForeColor = System.Drawing.Color.Red
            PictureBox10.Visible = False
            PictureBox16.Visible = False
            PictureBox9.Visible = True
            Me.Dispose()
        End Try
        Me.Refresh()
        clear()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        timer.Text = TimeOfDay
    End Sub
    Private Sub gridclick()
        Try

            Panel1.Visible = True
            Label25.Text = "Edit Values in above textboxes"
            PictureBox9.Visible = False
            PictureBox10.Visible = True
            PictureBox16.Visible = False
            Label18.Focus()
            Panel1.Enabled = True
            txtofftrnsctin.Enabled = False
            Btnsve.Enabled = False
            Btndel.Enabled = True
            btnupdte.Enabled = True
            Panel4.Visible = False
            Me.txtofftrnsctin.Text = DataGridView1.CurrentRow.Cells(0).Value.ToString
            Me.txtoffnme.Text = DataGridView1.CurrentRow.Cells(1).Value.ToString
            Me.txtoffdte.Text = DataGridView1.CurrentRow.Cells(2).Value.ToString
            Me.txtoffcash.Text = DataGridView1.CurrentRow.Cells(3).Value.ToString
            Me.txtoffbnk.Text = DataGridView1.CurrentRow.Cells(4).Value.ToString
            Me.txtacntnmbr.Text = DataGridView1.CurrentRow.Cells(5).Value.ToString
            Me.txtamnt.Text = DataGridView1.CurrentRow.Cells(6).Value.ToString
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
   
    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Panel1.Visible = True
        Panel1.Enabled = True
        Panel5.Enabled = True
        Panel5.Visible = True
        Panel7.Visible = True
        Panel7.Enabled = True
        'entry,unme,udte,upasprt,utrvlngdte,uexdte,uorgnlcst,usale,uprft,uduratin,upkge,urefrnce
        gridclick()
    End Sub

    Private Sub Btnsve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btnsve.Click
        If Len(Trim(txtofftrnsctin.Text)) = 0 Then
            MessageBox.Show("Please enter Transcation Number", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtofftrnsctin.Focus()
            Exit Sub
        End If
        If Len(Trim(txtoffnme.Text)) = 0 Then
            MessageBox.Show("Please enter Name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtoffnme.Focus()
            Exit Sub
        End If

        Me.Refresh()
        Try
            If Not txtofftrnsctin.Text = "" Then
                MessageBox.Show("Are you sure to add data", "Data Adding", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                insert()
                getdata()
                FillCombo()
                Label25.Text = "'" & txtofftrnsctin.Text & "' office details saved successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen
                PictureBox9.Visible = False
                PictureBox10.Visible = True
                PictureBox16.Visible = False
                Btnsve.Enabled = False
                Btnadd.Enabled = True
                Btndel.Enabled = True
                btnupdte.Enabled = True
                Button1.Enabled = True
                Button9.Enabled = True
                Panel1.Enabled = False

            Else

                MessageBox.Show("please fill all above textboxes", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Label25.Text = "Problem while saving '" & txtofftrnsctin.Text & "' office details"
                Label25.ForeColor = System.Drawing.Color.Red
                PictureBox10.Visible = False
                PictureBox16.Visible = False
                PictureBox9.Visible = True
            End If
        Catch ex As Exception
            Label25.Text = "Error while saving '" & txtofftrnsctin.Text & "' office details"
            Label25.ForeColor = System.Drawing.Color.Red
            MessageBox.Show("Data already exist, you again select office Details and Try other entry", "Data Invalid, Application is closing", MessageBoxButtons.OK, MessageBoxIcon.Error)
            PictureBox10.Visible = False
            PictureBox16.Visible = False
            PictureBox9.Visible = True
            Me.Dispose()
        End Try
        Panel1.Enabled = False
        clear()
        Me.Refresh()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If Not txtofftrnsctin.Text = "" Then
                MessageBox.Show("Are you sure to print data", "Data Adding", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                Dim report As New CRmainfoffce
                Dim objText As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(1).ReportObjects("Text1")
                Dim objText1 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text2")
                Dim objText2 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text3")
                Dim objText3 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text4")
                Dim objText4 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text5")
                Dim objText5 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text6")
                Dim objText6 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text7")

                objText.Text = Me.txtofftrnsctin.Text
                objText1.Text = Me.txtoffnme.Text
                objText2.Text = Me.txtoffdte.Text
                objText3.Text = Me.txtoffcash.Text
                objText4.Text = Me.txtoffbnk.Text
                objText5.Text = Me.txtacntnmbr.Text
                objText6.Text = Me.txtamnt.Text
                Rprtmainoffce.CrystalReportViewer1.ReportSource = report
                Rprtmainoffce.Show()
                Label25.Text = "'" & txtofftrnsctin.Text & "' office details printed successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen
                PictureBox10.Visible = False
                PictureBox16.Visible = False
                PictureBox9.Visible = True
            Else
                MessageBox.Show("Select rows from grid to edit", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Label25.Text = "Error while updating '" & txtofftrnsctin.Text & "' Office details"
                Label25.ForeColor = System.Drawing.Color.Red
                PictureBox10.Visible = False
                PictureBox16.Visible = False
                PictureBox9.Visible = True
                PictureBox8.Visible = False
                Panel4.Visible = True
                Panel4.Enabled = True
                Panel5.Visible = True
                Panel5.Enabled = True
                Panel7.Visible = False
                Panel7.Enabled = False
                Panel1.Enabled = False
                Panel1.Visible = False
                Label18.Focus()
            End If
            Catch ex As Exception
            MessageBox.Show("Reports are not loding properly,try again", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Label25.Text = "Error while printing'" & txtofftrnsctin.Text & "' office details"
            Label25.ForeColor = System.Drawing.Color.Red
            PictureBox10.Visible = False
            PictureBox16.Visible = False
            PictureBox9.Visible = True
            Me.Dispose()
        End Try
    End Sub

    Private Sub txtoffcash_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtoffcash.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        e.Graphics.DrawImage(bitmap, 0, 0)
    End Sub

    Private Sub PictureBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox4.Click
       
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
            ObjCommand.CommandText = "delete from mainoffice  where trnsctinnmbr='" & DataGridView1.SelectedRows(i).Cells("trnsctinnmbr").Value & "'"
            ObjConnection.Open()
            ObjCommand.ExecuteNonQuery()
            ObjConnection.Close()

            Me.DataGridView1.Rows.Remove(Me.DataGridView1.SelectedRows(i))
        Next

    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        Panel1.Enabled = True
        gridclick()
    End Sub

    Private Sub RemoveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveToolStripMenuItem.Click
        DeleteSelecedRows()
        PictureBox12.Visible = True
        PictureBox13.Visible = False
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        If Len(Trim(txtofftrnsctin.Text)) = 0 Then
            MessageBox.Show("Please enter Transcation Number", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtofftrnsctin.Focus()
            Exit Sub
        End If
        If Len(Trim(txtoffnme.Text)) = 0 Then
            MessageBox.Show("Please enter Name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtoffnme.Focus()
            Exit Sub
        End If
        Try

            MessageBox.Show("Are you sure to add data", "Data Adding", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            insert()
            Label25.Text = "'" & txtofftrnsctin.Text & "'office details saved successfully!"
            Label25.ForeColor = System.Drawing.Color.DarkGreen
            PictureBox10.Visible = False
            PictureBox16.Visible = False
            PictureBox9.Visible = True
            Btnsve.Enabled = False
            Button9.Enabled = False
            Panel1.Enabled = False
        Catch ex As Exception
            ' MessageBox.Show("Data is already exist", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Label25.Text = "Error while saving '" & txtofftrnsctin.Text & "'office details"
            Label25.ForeColor = System.Drawing.Color.Red
            PictureBox10.Visible = False
            PictureBox16.Visible = False
            PictureBox9.Visible = True
            MessageBox.Show("Data already exist, you again select office  Details and Try other entry", "Data Invalid, Application is closing", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Dispose()
        End Try
        clear()
        Me.Refresh()
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

    Private Sub PictureBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox6.Click
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
        PictureBox10.Visible = False
        PictureBox16.Visible = False
        PictureBox9.Visible = True
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        source1.Filter = "[nme] = '" & ComboBox1.Text & "'"
        source2.Filter = "[nme] = '" & ComboBox1.Text & "'"
        DataGridView1.Refresh()
        ComboBox1.Text = ""
    End Sub

    Private Sub PictureBox23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox23.Click
         txtofftrnsctin.Text = Val(txtofftrnsctin.Text) -1 
    End Sub
    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3.Click
        txtofftrnsctin.Text = Val(txtofftrnsctin.Text) + 1
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PictureBox11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox11.Click
        Panel1.Visible = False
        Panel6.Visible = False
        Panel4.Visible = False
        Panel1.Enabled = False
        Panel6.Enabled = False
        Panel4.Enabled = False
        PictureBox8.Visible = True
        Btndel.Enabled = True
        btnupdte.Enabled = True
        Button1.Enabled = True
        Btnsve.Enabled = False
        Panel8.Visible = False
        Panel5.Visible = False
        Panel7.Visible = False
        Panel8.Enabled = False
        Panel5.Enabled = False
        Panel7.Enabled = False
    End Sub

    Private Sub PictureBox21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox21.Click
        If MsgBox("Are you sure want to exit now?", MsgBoxStyle.YesNo, "Closing warning") = MsgBoxResult.Yes Then ' If you select yes in the MsgBox then it will close the window
            Me.Close()
            Me.Dispose() ' Close the window
        Else
            ' Will not close the application
        End If
    End Sub

    Private Sub PictureBox20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox20.Click
        Call CenterToScreen()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.WindowState = FormWindowState.Minimized
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
        strQ = "Select trnsctinnmbr as [Ticket Number], nme as [Name], dte AS [Date],cash as [Cash],bank as [Bank],amnt as [Amount] From mainoffice"
        cmd = New SqlCommand(strQ, conn)
        da = New SqlDataAdapter(cmd)
        ds = New DataSet
        da.Fill(ds, "mainoffice")
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

    Private Sub PictureBox24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox24.Click
        searchlstview()
        movoneitm()
    End Sub
End Class
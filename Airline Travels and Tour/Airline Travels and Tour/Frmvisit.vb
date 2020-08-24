Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.Data.DataTable
Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel '
Imports ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat
Public Class Frmvisit
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
    ' Dim con As New OleDb.OleDbConnection
    ' Dim cmd As New OleDb.OleDbCommand
    '////////////////////////////////////////////////////////////
    ' Private Sub dbaccessconnection()
    'Acces DataBase Connectivity and for MS Access 2003 PROVIDER=Microsoft.Jet.OLEDB.4.0
    ' Try
    ' con.ConnectionString = "provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb"
    ' cmd.Connection = con
    'MessageBox.Show("connection created")
    ' Catch ex As Exception
    ' MsgBox("DataBase not connected due to the reason because " & ex.Message)
    ' End Try
    ' End Sub
    '//////////////////////////////////////////////////////
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
        cmd.CommandText = "insert into visit( vientry,nme,dte,psprtno,trvelingdte,expirdte,vipresent,orgnlcst,sale,proft,viduration,refrnce)values('" & txtvientry.Text & "','" & txtvinme.Text & "','" & txtvidte.Value & "','" & txtvipsprt.Text & "','" & txtvitrvldte.Value & "','" & txtexvidte.Value & "','" & txtviprsnt.Text & "','" & txtorgnlcst.Text & "','" & txtsale.Text & "','" & txtprft.Text & "','" & txtvidurtin.Text & "','" & txtvirfrnce.Text & "')"
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub

    Private Sub Btnsve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btnsve.Click

        If Len(Trim(txtvientry.Text)) = 0 Then
            MessageBox.Show("Please enter Entry number", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtvientry.Focus()
            Exit Sub
        End If
        If Len(Trim(txtvinme.Text)) = 0 Then
            MessageBox.Show("Please enter name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtvinme.Focus()
            Exit Sub
        End If
        If Len(Trim(txtvipsprt.Text)) = 0 Then
            MessageBox.Show("Please enter Passport Number.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtvipsprt.Focus()
            Exit Sub
        End If
        Try
            If Not txtvinme.Text = "" Then
                MessageBox.Show("Are you sure to add data", "Data Adding", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                insert()
                getdata()
                FillCombo()
                Label25.Text = "'" & txtvientry.Text & "' Visit details saved successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen
                PictureBox5.Visible = True
                PictureBox4.Visible = False
                PictureBox6.Visible = False

                Btnsve.Enabled = False
                Button9.Enabled = True
                Panel2.Enabled = False

            Else

                MessageBox.Show("please fill all above textboxes", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Label25.Text = "Problem while saving '" & txtvientry.Text & "' visit details"
                Label25.ForeColor = System.Drawing.Color.Red
                PictureBox4.Visible = True
                PictureBox6.Visible = False
                PictureBox5.Visible = False

            End If
        Catch ex As Exception
            Label25.Text = "Error while saving '" & txtvientry.Text & "' visit details"
            Label25.ForeColor = System.Drawing.Color.Red
            MessageBox.Show("Data already exist, you again select visit Details and Try other entry", "Data Invalid, Application is closing", MessageBoxButtons.OK, MessageBoxIcon.Error)
            PictureBox4.Visible = True
            PictureBox6.Visible = False
            PictureBox5.Visible = False
            Me.Dispose()
        End Try
        Panel2.Enabled = False
        clear()
        Me.Refresh()
    End Sub
    Private Sub del()
        dbaccessconnection()
        con.Open()
        cmd.CommandText = "delete from visit where vientry=" & txtvientry.Text & ""
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
    Private Sub edit()
        dbaccessconnection()
        con.Open()
        cmd.CommandText = ("UPDATE visit SET nme= '" & txtvinme.Text & "',dte= '" & txtvidte.Value & "',psprtno= '" & txtvipsprt.Text & "',trvelingdte= '" & txtvitrvldte.Value & "',expirdte='" & txtexvidte.Value & "',vipresent='" & txtviprsnt.Text & "',orgnlcst='" & txtorgnlcst.Text & "',sale='" & txtsale.Text & "',proft='" & txtprft.Text & "',viduration='" & txtvidurtin.Text & "',refrnce= '" & txtvirfrnce.Text & "'  where vientry=" & txtvientry.Text & "")
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub

    Private Sub Frmvisit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If (DateTime.Now.Hour < 12) Then
            PictureBox2.Visible = False
            PictureBox3.Visible = True
            PictureBox1.Visible = False
            PictureBox7.Visible = False

            lblgrting.Text = "Good Morning"
            'Label12.Text = Convert.ToString(DateTime.Now)

        ElseIf (DateTime.Now.Hour < 17) Then
            PictureBox2.Visible = False
            PictureBox3.Visible = False
            PictureBox1.Visible = True
            PictureBox7.Visible = False
            lblgrting.Text = "Good Afternoon"
            ' Label12.Text = Convert.ToString(DateTime.Now)
        ElseIf (DateTime.Now.Hour < 19) Then
          PictureBox2.Visible = False
            PictureBox3.Visible = False
            PictureBox1.Visible = False
            PictureBox7.Visible = True
            lblgrting.Text = "Good Evening"
        Else
            PictureBox2.Visible = True
            PictureBox3.Visible = False
            PictureBox1.Visible = False
            PictureBox7.Visible = False
            lblgrting.Text = "Good Night"
            ' Label12.Text = Convert.ToString(DateTime.Now)
        End If
        Label12.Text = Date.Today.ToString("dddd")
        Try
            gridfill()
            dbaccessconnection()
            FillCombo()
            FillCombo2()
            Panel2.Enabled = False
            ComboBox1.Text = "----Select Or Type Name---"
            ' Call CenterToScreen()
            ' Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            'Me.WindowState = FormWindowState.Maximized
            Me.Label23.Text = Format(Now, "dd-MMM-yyyy")
            Timer1.Enabled = True
            Timer1.Start()

            ' DataGridView1.Sort(DataGridView1.Columns("vientry"), System.ComponentModel.ListSortDirection.Ascending)
        Catch ex As Exception
            MsgBox("DataBase not connected due to the reason because " & ex.Message)
            ' MessageBox.Show(" Not loaded successfully", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub
    Private Sub profitt()
        Try

            Dim minus As Integer
            Dim a As Integer
            Dim b As Integer
            a = txtorgnlcst.Text
            b = txtsale.Text
            minus = b - a
            txtprft.Text = minus
            'MessageBox.Show(minus)
        Catch ex As Exception
            MsgBox("Orignal cost and sale must not be empty to calculate commission")
        End Try
    End Sub

    Private Sub gridfill()
        Try
            'provider = "provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb"
            provider = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            'vientry,nme,dte,psprtno,trvelingdte,expirdte,vipresent,viduration,refrnce
            da = New SqlDataAdapter("Select vientry, nme , dte as [Date], psprtno as [Passport Number],trvelingdte as [Travelling Date],expirdte as [Expire Date],vipresent as [Present],orgnlcst as [Orginal cost],sale as [Sale],proft as[Profit],viduration as [Duration],refrnce as [Reference] from visit ", myConnection)
            da.Fill(ds, "visit")
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
        Dim da As New SqlDataAdapter("Select vientry, nme , dte as [Date], psprtno as [Passport Number],trvelingdte as [Travelling Date],expirdte as [Expire Date],vipresent as [Present],orgnlcst as [Orginal cost],sale as [Sale],proft as[Profit],viduration as [Duration],refrnce as [Reference] from visit ", con)
        Dim dt As New DataTable
        da.Fill(dt)
        'Dim view1 As New DataView(tables1(0))
        source2.DataSource = dt
        DataGridView1.DataSource = dt
        DataGridView1.Refresh()
    End Sub
    Private Sub datechk()
        Dim dteissue As Date
        Dim dteex As Date
        dteissue = txtvitrvldte.Value
        dteex = txtexvidte.Value
        If dteex < dteissue Then
            MessageBox.Show("Expire date must be greater than isssue date!", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
    Private Sub FillCombo()
        Try
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
                .SelectedIndex = -1
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .AutoCompleteSource = AutoCompleteSource.ListItems
            End With
        Catch ex As Exception
            MessageBox.Show("At least one entry", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub
    Private Sub FillCombo2()
        Try
            'Dim myConnToAccess As OleDbConnection
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
            da1 = New SqlDataAdapter("SELECT nme from visit", myConnToAccess)
            da1.Fill(ds1, "visit")
            Dim view1 As New DataView(tables(0))
            With txtvinme
                .DataSource = ds.Tables("visit")
                .DisplayMember = "nme"
                .ValueMember = "nme"
                .SelectedIndex = -1
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .AutoCompleteSource = AutoCompleteSource.ListItems
            End With
        Catch ex As Exception
            MessageBox.Show("At least one entry", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub Btndel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btndel.Click
        Try
            If Not txtvinme.Text = "" Then
                MessageBox.Show("Are you sure to delete data", "Data Deleting", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                del()
                FillCombo()
                getdata()
                Label25.Text = "'" & txtvientry.Text & "'visit details removed successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen
                PictureBox5.Visible = True
                PictureBox6.Visible = False
                PictureBox4.Visible = False
                Panel2.Enabled = False
                Panel2.Enabled = False
            Else
                MessageBox.Show("Select rows from grid to remove", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Label25.Text = "Error while removing '" & txtvientry.Text & "' visit details"
                Label25.ForeColor = System.Drawing.Color.Red
                PictureBox4.Visible = True
                PictureBox6.Visible = False
                PictureBox5.Visible = False
                TabPage1.Visible = True
                TabPage2.Visible = False
                TabPage1.Visible = False
                TabPage2.Visible = True
                TabControl1.SelectedTab = TabPage2
            End If
        Catch ex As Exception
            'MessageBox.Show("Data is not remove succesfully", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Label25.Text = "Error while removing '" & txtvientry.Text & "' ticket details"
            Label25.ForeColor = System.Drawing.Color.Red
            PictureBox4.Visible = True
            PictureBox6.Visible = False
            PictureBox5.Visible = False
            Me.Dispose()
        End Try

        Me.Refresh()
        clear()
    End Sub

    Private Sub btnupdte_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnupdte.Click
        Try
            If Not txtvinme.Text = "" Then
                MessageBox.Show("Are you sure to update data", "Data Updating", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                edit()
                FillCombo()
                getdata()
                Label25.Enabled = True
                Label25.Text = "'" & txtvientry.Text & "' visit details updated successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen
                PictureBox5.Visible = True
                PictureBox6.Visible = False
                PictureBox4.Visible = False
                Panel2.Enabled = False
                Me.Refresh()
            Else
                MessageBox.Show("Select rows from grid to edit", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Label25.Text = "Error while updating '" & txtvientry.Text & "' visit details"
                Label25.ForeColor = System.Drawing.Color.Red
                PictureBox4.Visible = True
                PictureBox6.Visible = False
                PictureBox5.Visible = False
                TabPage1.Visible = True
                TabPage2.Visible = False
                TabPage1.Visible = False
                TabPage2.Visible = True
                TabControl1.SelectedTab = TabPage2
            End If
        Catch ex As Exception
            MessageBox.Show("Data not updated successfully", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Label25.Text = "Error while updating '" & txtvientry.Text & "' visit details"
            Label25.ForeColor = System.Drawing.Color.Red
            PictureBox4.Visible = True
            PictureBox6.Visible = False
            PictureBox5.Visible = False

            Me.Dispose()
        End Try
        clear()
    End Sub

    Private Sub btnsearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsearch.Click
        source1.Filter = "[nme] = '" & ComboBox1.Text & "'"
        source2.Filter = "[nme] = '" & ComboBox1.Text & "'"
        DataGridView1.Refresh()
        ComboBox1.Text = ""
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        timer.Text = TimeOfDay
    End Sub
  
    Private Sub gridclick()
        Try
            TabPage1.Visible = True
            Label25.Text = "Edit Values in above textboxes"
            PictureBox4.Visible = False
            PictureBox6.Visible = True
            PictureBox5.Visible = False
            TabControl1.SelectedTab = TabPage1
            txtvientry.Enabled = False
            Panel2.Enabled = True
            Panel2.Enabled = True
            Btnsve.Enabled = False
            Btndel.Enabled = True
            btnupdte.Enabled = True
            Button1.Enabled = True
            Me.txtvientry.Text = DataGridView1.CurrentRow.Cells(0).Value.ToString
            Me.txtvinme.Text = DataGridView1.CurrentRow.Cells(1).Value.ToString
            Me.txtvidte.Text = DataGridView1.CurrentRow.Cells(2).Value.ToString
            Me.txtvipsprt.Text = DataGridView1.CurrentRow.Cells(3).Value.ToString
            Me.txtvitrvldte.Text = DataGridView1.CurrentRow.Cells(4).Value.ToString
            Me.txtexvidte.Text = DataGridView1.CurrentRow.Cells(5).Value.ToString
            Me.txtviprsnt.Text = DataGridView1.CurrentRow.Cells(6).Value.ToString
            Me.txtorgnlcst.Text = DataGridView1.CurrentRow.Cells(7).Value.ToString
            Me.txtsale.Text = DataGridView1.CurrentRow.Cells(8).Value.ToString
            Me.txtprft.Text = DataGridView1.CurrentRow.Cells(9).Value.ToString
            Me.txtvidurtin.Text = DataGridView1.CurrentRow.Cells(10).Value.ToString
            Me.txtvirfrnce.Text = DataGridView1.CurrentRow.Cells(11).Value.ToString
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try

            If Not txtvinme.Text = "" Then
                MessageBox.Show("Are you sure to print data", "Data Adding", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                Dim report As New CRvisit
                Dim objText As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(1).ReportObjects("Text1")
                Dim objText1 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text2")
                Dim objText2 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text3")
                Dim objText3 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text4")
                Dim objText4 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text5")
                Dim objText5 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text6")
                Dim objText6 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text7")
                Dim objText7 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text8")
                Dim objText8 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text9")
                objText.Text = Me.txtvientry.Text
                objText1.Text = Me.txtvinme.Text
                objText2.Text = Me.txtvidte.Text
                objText3.Text = Me.txtvipsprt.Text
                objText4.Text = Me.txtvitrvldte.Text
                objText5.Text = Me.txtexvidte.Text
                objText6.Text = Me.txtviprsnt.Text
                objText7.Text = Me.txtvidurtin.Text
                objText8.Text = Me.txtvirfrnce.Text
                Rprtvisit.CrystalReportViewer1.ReportSource = report
                Rprtvisit.Show()
                Label25.Text = "'" & txtvientry.Text & "' visit details printed successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen
                PictureBox5.Visible = True
                PictureBox4.Visible = False
                PictureBox6.Visible = False

            Else
                MessageBox.Show("Select value from gridview to print", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Label25.Text = "Problem while Printing '" & txtvientry.Text & "' visit details"
                Label25.ForeColor = System.Drawing.Color.Red
                PictureBox4.Visible = True

                PictureBox6.Visible = False
                PictureBox5.Visible = False
                TabPage1.Visible = True
                TabPage2.Visible = False
                TabPage1.Visible = False
                TabPage2.Visible = True
                TabControl1.SelectedTab = TabPage2
            End If
        Catch ex As Exception
            MessageBox.Show("Reports are not loding properly,try again", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Label25.Text = "Error while printing'" & txtvientry.Text & "' visit details"
            Label25.ForeColor = System.Drawing.Color.Red
            PictureBox4.Visible = True
            PictureBox6.Visible = False
            PictureBox5.Visible = False
            Me.Dispose()
        End Try
    End Sub
    Private Sub clear()
        txtvientry.Text = ""
        txtvinme.Text = ""
        txtvidte.Text = ""
        txtvipsprt.Text = ""
        txtvitrvldte.Text = ""
        txtexvidte.Text = ""
        txtviprsnt.Text = ""
        txtvidurtin.Text = ""
        txtvirfrnce.Text = ""
        txtsale.Text = ""
        txtprft.Text = ""
        txtorgnlcst.Text = ""
    End Sub


    Private Sub Btnadd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btnadd.Click
        Label25.Text = " Add new record!"
        Label25.ForeColor = System.Drawing.Color.DarkCyan
        PictureBox6.Visible = True
        PictureBox5.Visible = False
        PictureBox4.Visible = False
        Try
            Panel2.Enabled = True
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
            cmd.CommandText = "SELECT MAX(vientry) FROM visit"
            If (IsDBNull(cmd.ExecuteScalar)) Then
                num = 1
                txtvientry.Text = num.ToString
            Else
                'num = cmd.ExecuteScalar()
                'txtticket.Text = num + 1
                num = cmd.ExecuteScalar + 1
                txtvientry.Text = num.ToString
            End If
            con.Close()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        End Try
    End Sub
    Private Sub txtvientry_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtvientry.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtvientry, "Enter entry in numbers")
    End Sub

    Private Sub txtvinme_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtvinme, "Enter visa name in numbers")
    End Sub

    Private Sub txtvinme_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtvinme, "Enter visa name in letters,digits")
    End Sub

    Private Sub txtvidte_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtvidte.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtvidte, "select date from calender")
    End Sub

    Private Sub txtvidte_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtexvidte.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtvidte, "select date from calender")
    End Sub

    Private Sub txtvipsprt_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtvipsprt.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtvipsprt, "enter passport number in digits text or other symbols")
    End Sub

    Private Sub txtvipsprt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtvipsprt.TextChanged
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtvipsprt, "enter passport number in digits text or other symbols")
    End Sub

    Private Sub txtvitrvldte_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtvitrvldte.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtvitrvldte, "Select travel date from calender")
    End Sub

    Private Sub txtvitrvldte_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtvitrvldte.ValueChanged
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtvitrvldte, "Select travel date from calender")
    End Sub

    Private Sub txtexvidte_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtexvidte.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtvitrvldte, "Select expire date of visa from calender")
    End Sub

    Private Sub txtexvidte_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtexvidte.Validated
        datechk()
    End Sub

    Private Sub txtexvidte_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtexvidte.ValueChanged
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtvitrvldte, "Select expire date of visa from calender")
    End Sub

    Private Sub txtviprsnt_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtviprsnt.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtviprsnt, "Enter in digits,letters or #,., etc")
    End Sub

    Private Sub txtviprsnt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtviprsnt.TextChanged
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtviprsnt, "Enter in digits,letters or #,., etc")
    End Sub

    Private Sub txtvidurtin_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtvidurtin.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtvidurtin, "Enter visa duration in digits,letters or #,., etc")
    End Sub

    Private Sub txtvidurtin_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtvidurtin.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtvidurtin, "Enter visa duration in digits,letters or #,., etc")
    End Sub

    Private Sub txtvirfrnce_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtvirfrnce.MouseUp
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtvirfrnce, "Enter reference in digits,letters or #,., etc")
    End Sub

    Private Sub txtvirfrnce_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtvirfrnce.TextChanged
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtvirfrnce, "Enter reference in digits,letters or #,., etc")
    End Sub

    Private Sub txtvinme_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(txtvinme, "Enter visa name in letters,digits")
    End Sub
    Private Sub DataGridView1_SortCompare(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewSortCompareEventArgs)
        If e.Column.Index <> 0 Then
            Return
        End If
        Try
            e.SortResult = If(CInt(e.CellValue1) < CInt(e.CellValue2), -1, 1)
            e.Handled = True
        Catch
        End Try
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If MsgBox("Are you sure want to exit now?", MsgBoxStyle.YesNo, "Closing warning") = MsgBoxResult.Yes Then ' If you select yes in the MsgBox then it will close the window
            Me.Close()
            Me.Dispose() ' Close the window
        Else
            ' Will not close the application
        End If
    End Sub
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtvientry.Text = Val(txtvientry.Text) - 1
    End Sub
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        If Len(Trim(txtvientry.Text)) = 0 Then
            MessageBox.Show("Please enter Entry number", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtvientry.Focus()
            Exit Sub
        End If
        If Len(Trim(txtvinme.Text)) = 0 Then
            MessageBox.Show("Please enter name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtvinme.Focus()
            Exit Sub
        End If
        If Len(Trim(txtvipsprt.Text)) = 0 Then
            MessageBox.Show("Please enter Passport Number.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtvipsprt.Focus()
            Exit Sub
        End If
        Try

            MessageBox.Show("Are you sure to add data", "Data Adding", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            insert()
            Label25.Text = "'" & txtvientry.Text & "' Visit details saved successfully!"
            Label25.ForeColor = System.Drawing.Color.DarkGreen
            Btnsve.Enabled = False
            Button9.Enabled = False
            Panel2.Enabled = False
        Catch ex As Exception
            ' MessageBox.Show("Data is already exist", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Label25.Text = "Error while saving '" & txtvientry.Text & "' visit details"
            Label25.ForeColor = System.Drawing.Color.Red
            MessageBox.Show("Data already exist, you again select visit Details and Try other entry", "Data Invalid, Application is closing", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
                ' visit( vientry,nme,dte
                ' vientry,nme,dte,psprtno,trvelingdte,expirdte,vipresent,orgnlcst,sale,proft,viduration,refrnce
                Dim cmd As SqlCommand = SqlConn.CreateCommand()
                cmd.CommandText = "SELECT [nme]FROM [visit]"
                Using reader As SqlDataReader = cmd.ExecuteReader
                    While (reader.Read())

                        Me.ListBox1.Items.Add(reader("nme"))
                    End While
                End Using
                SqlConn.Close()
            End Using
        Catch ex As Exception
            MsgBox("Not Loading", MsgBoxStyle.OkOnly, "Error")

        End Try
    End Sub

    Private Sub loaddata_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ListBox1.Items.Clear()
        listboxfill()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call CenterToScreen()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub txtorgnlcst_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtorgnlcst.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub


    Private Sub txtsale_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtsale.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtsale_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtsale.Validated
        profitt()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        gridclick()
    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        e.Graphics.DrawImage(Bitmap, 0, 0)
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
        PrintPreviewDialog1.ShowDialog()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try

            If Not DataGridView1.CurrentRow.IsNewRow Then
                'Query string
                dbaccessconnection()
                con.Open()
                cmd.CommandText = "delete  from visit where vientry='" & DataGridView1.CurrentRow.Cells(0).Value & "'"
                cmd.ExecuteNonQuery()
                DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
                MessageBox.Show("Record Deleted")
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
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
        mResult = MsgBox("Want you really delete the selected records?", _
        vbYesNo + vbQuestion, "Removal confirmation")
        If mResult = vbNo Then
            Exit Sub
        End If
        ObjConnection.ConnectionString = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
        Dim ObjCommand As New SqlCommand()
        ObjCommand.Connection = ObjConnection
        For i = Me.DataGridView1.SelectedRows.Count - 1 To 0 Step -1
            ObjCommand.CommandText = "delete from visit where vientry='" & DataGridView1.SelectedRows(i).Cells("vientry").Value & "'"
            ObjConnection.Open()
            ObjCommand.ExecuteNonQuery()
            ObjConnection.Close()

            Me.DataGridView1.Rows.Remove(Me.DataGridView1.SelectedRows(i))
        Next

    End Sub

    Private Sub Button7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        DeleteSelecedRows()
    End Sub
    Private Sub Button8_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ListBox1.Items.Clear()
        listboxfill()
    End Sub

    Private Sub Label12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label12.Click
        Label12.Text = Date.Today.ToString("dddd")
    End Sub
    Private Sub RemoveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveToolStripMenuItem.Click
        DeleteSelecedRows()
    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        gridclick()
       
    End Sub
    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtvientry.Text = Val(txtvientry.Text) + 1
    End Sub

    Private Sub PictureBox8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox8.Click
        txtvientry.Text = Val(txtvientry.Text) - 1
    End Sub

    Private Sub PictureBox23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox23.Click
        txtvientry.Text = Val(txtvientry.Text) + 1
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
        strQ = " Select vientry as [Entry], nme as [Name] , dte as [Date], psprtno as [Passport Number],trvelingdte as [Travelling Date],expirdte as [Expire Date],vipresent as [Present],orgnlcst as [Orginal cost],viduration as [Duration],refrnce as [Reference]from visit"
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

    Private Sub Button2_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
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

    Private Sub TextBox1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox1.MouseClick
        TextBox1.Text = ""
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class
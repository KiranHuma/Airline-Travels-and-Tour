Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.Data.DataTable
Imports Excel = Microsoft.Office.Interop.Excel '
Imports ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat
Public Class frmumrah
    Private bitmap As Bitmap
    Dim con As New SqlClient.SqlConnection
    Dim cmd As New SqlClient.SqlCommand
    Dim provider As String
    Dim dataFile As String
    Dim connString As String
    Dim myConnection As SqlConnection = New SqlConnection
    Dim ds As DataSet = New DataSet
    Dim da As SqlDataAdapter
    Dim tables As DataTableCollection = ds.Tables
    Dim source1 As New BindingSource()
    Dim source2 As New BindingSource()
    Dim dt As New DataTable

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
        cmd.CommandText = "insert into umrahh(entry,unme,udte,upasprt,utrvlngdte,uexdte,uorgnlcst,usale,uprft,uduratin,upkge,urefrnce)values('" & txtuentry.Text & "','" & txtunme.Text & "','" & txtudte.Value & "','" & txtupsprt.Text & "','" & txtutrvldte.Value & "','" & txtuexdte.Value & "','" & txtuorgcst.Text & "','" & txtusle.Text & "','" & txtuprft.Text & "','" & txtuduratin.Text & "','" & txtupkge.Text & "','" & txturefrnce.Text & "')"
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
    Private Sub del()
        dbaccessconnection()
        con.Open()
        cmd.CommandText = "delete from umrahh where entry=" & txtuentry.Text & ""
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
    Private Sub edit()
        dbaccessconnection()
        con.Open()
        cmd.CommandText = ("UPDATE umrahh SET  unme = '" & txtunme.Text & "', udte= '" & txtudte.Value & "',upasprt= '" & txtupsprt.Text & "',utrvlngdte= '" & txtutrvldte.Value & "',uexdte= '" & txtuexdte.Value & "',uorgnlcst='" & txtuorgcst.Text & "',usale='" & txtusle.Text & "',uprft='" & txtuprft.Text & "',uduratin='" & txtuduratin.Text & "',upkge='" & txtupkge.Text & "',urefrnce='" & txturefrnce.Text & "'   where entry=" & txtuentry.Text & "")
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
    Private Sub Btnadd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btnadd.Click
        Label25.Text = " Add new record!"
        Label25.ForeColor = System.Drawing.Color.DarkCyan
        PictureBox6.Visible = True
        PictureBox5.Visible = False
        PictureBox8.Visible = False
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
            cmd.CommandText = "SELECT MAX(entry) FROM umrahh"
            If (IsDBNull(cmd.ExecuteScalar)) Then
                num = 1
                txtuentry.Text = num.ToString
            Else
                'num = cmd.ExecuteScalar()
                'txtticket.Text = num + 1
                num = cmd.ExecuteScalar + 1
                txtuentry.Text = num.ToString
            End If
            con.Close()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        End Try
    End Sub

    Private Sub btnupdte_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnupdte.Click
        Try
            If Not txtuentry.Text = "" Then
                MessageBox.Show("Are you sure to update data", "Data Updating", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                edit()
                FillCombo()
                getdata()
                Label25.Enabled = True
                Label25.Text = "'" & txtuentry.Text & "' Umrah details updated successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen
                PictureBox5.Visible = False
                PictureBox8.Visible = True
                PictureBox6.Visible = False
                GroupBox1.Enabled = False
                Btnadd.Enabled = True
                Me.Refresh()
            Else
                MessageBox.Show("Select rows from grid to edit", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Label25.Text = "Error while updating '" & txtuentry.Text & "' Umrah details"
                Label25.ForeColor = System.Drawing.Color.Red
                PictureBox8.Visible = False
                PictureBox6.Visible = False
                PictureBox5.Visible = True
                TabPage1.Visible = True
                TabPage2.Visible = False
                TabPage1.Visible = False
                TabPage2.Visible = True
                TabControl1.SelectedTab = TabPage2
            End If
        Catch ex As Exception
            MessageBox.Show("Data not updated successfully", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Label25.Text = "Error while updating '" & txtuentry.Text & "' Umrah details"
            Label25.ForeColor = System.Drawing.Color.Red
            PictureBox8.Visible = False
            PictureBox6.Visible = False
            PictureBox5.Visible = True
            Me.Dispose()
        End Try

        Me.Refresh()
        clear()
    End Sub

    Private Sub Btndel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btndel.Click
        Try
            If Not txtuentry.Text = "" Then
                MessageBox.Show("Are you sure to delete data", "Data Deleting", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                del()
                FillCombo()
                getdata()
                Label25.Text = "'" & txtuentry.Text & "'umrah details removed successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen
                PictureBox5.Visible = False
                PictureBox6.Visible = False
                PictureBox8.Visible = True
                GroupBox1.Enabled = False
                Btnadd.Enabled = True

            Else
                MessageBox.Show("Select rows from grid to remove", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Label25.Text = "Error while removing '" & txtuentry.Text & "' umrah details"
                Label25.ForeColor = System.Drawing.Color.Red
                PictureBox8.Visible = False
                PictureBox6.Visible = False
                PictureBox5.Visible = True
                TabPage1.Visible = True
                TabPage2.Visible = False
                TabPage1.Visible = False
                TabPage2.Visible = True
                TabControl1.SelectedTab = TabPage2
            End If
        Catch ex As Exception
            'MessageBox.Show("Data is not remove succesfully", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Label25.Text = "Error while removing '" & txtuentry.Text & "' Umrah details"
            Label25.ForeColor = System.Drawing.Color.Red
            PictureBox8.Visible = False
            PictureBox6.Visible = False
            PictureBox5.Visible = True
            Me.Dispose()
        End Try
        Me.Refresh()
        clear()
    End Sub


    Private Sub frmumrah_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button2.Visible = False
        Button2.Enabled = False
        If (DateTime.Now.Hour < 12) Then
            PictureBox2.Visible = False
            PictureBox3.Visible = False
            PictureBox4.Visible = True
            PictureBox7.Visible = False
            lblgrting.Text = "Good Morning"
            'Label12.Text = Convert.ToString(DateTime.Now)

        ElseIf (DateTime.Now.Hour < 17) Then
            PictureBox2.Visible = False
            PictureBox3.Visible = True
            PictureBox4.Visible = False
            PictureBox7.Visible = False
            lblgrting.Text = "Good Afternoon"
            ' Label12.Text = Convert.ToString(DateTime.Now)
        ElseIf (DateTime.Now.Hour < 19) Then
            PictureBox2.Visible = False
            PictureBox3.Visible = False
            PictureBox4.Visible = False
            PictureBox7.Visible = True
            lblgrting.Text = "Good Evening"
        Else
            PictureBox2.Visible = True
            PictureBox3.Visible = False
            PictureBox4.Visible = False
            PictureBox7.Visible = False
            Label17.Visible = False
            ' lblgrting.Text = "Good Night"
            ' Label12.Text = Convert.ToString(DateTime.Now)
        End If
        Label17.Text = Date.Today.ToString("dddd")
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
            Dim myConnToAccess As SqlConnection
            Dim ds As DataSet
            Dim da As SqlDataAdapter
            Dim tables As DataTableCollection
            myConnToAccess = New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
            myConnToAccess.Open()
            ds = New DataSet
            tables = ds.Tables
            da = New SqlDataAdapter("SELECT unme from umrahh ", myConnToAccess)
            da.Fill(ds, "umrahh ")
            Dim view1 As New DataView(tables(0))
            With ComboBox1
                .DataSource = ds.Tables("umrahh ")
                .DisplayMember = "unme"
                .ValueMember = "unme"
                .SelectedIndex = -1
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .AutoCompleteSource = AutoCompleteSource.ListItems
            End With
        Catch ex As Exception
            MsgBox("Search have problem!!!!")
        End Try

    End Sub
    Private Sub FillCombo2()
        Dim myConnToAccess As SqlConnection
        Dim ds As DataSet
        Dim da As SqlDataAdapter
        Dim tables As DataTableCollection
        myConnToAccess = New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
        myConnToAccess.Open()
        ds = New DataSet
        tables = ds.Tables
        da = New SqlDataAdapter("SELECT unme from umrahh ", myConnToAccess)
        da.Fill(ds, "umrahh ")
        Dim view1 As New DataView(tables(0))
        With txtunme
            .DataSource = ds.Tables("umrahh ")
            .DisplayMember = "unme"
            .ValueMember = "unme"
            .SelectedIndex = -1
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.ListItems
        End With
    End Sub
    Private Sub gridfill()
        Try
            provider = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            'entry,unme,udte,upasprt,utrvlngdte,uexdte,uorgnlcst,usale,uprft,uduratin,upkge,urefrnce)
            da = New SqlDataAdapter("Select [entry], [unme] , udte as [Date], upasprt as [Passport Number],utrvlngdte as [Travelling Date],uexdte as [Expire Date],uorgnlcst as [Orginal cost],usale as [Sale],uprft as[Profit],uduratin as [Duration],upkge as [Package],urefrnce as [Reference] from umrahh", myConnection)
            da.Fill(ds, "umrahh")
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
        Dim da As New SqlDataAdapter("Select [entry], [unme] , udte as [Date], upasprt as [Passport Number],utrvlngdte as [Travelling Date],uexdte as [Expire Date],uorgnlcst as [Orginal cost],usale as [Sale],uprft as[Profit],uduratin as [Duration],upkge as [Package],urefrnce as [Reference] from umrahh ", con)
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
        dteissue = txtutrvldte.Value
        dteex = txtuexdte.Value
        If dteex < dteissue Then
            MessageBox.Show("Expire date must be greater than isssue date!", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
    Private Sub profit()
        Try
            Dim minus As Integer
            Dim a As Integer
            Dim b As Integer
            a = txtuorgcst.Text
            b = txtusle.Text
            minus = b - a
            txtuprft.Text = minus
            'MessageBox.Show(minus)
        Catch ex As Exception
            MsgBox("Orignal cost and sale must not empty")
        End Try
    End Sub

    Private Sub btnsearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsearch.Click
        source1.Filter = "[unme] = '" & ComboBox1.Text & "'"
        source2.Filter = "[unme] = '" & ComboBox1.Text & "'"
        DataGridView1.Refresh()
        ComboBox1.Text = ""
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        timer.Text = TimeOfDay
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Panel2.Visible = True
        Panel2.Enabled = True
        'entry,unme,udte,upasprt,utrvlngdte,uexdte,uorgnlcst,usale,uprft,uduratin,upkge,urefrnce
        gridclick()
    End Sub
    Private Sub gridclick()
        Try
            TabPage1.Visible = True
            Label25.Text = "Edit Values in above textboxes"
            PictureBox8.Visible = False
            PictureBox6.Visible = True
            PictureBox5.Visible = False
            TabControl1.SelectedTab = TabPage1
            GroupBox1.Enabled = True
            Panel2.Visible = True
            txtuentry.Enabled = False
            Btnsve.Enabled = False
            Btndel.Enabled = True
            btnupdte.Enabled = True
            GroupBox3.Visible = False
            Me.txtuentry.Text = DataGridView1.CurrentRow.Cells(0).Value.ToString
            Me.txtunme.Text = DataGridView1.CurrentRow.Cells(1).Value.ToString
            Me.txtudte.Text = DataGridView1.CurrentRow.Cells(2).Value.ToString
            Me.txtupsprt.Text = DataGridView1.CurrentRow.Cells(3).Value.ToString
            Me.txtutrvldte.Text = DataGridView1.CurrentRow.Cells(4).Value.ToString
            Me.txtuexdte.Text = DataGridView1.CurrentRow.Cells(5).Value.ToString
            Me.txtuorgcst.Text = DataGridView1.CurrentRow.Cells(6).Value.ToString
            Me.txtusle.Text = DataGridView1.CurrentRow.Cells(7).Value.ToString
            Me.txtuprft.Text = DataGridView1.CurrentRow.Cells(8).Value.ToString
            Me.txtuduratin.Text = DataGridView1.CurrentRow.Cells(9).Value.ToString
            Me.txtupkge.Text = DataGridView1.CurrentRow.Cells(10).Value.ToString
            Me.txturefrnce.Text = DataGridView1.CurrentRow.Cells(11).Value.ToString
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub clear()
        txtuentry.Text = ""
        txtunme.Text = ""
        txtudte.Text = ""
        txtupsprt.Text = ""
        txtutrvldte.Text = ""
        txtuexdte.Text = ""
        txtuorgcst.Text = ""
        txtusle.Text = ""
        txtuprft.Text = ""
        txtuduratin.Text = ""
        txtupkge.Text = ""
        txturefrnce.Text = ""
    End Sub

    Private Sub Btnsve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btnsve.Click
        If Len(Trim(txtuentry.Text)) = 0 Then
            MessageBox.Show("Please enter Entry number", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtuentry.Focus()
            Exit Sub
        End If
        If Len(Trim(txtupsprt.Text)) = 0 Then
            MessageBox.Show("Please enter Passport Number", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtupsprt.Focus()
            Exit Sub
        End If
        If Len(Trim(txtuorgcst.Text)) = 0 Then
            MessageBox.Show("Please enter Original Cost.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtuorgcst.Focus()
            Exit Sub
        End If
        Me.Refresh()
        Try
            If Not txtuentry.Text = "" Then
                MessageBox.Show("Are you sure to add data", "Data Adding", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                insert()
                getdata()
                FillCombo()
                Label25.Text = "'" & txtuentry.Text & "' Umrah details saved successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen
                PictureBox5.Visible = False
                PictureBox8.Visible = True
                PictureBox6.Visible = False
                Btnsve.Enabled = False
                Btnadd.Enabled = True
                Btndel.Enabled = True
                btnupdte.Enabled = True
                Button1.Enabled = True
                Button9.Enabled = True
                GroupBox1.Enabled = False

            Else

                MessageBox.Show("please fill all above textboxes", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Label25.Text = "Problem while saving '" & txtuentry.Text & "' Umrah details"
                Label25.ForeColor = System.Drawing.Color.Red
                PictureBox8.Visible = False
                PictureBox6.Visible = False
                PictureBox5.Visible = True
            End If
        Catch ex As Exception
            Label25.Text = "Error while saving '" & txtuentry.Text & "' Umrah details"
            Label25.ForeColor = System.Drawing.Color.Red
            MessageBox.Show("Data already exist, you again select umrah Details and Try other entry", "Data Invalid, Application is closing", MessageBoxButtons.OK, MessageBoxIcon.Error)
            PictureBox8.Visible = False
            PictureBox6.Visible = False
            PictureBox5.Visible = True
            Me.Dispose()
        End Try
        GroupBox1.Enabled = False
        clear()
        Me.Refresh()
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If Not txtunme.Text = "" Then
                MessageBox.Show("Are you sure to print data", "Data Adding", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                Dim report As New CRumrah
                Dim objText As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(1).ReportObjects("Text1")
                Dim objText1 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text2")
                Dim objText2 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text3")
                Dim objText3 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text4")
                Dim objText4 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text5")
                Dim objText5 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text6")
                Dim objText6 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text7")
                Dim objText7 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text8")
                Dim objText8 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text27")
                Dim objText9 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text28")

                objText.Text = Me.txtuentry.Text
                objText1.Text = Me.txtunme.Text
                objText2.Text = Me.txtudte.Text
                objText3.Text = Me.txtupsprt.Text
                objText4.Text = Me.txtutrvldte.Text
                objText5.Text = Me.txtuexdte.Text
                objText6.Text = Me.txtuorgcst.Text
                objText7.Text = Me.txtuduratin.Text
                objText8.Text = Me.txtupkge.Text
                objText9.Text = Me.txturefrnce.Text
                Rprtumrah.CrystalReportViewer1.ReportSource = report
                Rprtumrah.Show()
                Label25.Text = "'" & txtuentry.Text & "' umrah details printed successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen
                PictureBox5.Visible = False
                PictureBox8.Visible = True
                PictureBox6.Visible = False

            Else
                MessageBox.Show("Select value from gridview to print", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Label25.Text = "Problem while Printing '" & txtuentry.Text & "' Umrah details"
                Label25.ForeColor = System.Drawing.Color.Red
                PictureBox8.Visible = False
                PictureBox6.Visible = False
                PictureBox5.Visible = True
                TabPage1.Visible = True
                TabPage2.Visible = False
                TabPage1.Visible = False
                TabPage2.Visible = True
                TabControl1.SelectedTab = TabPage2
            End If
        Catch ex As Exception
            MessageBox.Show("Reports are not loding properly,try again", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Label25.Text = "Error while printing'" & txtuentry.Text & "' Umrah details"
            Label25.ForeColor = System.Drawing.Color.Red
            PictureBox8.Visible = False
            PictureBox6.Visible = False
            PictureBox5.Visible = True
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
                cmd.CommandText = "Select unme FROM [umrahh]"
                Using reader As SqlDataReader = cmd.ExecuteReader
                    While (reader.Read())
                        Me.ListBox1.Items.Add(reader("unme"))
                    End While
                End Using
                SqlConn.Close()
            End Using
        Catch ex As Exception
            MsgBox("DataBase not connected due to the reason because " & ex.Message)
            ' MsgBox("Not Loading", MsgBoxStyle.OkOnly, "Error")

        End Try
    End Sub


    Private Sub TabPage2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage2.Click
        GroupBox3.Visible = True
    End Sub
    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        ' MessageBox.Show(TabControl1.SelectedTab.Text)

        Select Case TabControl1.SelectedIndex

            Case 0 ' User clicks on First Tab
                If Panel2.Enabled = True Then
                    Button10.Enabled = False
                    Button10.Visible = False
                    Button2.Visible = True
                    Button2.Enabled = True
                    Label12.Text = "Hide Details"
                Else
                    Button2.Visible = False
                    Button2.Enabled = False
                    Button10.Enabled = True
                    Button10.Visible = True
                    GroupBox3.Visible = False
                    GroupBox3.Enabled = False
                    GroupBox6.Visible = False
                    GroupBox6.Enabled = False
                    Panel2.Visible = False
                    Label12.Text = "Show Details"
                End If

            Case 1 ' User clicks on Second Tab
                GroupBox3.Visible = True
                GroupBox3.Enabled = True
                GroupBox6.Visible = False
                GroupBox6.Enabled = False
                Panel2.Visible = False
                Panel2.Enabled = False

            Case 2 ' User clicks on Third Tab
                GroupBox3.Visible = False
                GroupBox3.Enabled = False
                GroupBox6.Visible = True
                GroupBox6.Enabled = True
                Panel2.Visible = False
                Panel2.Enabled = False

            Case 4 ' User clicks on Fourth Tab
                ' code to do here

        End Select

    End Sub
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        GroupBox3.Visible = False
        GroupBox3.Enabled = False
        GroupBox6.Visible = True
        GroupBox6.Enabled = True
        Panel2.Visible = False
        Panel2.Enabled = False

    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Panel2.Visible = True
        Panel2.Enabled = True
        GroupBox3.Visible = False
        GroupBox3.Enabled = False
        GroupBox6.Visible = False
        GroupBox6.Enabled = False
        Btnadd.Enabled = True
        Btndel.Enabled = True
        btnupdte.Enabled = True
        Button1.Enabled = True
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Call CenterToScreen()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub txtuorgcst_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtuorgcst.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtusle_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtusle.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtusle_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtusle.Validated
        profit()
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
        PictureBox12.Visible = False
        PictureBox13.Visible = False
        PictureBox14.Visible = True
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
            ObjCommand.CommandText = "delete from umrahh where entry='" & DataGridView1.SelectedRows(i).Cells("entry").Value & "'"
            ObjConnection.Open()
            ObjCommand.ExecuteNonQuery()
            ObjConnection.Close()

            Me.DataGridView1.Rows.Remove(Me.DataGridView1.SelectedRows(i))
        Next

    End Sub
    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        GroupBox1.Enabled = True
        Panel2.Enabled = True
        gridclick()
    End Sub

    Private Sub RemoveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveToolStripMenuItem.Click
        DeleteSelecedRows()
        PictureBox12.Visible = True
        PictureBox13.Visible = False
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        If Len(Trim(txtuentry.Text)) = 0 Then
            MessageBox.Show("Please enter Entry number", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtuentry.Focus()
            Exit Sub
        End If
        If Len(Trim(txtupsprt.Text)) = 0 Then
            MessageBox.Show("Please enter Passport Number", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtupsprt.Focus()
            Exit Sub
        End If
        If Len(Trim(txtuorgcst.Text)) = 0 Then
            MessageBox.Show("Please enter Original Cost.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtuorgcst.Focus()
            Exit Sub
        End If
        Me.Refresh()
        Try

            MessageBox.Show("Are you sure to add data", "Data Adding", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            insert()
            Label25.Text = "'" & txtuentry.Text & "'umrah details saved successfully!"
            Label25.ForeColor = System.Drawing.Color.DarkGreen
            PictureBox5.Visible = False
            PictureBox8.Visible = True
            PictureBox6.Visible = False
            Btnsve.Enabled = False
            Button9.Enabled = False
            GroupBox1.Enabled = False
        Catch ex As Exception
            ' MessageBox.Show("Data is already exist", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Label25.Text = "Error while saving '" & txtuentry.Text & "'umrah details"
            Label25.ForeColor = System.Drawing.Color.Red
            PictureBox8.Visible = False
            PictureBox6.Visible = False
            PictureBox5.Visible = True
            MessageBox.Show("Data already exist, you again select umrah  Details and Try other entry", "Data Invalid, Application is closing", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Dispose()
        End Try
        clear()
        Me.Refresh()
    End Sub

    Private Sub txtuexdte_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtuexdte.Validated
        datechk()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Panel2.Visible = False
        Panel2.Enabled = False
        GroupBox3.Visible = False
        GroupBox3.Enabled = False
        GroupBox6.Visible = False
        GroupBox6.Enabled = False
        Btnadd.Enabled = False
        Btndel.Enabled = False
        btnupdte.Enabled = False
        Button1.Enabled = False
    End Sub

    Private Sub PictureBox11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox11.Click
        TabPage1.Visible = True
        TabPage2.Visible = False
        TabPage1.Visible = True
        TabPage2.Visible = False
        TabControl1.SelectedTab = TabPage1
        Button2.Visible = False
        Button2.Enabled = False
        PictureBox6.Visible = False
        PictureBox5.Visible = False
        PictureBox8.Visible = False
        Button10.Enabled = True
        Button10.Visible = True
        GroupBox3.Visible = False
        GroupBox3.Enabled = False
        GroupBox6.Visible = False
        GroupBox6.Enabled = False
        Panel2.Visible = False
        Label12.Text = "Show Details"
    End Sub

    Private Sub loaddata_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles loaddata.Click
        Try
            ListBox1.Items.Clear()
            listboxfill()
            listview()
            Label18.Text = "Umrah details loaded successfully!"
            Label18.ForeColor = System.Drawing.Color.DarkGreen
            
            Label20.Text = "Click to Hide Data"
            Label20.ForeColor = System.Drawing.Color.Black
        Catch ex As Exception
            Label18.Text = "Umrah details not loaded successfully!"
            Label18.ForeColor = System.Drawing.Color.Red
        End Try
    End Sub

    Private Sub Button3_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            ListBox1.Items.Clear()
            ListView1.Items.Clear()
            ListView2.Items.Clear()
            TextBox1.Text = ""
            Label20.Text = "Umrah details Unloaded successfully!"
            Label20.ForeColor = System.Drawing.Color.DarkGreen
            Label18.Text = "Click to view data!"
            Label18.ForeColor = System.Drawing.Color.Black
        Catch ex As Exception
            Label20.Text = "Umrah details not Unloaded successfully!"
            Label20.ForeColor = System.Drawing.Color.Red
        End Try
    End Sub

    Private Sub EmptyRecycleBinToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmptyRecycleBinToolStripMenuItem.Click
        PictureBox12.Visible = False
        PictureBox13.Visible = True
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        DeleteSelecedRows()
        FillCombo()
        PictureBox12.Visible = True
        PictureBox13.Visible = False
        PictureBox14.Visible = False
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        txtuentry.Text = Val(txtuentry.Text) + 1
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        txtuentry.Text = Val(txtuentry.Text) - 1
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
        strQ = "Select entry as [Entry],unme as [Name] , udte as [Date], upasprt as [Passport Number],utrvlngdte as [Travelling Date],uexdte as [Expire Date],uorgnlcst as [Orginal cost],uduratin as [Duration],upkge as [Package],urefrnce as [Reference] from umrahh"
        cmd = New SqlCommand(strQ, conn)
        da = New SqlDataAdapter(cmd)
        ds = New DataSet
        da.Fill(ds, "umrahh")
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

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        searchlstview()
        movoneitm()
    End Sub
End Class

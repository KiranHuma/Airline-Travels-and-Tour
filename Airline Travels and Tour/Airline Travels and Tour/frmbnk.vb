Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.Data.DataTable
Imports System.IO
Imports System.Drawing.Imaging
Imports System.Security.Cryptography
Imports System.Text
Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel '
Imports ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat
Public Class frmbnk
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
    'Dim con As New OleDb.OleDbConnection
    ' Dim cmd As New OleDb.OleDbCommand
    Dim con As New SqlClient.SqlConnection
    Dim cmd As New SqlClient.SqlCommand
    '///////////////////////////////////////////////////////
    'Private Sub dbaccessconnection()
    'Acces DataBase Connectivity and for MS Access 2003 PROVIDER=Microsoft.Jet.OLEDB.4.0
    ' Try
    '  con.ConnectionString = "provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb"
    '  cmd.Connection = con
    'MessageBox.Show("connection created")
    ' Catch ex As Exception
    '   MsgBox("DataBase not connected due to the reason because " & ex.Message)
    ' End Try
    ' End Sub
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
        cmd.CommandText = "insert into bankdetails( bnknme,Transctinnmbr,dte,accntnmbr,amount,accnthldr,photo)values('" & txtbnknme.Text & "','" & txttrnsctin.Text & "','" & txtbankdte.Value & "','" & txtaccnt.Text & "','" & txtamnt.Text & "','" & txtacnthodr.Text & "',@photo)"
        '?/////////////////////////////////////for sql//////////////////////////////////
        Dim ms As New MemoryStream()
        Dim bmpImage As New Bitmap(photo.Image)
        bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
        Dim data As Byte() = ms.GetBuffer()
        Dim p As New SqlParameter("@photo", SqlDbType.Image)
        p.Value = data
        cmd.Parameters.Add(p)
        cmd.ExecuteNonQuery()
        con.Close()



    End Sub
    Private Sub del()
        dbaccessconnection()
        con.Open()
        cmd.CommandText = "delete from bankdetails where Transctinnmbr=" & txttrnsctin.Text & ""
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
    Private Sub edit()

        dbaccessconnection()
        con.Open()
        cmd.CommandText = ("UPDATE bankdetails SET  bnknme = '" & txtbnknme.Text & "', dte= '" & txtbankdte.Value & "',accntnmbr= '" & txtaccnt.Text & "',amount= '" & txtamnt.Text & "',accnthldr= '" & txtacnthodr.Text & "', photo=@photo where Transctinnmbr=" & txttrnsctin.Text & "")
        Dim ms As New MemoryStream()
        Dim bmpImage As New Bitmap(photo.Image)
        bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
        Dim data As Byte() = ms.GetBuffer()
        Dim p As New SqlParameter("@photo", SqlDbType.Image)
        p.Value = data
        cmd.Parameters.Add(p)
        '/////////////////////////////////////for access image its type is varbinary//////////////////////////
        ' cmd.CommandText = ("UPDATE bankdetails SET  bnknme = '" & txtbnknme.Text & "', dte= '" & txtbankdte.Value & "',accntnmbr= '" & txtaccnt.Text & "',amount= '" & txtamnt.Text & "',accnthldr= '" & txtacnthodr.Text & "', photo=@d1 where Transctinnmbr=" & txttrnsctin.Text & "")
        'Dim ms As New MemoryStream()
        'Dim bmpImage As New Bitmap(photo.Image)
        'bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
        'Dim data As Byte() = ms.GetBuffer()
        ' Dim p As New SqlParameter("@d1", OleDbType.VarBinary)
        ' p.Value = data
        'cmd.Parameters.Add(p)
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
    Private Sub btnupload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnupload.Click
        Try
            With OpenFileDialog1
                .Filter = ("Images |*.png; *.bmp; *.jpg;*.jpeg; *.gif;")
                .FilterIndex = 4
            End With
            'Clear the file name
            OpenFileDialog1.FileName = ""
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                photo.Image = Image.FromFile(OpenFileDialog1.FileName)
            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub
    Private Sub frmbnk_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If (DateTime.Now.Hour < 12) Then
            lblgrting.Text = "Good Morning"
            'Label12.Text = Convert.ToString(DateTime.Now)

        ElseIf (DateTime.Now.Hour < 17) Then
            
            lblgrting.Text = "Good Afternoon"
            ' Label12.Text = Convert.ToString(DateTime.Now)
        ElseIf (DateTime.Now.Hour < 19) Then
           
            lblgrting.Text = "Good Evening"
        Else
            lblgrting.Text = "Good Night"
            ' Label12.Text = Convert.ToString(DateTime.Now)
        End If
        Label8.Text = Date.Today.ToString("dddd")
        'Call CenterToScreen()
        'Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        'Me.WindowState = FormWindowState.Maximized
        gridfill()
        dbaccessconnection()
        Panel1.Enabled = False
        FillCombo()
        FillCombo2()
        ComboBox1.Text = "----Select Or Type Name---"
        ' DataGridView1.Sort(DataGridView1.Columns("entry"), System.ComponentModel.ListSortDirection.Ascending)
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
            ' myConnToAccess = New OleDbConnection("provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb")
            myConnToAccess = New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
            myConnToAccess.Open()
            ds = New DataSet
            tables = ds.Tables
            'da = New OleDbDataAdapter("SELECT Transctinnmbr from bankdetails", myConnToAccess)
            da = New SqlDataAdapter("SELECT Transctinnmbr from bankdetails", myConnToAccess)
            da.Fill(ds, "bankdetails")
            Dim view1 As New DataView(tables(0))
            With ComboBox1
                .DataSource = ds.Tables("bankdetails")
                .DisplayMember = "Transctinnmbr"
                .ValueMember = "Transctinnmbr"
                .SelectedIndex = -1
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .AutoCompleteSource = AutoCompleteSource.ListItems
            End With
        Catch ex As Exception
            MsgBox("Search have problem!!!!")
            Me.Dispose()
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
            ' myConnToAccess = New OleDbConnection("provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb")
            myConnToAccess = New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
            myConnToAccess.Open()
            ds = New DataSet
            tables = ds.Tables
            'da = New OleDbDataAdapter("SELECT Transctinnmbr from bankdetails", myConnToAccess)
            da = New SqlDataAdapter("SELECT  bnknme from bankdetails", myConnToAccess)
            da.Fill(ds, "bankdetails")
            Dim view1 As New DataView(tables(0))
            With txtbnknme
                .DataSource = ds.Tables("bankdetails")
                .DisplayMember = "bnknme"
                .ValueMember = "bnknme"
                .SelectedIndex = -1
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .AutoCompleteSource = AutoCompleteSource.ListItems
            End With
        Catch ex As Exception
            MsgBox("Search have problem!!!!")
            Me.Dispose()
        End Try
    End Sub
    Private Sub gridfill()
        'dbaccessconnection()
        Try
            Me.Label23.Text = Format(Now, "dd-MMM-yyyy")
            Timer1.Enabled = True
            'provider = "provider=Microsoft.ACE.Oledb.12.0;Data Source=airline.accdb"
            provider = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            '( bnknme,Transctinnmbr,dte,accntnmbr,amount,accnthldr
            ' da = New OleDbDataAdapter("Select bnknme as [Bank Name], [Transctinnmbr], dte as[Date] ,accntnmbr as [Account Number],amount as[Amount],accnthldr as[Account Holder],photo as [Photo]from bankdetails ", myConnection)
            da = New SqlDataAdapter("Select bnknme as [Bank Name], [Transctinnmbr], dte as[Date] ,accntnmbr as [Account Number],amount as[Amount],accnthldr as[Account Holder],photo as [Photo]from bankdetails ", myConnection)
            da.Fill(ds, "bankdetails")
            Dim view1 As New DataView(tables(0))
            source1.DataSource = view1
            DataGridView1.DataSource = view1
            DataGridView1.Refresh()
        Catch ex As Exception
            MsgBox("Form is not loaded succesfuly!!!!!Try again")
        End Try
    End Sub
    Private Sub getdata()

        Dim con As New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
        con.Open()
        Dim da As New SqlDataAdapter("Select bnknme as [Bank Name], [Transctinnmbr], dte as[Date] ,accntnmbr as [Account Number],amount as[Amount],accnthldr as[Account Holder],photo as [Photo]from bankdetails ", con)
        Dim dt As New DataTable
        da.Fill(dt)
        'Dim view1 As New DataView(tables1(0))
        source2.DataSource = dt
        DataGridView1.DataSource = dt
        DataGridView1.Refresh()
    End Sub

    Private Sub Btnsve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btnsve.Click
        If Len(Trim(txtbnknme.Text)) = 0 Then
            MessageBox.Show("Please enter Bank Name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtbnknme.Focus()
            Exit Sub
        End If
        If Len(Trim(txtaccnt.Text)) = 0 Then
            MessageBox.Show("Please enter Account#", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtaccnt.Focus()
            Exit Sub
        End If
        If Len(Trim(txtbankdte.Text)) = 0 Then
            MessageBox.Show("Please enter date", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtbankdte.Focus()
            Exit Sub
        End If
        If Len(Trim(txtamnt.Text)) = 0 Then
            MessageBox.Show("Please enter Amount", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtamnt.Focus()
            Exit Sub
        End If
        If Len(Trim(txtacnthodr.Text)) = 0 Then
            MessageBox.Show("Please enter Account Holder", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtacnthodr.Focus()
            Exit Sub
        End If
        ' If photo.Image = Nothing Then
        'MessageBox.Show("Please enter Bank Name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'txtbnknme.Focus()
        'End If
        Try
            If Not photo.Image Is Nothing Then
                MessageBox.Show("Are you sure to add data", "Data Adding", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                insert()
                getdata()
                FillCombo()
                Label25.Text = "'" & txttrnsctin.Text & "' Bank details saved successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen
                Btnsve.Enabled = False
                Btnadd.Enabled = True
                Btndel.Enabled = True
                btnupdte.Enabled = True
                Button1.Enabled = True
                Button9.Enabled = True
                Panel1.Enabled = False
                clear()
            Else

                MessageBox.Show("please fill all above textboxes and picturebox", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Label25.Text = "Problem while saving '" & txttrnsctin.Text & "' Bank details"
                Label25.ForeColor = System.Drawing.Color.Red
            End If

        Catch ex As Exception
            Label25.Text = "Error while saving '" & txttrnsctin.Text & "' Bank details"
            Label25.ForeColor = System.Drawing.Color.Red
            MsgBox("DataBase not connected due to the reason because " & ex.Message)
            MessageBox.Show("Data already exist, you again select Bank Details and Try other entry", "Data Invalid, Application is closing", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Dispose()
        End Try

    End Sub

    Private Sub Btndel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btndel.Click

        Me.Refresh()
        Try
            If Not txtbnknme.Text = "" Then
                MessageBox.Show("Are you sure to delete data", "Data Deleting", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                del()
                FillCombo()
                getdata()
                Label25.Text = "'" & txttrnsctin.Text & "'Bank details removed successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen
                Panel2.Enabled = False
                Btnadd.Enabled = True

            Else
                MessageBox.Show("Select rows from grid to remove", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Label25.Text = "Error while removing '" & txttrnsctin.Text & "' Bank details"
                Label25.ForeColor = System.Drawing.Color.Red
                Panel1.Visible = False
                Panel1.Enabled = False
                GroupBox3.Visible = True
                GroupBox3.Enabled = True
                GroupBox5.Visible = False
                GroupBox5.Enabled = False
                Button5.Visible = False
                Button5.Enabled = False
                Label9.Visible = False
                Label7.Visible = False
                Btnadd.Visible = False
                Btnadd.Enabled = False
                Label19.Visible = False
                btnupdte.Enabled = False
                btnupdte.Visible = False
                Label7.Visible = False
                Btndel.Visible = False
                Btndel.Enabled = False
                Label15.Visible = False
                Btnsve.Visible = False
                Btnsve.Enabled = False
                Label14.Visible = False
                Button1.Visible = False
                Button1.Enabled = False
                Label5.Visible = False



                Label9.Visible = False
                Button6.Visible = False
                Button5.Visible = False
                Button2.Visible = True
                Label22.Visible = True
                Label10.Visible = False
                Button4.Visible = True

                Button6.Visible = False
                Button5.Enabled = False
                Button2.Enabled = False
                Button4.Enabled = True
                Label26.Visible = True
            End If
        Catch ex As Exception
            'MessageBox.Show("Data is not remove succesfully", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Label25.Text = "Error while removing '" & txttrnsctin.Text & "'Bank details"
            Label25.ForeColor = System.Drawing.Color.Red
            Me.Dispose()
        End Try
        clear()
    End Sub

    Private Sub btnupdte_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnupdte.Click
        Me.Refresh()
        Try
            If Not txtbnknme.Text = "" Then
                MessageBox.Show("Are you sure to update data", "Data Updating", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                edit()
                FillCombo()
                getdata()
                Label25.Enabled = True
                Label25.Text = "'" & txtbnknme.Text & "' Bank details updated successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen
                Panel2.Enabled = False
                Btnadd.Enabled = True
                Me.Refresh()
            Else
                MessageBox.Show("Select rows from grid to edit", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Label25.Text = "Error while updating '" & txtbnknme.Text & "' Bank details"
                Label25.ForeColor = System.Drawing.Color.Red
                Panel1.Visible = False
                Panel1.Enabled = False
                GroupBox3.Visible = True
                GroupBox3.Enabled = True
                GroupBox5.Visible = False
                GroupBox5.Enabled = False
                Button5.Visible = False
                Button5.Enabled = False
                Label9.Visible = False
                Label7.Visible = False
                Btnadd.Visible = False
                Btnadd.Enabled = False
                Label19.Visible = False
                btnupdte.Enabled = False
                btnupdte.Visible = False
                Label7.Visible = False
                Btndel.Visible = False
                Btndel.Enabled = False
                Label15.Visible = False
                Btnsve.Visible = False
                Btnsve.Enabled = False
                Label14.Visible = False
                Button1.Visible = False
                Button1.Enabled = False
                Label5.Visible = False



                Label9.Visible = False
                Button6.Visible = False
                Button5.Visible = False
                Button2.Visible = True
                Label22.Visible = True
                Label10.Visible = False
                Button4.Visible = True

                Button6.Visible = False
                Button5.Enabled = False
                Button2.Enabled = False
                Button4.Enabled = True
                Label26.Visible = True
            End If
        Catch ex As Exception
            MessageBox.Show("Data not updated successfully", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Label25.Text = "Error while updating '" & txtbnknme.Text & "' Bank details"
            Label25.ForeColor = System.Drawing.Color.Red
            Me.Dispose()
        End Try
        clear()
    End Sub

    Private Sub btnsearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        source1.Filter = "[Transctinnmbr] = '" & ComboBox1.Text & "'"
        source2.Filter = "[Transctinnmbr] = '" & ComboBox1.Text & "'"
        DataGridView1.Refresh()
        ComboBox1.Text = ""

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ' timer.Text = Date.Now.ToString(" hh:mm:ss")
        timer.Text = TimeOfDay
    End Sub

    Private Sub done_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Dispose()
        'photo.ImageLocation = " & 'E:\VBprojects\Airline Travels and Tour\Airline Travels and Tour\Resources '"" & txtimagepath.Text & ".jpg"
        Me.Refresh()

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Panel1.Visible = True
        Panel1.Enabled = True

        GroupBox3.Visible = False
        GroupBox5.Visible = False
        GroupBox3.Enabled = False
        GroupBox5.Enabled = False
        Button5.Visible = True
        Button5.Enabled = True
        Label9.Visible = True
        Label7.Visible = True
        Btnadd.Visible = True
        Btnadd.Enabled = True
        Label19.Visible = True
        btnupdte.Enabled = True
        btnupdte.Visible = True
        Label7.Visible = True
        Btndel.Visible = True
        Btndel.Enabled = True
        Label15.Visible = True
        Btnsve.Visible = True
        Btnsve.Enabled = True
        Label14.Visible = True
        Button1.Visible = True
        Button1.Enabled = True
        Label5.Visible = True

        Label9.Visible = False
        Button6.Visible = False
        Button5.Visible = False
        Button2.Visible = False
        Label22.Visible = False
        Label10.Visible = False
        Button4.Visible = False

        Button6.Visible = False
        Button5.Enabled = False
        Button2.Enabled = False
        Button4.Enabled = False
        Label26.Visible = False
        gridclick()
    End Sub
    Private Sub gridclick()
        Try
            Label25.Text = "Edit Values in above textboxes"
            Panel1.Enabled = True
            Btnsve.Enabled = False
            Btndel.Enabled = True
            btnupdte.Enabled = True
            Me.txttrnsctin.Text = DataGridView1.CurrentRow.Cells(1).Value.ToString
            Me.txtbnknme.Text = DataGridView1.CurrentRow.Cells(0).Value.ToString
            Me.txtbankdte.Text = DataGridView1.CurrentRow.Cells(2).Value.ToString
            Me.txtaccnt.Text = DataGridView1.CurrentRow.Cells(3).Value.ToString
            Me.txtamnt.Text = DataGridView1.CurrentRow.Cells(4).Value.ToString
            Me.txtacnthodr.Text = DataGridView1.CurrentRow.Cells(5).Value.ToString

            'image
            Dim i As Integer
            i = DataGridView1.CurrentRow.Index
            Dim bytes As [Byte]() = (Me.DataGridView1.Item(6, i).Value)
            Dim ms As New MemoryStream(bytes)
            photo.Image = Image.FromStream(ms)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub clear()
        Try
            txttrnsctin.Text = ""
            txtbnknme.Text = ""
            txtbankdte.Text = ""
            txtaccnt.Text = ""
            photo.Image = Nothing
            txtamnt.Text = ""
            txtacnthodr.Text = ""
        Catch ex As Exception
            MsgBox("Error:Some thing is going wrong,Close application and try again")
        End Try
    End Sub
    Private Sub upenble()
        Try
            txttrnsctin.Enabled = False
            txtbnknme.Enabled = True
            txtbankdte.Enabled = True
            txtaccnt.Enabled = True
            txtamnt.Enabled = True
            txtacnthodr.Enabled = True
            photo.Enabled = True
        Catch ex As Exception
            MsgBox("Error:Some thing is going wrong,Close application and try again")
        End Try

    End Sub

    Private Sub Btnadd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btnadd.Click
        Label25.Text = " Add new record!"
        Label25.ForeColor = System.Drawing.Color.DeepPink
        Try
            Btnsve.Enabled = True
            Button9.Enabled = True
            Btndel.Enabled = False
            btnupdte.Enabled = False
            Button2.Enabled = False
            Panel1.Enabled = True
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

            cmd.CommandText = "SELECT MAX(Transctinnmbr) FROM bankdetails"
            If (IsDBNull(cmd.ExecuteScalar)) Then
                num = 1
                txttrnsctin.Text = num.ToString
            Else
                num = cmd.ExecuteScalar + 1
                txttrnsctin.Text = num.ToString
            End If
            con.Close()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        End Try
    End Sub
    Private Sub photo_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles photo.MouseHover
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(photo, "To see in large size double click on picture")
    End Sub
    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If Not txttrnsctin.Text = "" Then
                MessageBox.Show("Are you sure to print data", "Data Adding", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                Dim report As New CRBank
                Dim objText As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(1).ReportObjects("Text1")
                Dim objText1 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text2")
                Dim objText2 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text3")
                Dim objText3 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text4")
                Dim objText4 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text5")
                Dim objText5 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Text6")
                'Dim objText6 As CrystalDecisions.CrystalReports.Engine.TextObject = report.ReportDefinition.Sections(2).ReportObjects("Picture1")

                objText.Text = Me.txttrnsctin.Text
                objText1.Text = Me.txtbnknme.Text
                objText2.Text = Me.txtbankdte.Text
                objText3.Text = Me.txtaccnt.Text
                objText4.Text = Me.txtamnt.Text
                objText5.Text = Me.txtacnthodr.Text
                'objText6.Text = Me.photo.Text
                Rprtbnk.CrystalReportViewer1.ReportSource = report
                Rprtbnk.Show()
                Label25.Text = "'" & txttrnsctin.Text & "' Bank details printed successfully!"
                Label25.ForeColor = System.Drawing.Color.DarkGreen


            Else
                MessageBox.Show("Select value from gridview to print", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Label25.Text = "Problem while Printing '" & txttrnsctin.Text & "' Bankdetails"
                Label25.ForeColor = System.Drawing.Color.Red
                Panel1.Visible = False
                Panel1.Enabled = False
                GroupBox3.Visible = True
                GroupBox3.Enabled = True
                GroupBox5.Visible = False
                GroupBox5.Enabled = False
                Button5.Visible = False
                Button5.Enabled = False
                Label9.Visible = False
                Label7.Visible = False
                Btnadd.Visible = False
                Btnadd.Enabled = False
                Label19.Visible = False
                btnupdte.Enabled = False
                btnupdte.Visible = False
                Label7.Visible = False
                Btndel.Visible = False
                Btndel.Enabled = False
                Label15.Visible = False
                Btnsve.Visible = False
                Btnsve.Enabled = False
                Label14.Visible = False
                Button1.Visible = False
                Button1.Enabled = False
                Label5.Visible = False



                Label9.Visible = False
                Button6.Visible = False
                Button5.Visible = False
                Button2.Visible = True
                Label22.Visible = True
                Label10.Visible = False
                Button4.Visible = True

                Button6.Visible = False
                Button5.Enabled = False
                Button2.Enabled = False
                Button4.Enabled = True
                Label26.Visible = True

            End If
        Catch ex As Exception
            MessageBox.Show("Reports are not loding properly,try again", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Label25.Text = "Error while printing'" & txttrnsctin.Text & "' Bank details"
            Label25.ForeColor = System.Drawing.Color.Red

            Me.Dispose()
        End Try
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
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
                cmd.CommandText = "Select bnknme as [Bank Name], [Transctinnmbr], dte as[Date] ,accntnmbr as [Account Number],amount as[Amount],accnthldr as[Account Holder]FROM [bankdetails]"
                Using reader As SqlDataReader = cmd.ExecuteReader
                    While (reader.Read())
                        Me.ListBox1.Items.Add(reader("Bank Name"))
                       
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
        Panel1.Visible = True
        clear()
        Label25.Text = "Add New Records"
        Label25.ForeColor = System.Drawing.Color.BlueViolet
        GroupBox3.Visible = False
        GroupBox5.Visible = False
        GroupBox3.Enabled = False
        GroupBox5.Enabled = False
        Button5.Visible = True
        Button5.Enabled = True
        Label9.Visible = True
        Label7.Visible = True
        Btnadd.Visible = True
        Btnadd.Enabled = True
        Label19.Visible = True
        btnupdte.Enabled = True
        btnupdte.Visible = True
        Label7.Visible = True
        Btndel.Visible = True
        Btndel.Enabled = True
        Label15.Visible = True
        Btnsve.Visible = True
        Btnsve.Enabled = True
        Label14.Visible = True
        Button1.Visible = True
        Button1.Enabled = True
        Label5.Visible = True

        Label9.Visible = False
        Button6.Visible = False
        Button5.Visible = False
        Button2.Visible = False
        Label22.Visible = False
        Label10.Visible = False
        Button4.Visible = False

        Button6.Visible = False
        Button5.Enabled = False
        Button2.Enabled = False
        Button4.Enabled = False
        Label26.Visible = False
    End Sub

    Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Panel1.Visible = False
        Panel1.Enabled = False
        Label25.Text = " View All records"
        Label25.ForeColor = System.Drawing.Color.DeepPink
        GroupBox3.Visible = True
        GroupBox3.Enabled = True
        GroupBox5.Visible = False
        GroupBox5.Enabled = False
        Button5.Visible = False
        Button5.Enabled = False
        Label9.Visible = False
        Label7.Visible = False
        Btnadd.Visible = False
        Btnadd.Enabled = False
        Label19.Visible = False
        btnupdte.Enabled = False
        btnupdte.Visible = False
        Label7.Visible = False
        Btndel.Visible = False
        Btndel.Enabled = False
        Label15.Visible = False
        Btnsve.Visible = False
        Btnsve.Enabled = False
        Label14.Visible = False
        Button1.Visible = False
        Button1.Enabled = False
        Label5.Visible = False



        Label9.Visible = False
        Button6.Visible = False
        Button5.Visible = False
        Button2.Visible = True
        Label22.Visible = True
        Label10.Visible = False
        Button4.Visible = True

        Button6.Visible = False
        Button5.Enabled = False
        Button2.Enabled = True
        Button4.Enabled = True
        Label26.Visible = True
    End Sub

    Private Sub LinkLabel4_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        Panel1.Visible = False
        Panel1.Enabled = False
        Label25.Text = " View Information"
        Label25.ForeColor = System.Drawing.Color.DeepPink
        GroupBox3.Visible = False
        GroupBox3.Enabled = False
        GroupBox5.Visible = True
        GroupBox5.Enabled = True
        Button5.Visible = False
        Button5.Enabled = False
        Label9.Visible = False
        Label7.Visible = False
        Btnadd.Visible = False
        Btnadd.Enabled = False
        Label19.Visible = False
        btnupdte.Enabled = False
        btnupdte.Visible = False
        Label7.Visible = False
        Btndel.Visible = False
        Btndel.Enabled = False
        Label15.Visible = False
        Btnsve.Visible = False
        Btnsve.Enabled = False
        Label14.Visible = False
        Button1.Visible = False
        Button1.Enabled = False
        Label5.Visible = False



        Label9.Visible = True
        Button6.Visible = True
        Button5.Visible = True
        Button2.Visible = False
        Label22.Visible = False
        Label10.Visible = True
        Button4.Visible = False

        Button6.Visible = True
        Button6.Enabled = True
        Button5.Enabled = True
        Button2.Enabled = False
        Button4.Enabled = False
        Label26.Visible = False
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Panel1.Visible = False
        Label25.Text = "Welcome!"
        Label25.ForeColor = System.Drawing.Color.Black
        GroupBox3.Visible = False
        GroupBox5.Visible = False
        GroupBox3.Enabled = False
        GroupBox5.Enabled = False
        Button5.Visible = False
        Button5.Enabled = False
        Label9.Visible = False
        Label7.Visible = False
        Btnadd.Visible = False
        Btnadd.Enabled = False
        Label19.Visible = False
        btnupdte.Enabled = False
        btnupdte.Visible = False
        Label7.Visible = False
        Btndel.Visible = False
        Btndel.Enabled = False
        Label15.Visible = False
        Btnsve.Visible = False
        Btnsve.Enabled = False
        Label14.Visible = False
        Button1.Visible = False
        Button1.Enabled = False
        Label5.Visible = False

        Label9.Visible = False
        Button6.Visible = False
        Button5.Visible = False
        Button2.Visible = False
        Label22.Visible = False
        Label10.Visible = False
        Button4.Visible = False

        Button6.Visible = False
        Button5.Enabled = False
        Button2.Enabled = False
        Button4.Enabled = False
        Label26.Visible = False
    End Sub

    Private Sub LargeSizeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LargeSizeToolStripMenuItem.Click
        frmbpic.PictureBox1.Image = photo.Image
        frmbpic.Show()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Try

            MessageBox.Show("Are you sure to add data", "Data Adding", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            insert()
            Label25.Text = "'" & txttrnsctin.Text & "'Bank details saved successfully!"
            Label25.ForeColor = System.Drawing.Color.DarkGreen

            Btnsve.Enabled = False
            Button9.Enabled = False
            Panel2.Enabled = False
        Catch ex As Exception
            ' MessageBox.Show("Data is already exist", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Label25.Text = "Error while saving '" & txttrnsctin.Text & "'Bank details"
            Label25.ForeColor = System.Drawing.Color.Red

            MessageBox.Show("Data already exist, you again select Bank Details and Try other entry", "Data Invalid, Application is closing", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Dispose()
        End Try
        clear()
        Me.Refresh()
    End Sub

    Private Sub txtamnt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtamnt.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        e.Graphics.DrawImage(bitmap, 0, 0)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
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
            ObjCommand.CommandText = "delete from bankdetails where Transctinnmbr='" & DataGridView1.SelectedRows(i).Cells("Transctinnmbr").Value & "'"
            ObjConnection.Open()
            ObjCommand.ExecuteNonQuery()
            ObjConnection.Close()

            Me.DataGridView1.Rows.Remove(Me.DataGridView1.SelectedRows(i))
        Next

    End Sub

    Private Sub RemoveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveToolStripMenuItem.Click
        DeleteSelecedRows()
        FillCombo()
        FillCombo2()
    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        gridclick()
        Panel1.Visible = True
        Panel1.Enabled = True
        GroupBox3.Visible = False
        GroupBox5.Visible = False
        GroupBox3.Enabled = False
        GroupBox5.Enabled = False
        Button5.Visible = True
        Button5.Enabled = True
        Label9.Visible = True
        Label7.Visible = True
        Btnadd.Visible = True
        Btnadd.Enabled = True
        Label19.Visible = True
        btnupdte.Enabled = True
        btnupdte.Visible = True
        Label7.Visible = True
        Btndel.Visible = True
        Btndel.Enabled = True
        Label15.Visible = True
        Btnsve.Visible = True
        Btnsve.Enabled = False
        Label14.Visible = True
        Button1.Visible = True
        Button1.Enabled = True
        Label5.Visible = True


        Label9.Visible = False
        Button6.Visible = False
        Button5.Visible = False
        Button2.Visible = False
        Label22.Visible = False
        Label10.Visible = False
        Button4.Visible = False

        Button6.Visible = False
        Button5.Enabled = False
        Button2.Enabled = False
        Button4.Enabled = False
        Label26.Visible = False
    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub PictureBox23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox23.Click
        txttrnsctin.Text = Val(txttrnsctin.Text) + 1
    End Sub

    Private Sub PictureBox8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox8.Click
        txttrnsctin.Text = Val(txttrnsctin.Text) - 1
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try

            ListBox1.Items.Clear()
            listboxfill()
            listview()
            Label11.Text = "Umrah details loaded successfully!"
            Label11.ForeColor = System.Drawing.Color.DarkGreen
            Label27.Text = ""
            Label27.ForeColor = System.Drawing.Color.Black
        Catch ex As Exception
            Label11.Text = "Umrah details not loaded successfully!"
            Label11.ForeColor = System.Drawing.Color.Red
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        source1.Filter = "[Transctinnmbr] = '" & ComboBox1.Text & "'"
        source2.Filter = "[Transctinnmbr] = '" & ComboBox1.Text & "'"
        DataGridView1.Refresh()
        ComboBox1.Text = ""

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Try
            ListBox1.Items.Clear()
            ListView1.Items.Clear()
            ListView2.Items.Clear()
            TextBox1.Text = ""
            Label27.Text = "Umrah details Unloaded successfully!"
            Label27.ForeColor = System.Drawing.Color.DarkGreen
            Label11.Text = "!"
            Label11.ForeColor = System.Drawing.Color.Black
        Catch ex As Exception
            Label27.Text = "Umrah details not Unloaded successfully!"
            Label27.ForeColor = System.Drawing.Color.Red
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        DeleteSelecedRows()
        FillCombo()
        FillCombo2()
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
        strQ = "Select bnknme as [Bank Name], [Transctinnmbr], dte as[Date] ,accntnmbr as [Account Number],amount as[Amount],accnthldr as[Account Holder]FROM [bankdetails]"
        cmd = New SqlCommand(strQ, conn)
        da = New SqlDataAdapter(cmd)
        ds = New DataSet
        da.Fill(ds, "bankdetails")
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
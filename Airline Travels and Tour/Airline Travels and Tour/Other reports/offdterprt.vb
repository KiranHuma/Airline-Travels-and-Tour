Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.Data.DataTable
Imports System.Data.SqlClient
Public Class offdterprt
    Dim con As New SqlClient.SqlConnection
    Dim cmd As New SqlClient.SqlCommand
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim cn As New SqlConnection
        Dim ds As New DataSet
        Dim dt As New DataTable
        Dim dfrom As DateTime = DateTimePicker1.Value
        Dim dto As DateTime = DateTimePicker2.Value
        cn.ConnectionString = "Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True"
        cn.Open()
        Dim str As String = "select * from mainoffice where dte >= '" & Format(dfrom, "MM-dd-yyyy") & "' and dte <='" & Format(dto, "MM-dd-yyyy") & "'"

        Dim da As SqlDataAdapter = New SqlDataAdapter(str, cn)
        da.Fill(dt)
        DataGridView1.DataSource = dt
    End Sub
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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Cursor = Cursors.WaitCursor

            Dim rprt As New CRDoffce 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand As New SqlCommand()
            Dim myDA As New SqlDataAdapter()
            Dim myDS As New airlineeDataSet 'The DataSet you created.
            myConnection = New SqlConnection("Data Source=MEERHAMZA;Initial Catalog=airlinee;Integrated Security=True")
            Dim dfrom As DateTime = DateTimePicker1.Value
            Dim dto As DateTime = DateTimePicker2.Value
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "select * from mainoffice where dte  >= '" & Format(dfrom, "MM-dd-yyyy") & "' and dte <='" & Format(dto, "MM-dd-yyyy") & "'"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "mainoffice")
            rprt.SetDataSource(myDS)
            viewoffcerprt.CrystalReportViewer1.ReportSource = rprt
            viewoffcerprt.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
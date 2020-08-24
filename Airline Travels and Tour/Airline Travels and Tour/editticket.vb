Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.Data.DataTable
Imports System.Data.SqlClient
Public Class editticket
    'Dim con As New OleDb.OleDbConnection
    'Dim cmd As New OleDb.OleDbCommand
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
    Private Sub edit()
        dbaccessconnection()
        con.Open()
        cmd.CommandText = ("UPDATE ticket SET nme='" & txtnme.Text & "',  sector= '" & txtsectr.Text & "',dte= '" & txtdte.Text & "',pnr = '" & txtpnr.Text & "',issuedte = '" & Txtissdte.Text & "',expiredte = '" & Txtexdte.Text & "',mobileno = '" & txtmobile.Text & "',flight = '" & txtflight.Text & "',pasprtnmbr = '" & txtpasprt.Text & "',comissin = '" & txtcommissin.Text & "',basicfair= '" & txtbasicfair.Text & "',deprttme= '" & txtdeptme.Text & "',arrtme= '" & txtarrtime.Text & "',totalamount= '" & txttotal.Text & "',status= '" & txtstatus.Text & "'  where ticketnmbr = " & txtticket.Text & "")
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub

    Private Sub Btnsve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btnsve.Click
        If Not txtticket.Text = "" Then
            edit()
            MsgBox("data update successfully")

        Else
            MsgBox("data not update successfully")
        End If
    End Sub



    Private Sub editticket_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dbaccessconnection()
    End Sub
End Class
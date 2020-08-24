
Public Class Rprtticket

   
    Private Sub Rprtticket_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        
    End Sub
    Private Sub frmload()
        Dim Report1 As New CRticket

        Report1.SetParameterValue("", Ticket.txtticket.Text)
        CrystalReportViewer1.ReportSource = Report1
    End Sub
End Class
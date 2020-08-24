Public Class choose

    

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Button10.Enabled = False Then
            Ticket.txtcommissin.ForeColor = Color.White
            Ticket.txtcommissin.Enabled = False
            Ticket.txtbasicfair.Enabled = False
            Ticket.DataGridView1.Visible = False
            Ticket.btnupdte.Enabled = False
            Ticket.Btnsve.Enabled = False
            Ticket.Label14.Visible = False
            Ticket.Btnsve.Visible = False
            Ticket.Button9.Enabled = False
            Ticket.Button9.Visible = True
            Ticket.Label16.Visible = True
            Ticket.Button1.Visible = False
            Ticket.Panel3.Visible = False
            Ticket.Button13.Visible = True
            Ticket.Label26.Visible = True
            Ticket.Label18.Visible = False
            Ticket.Button1.Enabled = False
            Ticket.GroupBox3.Visible = False
            Ticket.GroupBox3.Enabled = False

            Ticket.ShowDialog()
        Else
            Ticket.txtcommissin.ForeColor = Color.Black
            Ticket.txttotal.ForeColor = Color.Black
            Ticket.Btnsve.Visible = True
            Ticket.Button1.Visible = True
            Ticket.Label18.Visible = True
            Ticket.GroupBox3.Visible = True
            Ticket.GroupBox3.Enabled = True
           Ticket.ShowDialog()
        End If
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        frmbnk.ShowDialog()
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Frmvisa.ShowDialog()
       
    End Sub

    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If Button10.Enabled = False Then
            Frmvisit.txtprft.ForeColor = Color.White
            Frmvisit.txtprft.Enabled = False
            Frmvisit.txtsale.Enabled = False
            Frmvisit.DataGridView1.Visible = False
            Frmvisit.btnupdte.Enabled = False
            Frmvisit.Btnsve.Enabled = False
            Frmvisit.Label14.Visible = False
            Frmvisit.Btnsve.Visible = False
            Frmvisit.Button9.Enabled = False
            Frmvisit.Button9.Visible = True
            Frmvisit.Label16.Visible = True
            Frmvisit.Button1.Visible = False
            Frmvisit.Panel3.Visible = False
            Frmvisit.Button13.Visible = True
            Frmvisit.Label26.Visible = True
            Frmvisit.Label8.Visible = False
            Frmvisit.Button1.Enabled = False
            Frmvisit.GroupBox3.Visible = False
            Frmvisit.GroupBox3.Enabled = False

            Frmvisit.ShowDialog()
        Else
            Frmvisit.txtprft.ForeColor = Color.Black
            Frmvisit.txtsale.ForeColor = Color.Black
            Frmvisit.Btnsve.Visible = True
            Frmvisit.Button1.Visible = True
            Frmvisit.Label8.Visible = True
            Frmvisit.GroupBox3.Visible = True
            Frmvisit.GroupBox3.Enabled = True
            Frmvisit.ShowDialog()
        End If
    End Sub

    Private Sub Button5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        frmumrah.ShowDialog()
    End Sub

    Private Sub Button6_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Frmmainoffice.ShowDialog()
    End Sub

    Private Sub Button7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        frmexpensis.ShowDialog()
    End Sub

    Private Sub choose_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        editbnk.ShowDialog()

        Me.Label3.Text = Format(Now, "dd-MMM-yyyy")
        ' Me.Label4.Text = Format(Now, "hh:mm")
        'Timer1.Start()
        Call CenterToScreen()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        tprft.ShowDialog()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        If MsgBox("Are you sure want to exit now?", MsgBoxStyle.YesNo, "Closing warning") = MsgBoxResult.Yes Then ' If you select yes in the MsgBox then it will close the window
            End ' Close the window
        Else
            ' Will not close the application
        End If
    End Sub

    Private Sub choose_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDoubleClick

    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Choosereports.ShowDialog()
    End Sub

    Private Sub PictureBox5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox5.Click

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Label4.Text = TimeOfDay
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Call CenterToScreen()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        If MsgBox("Are you sure want to exit now?", MsgBoxStyle.YesNo, "Closing warning") = MsgBoxResult.Yes Then ' If you select yes in the MsgBox then it will close the window
            End
            '  Me.Dispose() ' Close the window
        Else
            ' Will not close the application
        End If
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Call CenterToScreen()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Call CenterToScreen()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.WindowState = FormWindowState.Normal
    End Sub
End Class
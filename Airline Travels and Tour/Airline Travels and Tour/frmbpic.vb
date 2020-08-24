Public Class frmbpic

   
  

   
    Private Sub frmbpic_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call CenterToScreen()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub frmbpic_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDoubleClick
        Me.Dispose()
    End Sub
   

    Private Sub PictureBox1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox1.DoubleClick
        Me.Dispose()
    End Sub

    Private Sub PictureBox1_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox1.MouseHover
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(PictureBox1, "to leave pic double cilck on image")
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        Call CenterToScreen()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub PictureBox1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseDoubleClick
        Try
            Call CenterToScreen()
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            Me.WindowState = FormWindowState.Normal
        Catch ex As Exception
            End Try
    End Sub

    Private Sub PictureBox1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseClick
        Call CenterToScreen()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        printDoc.Print()
    End Sub
    Dim WithEvents printDoc As New Printing.PrintDocument()
    Private Sub PrintImage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles printDoc.PrintPage
        e.Graphics.DrawImage(PictureBox1.Image, e.MarginBounds.Left, e.MarginBounds.Top)
    End Sub

End Class
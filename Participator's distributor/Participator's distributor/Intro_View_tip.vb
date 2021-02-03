Public Class Intro_View_tip
    Private Sub Intro_View_tip_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Left = MousePosition.X - Me.Width / 2
        Me.Top = MousePosition.Y - Me.Height / 2
        Me.BackColor = Color.DarkRed
    End Sub
End Class

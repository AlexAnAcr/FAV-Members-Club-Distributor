Public Class Intro_View
    Dim x, y, w, h As Integer, to_close As Boolean = False, painted As Boolean = True
    Private Sub Motion_tip_pos_Tick(sender As Object, e As EventArgs) Handles Motion_tip_pos.Tick
        If to_close = True Then
            Motion_tip_pos.Stop()
            Me.Close()
        Else
            If Intro_View_tip.Left = Me.Left And Intro_View_tip.Top = Me.Top Then
                Intro_View_tip.Left = MousePosition.X - Intro_View_tip.Width / 2
                Intro_View_tip.Top = MousePosition.Y - Intro_View_tip.Height / 2
                Intro_View_tip.BackColor = Color.DarkRed
                painted = True
            Else
                If Math.Max(Intro_View_tip.Left, Me.Left) - Math.Min(Intro_View_tip.Left, Me.Left) >= 20 Then
                    If Intro_View_tip.Left > Me.Left Then
                        Intro_View_tip.Left -= 20
                    Else
                        Intro_View_tip.Left += 20
                    End If
                Else
                    Intro_View_tip.Left = Me.Left
                End If
                If Math.Max(Intro_View_tip.Top, Me.Top) - Math.Min(Intro_View_tip.Top, Me.Top) >= 20 Then
                    If Intro_View_tip.Top > Me.Top Then
                        Intro_View_tip.Top -= 20
                    Else
                        Intro_View_tip.Top += 20
                    End If
                Else
                    Intro_View_tip.Top = Me.Top
                End If
                If painted = True Then
                    Intro_View_tip.BackColor = Color.DarkBlue
                    painted = False
                End If
            End If
            If MousePosition.X > Me.Left And MousePosition.X < Me.Left + Me.Width And MousePosition.Y > Me.Top And MousePosition.Y < Me.Top + Me.Height Then
                Intro_View_tip.Close()
                to_close = True
                Me.BackColor = Color.DarkRed
            End If
        End If
        Try
            If ActiveForm.Name = "Intro_View" Then
                Intro_View_tip.Activate()
            End If
        Catch
            Motion_tip_pos.Stop()
            Intro_View_tip.Close()
            Me.Close()
        End Try
    End Sub
    Private Sub Intro_View_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Left = Main.Left + x
        Me.Top = Main.Top + y
        Me.Width = w
        Me.Height = h
        Intro_View_tip.Width = Me.Width
        Intro_View_tip.Height = Me.Height
        to_close = False
        painted = True
        Me.BackColor = Color.DarkBlue
        Intro_View_tip.Show()
        Motion_tip_pos.Start()
    End Sub
    Public Shared Sub Set_position(x As Integer, y As Integer, width As Integer, height As Integer)
        Intro_View.x = x
        Intro_View.y = y
        Intro_View.w = width
        Intro_View.h = height
    End Sub
End Class

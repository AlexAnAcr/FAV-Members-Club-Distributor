Public Class Progress
    Public Shared progress_mode(1) As Single
    Public Shared Sub Progress_tick()
        If progress_mode(0) = 1 Then
            If Progress.ProgressBar1.Value + 1 >= 100 Then
                Progress.ProgressBar1.Value = 100
            Else
                Progress.ProgressBar1.Value += 1
            End If
        Else
            If Progress.ProgressBar1.Value + progress_mode(1) >= 100 Then
                Progress.ProgressBar1.Value = 100
            Else
                Progress.ProgressBar1.Value += progress_mode(1)
            End If
        End If
        Progress.Label1.Text = Progress.ProgressBar1.Value & "%"
    End Sub
    Public Shared Sub Reset()
        Progress.ProgressBar1.Value = 0
        Progress.Label1.Text = "0%"
    End Sub
    Dim p_form As Form = Main
    Public Shared Sub Set_parent_(form As Form)
        Progress.p_form = form
    End Sub
    Private Sub Progress_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Left = Int(p_form.Left + p_form.Width / 2 - Me.Width / 2)
        Me.Top = Int(p_form.Top + p_form.Height / 2 - Me.Height / 2)
    End Sub
End Class

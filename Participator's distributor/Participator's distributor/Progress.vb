Public Class Progress
    Public Shared progress_mode(1) As Integer
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
    Private Sub Progress_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Left = Int(Main.Left + Main.Width / 2 - Me.Width / 2)
        Me.Top = Int(Main.Top + Main.Height / 2 - Me.Height / 2)
    End Sub
End Class

Public Class full_screen
    Public font_loaded As Boolean = False
    Private Sub full_screen_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Left = -10
        Me.Top = -10
        Me.Width = My.Computer.Screen.Bounds.Width + 20
        Me.Height = My.Computer.Screen.Bounds.Height + 20
        If font_loaded = False Then
            font_loaded = True
            Dim f_finded As Boolean = False, f_size As UShort = 72, f_last_size As UShort = 72
            While f_finded = False
                Dim fnt As New Font("Lucida Console", f_size, FontStyle.Regular)
                Label2.Font = fnt
                If Label2.Height >= Label1.Height Or Label2.Width >= Label1.Width Then
                    f_size = f_last_size
                    f_finded = True
                Else
                    f_last_size = f_size
                    f_size += 4
                End If
            End While
            Label2.Visible = False
            Dim new_font As New Font("Lucida Console", f_size, FontStyle.Regular)
            Label1.Font = new_font
            Button2.Width = Panel1.Width - 2
            Button1.Width = Panel2.Width - 2
            Button2.Height = Panel1.Height - 2
            Button1.Height = Panel2.Height - 2
            Button2.Left = 0
            Button1.Left = 0
            Button2.Top = 0
            Button1.Top = 0
        End If
        button_y = Me.Height - Button1.Height * 1.5
    End Sub
    Private Sub full_screen_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Active_C.Start()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Active_C.Stop()
        Me.Close()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Main.timer_button_mode = 1
        Main.player1.controls.stop()
        Main.Button24.Text = "Старт"
        Main.Button27.Enabled = True
        Main.Button25.Enabled = True
        Main.Label17.ForeColor = Color.Black
        Label1.ForeColor = Color.Black
        Main.Button24.Enabled = False
        Button2.Enabled = False
        Main.Button26.Enabled = True
    End Sub
    Private Sub cont_KeyDown(sender As Object, e As KeyEventArgs) Handles active_control.KeyDown
        If e.KeyCode = 27 Then
            Active_C.Stop()
            Me.Close()
        End If
    End Sub
    WithEvents active_control As Control
    Dim button_y As Integer, on_pos As Boolean = True, last_on_pos As Boolean = True
    Private Sub Active_C_Tick(sender As Object, e As EventArgs) Handles Active_C.Tick
        active_control = Me.ActiveControl
        Dim mouse_position As Integer = MousePosition.Y - Me.Top
        If mouse_position > button_y Then
            If on_pos = False Or last_on_pos = False Then
                If Button1.Top > 0 Then
                    Button1.Top -= 10
                    Button2.Top -= 10
                    on_pos = False
                Else
                    on_pos = True
                End If
            End If
            last_on_pos = True
        Else
            If on_pos = True Or last_on_pos = True Then
                If Button1.Top < Panel1.Height Then
                    Button1.Top += 10
                    Button2.Top += 10
                    on_pos = True
                Else
                    on_pos = False
                End If
            End If
            last_on_pos = False
        End If
    End Sub
    Private Sub Panel1_MouseMove(sender As Object, e As MouseEventArgs) Handles Panel1.MouseMove

    End Sub
End Class